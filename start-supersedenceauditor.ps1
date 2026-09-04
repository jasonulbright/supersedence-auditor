<#
.SYNOPSIS
    MahApps.Metro WPF shell for the Supersedence and Dependency Auditor.

.DESCRIPTION
    Replaces the v1.0.x WinForms shell with a brand-aligned WPF UI: sidebar
    navigation across four views (Supersedence, Dependencies, Broken Rules,
    Tree View), inline action bar (Scan, filter, status filter, exports),
log drawer, and status bar. Status conveyed via glyph, not row color
    coloring). Site / SMS Provider configured from the Options sidebar
    button (no File menu).

    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.7.2+
      - MahApps.Metro DLLs in .\Lib\
      - SupersedenceAuditorCommon module under .\Module\
      - ConfigurationManager console (provides Get-CMApplication / Get-CMSite)

.NOTES
    ScriptName : start-supersedenceauditor.ps1
    Version    : 1.0.1
    Updated    : 2026-05-02
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification='Per feedback_ps_wpf_handler_rules.md and PS51-WPF-001..003: flat-.ps1 GetNewClosure strips $script: scope. $global: survives closure scope-strip and keeps shared mutable state reachable from closure-captured handlers.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification='WPF event handler scriptblocks bind positional sender/args ($s, $e). The sender is required to fulfill the signature even when the handler body does not read it.')]
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# A Windows PowerShell process launched from PowerShell 7 inherits the 7.x
# module directories at the front of PSModulePath; the background runspace
# opened later would autoload Microsoft.PowerShell.Utility from the 7.x
# manifest, which carries no Get-FileHash / ConvertFrom-Json under 5.1.
# Strip those roots from the process environment before any runspace opens.
$__winPsModules = Join-Path $PSHOME 'Modules'
$__moduleRoots = @($env:PSModulePath -split ';' | Where-Object { $_ -and $_ -notmatch '(?i)[\\/]PowerShell[\\/](7[\\/]|Modules)|microsoft\.powershell_' })
if ($__moduleRoots -notcontains $__winPsModules) { $__moduleRoots = @($__winPsModules) + $__moduleRoots }
$env:PSModulePath = ($__moduleRoots -join ';')

# =============================================================================
# Startup transcript (best-effort).
# =============================================================================
$__txDir = Join-Path $PSScriptRoot 'Logs'
try {
    if (-not (Test-Path -LiteralPath $__txDir)) { New-Item -ItemType Directory -Path $__txDir -Force | Out-Null }
    $__tx = Join-Path $__txDir ('SupersedenceAuditor-startup-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    Start-Transcript -LiteralPath $__tx -Force | Out-Null
} catch { $null = $_ }

# =============================================================================
# STA guard. WPF requires STA. PS51-WPF-009.
# =============================================================================
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $psExe = (Get-Process -Id $PID).Path
    $fwd   = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',$PSCommandPath)
    Start-Process -FilePath $psExe -ArgumentList $fwd | Out-Null
    try { Stop-Transcript | Out-Null } catch { $null = $_ }
    exit 0
}

# =============================================================================
# Assemblies.
# =============================================================================
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$libDir = Join-Path $PSScriptRoot 'Lib'
if (-not (Test-Path -LiteralPath $libDir)) {
    throw "Lib/ directory not found at: $libDir. Re-extract the release zip."
}

Get-ChildItem -LiteralPath $libDir -File -ErrorAction SilentlyContinue |
    Unblock-File -ErrorAction SilentlyContinue

[void][System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'Microsoft.Xaml.Behaviors.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'ControlzEx.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'MahApps.Metro.dll'))

# =============================================================================
# Module import.
# =============================================================================
$__modulePath = Join-Path $PSScriptRoot 'Module\SupersedenceAuditorCommon.psd1'
if (-not (Test-Path -LiteralPath $__modulePath)) {
    throw "Shared module not found at: $__modulePath"
}
Import-Module -Name $__modulePath -Force -DisableNameChecking
if (-not (Get-Command Initialize-Logging -ErrorAction SilentlyContinue)) {
    throw "SupersedenceAuditorCommon imported but Initialize-Logging is not exported."
}

# =============================================================================
# Preferences (SupersedenceAuditor.prefs.json next to the script).
# Closure-safe via $global:.
# =============================================================================
$global:PrefsPath = Join-Path $PSScriptRoot 'SupersedenceAuditor.prefs.json'

function Get-SaPreferences {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns the full preferences hashtable by design.')]
    param()
    $defaults = @{
        DarkMode    = $true
        SiteCode    = ''
        SMSProvider = ''
    }
    if (Test-Path -LiteralPath $global:PrefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $global:PrefsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            foreach ($k in @($defaults.Keys)) {
                $val = $loaded.$k
                if ($null -ne $val) { $defaults[$k] = $val }
            }
        } catch { $null = $_ }
    }
    return $defaults
}

function Save-SaPreferences {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Writes the full preferences hashtable by design.')]
    param([Parameter(Mandatory)][hashtable]$Prefs)
    try {
        $Prefs | ConvertTo-Json | Set-Content -LiteralPath $global:PrefsPath -Encoding UTF8
    } catch { $null = $_ }
}

$global:Prefs = Get-SaPreferences

# =============================================================================
# Tool log.
# =============================================================================
$toolLogPath = Join-Path $__txDir ('SupersedenceAuditor-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $toolLogPath

# =============================================================================
# Load XAML and resolve named elements.
# =============================================================================
$xamlPath = Join-Path $PSScriptRoot 'MainWindow.xaml'
if (-not (Test-Path -LiteralPath $xamlPath)) {
    throw "MainWindow.xaml not found at: $xamlPath"
}
[xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$txtAppTitle        = $window.FindName('txtAppTitle')
$txtVersion         = $window.FindName('txtVersion')
# Installed version: the script header is the single source of truth for the
# sidebar label and the About panel.
$script:AppVersion = '0.0.0'
foreach ($headerLine in (Get-Content -LiteralPath $PSCommandPath -TotalCount 80)) {
    if ($headerLine -match '^\s*Version\s*:\s*([0-9][0-9\.]*[0-9])\s*$') { $script:AppVersion = $Matches[1]; break }
}
if ($txtVersion) { $txtVersion.Text = 'v' + $script:AppVersion }
$txtThemeLabel      = $window.FindName('txtThemeLabel')
$toggleTheme        = $window.FindName('toggleTheme')

$btnViewSupersedence = $window.FindName('btnViewSupersedence')
$btnViewDependencies = $window.FindName('btnViewDependencies')
$btnViewBroken       = $window.FindName('btnViewBroken')
$btnViewTree         = $window.FindName('btnViewTree')
$btnOptions          = $window.FindName('btnOptions')

$txtModuleTitle    = $window.FindName('txtModuleTitle')
$txtModuleSubtitle = $window.FindName('txtModuleSubtitle')

$btnScan         = $window.FindName('btnScan')
$txtFilter       = $window.FindName('txtFilter')
$cboStatus       = $window.FindName('cboStatus')
$btnExportCsv    = $window.FindName('btnExportCsv')
$btnExportHtml   = $window.FindName('btnExportHtml')
$btnCopySummary  = $window.FindName('btnCopySummary')

$viewSupersedence = $window.FindName('viewSupersedence')
$viewDependencies = $window.FindName('viewDependencies')
$viewBroken       = $window.FindName('viewBroken')
$viewTree         = $window.FindName('viewTree')

$gridSupersedence      = $window.FindName('gridSupersedence')
$txtSupersedenceDetail = $window.FindName('txtSupersedenceDetail')
$gridDependencies      = $window.FindName('gridDependencies')
$txtDependencyDetail   = $window.FindName('txtDependencyDetail')
$gridBroken            = $window.FindName('gridBroken')
$txtBrokenDetail       = $window.FindName('txtBrokenDetail')
$treeRelationships     = $window.FindName('treeRelationships')
$txtTreeDetail         = $window.FindName('txtTreeDetail')

$progressOverlay   = $window.FindName('progressOverlay')
$txtProgressTitle  = $window.FindName('txtProgressTitle')
$txtProgressStep   = $window.FindName('txtProgressStep')

$lblLogOutput = $window.FindName('lblLogOutput')
$txtLog       = $window.FindName('txtLog')
$txtStatus    = $window.FindName('txtStatus')

$null = $txtAppTitle, $txtVersion

# =============================================================================
# Log drawer + status bar helpers.
# =============================================================================
function Add-LogLine {
    param([Parameter(Mandatory)][string]$Message)
    $ts = (Get-Date).ToString('HH:mm:ss')
    $line = '{0}  {1}' -f $ts, $Message
    if ([string]::IsNullOrWhiteSpace($txtLog.Text)) {
        $txtLog.Text = $line
    } else {
        $txtLog.AppendText([Environment]::NewLine + $line)
    }
    $txtLog.ScrollToEnd()
}

function Set-StatusText {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates an in-window TextBlock only; no external state.')]
    param([Parameter(Mandatory)][string]$Text)
    $txtStatus.Text = $Text
}

# =============================================================================
# Title-bar drag fallback. PS51-WPF-033.
# Some VS Code PowerShell launch contexts can leave MahApps' custom title
# thumb unable to initiate native window move. Install a WM_NCHITTEST hook
# returning HTCAPTION for the title band, plus a managed DragMove fallback
# for hosts where HwndSource cannot be hooked. Wire on every MetroWindow
# (main and every modal popup).
# =============================================================================
$script:TitleBarHitTestWindows = @{}
$script:TitleBarHitTestHooks   = @{}

function Get-TitleBarDragHeight {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $h = [double]$Window.TitleBarHeight
        if ($h -gt 0 -and -not [double]::IsNaN($h)) { return $h }
    } catch { $null = $_ }
    return 30.0
}

function Get-InputAncestors {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Private visual-tree helper yields an ancestor chain.')]
    param([System.Windows.DependencyObject]$Start)
    $cur = $Start
    while ($cur) {
        $cur
        $parent = $null
        if ($cur -is [System.Windows.Media.Visual] -or $cur -is [System.Windows.Media.Media3D.Visual3D]) {
            try { $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($cur) } catch { $parent = $null }
        }
        if (-not $parent -and $cur -is [System.Windows.FrameworkElement]) { $parent = $cur.Parent }
        if (-not $parent -and $cur -is [System.Windows.FrameworkContentElement]) { $parent = $cur.Parent }
        if (-not $parent -and $cur -is [System.Windows.ContentElement]) {
            try { $parent = [System.Windows.ContentOperations]::GetParent($cur) } catch { $parent = $null }
        }
        $cur = $parent
    }
}

function Test-IsWindowCommandPoint {
    param([MahApps.Metro.Controls.MetroWindow]$Window, [System.Windows.Point]$Point)
    try {
        [void]$Window.ApplyTemplate()
        $commands = $Window.Template.FindName('PART_WindowButtonCommands', $Window)
        if ($commands -and $commands.IsVisible -and $commands.ActualWidth -gt 0 -and $commands.ActualHeight -gt 0) {
            $origin = $commands.TransformToAncestor($Window).Transform([System.Windows.Point]::new(0, 0))
            if ($Point.X -ge $origin.X -and $Point.X -le ($origin.X + $commands.ActualWidth) -and
                $Point.Y -ge $origin.Y -and $Point.Y -le ($origin.Y + $commands.ActualHeight)) {
                return $true
            }
        }
    } catch { $null = $_ }
    return ($Window.ActualWidth -gt 150 -and $Point.X -ge ($Window.ActualWidth - 150))
}

function Add-NativeTitleBarHitTestHook {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Installs an in-process HWND hook for this WPF window only.')]
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
        if (-not $source) { return }
        $key = $helper.Handle.ToInt64().ToString()
        if ($script:TitleBarHitTestHooks.ContainsKey($key)) { return }
        $script:TitleBarHitTestWindows[$key] = $Window
        $hook = [System.Windows.Interop.HwndSourceHook]{
            param([IntPtr]$hwnd, [int]$msg, [IntPtr]$wParam, [IntPtr]$lParam, [ref]$handled)
            $WM_NCHITTEST = 0x0084; $HTCAPTION = 2
            if ($msg -ne $WM_NCHITTEST) { return [IntPtr]::Zero }
            try {
                $target = $script:TitleBarHitTestWindows[$hwnd.ToInt64().ToString()]
                if (-not $target) { return [IntPtr]::Zero }
                $raw = $lParam.ToInt64()
                $screenX = [int]($raw -band 0xffff); if ($screenX -ge 0x8000) { $screenX -= 0x10000 }
                $screenY = [int](($raw -shr 16) -band 0xffff); if ($screenY -ge 0x8000) { $screenY -= 0x10000 }
                $pt = $target.PointFromScreen([System.Windows.Point]::new($screenX, $screenY))
                $titleBarH = Get-TitleBarDragHeight -Window $target
                if ($pt.X -lt 0 -or $pt.X -gt $target.ActualWidth) { return [IntPtr]::Zero }
                if ($pt.Y -lt 4 -or $pt.Y -gt $titleBarH) { return [IntPtr]::Zero }
                if (Test-IsWindowCommandPoint -Window $target -Point $pt) { return [IntPtr]::Zero }
                $handled.Value = $true
                return [IntPtr]$HTCAPTION
            } catch { return [IntPtr]::Zero }
        }
        $script:TitleBarHitTestHooks[$key] = $hook
        $source.AddHook($hook)
    } catch { $null = $_ }
}

function Remove-NativeTitleBarHitTestHook {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Removes an in-process HWND hook for this WPF window only.')]
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $key = $helper.Handle.ToInt64().ToString()
        if ($script:TitleBarHitTestHooks.ContainsKey($key)) {
            $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
            if ($source) { $source.RemoveHook($script:TitleBarHitTestHooks[$key]) }
            $script:TitleBarHitTestHooks.Remove($key)
        }
        if ($script:TitleBarHitTestWindows.ContainsKey($key)) {
            $script:TitleBarHitTestWindows.Remove($key)
        }
    } catch { $null = $_ }
}

function Install-TitleBarDragFallback {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Registers window-local WPF event handlers for title-bar drag fallback.')]
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    $Window.Add_SourceInitialized({ param($s, $e) Add-NativeTitleBarHitTestHook -Window $s })
    $Window.Add_Closed({ param($s, $e) Remove-NativeTitleBarHitTestHook -Window $s })
    $Window.Add_PreviewMouseLeftButtonDown({
        param($s, $e)
        try {
            if ($s.WindowState -eq [System.Windows.WindowState]::Maximized) { return }
            $titleBarH = Get-TitleBarDragHeight -Window $s
            $pos = $e.GetPosition($s)
            if ($pos.Y -lt 4 -or $pos.Y -gt $titleBarH) { return }
            if (Test-IsWindowCommandPoint -Window $s -Point $pos) { return }
            foreach ($ancestor in Get-InputAncestors -Start ($e.OriginalSource -as [System.Windows.DependencyObject])) {
                if ($ancestor -is [System.Windows.Controls.Primitives.ButtonBase]) { return }
            }
            $s.DragMove()
            $e.Handled = $true
        } catch { $null = $_ }
    })
}

Install-TitleBarDragFallback -Window $window

# =============================================================================
# Theme setup and toggle.
# =============================================================================
[void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Dark.Steel')

$script:DarkButtonBg      = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#1E1E1E')
$script:DarkButtonBorder  = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#555555')
# Active-view indicator: shade-shift the BACKGROUND, not the border. White-on-blue
# borders in Light theme are only 4.53:1 contrast, same as the white button text,
# so a thicker border just shrinks the visible button. Background shifts use the
# brand's pressed-state shades (#3A3A3A dark / #005A9E light) so the active button
# reads as one visible step darker than its inactive siblings under either theme.
$script:DarkActiveBg      = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#3A3A3A')
$script:LightWfBg         = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:LightWfBorder     = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#006CBE')
$script:LightActiveBg     = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#005A9E')

$script:TitleBarBlue         = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:TitleBarBlueInactive = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#4BA3E0')

$script:LogLabelDark  = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#B0B0B0')
$script:LogLabelLight = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#595959')

$script:ViewButtons = @(
    @{ Name = 'Supersedence'; Button = $btnViewSupersedence },
    @{ Name = 'Dependencies'; Button = $btnViewDependencies },
    @{ Name = 'Broken Rules'; Button = $btnViewBroken       },
    @{ Name = 'Tree View';    Button = $btnViewTree         }
)
$script:ActiveView = 'Supersedence'

function Update-SidebarButtonTheme {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only.')]
    param()
    $isDark   = [bool]$global:Prefs['DarkMode']
    $idleBg   = if ($isDark) { $script:DarkButtonBg }     else { $script:LightWfBg }
    $activeBg = if ($isDark) { $script:DarkActiveBg }     else { $script:LightActiveBg }
    $border   = if ($isDark) { $script:DarkButtonBorder } else { $script:LightWfBorder }
    $thickness = [System.Windows.Thickness]::new(1)

    foreach ($v in $script:ViewButtons) {
        if (-not $v.Button) { continue }
        $isActive = ($v.Name -eq $script:ActiveView)
        $v.Button.Background      = if ($isActive) { $activeBg } else { $idleBg }
        $v.Button.BorderBrush     = $border
        $v.Button.BorderThickness = $thickness
    }
    if ($btnOptions) {
        $btnOptions.Background      = $idleBg
        $btnOptions.BorderBrush     = $border
        $btnOptions.BorderThickness = $thickness
    }
    if ($lblLogOutput) {
        $lblLogOutput.Foreground = if ($isDark) { $script:LogLabelDark } else { $script:LogLabelLight }
    }
}

function Update-TitleBarBrushes {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Sets both active and non-active title brushes per theme.')]
    param()
    $isDark = [bool]$global:Prefs['DarkMode']
    if ($isDark) {
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::WindowTitleBrushProperty)
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::NonActiveWindowTitleBrushProperty)
    } else {
        $window.WindowTitleBrush          = $script:TitleBarBlue
        $window.NonActiveWindowTitleBrush = $script:TitleBarBlueInactive
    }
}

$__startIsDark = [bool]$global:Prefs['DarkMode']
$toggleTheme.IsOn = $__startIsDark
$txtThemeLabel.Text = if ($__startIsDark) { 'Dark Theme' } else { 'Light Theme' }
Update-SidebarButtonTheme
# NOTE: ChangeTheme to a non-default theme + WindowTitleBrush mutation are
# DEFERRED to $window.Add_Loaded. Calling them at script-top (before the
# MetroWindow's WindowChromeBehavior attaches during the first layout pass)
# leaves the title bar's NCHITTEST routing client-area instead of caption,
# which silently breaks mouse-drag of the title bar. app-packager's reference
# pattern only mutates these DPs from inside the post-Loaded toggle handler.

$toggleTheme.Add_Toggled({
    $isDark = [bool]$toggleTheme.IsOn
    if ($isDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Dark.Steel')
        $txtThemeLabel.Text = 'Dark Theme'
    } else {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Light.Blue')
        $txtThemeLabel.Text = 'Light Theme'
    }
    $global:Prefs['DarkMode'] = $isDark
    Save-SaPreferences -Prefs $global:Prefs
    Update-SidebarButtonTheme
    Update-TitleBarBrushes
    Add-LogLine ('Theme: {0}' -f $(if ($isDark) { 'dark' } else { 'light' }))
})

# =============================================================================
# View switching.
# =============================================================================
$script:ViewMeta = @{
    'Supersedence' = @{ Title = 'Supersedence'; Subtitle = 'Superseding / superseded application pairs with chain depth and status. Scan to populate.' }
    'Dependencies' = @{ Title = 'Dependencies'; Subtitle = 'Parent / dependency pairs by type (Required, Optional, App). Scan to populate.' }
    'Broken Rules' = @{ Title = 'Broken Rules'; Subtitle = 'Orphaned, circular, expired, disabled, missing-content, and undocumented findings across both relationship types.' }
    'Tree View'    = @{ Title = 'Tree View';    Subtitle = 'Hierarchical visualization. Two roots: Supersedence Chains and Dependency Trees. Click a node for application details.' }
}

function Set-ActiveView {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates in-window Visibility + header text only.')]
    param([Parameter(Mandatory)][ValidateSet('Supersedence','Dependencies','Broken Rules','Tree View')][string]$View)

    $script:ActiveView = $View

    $viewSupersedence.Visibility = if ($View -eq 'Supersedence') { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewDependencies.Visibility = if ($View -eq 'Dependencies') { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewBroken.Visibility       = if ($View -eq 'Broken Rules') { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewTree.Visibility         = if ($View -eq 'Tree View')    { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }

    $meta = $script:ViewMeta[$View]
    if ($meta) {
        $txtModuleTitle.Text    = $meta.Title
        $txtModuleSubtitle.Text = $meta.Subtitle
    }

    Update-SidebarButtonTheme
    Update-ActionBarVisibility
    Update-Filter
    Update-StatusBarSummary
}

$btnViewSupersedence.Add_Click({ Set-ActiveView -View 'Supersedence' })
$btnViewDependencies.Add_Click({ Set-ActiveView -View 'Dependencies' })
$btnViewBroken.Add_Click({       Set-ActiveView -View 'Broken Rules' })
$btnViewTree.Add_Click({         Set-ActiveView -View 'Tree View' })

# =============================================================================
# Crash handlers (PS51-WPF-010, PS51-WPF-011, PS51-WPF-025).
# =============================================================================
$global:__crashLog = Join-Path $__txDir ('SupersedenceAuditor-crash-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

$global:__writeCrash = {
    param($Source, $Exception)
    try {
        $lines = @()
        $lines += ('=== ' + $Source + ' @ ' + (Get-Date -Format 'o') + ' ===')
        $lines += ('Type   : ' + $Exception.GetType().FullName)
        $lines += ('Message: ' + $Exception.Message)
        $lines += ('Stack  :')
        $lines += ([string]$Exception.StackTrace).Split([Environment]::NewLine)
        $inner = $Exception.InnerException
        $depth = 1
        while ($inner) {
            $lines += ('--- InnerException depth ' + $depth + ' ---')
            $lines += ('Type   : ' + $inner.GetType().FullName)
            $lines += ('Message: ' + $inner.Message)
            $lines += ('Stack  :')
            $lines += ([string]$inner.StackTrace).Split([Environment]::NewLine)
            $inner = $inner.InnerException
            $depth++
        }
        [System.IO.File]::AppendAllText($global:__crashLog, (($lines -join [Environment]::NewLine) + [Environment]::NewLine))
    } catch { $null = $_ }
}

$window.Dispatcher.Add_UnhandledException({
    param($s, $e)
    & $global:__writeCrash 'DispatcherUnhandledException' $e.Exception
    $e.Handled = $false
})

[AppDomain]::CurrentDomain.Add_UnhandledException({
    param($s, $e)
    & $global:__writeCrash 'AppDomainUnhandledException' ([Exception]$e.ExceptionObject)
})

# =============================================================================
# Glyph mapping for status / severity (per brand: check, x, ellipsis, warn at ThemeForeground).
# =============================================================================
function Get-StatusGlyph {
    param([string]$Status)
    # 0x2713=check, 0x2717=x, 0x26A0=warn, 0x22EF=ellipsis
    switch ($Status) {
        'Healthy'         { return [char]0x2713 }
        'Orphaned'        { return [char]0x2717 }
        'Circular'        { return [char]0x2717 }
        'Missing Content' { return [char]0x2717 }
        'Expired Target'  { return [char]0x26A0 }
        'Disabled Source' { return [char]0x26A0 }
        'Disabled Target' { return [char]0x26A0 }
        'Undocumented'    { return [char]0x22EF }
        default           { return '' }
    }
}

function Get-SeverityGlyph {
    param([string]$Severity)
    switch ($Severity) {
        'Error'   { return [char]0x2717 }
        'Warning' { return [char]0x26A0 }
        'Info'    { return [char]0x22EF }
        default   { return '' }
    }
}

# =============================================================================
# Scan state.
# =============================================================================
$script:AppLookup        = @{}
$script:SupersedenceData = @()
$script:DependencyData   = @()
$script:BrokenData       = @()
$script:ScanCounts       = $null
$script:LastScanTime     = $null
# Connection lives in the background scan runspace, so Test-CMConnection on the
# UI thread always returns false. Track scan-side success here for status text.
$script:IsConnectedFromScan = $false

# Decorated ItemsSource collections (with StatusGlyph / SeverityGlyph fields).
$script:SupersedenceRows = @()
$script:DependencyRows   = @()
$script:BrokenRows       = @()

# =============================================================================
# Filter / status combobox / detail panel wiring.
# =============================================================================
function Get-StatusFilterValue {
    if (-not $cboStatus.SelectedItem) { return 'All' }
    $item = $cboStatus.SelectedItem
    if ($item -is [System.Windows.Controls.ComboBoxItem]) { return [string]$item.Content }
    return [string]$item
}

function Test-StatusFilterMatch {
    param([string]$Status, [string]$Severity, [string]$Filter)
    switch ($Filter) {
        'All' { return $true }
        'Healthy' {
            return ($Status -eq 'Healthy')
        }
        'Broken / Warning' {
            if ($Severity) { return ($Severity -in @('Warning','Error')) }
            return ($Status -ne 'Healthy')
        }
        'Error' {
            if ($Severity) { return ($Severity -eq 'Error') }
            return ($Status -in @('Orphaned','Circular','Missing Content'))
        }
        default { return $true }
    }
}

function Update-Filter {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Recomputes ItemsSource on the active grid only; no external state.')]
    param()

    $needle = ([string]$txtFilter.Text).Trim().ToLowerInvariant()
    $statusFilter = Get-StatusFilterValue

    switch ($script:ActiveView) {
        'Supersedence' {
            $rows = $script:SupersedenceRows
            if ($needle) {
                $rows = @($rows | Where-Object {
                    ([string]$_.SupersedingApp).ToLowerInvariant().Contains($needle) -or
                    ([string]$_.SupersededApp).ToLowerInvariant().Contains($needle)
                })
            }
            if ($statusFilter -ne 'All') {
                $rows = @($rows | Where-Object { Test-StatusFilterMatch -Status $_.Status -Severity '' -Filter $statusFilter })
            }
            $gridSupersedence.ItemsSource = $rows
        }
        'Dependencies' {
            $rows = $script:DependencyRows
            if ($needle) {
                $rows = @($rows | Where-Object {
                    ([string]$_.ParentApp).ToLowerInvariant().Contains($needle) -or
                    ([string]$_.DependencyApp).ToLowerInvariant().Contains($needle)
                })
            }
            if ($statusFilter -ne 'All') {
                $rows = @($rows | Where-Object { Test-StatusFilterMatch -Status $_.Status -Severity '' -Filter $statusFilter })
            }
            $gridDependencies.ItemsSource = $rows
        }
        'Broken Rules' {
            $rows = $script:BrokenRows
            if ($needle) {
                $rows = @($rows | Where-Object {
                    ([string]$_.FromApp).ToLowerInvariant().Contains($needle) -or
                    ([string]$_.ToApp).ToLowerInvariant().Contains($needle) -or
                    ([string]$_.Description).ToLowerInvariant().Contains($needle)
                })
            }
            if ($statusFilter -ne 'All') {
                $rows = @($rows | Where-Object { Test-StatusFilterMatch -Status '' -Severity $_.Severity -Filter $statusFilter })
            }
            $gridBroken.ItemsSource = $rows
        }
        'Tree View' {
            # Tree view is not filtered; scan rebuilds it whole.
        }
    }
}

$txtFilter.Add_TextChanged({ Update-Filter })
$cboStatus.Add_SelectionChanged({ Update-Filter })

# Detail panels.
$gridSupersedence.Add_SelectionChanged({
    $row = $gridSupersedence.SelectedItem
    if (-not $row) { $txtSupersedenceDetail.Text = 'Select a row to see relationship details.'; return }

    $lines = @(
        'SUPERSEDENCE RELATIONSHIP',
        ('-' * 40),
        '',
        ('Superseding: {0} ({1})' -f $row.SupersedingApp, $row.SupersedingVersion),
        ('  CI_ID:     {0}' -f $row.SupersedingCIID),
        '',
        ('Superseded:  {0} ({1})' -f $row.SupersededApp, $row.SupersededVersion),
        ('  CI_ID:     {0}' -f $row.SupersededCIID),
        '',
        ('Chain Depth: {0}' -f $row.ChainDepth),
        ('Status:      {0}' -f $row.Status)
    )

    $supCIID = [int]$row.SupersedingCIID
    $sedCIID = [int]$row.SupersededCIID
    if ($script:AppLookup.ContainsKey($supCIID)) {
        $app = $script:AppLookup[$supCIID]
        $lines += '', 'Superseding App Details:'
        $lines += ('  Manufacturer: {0}' -f $app.Manufacturer)
        $lines += ('  Enabled:      {0}' -f $app.IsEnabled)
        $lines += ('  Deployments:  {0}' -f $app.NumberOfDeployments)
        $lines += ('  Created:      {0}' -f $app.DateCreated)
        $lines += ('  Created By:   {0}' -f $app.CreatedBy)
    }
    if ($script:AppLookup.ContainsKey($sedCIID)) {
        $app = $script:AppLookup[$sedCIID]
        $lines += '', 'Superseded App Details:'
        $lines += ('  Manufacturer: {0}' -f $app.Manufacturer)
        $lines += ('  Expired:      {0}' -f $app.IsExpired)
        $lines += ('  Enabled:      {0}' -f $app.IsEnabled)
        $lines += ('  Deployments:  {0}' -f $app.NumberOfDeployments)
    }
    $txtSupersedenceDetail.Text = $lines -join [Environment]::NewLine
})

$gridDependencies.Add_SelectionChanged({
    $row = $gridDependencies.SelectedItem
    if (-not $row) { $txtDependencyDetail.Text = 'Select a row to see dependency details.'; return }

    $lines = @(
        'DEPENDENCY RELATIONSHIP',
        ('-' * 40),
        '',
        ('Parent:     {0} ({1})' -f $row.ParentApp, $row.ParentVersion),
        ('  CI_ID:    {0}' -f $row.ParentCIID),
        '',
        ('Dependency: {0} ({1})' -f $row.DependencyApp, $row.DependencyVersion),
        ('  CI_ID:    {0}' -f $row.DependencyCIID),
        '',
        ('Type:       {0}' -f $row.DependencyType),
        ('Level:      {0}' -f $row.Level),
        ('Status:     {0}' -f $row.Status)
    )

    $depCIID = [int]$row.DependencyCIID
    if ($script:AppLookup.ContainsKey($depCIID)) {
        $app = $script:AppLookup[$depCIID]
        $lines += '', 'Dependency App Details:'
        $lines += ('  Manufacturer: {0}' -f $app.Manufacturer)
        $lines += ('  Enabled:      {0}' -f $app.IsEnabled)
        $lines += ('  Expired:      {0}' -f $app.IsExpired)
        $lines += ('  Has Content:  {0}' -f $app.HasContent)
        $lines += ('  Deployments:  {0}' -f $app.NumberOfDeployments)
    }
    $txtDependencyDetail.Text = $lines -join [Environment]::NewLine
})

$gridBroken.Add_SelectionChanged({
    $row = $gridBroken.SelectedItem
    if (-not $row) { $txtBrokenDetail.Text = 'Select a row to see issue details and remediation.'; return }

    $lines = @(
        'BROKEN RULE DETAILS',
        ('-' * 40),
        '',
        ('Issue:       {0}' -f $row.IssueType),
        ('Severity:    {0}' -f $row.Severity),
        ('Category:    {0}' -f $row.Category),
        ('From App:    {0}' -f $row.FromApp),
        ('To App:      {0}' -f $row.ToApp),
        '',
        'Description:',
        $row.Description,
        '',
        'Remediation:',
        $row.Remediation
    )
    $txtBrokenDetail.Text = $lines -join [Environment]::NewLine
})

$treeRelationships.Add_SelectedItemChanged({
    $node = $treeRelationships.SelectedItem
    if (-not $node -or -not $node.Tag) { $txtTreeDetail.Text = 'Select a node to see application details.'; return }

    $ciid = 0
    if (-not [int]::TryParse([string]$node.Tag, [ref]$ciid)) {
        $txtTreeDetail.Text = 'Select a node to see application details.'
        return
    }
    if (-not $script:AppLookup.ContainsKey($ciid)) {
        $txtTreeDetail.Text = "Application CI_ID: $ciid (not found in current scan)"
        return
    }

    $app = $script:AppLookup[$ciid]
    $lines = @(
        'APPLICATION DETAILS',
        ('-' * 40),
        '',
        ('Name:           {0}' -f $app.LocalizedDisplayName),
        ('Version:        {0}' -f $app.SoftwareVersion),
        ('Manufacturer:   {0}' -f $app.Manufacturer),
        ('CI_ID:          {0}' -f $app.CI_ID),
        '',
        ('Enabled:        {0}' -f $app.IsEnabled),
        ('Expired:        {0}' -f $app.IsExpired),
        ('Is Superseded:  {0}' -f $app.IsSuperseded),
        ('Is Superseding: {0}' -f $app.IsSuperseding),
        ('Has Content:    {0}' -f $app.HasContent),
        '',
        ('Deploy Types:   {0}' -f $app.NumberOfDeploymentTypes),
        ('Deployments:    {0}' -f $app.NumberOfDeployments),
        '',
        ('Created:        {0}' -f $app.DateCreated),
        ('Created By:     {0}' -f $app.CreatedBy),
        ('Modified:       {0}' -f $app.DateLastModified),
        ('Modified By:    {0}' -f $app.LastModifiedBy)
    )
    $txtTreeDetail.Text = $lines -join [Environment]::NewLine
})

# =============================================================================
# Action bar visibility and status bar summary.
# =============================================================================
function Update-ActionBarVisibility {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Toggles in-window Visibility only.')]
    param()

    $hasData = ($null -ne $script:ScanCounts)
    $isTree  = ($script:ActiveView -eq 'Tree View')

    # CSV / HTML available on the three grid views once data is loaded.
    $exportVis = if ($hasData -and -not $isTree) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $btnExportCsv.Visibility  = $exportVis
    $btnExportHtml.Visibility = $exportVis

    # Copy Summary requires scan data; not view-specific.
    $btnCopySummary.Visibility = if ($hasData) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
}

function Update-StatusBarSummary {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates an in-window TextBlock only.')]
    param()

    $parts = @()
    if ($script:IsConnectedFromScan -and $global:Prefs.SiteCode) {
        $parts += "Connected to $($global:Prefs.SiteCode)"
    } elseif (-not $global:Prefs.SiteCode -or -not $global:Prefs.SMSProvider) {
        $parts += 'Open Options to configure site code and SMS provider'
    } else {
        $parts += 'Ready. Click Scan Environment.'
    }
    if ($script:ScanCounts) {
        $c = $script:ScanCounts
        $parts += ('{0} apps' -f $c.AppCount)
        $parts += ('{0} supersedence' -f $c.SupersedenceTotal)
        $parts += ('{0} dependencies' -f $c.DependencyTotal)
        $parts += ('{0} broken' -f $c.BrokenRulesTotal)
    }
    if ($script:LastScanTime) {
        $parts += ('last scan {0}' -f $script:LastScanTime.ToString('HH:mm:ss'))
    }
    Set-StatusText ($parts -join '   |   ')
}

# =============================================================================
# Decorate raw module rows with glyphs for grid display.
# =============================================================================
function ConvertTo-SupersedenceGridRows {
    param($Rows)
    $out = @()
    foreach ($r in @($Rows)) {
        $out += [PSCustomObject]@{
            StatusGlyph        = Get-StatusGlyph -Status $r.Status
            SupersedingApp     = $r.SupersedingApp
            SupersedingVersion = $r.SupersedingVersion
            SupersedingCIID    = $r.SupersedingCIID
            SupersededApp      = $r.SupersededApp
            SupersededVersion  = $r.SupersededVersion
            SupersededCIID     = $r.SupersededCIID
            ChainDepth         = $r.ChainDepth
            Status             = $r.Status
        }
    }
    return ,$out
}

function ConvertTo-DependencyGridRows {
    param($Rows)
    $out = @()
    foreach ($r in @($Rows)) {
        $out += [PSCustomObject]@{
            StatusGlyph       = Get-StatusGlyph -Status $r.Status
            ParentApp         = $r.ParentApp
            ParentVersion     = $r.ParentVersion
            ParentCIID        = $r.ParentCIID
            DependencyApp     = $r.DependencyApp
            DependencyVersion = $r.DependencyVersion
            DependencyCIID    = $r.DependencyCIID
            DependencyType    = $r.DependencyType
            Level             = $r.Level
            Status            = $r.Status
        }
    }
    return ,$out
}

function ConvertTo-BrokenGridRows {
    param($Rows)
    $out = @()
    foreach ($r in @($Rows)) {
        $out += [PSCustomObject]@{
            SeverityGlyph = Get-SeverityGlyph -Severity $r.Severity
            IssueType     = $r.IssueType
            Severity      = $r.Severity
            Category      = $r.Category
            FromApp       = $r.FromApp
            ToApp         = $r.ToApp
            Description   = $r.Description
            Remediation   = $r.Remediation
        }
    }
    return ,$out
}

# =============================================================================
# Tree population. Uses existing Build-* functions; nodes carry CIID in Tag.
# =============================================================================
function Add-TreeNode {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates the in-window TreeView only.')]
    param(
        [Parameter(Mandatory)]$Parent,
        [Parameter(Mandatory)][string]$Header,
        $Tag = $null,
        [bool]$Bold = $false
    )
    $node = New-Object System.Windows.Controls.TreeViewItem
    $node.Header = $Header
    if ($null -ne $Tag) { $node.Tag = $Tag }
    if ($Bold) { $node.FontWeight = [System.Windows.FontWeights]::SemiBold }
    [void]$Parent.Items.Add($node)
    return $node
}

function Add-SupersedenceChildNodes {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates the in-window TreeView only.')]
    param($Parent, $Children)
    foreach ($child in @($Children)) {
        $glyph = Get-StatusGlyph -Status $child.Status
        $prefix = if ($glyph) { ('{0}  ' -f $glyph) } else { '' }
        $header = '{0}{1} ({2})' -f $prefix, $child.Name, $child.Version
        $childNode = Add-TreeNode -Parent $Parent -Header $header -Tag $child.CIID
        if ($child.Children -and $child.Children.Count -gt 0) {
            Add-SupersedenceChildNodes -Parent $childNode -Children $child.Children
        }
    }
}

function Set-TreeViewData {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates the in-window TreeView only.')]
    param($SupersedenceRoots, $DependencyRoots)

    $treeRelationships.Items.Clear()

    $supRoot = New-Object System.Windows.Controls.TreeViewItem
    $supRoot.Header = ('Supersedence Chains ({0})' -f @($SupersedenceRoots).Count)
    $supRoot.FontWeight = [System.Windows.FontWeights]::Bold
    [void]$treeRelationships.Items.Add($supRoot)

    foreach ($root in @($SupersedenceRoots)) {
        $glyph = Get-StatusGlyph -Status $root.Status
        $prefix = if ($glyph) { ('{0}  ' -f $glyph) } else { '' }
        $header = '{0}{1} ({2})' -f $prefix, $root.Name, $root.Version
        $rootNode = Add-TreeNode -Parent $supRoot -Header $header -Tag $root.CIID -Bold:$true
        if ($root.Children -and $root.Children.Count -gt 0) {
            Add-SupersedenceChildNodes -Parent $rootNode -Children $root.Children
        }
    }
    $supRoot.IsExpanded = $true

    $depRoot = New-Object System.Windows.Controls.TreeViewItem
    $depRoot.Header = ('Dependency Trees ({0})' -f @($DependencyRoots).Count)
    $depRoot.FontWeight = [System.Windows.FontWeights]::Bold
    [void]$treeRelationships.Items.Add($depRoot)

    foreach ($root in @($DependencyRoots)) {
        $glyph = Get-StatusGlyph -Status $root.Status
        $prefix = if ($glyph) { ('{0}  ' -f $glyph) } else { '' }
        $header = '{0}{1} ({2})' -f $prefix, $root.Name, $root.Version
        $rootNode = Add-TreeNode -Parent $depRoot -Header $header -Tag $root.CIID -Bold:$true
        if ($root.Children -and $root.Children.Count -gt 0) {
            foreach ($child in @($root.Children)) {
                $cGlyph = Get-StatusGlyph -Status $child.Status
                $cPrefix = if ($cGlyph) { ('{0}  ' -f $cGlyph) } else { '' }
                $cHeader = '{0}{1} ({2}) [{3}]' -f $cPrefix, $child.Name, $child.Version, $child.Type
                [void](Add-TreeNode -Parent $rootNode -Header $cHeader -Tag $child.CIID)
            }
        }
    }
    $depRoot.IsExpanded = $true
}

# =============================================================================
# Background scan runspace. Scan is potentially long (large environments parse
# thousands of SDMPackageXML blobs). Run in a STA runspace and poll a
# DispatcherTimer so the spinner animates and the UI remains responsive.
# =============================================================================
$script:BgRunspace     = $null
$script:BgPowerShell   = $null
$script:BgInvokeHandle = $null
$script:ScanState      = $null
$script:ScanTimer      = $null

function Initialize-ScanRunspace {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Lazy-init of the background scan runspace; idempotent.')]
    param()
    if ($script:BgRunspace -and $script:BgRunspace.RunspaceStateInfo.State -eq 'Opened') { return }

    $script:BgRunspace = [runspacefactory]::CreateRunspace()
    $script:BgRunspace.ApartmentState = 'STA'
    $script:BgRunspace.ThreadOptions  = 'ReuseThread'
    $script:BgRunspace.Open()

    $modulePath = Join-Path $PSScriptRoot 'Module\SupersedenceAuditorCommon.psd1'

    # The bg runspace gets its OWN module instance with its own module-scoped
    # state, so the UI thread's Initialize-Logging call doesn't reach it.
    # Without this second Initialize-Logging, every Write-Log inside the bg
    # scan path is silently dropped (module-scoped log path stays $null).
    $initPS = [powershell]::Create()
    $initPS.Runspace = $script:BgRunspace
    [void]$initPS.AddScript({
        param($ModulePath, $LogPath)
        Import-Module -Name $ModulePath -Force -DisableNameChecking
        if ($LogPath) { Initialize-Logging -LogPath $LogPath -Attach }
    }).AddArgument($modulePath).AddArgument($script:toolLogPath)
    [void]$initPS.Invoke()
    $initPS.Dispose()
}

function Invoke-ScanEnvironment {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Posts work to the background runspace and arms a DispatcherTimer.')]
    param()

    if (-not $global:Prefs.SiteCode -or -not $global:Prefs.SMSProvider) {
        Add-LogLine 'Scan: site code and SMS provider must be set in Options first.'
        Set-StatusText 'Open Options to configure site code and SMS provider, then scan.'
        return
    }

    Initialize-ScanRunspace

    if ($script:ScanTimer)    { try { $script:ScanTimer.Stop() } catch { $null = $_ } }
    if ($script:BgPowerShell) {
        try { [void]$script:BgPowerShell.Stop() } catch { $null = $_ }
        try { $script:BgPowerShell.Dispose() }   catch { $null = $_ }
        $script:BgPowerShell = $null
    }

    $script:ScanState = [hashtable]::Synchronized(@{
        Step     = 'Connecting...'
        Done     = $false
        Result   = $null
        ErrorMsg = $null
    })

    $btnScan.IsEnabled = $false
    $txtProgressTitle.Text = 'Scanning environment'
    $txtProgressStep.Text  = 'Connecting...'
    $progressOverlay.Visibility = [System.Windows.Visibility]::Visible
    Add-LogLine ('Scan: site={0} provider={1}' -f $global:Prefs.SiteCode, $global:Prefs.SMSProvider)
    Set-StatusText 'Scanning...'

    $siteCode    = [string]$global:Prefs.SiteCode
    $smsProvider = [string]$global:Prefs.SMSProvider

    $script:BgPowerShell = [powershell]::Create()
    $script:BgPowerShell.Runspace = $script:BgRunspace
    [void]$script:BgPowerShell.AddScript({
        param($SiteCode, $SMSProvider, $State)
        try {
            if (-not (Test-CMConnection)) {
                $State.Step = "Connecting to $SiteCode..."
                $ok = Connect-CMSite -SiteCode $SiteCode -SMSProvider $SMSProvider
                if (-not $ok) {
                    $State.ErrorMsg = "Failed to connect to site $SiteCode (provider $SMSProvider)."
                    return
                }
            }

            $State.Step = 'Loading applications (Get-CMApplication, bulk)...'
            $appLookup = Get-AllApplicationSummary

            $State.Step = ('Parsing SDMPackageXML for {0} applications...' -f $appLookup.Count)
            $resolved = Get-AllResolvedRelationships -AppLookup $appLookup
            if (-not $resolved) { $resolved = @() }

            $State.Step = 'Analyzing supersedence chains...'
            $supersedence = @(Find-SupersedenceChains -ResolvedRelationships $resolved -AppLookup $appLookup)

            $State.Step = 'Analyzing dependency groups...'
            $dependencies = @(Find-DependencyGroups -ResolvedRelationships $resolved -AppLookup $appLookup)

            $State.Step = 'Detecting broken rules...'
            $brokenSup    = @(Find-BrokenSupersedence -SupersedenceData $supersedence -ResolvedRelationships $resolved -AppLookup $appLookup)
            $brokenDep    = @(Find-BrokenDependencies -DependencyData $dependencies -ResolvedRelationships $resolved -AppLookup $appLookup)
            $undocumented = @(Find-UndocumentedRelationships -SupersedenceData $supersedence -DependencyData $dependencies -AppLookup $appLookup)
            $broken       = @($brokenSup) + @($brokenDep) + @($undocumented)

            $State.Step = 'Building tree view...'
            $supRoots = Build-SupersedenceTree -SupersedenceData $supersedence -AppLookup $appLookup
            $depRoots = Build-DependencyTree   -DependencyData   $dependencies -AppLookup $appLookup

            $counts = Get-ScanSummaryCounts `
                -AppCount        $appLookup.Count `
                -SupersedenceData $supersedence `
                -DependencyData   $dependencies `
                -BrokenRules      $broken

            $State.Result = [PSCustomObject]@{
                AppLookup        = $appLookup
                SupersedenceData = $supersedence
                DependencyData   = $dependencies
                BrokenData       = $broken
                SupersedenceRoots = $supRoots
                DependencyRoots   = $depRoots
                Counts            = $counts
            }
        }
        catch {
            $State.ErrorMsg = $_.Exception.Message
        }
        finally {
            $State.Done = $true
        }
    }).AddArgument($siteCode).AddArgument($smsProvider).AddArgument($script:ScanState)

    $script:BgInvokeHandle = $script:BgPowerShell.BeginInvoke()

    $script:ScanTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:ScanTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:ScanTimer.Add_Tick({
        if ($script:ScanState) {
            $current = [string]$script:ScanState.Step
            if ($txtProgressStep.Text -ne $current) { $txtProgressStep.Text = $current }
        }
        if ($script:ScanState -and $script:ScanState.Done) {
            $script:ScanTimer.Stop()
            try { [void]$script:BgPowerShell.EndInvoke($script:BgInvokeHandle) } catch { $null = $_ }
            try { $script:BgPowerShell.Dispose() } catch { $null = $_ }
            $script:BgPowerShell   = $null
            $script:BgInvokeHandle = $null

            if ($script:ScanState.ErrorMsg) {
                $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
                $btnScan.IsEnabled = $true
                $script:IsConnectedFromScan = $false
                Add-LogLine ('Scan failed: {0}' -f $script:ScanState.ErrorMsg)
                Set-StatusText 'Scan failed.'
                return
            }

            $script:IsConnectedFromScan = $true
            $r = $script:ScanState.Result
            $script:AppLookup        = $r.AppLookup
            $script:SupersedenceData = @($r.SupersedenceData)
            $script:DependencyData   = @($r.DependencyData)
            $script:BrokenData       = @($r.BrokenData)
            $script:ScanCounts       = $r.Counts
            $script:LastScanTime     = Get-Date

            $script:SupersedenceRows = ConvertTo-SupersedenceGridRows -Rows $script:SupersedenceData
            $script:DependencyRows   = ConvertTo-DependencyGridRows   -Rows $script:DependencyData
            $script:BrokenRows       = ConvertTo-BrokenGridRows       -Rows $script:BrokenData

            $gridSupersedence.ItemsSource = $script:SupersedenceRows
            $gridDependencies.ItemsSource = $script:DependencyRows
            $gridBroken.ItemsSource       = $script:BrokenRows
            Set-TreeViewData -SupersedenceRoots $r.SupersedenceRoots -DependencyRoots $r.DependencyRoots

            Update-Filter
            Update-ActionBarVisibility
            Update-StatusBarSummary

            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $btnScan.IsEnabled = $true

            Add-LogLine ('Scan complete: {0} apps, {1} supersedence, {2} dependencies, {3} broken' -f $r.Counts.AppCount, $r.Counts.SupersedenceTotal, $r.Counts.DependencyTotal, $r.Counts.BrokenRulesTotal)
        }
    })
    $script:ScanTimer.Start()
}

$btnScan.Add_Click({ Invoke-ScanEnvironment })

# =============================================================================
# Export and copy summary.
# =============================================================================
function ConvertTo-DataTable {
    param([Parameter(Mandatory)]$Rows, [string[]]$Columns)
    $dt = New-Object System.Data.DataTable
    if ($Columns -and $Columns.Count -gt 0) {
        foreach ($c in $Columns) { [void]$dt.Columns.Add($c, [string]) }
    } elseif (@($Rows).Count -gt 0) {
        $first = @($Rows)[0]
        foreach ($p in $first.PSObject.Properties) { [void]$dt.Columns.Add($p.Name, [string]) }
    } else {
        return $dt
    }
    foreach ($row in @($Rows)) {
        $vals = @()
        foreach ($col in $dt.Columns) {
            $val = $row.PSObject.Properties[$col.ColumnName].Value
            $vals += [string]$val
        }
        [void]$dt.Rows.Add($vals)
    }
    return $dt
}

function Get-ActiveExportInfo {
    switch ($script:ActiveView) {
        'Supersedence' {
            return @{
                Name    = 'Supersedence'
                Columns = @('SupersedingApp','SupersedingVersion','SupersededApp','SupersededVersion','ChainDepth','Status')
                Rows    = $gridSupersedence.ItemsSource
            }
        }
        'Dependencies' {
            return @{
                Name    = 'Dependencies'
                Columns = @('ParentApp','ParentVersion','DependencyApp','DependencyVersion','DependencyType','Level','Status')
                Rows    = $gridDependencies.ItemsSource
            }
        }
        'Broken Rules' {
            return @{
                Name    = 'BrokenRules'
                Columns = @('IssueType','Severity','Category','FromApp','ToApp','Description','Remediation')
                Rows    = $gridBroken.ItemsSource
            }
        }
        default { return $null }
    }
}

$btnExportCsv.Add_Click({
    $info = Get-ActiveExportInfo
    if (-not $info -or -not $info.Rows -or @($info.Rows).Count -eq 0) {
        Add-LogLine 'Export CSV: nothing to export.'
        return
    }
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = 'CSV files (*.csv)|*.csv'
    $sfd.FileName = ('Audit-{0}-{1}.csv' -f $info.Name, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $reportsDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq $true) {
        $dt = ConvertTo-DataTable -Rows $info.Rows -Columns $info.Columns
        Export-AuditCsv -DataTable $dt -OutputPath $sfd.FileName
        Add-LogLine ('Exported CSV: {0}' -f $sfd.FileName)
    }
})

$btnExportHtml.Add_Click({
    $info = Get-ActiveExportInfo
    if (-not $info -or -not $info.Rows -or @($info.Rows).Count -eq 0) {
        Add-LogLine 'Export HTML: nothing to export.'
        return
    }
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = 'HTML files (*.html)|*.html'
    $sfd.FileName = ('Audit-{0}-{1}.html' -f $info.Name, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $reportsDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq $true) {
        $dt = ConvertTo-DataTable -Rows $info.Rows -Columns $info.Columns
        Export-AuditHtml -DataTable $dt -OutputPath $sfd.FileName -ReportTitle ('Audit Report: {0}' -f $info.Name)
        Add-LogLine ('Exported HTML: {0}' -f $sfd.FileName)
    }
})

$btnCopySummary.Add_Click({
    if (-not $script:ScanCounts) {
        Add-LogLine 'Copy Summary: no scan data yet.'
        return
    }
    $summary = New-AuditSummaryText -Counts $script:ScanCounts
    [System.Windows.Clipboard]::SetText($summary)
    Add-LogLine 'Summary copied to clipboard.'
})

# =============================================================================
# Options dialog (Site / Provider / About). MetroWindow inline XAML; no File
# menu per brand.
# =============================================================================
function Show-OptionsDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; verb-noun reads as a single action.')]
    param()

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Options"
    Width="640" Height="380"
    MinWidth="560" MinHeight="380"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1"
    ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="CategoryRowStyle" TargetType="Button"
                   BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="Height" Value="36"/>
                <Setter Property="HorizontalContentAlignment" Value="Left"/>
                <Setter Property="Padding" Value="14,0,14,0"/>
                <Setter Property="FontSize" Value="13"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
                <Setter Property="Margin" Value="0"/>
            </Style>
            <Style x:Key="DialogButton" TargetType="Button"
                   BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/>
                <Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button"
                   BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/>
                <Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="180"/>
            <ColumnDefinition Width="1"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Column="0" Grid.Row="0" Padding="6,12,0,12">
            <StackPanel>
                <Button x:Name="btnCatConnection" Content="Connection" Style="{StaticResource CategoryRowStyle}"/>
                <Button x:Name="btnCatAbout"      Content="About"      Style="{StaticResource CategoryRowStyle}"/>
            </StackPanel>
        </Border>

        <Border Grid.Column="1" Grid.Row="0" Background="{DynamicResource MahApps.Brushes.Gray8}"/>

        <Grid Grid.Column="2" Grid.Row="0" Margin="20,16,20,16">

            <StackPanel x:Name="paneConnection" Visibility="Visible">
                <TextBlock Text="MECM Connection" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <TextBlock Text="Site Code" FontSize="11" Margin="0,4,0,2"
                           Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <TextBox x:Name="txtSiteCode" FontSize="12" Padding="6,4,6,4"
                         Controls:TextBoxHelper.Watermark="e.g. P01" Width="120" HorizontalAlignment="Left"/>
                <TextBlock Text="SMS Provider FQDN" FontSize="11" Margin="0,12,0,2"
                           Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <TextBox x:Name="txtSmsProvider" FontSize="12" Padding="6,4,6,4"
                         Controls:TextBoxHelper.Watermark="e.g. cm01.contoso.com"/>
                <TextBlock Text="Used for the CM PSDrive root. Read-only access only -- the Auditor calls Get-CMSite and Get-CMApplication, never any Set / New / Remove cmdlet."
                           FontSize="11" TextWrapping="Wrap" Margin="0,16,0,0"
                           Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            </StackPanel>

            <StackPanel x:Name="paneAbout" Visibility="Collapsed">
                <TextBlock Text="About" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <TextBlock x:Name="txtAboutVersion" Text="Supersedence and Dependency Auditor v1.0.0"
                           FontSize="13" FontWeight="SemiBold"/>
                <TextBlock Text="Maps every supersedence and dependency relationship in your MECM environment from a single bulk Get-CMApplication query plus in-memory SDMPackageXML parse. Detects orphans, circular chains, expired targets, disabled sources, missing content, and undocumented apps."
                           FontSize="12" TextWrapping="Wrap" Margin="0,8,0,0"/>
                <TextBlock Text="Read-only. No Set / New / Remove / Add cmdlets are called against the SMS Provider."
                           FontSize="12" TextWrapping="Wrap" Margin="0,12,0,0"/>
                <TextBlock Text="Author: Jason Ulbright. License: MIT."
                           FontSize="11" Margin="0,16,0,0"
                           Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            </StackPanel>
        </Grid>

        <Border Grid.Row="1" Grid.ColumnSpan="3" Padding="16,12,16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnOk"     Content="OK"     Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
                <Button x:Name="btnCancel" Content="Cancel" Style="{StaticResource DialogButton}" IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    Install-TitleBarDragFallback -Window $dlg

    $isDark = [bool]$global:Prefs['DarkMode']
    if ($isDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, 'Dark.Steel')
    } else {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, 'Light.Blue')
        $dlg.WindowTitleBrush          = $script:TitleBarBlue
        $dlg.NonActiveWindowTitleBrush = $script:TitleBarBlueInactive
    }

    $btnCatConnection = $dlg.FindName('btnCatConnection')
    $btnCatAbout      = $dlg.FindName('btnCatAbout')
    $paneConnection   = $dlg.FindName('paneConnection')
    $paneAbout        = $dlg.FindName('paneAbout')
    $txtAboutVersion  = $dlg.FindName('txtAboutVersion')
    if ($txtAboutVersion) { $txtAboutVersion.Text = ($txtAboutVersion.Text -replace 'v[0-9][0-9\.]*[0-9]\s*$', ('v' + $script:AppVersion)) }
    $txtSiteCode      = $dlg.FindName('txtSiteCode')
    $txtSmsProvider   = $dlg.FindName('txtSmsProvider')
    $btnOk            = $dlg.FindName('btnOk')
    $btnCancel        = $dlg.FindName('btnCancel')

    $txtSiteCode.Text    = [string]$global:Prefs.SiteCode
    $txtSmsProvider.Text = [string]$global:Prefs.SMSProvider

    $btnCatConnection.Add_Click({
        $paneConnection.Visibility = [System.Windows.Visibility]::Visible
        $paneAbout.Visibility      = [System.Windows.Visibility]::Collapsed
    })
    $btnCatAbout.Add_Click({
        $paneConnection.Visibility = [System.Windows.Visibility]::Collapsed
        $paneAbout.Visibility      = [System.Windows.Visibility]::Visible
    })

    $btnOk.Add_Click({
        $newSite     = ([string]$txtSiteCode.Text).Trim()
        $newProvider = ([string]$txtSmsProvider.Text).Trim()
        $connectionChanged = ($newSite -ne [string]$global:Prefs.SiteCode) -or
                             ($newProvider -ne [string]$global:Prefs.SMSProvider)

        $global:Prefs.SiteCode    = $newSite
        $global:Prefs.SMSProvider = $newProvider
        Save-SaPreferences -Prefs $global:Prefs

        if ($connectionChanged) {
            # Bg runspace caches the prior CM connection. Recycle it so the
            # next scan reconnects with the new site / provider values.
            #
            # If a scan was in flight when Options was opened (Options stays
            # clickable during scan), we ALSO need to tear down the scan
            # timer + reset scan UI state. Otherwise the timer keeps polling
            # a disposed runspace's ScanState (Done stays false forever) and
            # the user is stuck on the progress overlay with the Scan button
            # permanently disabled.
            if ($script:ScanTimer) {
                try { $script:ScanTimer.Stop() } catch { $null = $_ }
                $script:ScanTimer = $null
            }
            if ($script:BgPowerShell) {
                try { [void]$script:BgPowerShell.Stop() } catch { $null = $_ }
                try { $script:BgPowerShell.Dispose() }   catch { $null = $_ }
                $script:BgPowerShell = $null
            }
            if ($script:BgRunspace) {
                try { $script:BgRunspace.Close() }  catch { $null = $_ }
                try { $script:BgRunspace.Dispose() } catch { $null = $_ }
                $script:BgRunspace = $null
            }
            $script:BgInvokeHandle      = $null
            $script:ScanState           = $null
            $script:IsConnectedFromScan = $false

            # Reset scan UI in case a scan was running.
            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $btnScan.IsEnabled          = $true
        }

        $dlg.DialogResult = $true
        $dlg.Close()
    })
    $btnCancel.Add_Click({
        $dlg.DialogResult = $false
        $dlg.Close()
    })

    [void]$dlg.ShowDialog()

    Update-StatusBarSummary
}

$btnOptions.Add_Click({ Show-OptionsDialog })

# =============================================================================
# Window state persistence.
# =============================================================================
$global:WindowStatePath = Join-Path $PSScriptRoot 'SupersedenceAuditor.windowstate.json'

function Save-WindowState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Writes a small JSON state file; idempotent.')]
    param()
    try {
        $state = @{
            Left       = [int]$window.Left
            Top        = [int]$window.Top
            Width      = [int]$window.Width
            Height     = [int]$window.Height
            Maximized  = ($window.WindowState -eq [System.Windows.WindowState]::Maximized)
            ActiveView = $script:ActiveView
        }
        $state | ConvertTo-Json | Set-Content -LiteralPath $global:WindowStatePath -Encoding UTF8
    } catch { $null = $_ }
}

function Restore-WindowState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Reads the JSON state file and applies geometry; idempotent.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification='Restore is intentional and reads as a single action.')]
    param()
    if (-not (Test-Path -LiteralPath $global:WindowStatePath)) { return }
    try {
        $s = Get-Content -LiteralPath $global:WindowStatePath -Raw | ConvertFrom-Json -ErrorAction Stop

        # Schema bridge: WinForms 1.0 used X/Y/ActiveTab; WPF refresh uses
        # Left/Top/ActiveView. Read both so legacy files don't snap the window
        # to (0,0) at MinSize after the upgrade.
        $left = if ($null -ne $s.Left) { [int]$s.Left } elseif ($null -ne $s.X) { [int]$s.X } else { $null }
        $top  = if ($null -ne $s.Top)  { [int]$s.Top  } elseif ($null -ne $s.Y) { [int]$s.Y } else { $null }
        $w    = if ($null -ne $s.Width)  { [int]$s.Width  } else { $null }
        $h    = if ($null -ne $s.Height) { [int]$s.Height } else { $null }

        if ($s.Maximized) {
            $window.WindowState = [System.Windows.WindowState]::Maximized
        } elseif ($null -ne $left -and $null -ne $top -and $null -ne $w -and $null -ne $h) {
            $screen = [System.Windows.Forms.Screen]::FromPoint([System.Drawing.Point]::new($left, $top))
            $bounds = $screen.WorkingArea
            $left = [Math]::Max($bounds.X, [Math]::Min($left, $bounds.Right - 200))
            $top  = [Math]::Max($bounds.Y, [Math]::Min($top,  $bounds.Bottom - 100))
            $window.Left   = $left
            $window.Top    = $top
            $window.Width  = [Math]::Max($window.MinWidth,  $w)
            $window.Height = [Math]::Max($window.MinHeight, $h)
        }

        if ($s.ActiveView -in @('Supersedence','Dependencies','Broken Rules','Tree View')) {
            Set-ActiveView -View ([string]$s.ActiveView)
        }
    } catch { $null = $_ }
}

$window.Add_Closing({
    Save-WindowState
    if ($script:ScanTimer)    { try { $script:ScanTimer.Stop() } catch { $null = $_ } }
    if ($script:BgPowerShell) {
        try { [void]$script:BgPowerShell.Stop() } catch { $null = $_ }
        try { $script:BgPowerShell.Dispose() }   catch { $null = $_ }
    }
    if ($script:BgRunspace) {
        try { $script:BgRunspace.Close() }  catch { $null = $_ }
        try { $script:BgRunspace.Dispose() } catch { $null = $_ }
    }
    if (Test-CMConnection) { try { Disconnect-CMSite } catch { $null = $_ } }
})

$window.Add_Loaded({
    Restore-WindowState

    # Apply user theme prefs AFTER the chrome has fully attached. Calling
    # ChangeTheme + WindowTitleBrush mutation at script-top breaks the title
    # bar's NCHITTEST routing (drag becomes a no-op). See the note where
    # Update-TitleBarBrushes is defined.
    $isDark = [bool]$global:Prefs['DarkMode']
    if (-not $isDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Light.Blue')
    }
    Update-TitleBarBrushes

    Update-ActionBarVisibility
    Update-StatusBarSummary
    Add-LogLine 'Supersedence and Dependency Auditor ready. Configure Site / Provider in Options, then click Scan Environment.'
})

# =============================================================================
# Run.
# =============================================================================
[void]$window.ShowDialog()
try { Stop-Transcript | Out-Null } catch { $null = $_ }
