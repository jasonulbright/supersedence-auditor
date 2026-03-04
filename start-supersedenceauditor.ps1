<#
.SYNOPSIS
    WinForms front-end for MECM Supersedence & Dependency Auditor.

.DESCRIPTION
    Provides a GUI for discovering and auditing all application supersedence and
    dependency relationships in an MECM environment. Detects broken rules
    (orphaned references, circular chains, expired targets, disabled sources,
    missing content) and visualizes relationship trees.

    Features:
      - Bulk WMI scan of all applications, deployment types, and relationships
      - Supersedence chain discovery with depth calculation
      - Dependency group analysis (Required/Optional)
      - Broken rule detection with remediation guidance
      - TreeView visualization of relationship hierarchies
      - Summary cards with at-a-glance status
      - Export to CSV, HTML, and clipboard
      - Dark mode / light mode

.EXAMPLE
    .\start-supersedenceauditor.ps1

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8+
      - Windows Forms (System.Windows.Forms)
      - Configuration Manager console installed

    ScriptName : start-supersedenceauditor.ps1
    Purpose    : WinForms front-end for MECM supersedence and dependency auditing
    Version    : 1.0.0
    Updated    : 2026-03-03
#>

param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

$moduleRoot = Join-Path $PSScriptRoot "Module"
Import-Module (Join-Path $moduleRoot "SupersedenceAuditorCommon.psd1") -Force -DisableNameChecking

# Initialize tool logging
$toolLogFolder = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $toolLogFolder)) {
    New-Item -ItemType Directory -Path $toolLogFolder -Force | Out-Null
}
$toolLogPath = Join-Path $toolLogFolder ("Audit-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $toolLogPath

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Set-ModernButtonStyle {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Button]$Button,
        [Parameter(Mandatory)][System.Drawing.Color]$BackColor
    )

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $BackColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.UseVisualStyleBackColor = $false
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand

    $hover = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BackColor.R - 18),
        [Math]::Max(0, $BackColor.G - 18),
        [Math]::Max(0, $BackColor.B - 18)
    )
    $down = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BackColor.R - 36),
        [Math]::Max(0, $BackColor.G - 36),
        [Math]::Max(0, $BackColor.B - 36)
    )

    $Button.FlatAppearance.MouseOverBackColor = $hover
    $Button.FlatAppearance.MouseDownBackColor = $down
}

function Enable-DoubleBuffer {
    param([Parameter(Mandatory)][System.Windows.Forms.Control]$Control)
    $prop = $Control.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
    if ($prop) { $prop.SetValue($Control, $true, $null) | Out-Null }
}

function Add-LogLine {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TextBox,
        [Parameter(Mandatory)][string]$Message
    )
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "{0}  {1}" -f $ts, $Message

    if ([string]::IsNullOrWhiteSpace($TextBox.Text)) {
        $TextBox.Text = $line
    }
    else {
        $TextBox.AppendText([Environment]::NewLine + $line)
    }

    $TextBox.SelectionStart = $TextBox.TextLength
    $TextBox.ScrollToCaret()
}

function Save-WindowState {
    $statePath = Join-Path $PSScriptRoot "SupersedenceAuditor.windowstate.json"
    $state = @{
        X         = $form.Location.X
        Y         = $form.Location.Y
        Width     = $form.Size.Width
        Height    = $form.Size.Height
        Maximized = ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Maximized)
        ActiveTab = $tabMain.SelectedIndex
    }

    # Save splitter distances per tab
    if ($splitSupersedence.SplitterDistance -gt 0) { $state.SplitSupersedence = $splitSupersedence.SplitterDistance }
    if ($splitDependencies.SplitterDistance -gt 0) { $state.SplitDependencies = $splitDependencies.SplitterDistance }
    if ($splitBroken.SplitterDistance -gt 0)       { $state.SplitBroken = $splitBroken.SplitterDistance }
    if ($splitTree.SplitterDistance -gt 0)         { $state.SplitTree = $splitTree.SplitterDistance }

    $state | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
}

function Restore-WindowState {
    $statePath = Join-Path $PSScriptRoot "SupersedenceAuditor.windowstate.json"
    if (-not (Test-Path -LiteralPath $statePath)) { return }

    try {
        $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        if ($state.Maximized) {
            $form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
        } else {
            # Validate against screen bounds
            $screen = [System.Windows.Forms.Screen]::FromPoint((New-Object System.Drawing.Point($state.X, $state.Y)))
            $bounds = $screen.WorkingArea
            $x = [Math]::Max($bounds.X, [Math]::Min($state.X, $bounds.Right - 200))
            $y = [Math]::Max($bounds.Y, [Math]::Min($state.Y, $bounds.Bottom - 100))
            $form.Location = New-Object System.Drawing.Point($x, $y)
            $form.Size = New-Object System.Drawing.Size(
                [Math]::Max($form.MinimumSize.Width, $state.Width),
                [Math]::Max($form.MinimumSize.Height, $state.Height)
            )
        }
        if ($null -ne $state.ActiveTab -and $state.ActiveTab -ge 0 -and $state.ActiveTab -lt $tabMain.TabCount) {
            $tabMain.SelectedIndex = [int]$state.ActiveTab
        }
        if ($null -ne $state.SplitSupersedence) { try { $splitSupersedence.SplitterDistance = [int]$state.SplitSupersedence } catch {} }
        if ($null -ne $state.SplitDependencies) { try { $splitDependencies.SplitterDistance = [int]$state.SplitDependencies } catch {} }
        if ($null -ne $state.SplitBroken)       { try { $splitBroken.SplitterDistance = [int]$state.SplitBroken } catch {} }
        if ($null -ne $state.SplitTree)         { try { $splitTree.SplitterDistance = [int]$state.SplitTree } catch {} }
    } catch { }
}

# ---------------------------------------------------------------------------
# Preferences
# ---------------------------------------------------------------------------

function Get-SaPreferences {
    $prefsPath = Join-Path $PSScriptRoot "SupersedenceAuditor.prefs.json"
    $defaults = @{
        DarkMode    = $false
        SiteCode    = ''
        SMSProvider = ''
    }

    if (Test-Path -LiteralPath $prefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
            if ($null -ne $loaded.DarkMode)  { $defaults.DarkMode    = [bool]$loaded.DarkMode }
            if ($loaded.SiteCode)            { $defaults.SiteCode    = $loaded.SiteCode }
            if ($loaded.SMSProvider)          { $defaults.SMSProvider = $loaded.SMSProvider }
        } catch { }
    }

    return $defaults
}

function Save-SaPreferences {
    param([hashtable]$Prefs)
    $prefsPath = Join-Path $PSScriptRoot "SupersedenceAuditor.prefs.json"
    $Prefs | ConvertTo-Json | Set-Content -LiteralPath $prefsPath -Encoding UTF8
}

$script:Prefs = Get-SaPreferences

# ---------------------------------------------------------------------------
# Colors (theme-aware)
# ---------------------------------------------------------------------------

$clrAccent = [System.Drawing.Color]::FromArgb(0, 120, 212)

if ($script:Prefs.DarkMode) {
    $clrFormBg     = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $clrPanelBg    = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $clrHint       = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle   = [System.Drawing.Color]::FromArgb(180, 200, 220)
    $clrGridAlt    = [System.Drawing.Color]::FromArgb(48, 48, 48)
    $clrGridLine   = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $clrDetailBg   = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $clrSepLine    = [System.Drawing.Color]::FromArgb(55, 55, 55)
    $clrLogBg      = [System.Drawing.Color]::FromArgb(35, 35, 35)
    $clrLogFg      = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $clrText       = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrGridText   = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrErrText    = [System.Drawing.Color]::FromArgb(255, 100, 100)
    $clrWarnText   = [System.Drawing.Color]::FromArgb(255, 200, 80)
    $clrOkText     = [System.Drawing.Color]::FromArgb(80, 200, 80)
    $clrInfoText   = [System.Drawing.Color]::FromArgb(100, 180, 255)
    $clrCardGreen  = [System.Drawing.Color]::FromArgb(30, 60, 30)
    $clrCardYellow = [System.Drawing.Color]::FromArgb(60, 50, 20)
    $clrCardRed    = [System.Drawing.Color]::FromArgb(60, 25, 25)
    $clrCardBlue   = [System.Drawing.Color]::FromArgb(25, 40, 60)
    $clrTreeBg     = [System.Drawing.Color]::FromArgb(38, 38, 38)
} else {
    $clrFormBg     = [System.Drawing.Color]::FromArgb(245, 246, 248)
    $clrPanelBg    = [System.Drawing.Color]::White
    $clrHint       = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle   = [System.Drawing.Color]::FromArgb(220, 230, 245)
    $clrGridAlt    = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $clrGridLine   = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $clrDetailBg   = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $clrSepLine    = [System.Drawing.Color]::FromArgb(218, 220, 224)
    $clrLogBg      = [System.Drawing.Color]::White
    $clrLogFg      = [System.Drawing.Color]::Black
    $clrText       = [System.Drawing.Color]::Black
    $clrGridText   = [System.Drawing.Color]::Black
    $clrErrText    = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $clrWarnText   = [System.Drawing.Color]::FromArgb(180, 120, 0)
    $clrOkText     = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $clrInfoText   = [System.Drawing.Color]::FromArgb(0, 100, 180)
    $clrCardGreen  = [System.Drawing.Color]::FromArgb(220, 245, 220)
    $clrCardYellow = [System.Drawing.Color]::FromArgb(255, 248, 220)
    $clrCardRed    = [System.Drawing.Color]::FromArgb(255, 225, 225)
    $clrCardBlue   = [System.Drawing.Color]::FromArgb(220, 235, 255)
    $clrTreeBg     = [System.Drawing.Color]::White
}

# Custom dark mode ToolStrip renderer
if ($script:Prefs.DarkMode) {
    if (-not ('DarkToolStripRenderer' -as [type])) {
        $rendererCs = (
            'using System.Drawing;',
            'using System.Windows.Forms;',
            'public class DarkToolStripRenderer : ToolStripProfessionalRenderer {',
            '    private Color _bg;',
            '    public DarkToolStripRenderer(Color bg) : base() { _bg = bg; }',
            '    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { }',
            '    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) {',
            '        using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); }',
            '    }',
            '    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) {',
            '        if (e.Item.Selected || e.Item.Pressed) {',
            '            using (var b = new SolidBrush(Color.FromArgb(60, 60, 60)))',
            '            { e.Graphics.FillRectangle(b, new Rectangle(Point.Empty, e.Item.Size)); }',
            '        }',
            '    }',
            '    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e) {',
            '        int y = e.Item.Height / 2;',
            '        using (var p = new Pen(Color.FromArgb(70, 70, 70)))',
            '        { e.Graphics.DrawLine(p, 0, y, e.Item.Width, y); }',
            '    }',
            '    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) {',
            '        using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); }',
            '    }',
            '}'
        ) -join "`r`n"
        Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition $rendererCs
    }
    $script:DarkRenderer = New-Object DarkToolStripRenderer($clrPanelBg)
}

# ---------------------------------------------------------------------------
# Preferences dialog
# ---------------------------------------------------------------------------

function Show-PreferencesDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Preferences"
    $dlg.Size = New-Object System.Drawing.Size(440, 300)
    $dlg.MinimumSize = $dlg.Size
    $dlg.MaximumSize = $dlg.Size
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $dlg.BackColor = $clrFormBg

    # Appearance
    $grpAppearance = New-Object System.Windows.Forms.GroupBox
    $grpAppearance.Text = "Appearance"
    $grpAppearance.SetBounds(16, 12, 392, 60)
    $grpAppearance.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpAppearance.ForeColor = $clrText
    $grpAppearance.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpAppearance.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpAppearance.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpAppearance)

    $chkDark = New-Object System.Windows.Forms.CheckBox
    $chkDark.Text = "Enable dark mode (requires restart)"
    $chkDark.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkDark.AutoSize = $true
    $chkDark.Location = New-Object System.Drawing.Point(14, 24)
    $chkDark.Checked = $script:Prefs.DarkMode
    $chkDark.ForeColor = $clrText
    $chkDark.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $chkDark.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkDark.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
    $grpAppearance.Controls.Add($chkDark)

    # MECM Connection
    $grpConn = New-Object System.Windows.Forms.GroupBox
    $grpConn.Text = "MECM Connection"
    $grpConn.SetBounds(16, 82, 392, 110)
    $grpConn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpConn.ForeColor = $clrText
    $grpConn.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpConn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpConn.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpConn)

    $lblSiteCode = New-Object System.Windows.Forms.Label
    $lblSiteCode.Text = "Site Code:"
    $lblSiteCode.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSiteCode.Location = New-Object System.Drawing.Point(14, 30)
    $lblSiteCode.AutoSize = $true
    $lblSiteCode.ForeColor = $clrText
    $grpConn.Controls.Add($lblSiteCode)

    $txtSiteCodePref = New-Object System.Windows.Forms.TextBox
    $txtSiteCodePref.SetBounds(130, 27, 80, 24)
    $txtSiteCodePref.Text = $script:Prefs.SiteCode
    $txtSiteCodePref.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtSiteCodePref.BackColor = $clrDetailBg
    $txtSiteCodePref.ForeColor = $clrText
    $grpConn.Controls.Add($txtSiteCodePref)

    $lblSMSProv = New-Object System.Windows.Forms.Label
    $lblSMSProv.Text = "SMS Provider:"
    $lblSMSProv.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSMSProv.Location = New-Object System.Drawing.Point(14, 64)
    $lblSMSProv.AutoSize = $true
    $lblSMSProv.ForeColor = $clrText
    $grpConn.Controls.Add($lblSMSProv)

    $txtSMSProvPref = New-Object System.Windows.Forms.TextBox
    $txtSMSProvPref.SetBounds(130, 61, 240, 24)
    $txtSMSProvPref.Text = $script:Prefs.SMSProvider
    $txtSMSProvPref.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtSMSProvPref.BackColor = $clrDetailBg
    $txtSMSProvPref.ForeColor = $clrText
    $grpConn.Controls.Add($txtSMSProvPref)

    # Buttons
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save"
    $btnSave.SetBounds(220, 210, 90, 32)
    $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    Set-ModernButtonStyle -Button $btnSave -BackColor $clrAccent
    $dlg.Controls.Add($btnSave)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.SetBounds(318, 210, 90, 32)
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.ForeColor = $clrText
    $btnCancel.BackColor = $clrFormBg
    $dlg.Controls.Add($btnCancel)

    $btnSave.Add_Click({
        $script:Prefs.DarkMode    = $chkDark.Checked
        $script:Prefs.SiteCode    = $txtSiteCodePref.Text.Trim()
        $script:Prefs.SMSProvider = $txtSMSProvPref.Text.Trim()
        Save-SaPreferences -Prefs $script:Prefs

        # Update connection bar labels
        $lblSiteVal.Text    = if ($script:Prefs.SiteCode) { $script:Prefs.SiteCode } else { "(not set)" }
        $lblProviderVal.Text = if ($script:Prefs.SMSProvider) { $script:Prefs.SMSProvider } else { "(not set)" }
        Update-StatusBar

        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlg.Close() })

    $dlg.AcceptButton = $btnSave
    $dlg.CancelButton = $btnCancel
    $dlg.ShowDialog($form) | Out-Null
    $dlg.Dispose()
}

# ---------------------------------------------------------------------------
# Form
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Supersedence & Dependency Auditor"
$form.Size = New-Object System.Drawing.Size(1280, 860)
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$form.BackColor = $clrFormBg

# Try to load icon
$icoPath = Join-Path $PSScriptRoot "supersedenceauditor.ico"
if (Test-Path -LiteralPath $icoPath) {
    try { $form.Icon = New-Object System.Drawing.Icon($icoPath) } catch { }
}

# ---------------------------------------------------------------------------
# StatusStrip (Dock:Bottom) -- add early so it stays at very bottom
# ---------------------------------------------------------------------------

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $clrPanelBg
$statusStrip.ForeColor = $clrText
$statusStrip.SizingGrip = $false
if ($script:Prefs.DarkMode -and $script:DarkRenderer) { $statusStrip.Renderer = $script:DarkRenderer }

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Disconnected"
$statusLabel.Spring = $true
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusLabel.ForeColor = $clrText
$statusStrip.Items.Add($statusLabel) | Out-Null

$statusRowCount = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusRowCount.Text = ""
$statusRowCount.ForeColor = $clrHint
$statusStrip.Items.Add($statusRowCount) | Out-Null

$form.Controls.Add($statusStrip)

# ---------------------------------------------------------------------------
# Log console (Dock:Bottom)
# ---------------------------------------------------------------------------

$pnlLog = New-Object System.Windows.Forms.Panel
$pnlLog.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlLog.Height = 95
$pnlLog.BackColor = $clrLogBg
$form.Controls.Add($pnlLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.WordWrap = $true
$txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor = $clrLogBg
$txtLog.ForeColor = $clrLogFg
$txtLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$pnlLog.Controls.Add($txtLog)

# Separator above log
$pnlLogSep = New-Object System.Windows.Forms.Panel
$pnlLogSep.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlLogSep.Height = 1
$pnlLogSep.BackColor = $clrSepLine
$form.Controls.Add($pnlLogSep)

# ---------------------------------------------------------------------------
# Button panel (Dock:Bottom)
# ---------------------------------------------------------------------------

$pnlButtons = New-Object System.Windows.Forms.Panel
$pnlButtons.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlButtons.Height = 56
$pnlButtons.BackColor = $clrPanelBg
$pnlButtons.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 10)
$form.Controls.Add($pnlButtons)

$flowButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$flowButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowButtons.WrapContents = $false
$flowButtons.BackColor = $clrPanelBg
$pnlButtons.Controls.Add($flowButtons)

$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Text = "Export CSV"
$btnExportCsv.Size = New-Object System.Drawing.Size(120, 34)
$btnExportCsv.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnExportCsv.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
Set-ModernButtonStyle -Button $btnExportCsv -BackColor ([System.Drawing.Color]::FromArgb(50, 130, 50))

$btnExportHtml = New-Object System.Windows.Forms.Button
$btnExportHtml.Text = "Export HTML"
$btnExportHtml.Size = New-Object System.Drawing.Size(120, 34)
$btnExportHtml.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnExportHtml.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
Set-ModernButtonStyle -Button $btnExportHtml -BackColor ([System.Drawing.Color]::FromArgb(50, 130, 50))

$btnCopySummary = New-Object System.Windows.Forms.Button
$btnCopySummary.Text = "Copy Summary"
$btnCopySummary.Size = New-Object System.Drawing.Size(130, 34)
$btnCopySummary.Font = New-Object System.Drawing.Font("Segoe UI", 9)
Set-ModernButtonStyle -Button $btnCopySummary -BackColor ([System.Drawing.Color]::FromArgb(100, 100, 100))

$flowButtons.Controls.Add($btnExportCsv)
$flowButtons.Controls.Add($btnExportHtml)
$flowButtons.Controls.Add($btnCopySummary)

# Separator above buttons
$pnlBtnSep = New-Object System.Windows.Forms.Panel
$pnlBtnSep.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlBtnSep.Height = 1
$pnlBtnSep.BackColor = $clrSepLine
$form.Controls.Add($pnlBtnSep)

# ---------------------------------------------------------------------------
# MenuStrip
# ---------------------------------------------------------------------------

$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.BackColor = $clrPanelBg
$menuStrip.ForeColor = $clrText
if ($script:Prefs.DarkMode -and $script:DarkRenderer) { $menuStrip.Renderer = $script:DarkRenderer }

# File menu
$mnuFile = New-Object System.Windows.Forms.ToolStripMenuItem("&File")
$mnuFile.ForeColor = $clrText
$mnuPrefs = New-Object System.Windows.Forms.ToolStripMenuItem("&Preferences...")
$mnuPrefs.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Oemcomma
$mnuPrefs.ForeColor = $clrText
$mnuPrefs.Add_Click({ Show-PreferencesDialog })
$mnuFile.DropDownItems.Add($mnuPrefs) | Out-Null
$mnuFile.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
$mnuExit = New-Object System.Windows.Forms.ToolStripMenuItem("E&xit")
$mnuExit.ForeColor = $clrText
$mnuExit.Add_Click({ $form.Close() })
$mnuFile.DropDownItems.Add($mnuExit) | Out-Null
$menuStrip.Items.Add($mnuFile) | Out-Null

# View menu
$mnuView = New-Object System.Windows.Forms.ToolStripMenuItem("&View")
$mnuView.ForeColor = $clrText
$tabNames = @('Supersedence', 'Dependencies', 'Broken Rules', 'Tree View')
for ($idx = 0; $idx -lt $tabNames.Count; $idx++) {
    $mnuItem = New-Object System.Windows.Forms.ToolStripMenuItem($tabNames[$idx])
    $mnuItem.ForeColor = $clrText
    $mnuItem.Tag = $idx
    $mnuItem.Add_Click({ $tabMain.SelectedIndex = [int]$this.Tag }.GetNewClosure())
    $mnuView.DropDownItems.Add($mnuItem) | Out-Null
}
$menuStrip.Items.Add($mnuView) | Out-Null

# Help menu
$mnuHelp = New-Object System.Windows.Forms.ToolStripMenuItem("&Help")
$mnuHelp.ForeColor = $clrText
$mnuAbout = New-Object System.Windows.Forms.ToolStripMenuItem("&About")
$mnuAbout.ForeColor = $clrText
$mnuAbout.Add_Click({
    [System.Windows.Forms.MessageBox]::Show(
        "Supersedence & Dependency Auditor v1.0.0`r`n`r`nMaps all application supersedence and dependency relationships in your MECM environment. Detects broken rules and visualizes relationship hierarchies.`r`n`r`nRequires: ConfigMgr console, WMI access to SMS Provider.",
        "About", "OK", "Information") | Out-Null
})
$mnuHelp.DropDownItems.Add($mnuAbout) | Out-Null
$menuStrip.Items.Add($mnuHelp) | Out-Null

# ---------------------------------------------------------------------------
# Header panel (Dock:Top)
# ---------------------------------------------------------------------------

$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlHeader.Height = 60
$pnlHeader.BackColor = $clrAccent

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Supersedence & Dependency Auditor"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(16, 6)
$lblTitle.BackColor = [System.Drawing.Color]::Transparent
$pnlHeader.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = "Application Relationship Analysis"
$lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSubtitle.ForeColor = $clrSubtitle
$lblSubtitle.AutoSize = $true
$lblSubtitle.Location = New-Object System.Drawing.Point(18, 36)
$lblSubtitle.BackColor = [System.Drawing.Color]::Transparent
$pnlHeader.Controls.Add($lblSubtitle)

$form.Controls.Add($pnlHeader)

# ---------------------------------------------------------------------------
# Connection bar (Dock:Top)
# ---------------------------------------------------------------------------

$pnlConnBar = New-Object System.Windows.Forms.Panel
$pnlConnBar.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlConnBar.Height = 40
$pnlConnBar.BackColor = $clrPanelBg
$pnlConnBar.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)
$form.Controls.Add($pnlConnBar)

$flowConn = New-Object System.Windows.Forms.FlowLayoutPanel
$flowConn.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowConn.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowConn.WrapContents = $false
$flowConn.BackColor = $clrPanelBg
$pnlConnBar.Controls.Add($flowConn)

$lblSite = New-Object System.Windows.Forms.Label
$lblSite.Text = "Site:"
$lblSite.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblSite.AutoSize = $true
$lblSite.Margin = New-Object System.Windows.Forms.Padding(0, 5, 2, 0)
$lblSite.ForeColor = $clrText
$lblSite.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSite)

$lblSiteVal = New-Object System.Windows.Forms.Label
$lblSiteVal.Text = if ($script:Prefs.SiteCode) { $script:Prefs.SiteCode } else { "(not set)" }
$lblSiteVal.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSiteVal.AutoSize = $true
$lblSiteVal.Margin = New-Object System.Windows.Forms.Padding(0, 5, 16, 0)
$lblSiteVal.ForeColor = $clrHint
$lblSiteVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSiteVal)

$lblProvider = New-Object System.Windows.Forms.Label
$lblProvider.Text = "Provider:"
$lblProvider.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblProvider.AutoSize = $true
$lblProvider.Margin = New-Object System.Windows.Forms.Padding(0, 5, 2, 0)
$lblProvider.ForeColor = $clrText
$lblProvider.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblProvider)

$lblProviderVal = New-Object System.Windows.Forms.Label
$lblProviderVal.Text = if ($script:Prefs.SMSProvider) { $script:Prefs.SMSProvider } else { "(not set)" }
$lblProviderVal.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblProviderVal.AutoSize = $true
$lblProviderVal.Margin = New-Object System.Windows.Forms.Padding(0, 5, 24, 0)
$lblProviderVal.ForeColor = $clrHint
$lblProviderVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblProviderVal)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan Environment"
$btnScan.Size = New-Object System.Drawing.Size(150, 26)
$btnScan.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnScan.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
Set-ModernButtonStyle -Button $btnScan -BackColor $clrAccent

$flowConn.Controls.Add($btnScan)

# Separator below connection bar
$pnlSep1 = New-Object System.Windows.Forms.Panel
$pnlSep1.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep1.Height = 1
$pnlSep1.BackColor = $clrSepLine
$form.Controls.Add($pnlSep1)

# ---------------------------------------------------------------------------
# Summary cards (Dock:Top)
# ---------------------------------------------------------------------------

$pnlCards = New-Object System.Windows.Forms.Panel
$pnlCards.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlCards.Height = 56
$pnlCards.BackColor = $clrFormBg
$pnlCards.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
$form.Controls.Add($pnlCards)

$flowCards = New-Object System.Windows.Forms.FlowLayoutPanel
$flowCards.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowCards.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowCards.WrapContents = $false
$flowCards.BackColor = $clrFormBg
$pnlCards.Controls.Add($flowCards)

function New-SummaryCard {
    param(
        [string]$Title,
        [int]$TabIndex
    )

    $card = New-Object System.Windows.Forms.Panel
    $card.Size = New-Object System.Drawing.Size(220, 44)
    $card.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 2)
    $card.BackColor = $clrPanelBg
    $card.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $card.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Tag = $TabIndex

    # Left color indicator bar
    $bar = New-Object System.Windows.Forms.Panel
    $bar.Dock = [System.Windows.Forms.DockStyle]::Left
    $bar.Width = 4
    $bar.BackColor = $clrHint
    $card.Controls.Add($bar)

    $lblCardTitle = New-Object System.Windows.Forms.Label
    $lblCardTitle.Text = $Title
    $lblCardTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $lblCardTitle.ForeColor = $clrText
    $lblCardTitle.AutoSize = $true
    $lblCardTitle.Location = New-Object System.Drawing.Point(10, 4)
    $lblCardTitle.BackColor = [System.Drawing.Color]::Transparent
    $card.Controls.Add($lblCardTitle)

    $lblCardValue = New-Object System.Windows.Forms.Label
    $lblCardValue.Text = "--"
    $lblCardValue.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblCardValue.ForeColor = $clrHint
    $lblCardValue.AutoSize = $true
    $lblCardValue.Location = New-Object System.Drawing.Point(10, 22)
    $lblCardValue.BackColor = [System.Drawing.Color]::Transparent
    $lblCardValue.Tag = "value"
    $card.Controls.Add($lblCardValue)

    $clickHandler = { $tabMain.SelectedIndex = [int]$this.Parent.Tag }
    $cardClickHandler = { $tabMain.SelectedIndex = [int]$this.Tag }
    $card.Add_Click($cardClickHandler)
    $lblCardTitle.Add_Click($clickHandler)
    $lblCardValue.Add_Click($clickHandler)

    return $card
}

$cardApps        = New-SummaryCard -Title "Applications Scanned" -TabIndex 0
$cardSupersede   = New-SummaryCard -Title "Supersedence Rules"   -TabIndex 0
$cardDependency  = New-SummaryCard -Title "Dependency Rules"     -TabIndex 1
$cardBroken      = New-SummaryCard -Title "Broken Rules"         -TabIndex 2

$flowCards.Controls.Add($cardApps)
$flowCards.Controls.Add($cardSupersede)
$flowCards.Controls.Add($cardDependency)
$flowCards.Controls.Add($cardBroken)

function Update-Card {
    param(
        [System.Windows.Forms.Panel]$Card,
        [string]$ValueText,
        [string]$Severity   # 'ok', 'warn', 'critical', 'info'
    )

    $bar = $Card.Controls[0]
    $valLabel = $Card.Controls | Where-Object { $_.Tag -eq 'value' }

    switch ($Severity) {
        'ok'       { $bar.BackColor = $clrOkText;   $Card.BackColor = $clrCardGreen;  if ($valLabel) { $valLabel.ForeColor = $clrOkText } }
        'warn'     { $bar.BackColor = $clrWarnText;  $Card.BackColor = $clrCardYellow; if ($valLabel) { $valLabel.ForeColor = $clrWarnText } }
        'critical' { $bar.BackColor = $clrErrText;   $Card.BackColor = $clrCardRed;    if ($valLabel) { $valLabel.ForeColor = $clrErrText } }
        'info'     { $bar.BackColor = $clrInfoText;  $Card.BackColor = $clrCardBlue;   if ($valLabel) { $valLabel.ForeColor = $clrInfoText } }
        default    { $bar.BackColor = $clrHint;      $Card.BackColor = $clrPanelBg;    if ($valLabel) { $valLabel.ForeColor = $clrHint } }
    }

    if ($valLabel) { $valLabel.Text = $ValueText }
}

# Separator below cards
$pnlSep2 = New-Object System.Windows.Forms.Panel
$pnlSep2.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep2.Height = 1
$pnlSep2.BackColor = $clrSepLine
$form.Controls.Add($pnlSep2)

# ---------------------------------------------------------------------------
# Filter bar (Dock:Top)
# ---------------------------------------------------------------------------

$pnlFilter = New-Object System.Windows.Forms.Panel
$pnlFilter.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlFilter.Height = 44
$pnlFilter.BackColor = $clrPanelBg
$pnlFilter.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)
$form.Controls.Add($pnlFilter)

$flowFilter = New-Object System.Windows.Forms.FlowLayoutPanel
$flowFilter.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowFilter.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowFilter.WrapContents = $false
$flowFilter.BackColor = $clrPanelBg
$pnlFilter.Controls.Add($flowFilter)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Filter:"
$lblFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblFilter.AutoSize = $true
$lblFilter.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)
$lblFilter.ForeColor = $clrText
$lblFilter.BackColor = $clrPanelBg
$flowFilter.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$txtFilter.Width = 300
$txtFilter.Margin = New-Object System.Windows.Forms.Padding(0, 2, 16, 0)
$txtFilter.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
$txtFilter.BackColor = $clrDetailBg
$txtFilter.ForeColor = $clrText
$flowFilter.Controls.Add($txtFilter)

$lblStatusFilter = New-Object System.Windows.Forms.Label
$lblStatusFilter.Text = "Status:"
$lblStatusFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblStatusFilter.AutoSize = $true
$lblStatusFilter.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)
$lblStatusFilter.ForeColor = $clrText
$lblStatusFilter.BackColor = $clrPanelBg
$flowFilter.Controls.Add($lblStatusFilter)

$cboStatus = New-Object System.Windows.Forms.ComboBox
$cboStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboStatus.Width = 140
$cboStatus.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$cboStatus.BackColor = $clrDetailBg
$cboStatus.ForeColor = $clrText
$cboStatus.FlatStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.FlatStyle]::Flat } else { [System.Windows.Forms.FlatStyle]::Standard }
[void]$cboStatus.Items.AddRange(@('All', 'Healthy', 'Broken/Warning', 'Error'))
$cboStatus.SelectedIndex = 0
$flowFilter.Controls.Add($cboStatus)

# Separator below filter bar
$pnlSep3 = New-Object System.Windows.Forms.Panel
$pnlSep3.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep3.Height = 1
$pnlSep3.BackColor = $clrSepLine
$form.Controls.Add($pnlSep3)

# ---------------------------------------------------------------------------
# Helper: Create a themed DataGridView
# ---------------------------------------------------------------------------

function New-ThemedGrid {
    param([switch]$MultiSelect)

    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = [System.Windows.Forms.DockStyle]::Fill
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false
    $g.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $g.MultiSelect = [bool]$MultiSelect
    $g.AutoGenerateColumns = $false
    $g.RowHeadersVisible = $false
    $g.BackgroundColor = $clrPanelBg
    $g.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $g.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $g.GridColor = $clrGridLine
    $g.ColumnHeadersDefaultCellStyle.BackColor = $clrAccent
    $g.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $g.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $g.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4)
    $g.ColumnHeadersDefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
    $g.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $g.ColumnHeadersHeight = 32
    $g.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $g.EnableHeadersVisualStyles = $false
    $g.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $g.DefaultCellStyle.ForeColor = $clrGridText
    $g.DefaultCellStyle.BackColor = $clrPanelBg
    $g.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(2)
    $g.DefaultCellStyle.SelectionBackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(38, 79, 120) } else { [System.Drawing.Color]::FromArgb(0, 120, 215) }
    $g.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $g.RowTemplate.Height = 26
    $g.AlternatingRowsDefaultCellStyle.BackColor = $clrGridAlt

    Enable-DoubleBuffer -Control $g
    return $g
}

# ---------------------------------------------------------------------------
# TabControl (Fill)
# ---------------------------------------------------------------------------

$tabMain = New-Object System.Windows.Forms.TabControl
$tabMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabMain.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$tabMain.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$tabMain.ItemSize = New-Object System.Drawing.Size(140, 30)
$tabMain.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$tabMain.Add_DrawItem({
    param($s, $e)
    $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $tab = $s.TabPages[$e.Index]
    $isSelected = ($s.SelectedIndex -eq $e.Index)

    $bgColor = if ($script:Prefs.DarkMode) {
        if ($isSelected) { $clrAccent } else { $clrPanelBg }
    } else {
        if ($isSelected) { $clrAccent } else { [System.Drawing.Color]::FromArgb(240, 240, 240) }
    }
    $fgColor = if ($isSelected) { [System.Drawing.Color]::White } else { $clrText }

    $bgBrush = New-Object System.Drawing.SolidBrush($bgColor)
    $e.Graphics.FillRectangle($bgBrush, $e.Bounds)

    $font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Near
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Far
    $sf.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap

    $textRect = New-Object System.Drawing.RectangleF(
        ($e.Bounds.X + 8),
        $e.Bounds.Y,
        ($e.Bounds.Width - 12),
        ($e.Bounds.Height - 3)
    )

    $textBrush = New-Object System.Drawing.SolidBrush($fgColor)
    $e.Graphics.DrawString($tab.Text, $font, $textBrush, $textRect, $sf)

    $bgBrush.Dispose(); $textBrush.Dispose(); $font.Dispose(); $sf.Dispose()
})

$form.Controls.Add($tabMain)

# ===================== TAB 0: Supersedence =====================

$tabSupersedence = New-Object System.Windows.Forms.TabPage
$tabSupersedence.Text = "Supersedence"
$tabSupersedence.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabSupersedence)

$splitSupersedence = New-Object System.Windows.Forms.SplitContainer
$splitSupersedence.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitSupersedence.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitSupersedence.SplitterDistance = 350
$splitSupersedence.SplitterWidth = 6
$splitSupersedence.BackColor = $clrSepLine
$splitSupersedence.Panel1.BackColor = $clrPanelBg
$splitSupersedence.Panel2.BackColor = $clrPanelBg
$splitSupersedence.Panel1MinSize = 100
$splitSupersedence.Panel2MinSize = 80
$tabSupersedence.Controls.Add($splitSupersedence)

$gridSupersedence = New-ThemedGrid

$colSSupApp  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSSupApp.HeaderText = "Superseding App";   $colSSupApp.DataPropertyName = "SupersedingApp";     $colSSupApp.Width = 200
$colSSupVer  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSSupVer.HeaderText = "Version";           $colSSupVer.DataPropertyName = "SupersedingVersion"; $colSSupVer.Width = 90
$colSSedApp  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSSedApp.HeaderText = "Superseded App";    $colSSedApp.DataPropertyName = "SupersededApp";      $colSSedApp.Width = 200
$colSSedVer  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSSedVer.HeaderText = "Version";           $colSSedVer.DataPropertyName = "SupersededVersion";  $colSSedVer.Width = 90
$colSDepth   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSDepth.HeaderText = "Chain Depth";        $colSDepth.DataPropertyName = "ChainDepth";          $colSDepth.Width = 80
$colSStatus  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSStatus.HeaderText = "Status";            $colSStatus.DataPropertyName = "Status";             $colSStatus.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

$gridSupersedence.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colSSupApp, $colSSupVer, $colSSedApp, $colSSedVer, $colSDepth, $colSStatus))
$splitSupersedence.Panel1.Controls.Add($gridSupersedence)

$dtSupersedence = New-Object System.Data.DataTable
[void]$dtSupersedence.Columns.Add("SupersedingApp", [string])
[void]$dtSupersedence.Columns.Add("SupersedingVersion", [string])
[void]$dtSupersedence.Columns.Add("SupersedingCIID", [string])
[void]$dtSupersedence.Columns.Add("SupersededApp", [string])
[void]$dtSupersedence.Columns.Add("SupersededVersion", [string])
[void]$dtSupersedence.Columns.Add("SupersededCIID", [string])
[void]$dtSupersedence.Columns.Add("ChainDepth", [int])
[void]$dtSupersedence.Columns.Add("Status", [string])
$gridSupersedence.DataSource = $dtSupersedence

$txtSupersedenceDetail = New-Object System.Windows.Forms.RichTextBox
$txtSupersedenceDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtSupersedenceDetail.ReadOnly = $true
$txtSupersedenceDetail.BackColor = $clrDetailBg
$txtSupersedenceDetail.ForeColor = $clrText
$txtSupersedenceDetail.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$txtSupersedenceDetail.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$splitSupersedence.Panel2.Controls.Add($txtSupersedenceDetail)

# Row color coding
$gridSupersedence.Add_RowPrePaint({
    param($s, $e)
    try {
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $dtSupersedence.DefaultView.Count) {
            $rowView = $dtSupersedence.DefaultView[$e.RowIndex]
            $st = [string]$rowView["Status"]
            if ($st -in 'Orphaned', 'Circular')         { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrErrText }
            elseif ($st -in 'Expired Target', 'Disabled Source') { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrWarnText }
            else                                          { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrGridText }
        }
    } catch {}
})

# Selection handler
$gridSupersedence.Add_SelectionChanged({
    if ($gridSupersedence.SelectedRows.Count -eq 0) { $txtSupersedenceDetail.Text = ''; return }
    $rowIdx = $gridSupersedence.SelectedRows[0].Index
    if ($rowIdx -lt 0 -or $rowIdx -ge $dtSupersedence.DefaultView.Count) { return }
    $row = $dtSupersedence.DefaultView[$rowIdx]

    $lines = @(
        "SUPERSEDENCE RELATIONSHIP",
        ("-" * 40),
        "",
        "Superseding: $($row['SupersedingApp']) ($($row['SupersedingVersion']))",
        "  CI_ID:     $($row['SupersedingCIID'])",
        "",
        "Superseded:  $($row['SupersededApp']) ($($row['SupersededVersion']))",
        "  CI_ID:     $($row['SupersededCIID'])",
        "",
        "Chain Depth: $($row['ChainDepth'])",
        "Status:      $($row['Status'])"
    )

    $supCIID = [int]$row['SupersedingCIID']
    $sedCIID = [int]$row['SupersededCIID']
    if ($script:AppLookup.ContainsKey($supCIID)) {
        $app = $script:AppLookup[$supCIID]
        $lines += ""
        $lines += "Superseding App Details:"
        $lines += "  Manufacturer: $($app.Manufacturer)"
        $lines += "  Enabled:      $($app.IsEnabled)"
        $lines += "  Deployments:  $($app.NumberOfDeployments)"
        $lines += "  Created:      $($app.DateCreated)"
        $lines += "  Created By:   $($app.CreatedBy)"
    }
    if ($script:AppLookup.ContainsKey($sedCIID)) {
        $app = $script:AppLookup[$sedCIID]
        $lines += ""
        $lines += "Superseded App Details:"
        $lines += "  Manufacturer: $($app.Manufacturer)"
        $lines += "  Expired:      $($app.IsExpired)"
        $lines += "  Enabled:      $($app.IsEnabled)"
        $lines += "  Deployments:  $($app.NumberOfDeployments)"
    }

    $txtSupersedenceDetail.Text = $lines -join "`r`n"
})

# ===================== TAB 1: Dependencies =====================

$tabDependencies = New-Object System.Windows.Forms.TabPage
$tabDependencies.Text = "Dependencies"
$tabDependencies.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabDependencies)

$splitDependencies = New-Object System.Windows.Forms.SplitContainer
$splitDependencies.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitDependencies.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitDependencies.SplitterDistance = 350
$splitDependencies.SplitterWidth = 6
$splitDependencies.BackColor = $clrSepLine
$splitDependencies.Panel1.BackColor = $clrPanelBg
$splitDependencies.Panel2.BackColor = $clrPanelBg
$splitDependencies.Panel1MinSize = 100
$splitDependencies.Panel2MinSize = 80
$tabDependencies.Controls.Add($splitDependencies)

$gridDependencies = New-ThemedGrid

$colDParent  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDParent.HeaderText = "Parent App";        $colDParent.DataPropertyName = "ParentApp";        $colDParent.Width = 200
$colDPVer    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDPVer.HeaderText = "Version";             $colDPVer.DataPropertyName = "ParentVersion";      $colDPVer.Width = 90
$colDDepApp  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDDepApp.HeaderText = "Dependency App";    $colDDepApp.DataPropertyName = "DependencyApp";    $colDDepApp.Width = 200
$colDDVer    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDDVer.HeaderText = "Version";             $colDDVer.DataPropertyName = "DependencyVersion"; $colDDVer.Width = 90
$colDType    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDType.HeaderText = "Type";                $colDType.DataPropertyName = "DependencyType";     $colDType.Width = 100
$colDLevel   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDLevel.HeaderText = "Level";              $colDLevel.DataPropertyName = "Level";             $colDLevel.Width = 50
$colDStatus  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDStatus.HeaderText = "Status";            $colDStatus.DataPropertyName = "Status";           $colDStatus.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

$gridDependencies.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colDParent, $colDPVer, $colDDepApp, $colDDVer, $colDType, $colDLevel, $colDStatus))
$splitDependencies.Panel1.Controls.Add($gridDependencies)

$dtDependencies = New-Object System.Data.DataTable
[void]$dtDependencies.Columns.Add("ParentApp", [string])
[void]$dtDependencies.Columns.Add("ParentVersion", [string])
[void]$dtDependencies.Columns.Add("ParentCIID", [string])
[void]$dtDependencies.Columns.Add("DependencyApp", [string])
[void]$dtDependencies.Columns.Add("DependencyVersion", [string])
[void]$dtDependencies.Columns.Add("DependencyCIID", [string])
[void]$dtDependencies.Columns.Add("DependencyType", [string])
[void]$dtDependencies.Columns.Add("Level", [int])
[void]$dtDependencies.Columns.Add("Status", [string])
$gridDependencies.DataSource = $dtDependencies

$txtDependencyDetail = New-Object System.Windows.Forms.RichTextBox
$txtDependencyDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtDependencyDetail.ReadOnly = $true
$txtDependencyDetail.BackColor = $clrDetailBg
$txtDependencyDetail.ForeColor = $clrText
$txtDependencyDetail.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$txtDependencyDetail.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$splitDependencies.Panel2.Controls.Add($txtDependencyDetail)

$gridDependencies.Add_RowPrePaint({
    param($s, $e)
    try {
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $dtDependencies.DefaultView.Count) {
            $rowView = $dtDependencies.DefaultView[$e.RowIndex]
            $st = [string]$rowView["Status"]
            if ($st -in 'Orphaned', 'Missing Content')           { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrErrText }
            elseif ($st -in 'Expired Target', 'Disabled Target') { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrWarnText }
            else                                                  { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrGridText }
        }
    } catch {}
})

$gridDependencies.Add_SelectionChanged({
    if ($gridDependencies.SelectedRows.Count -eq 0) { $txtDependencyDetail.Text = ''; return }
    $rowIdx = $gridDependencies.SelectedRows[0].Index
    if ($rowIdx -lt 0 -or $rowIdx -ge $dtDependencies.DefaultView.Count) { return }
    $row = $dtDependencies.DefaultView[$rowIdx]

    $lines = @(
        "DEPENDENCY RELATIONSHIP",
        ("-" * 40),
        "",
        "Parent:     $($row['ParentApp']) ($($row['ParentVersion']))",
        "  CI_ID:    $($row['ParentCIID'])",
        "",
        "Dependency: $($row['DependencyApp']) ($($row['DependencyVersion']))",
        "  CI_ID:    $($row['DependencyCIID'])",
        "",
        "Type:       $($row['DependencyType'])",
        "Level:      $($row['Level'])",
        "Status:     $($row['Status'])"
    )

    $depCIID = [int]$row['DependencyCIID']
    if ($script:AppLookup.ContainsKey($depCIID)) {
        $app = $script:AppLookup[$depCIID]
        $lines += ""
        $lines += "Dependency App Details:"
        $lines += "  Manufacturer: $($app.Manufacturer)"
        $lines += "  Enabled:      $($app.IsEnabled)"
        $lines += "  Expired:      $($app.IsExpired)"
        $lines += "  Has Content:  $($app.HasContent)"
        $lines += "  Deployments:  $($app.NumberOfDeployments)"
    }

    $txtDependencyDetail.Text = $lines -join "`r`n"
})

# ===================== TAB 2: Broken Rules =====================

$tabBroken = New-Object System.Windows.Forms.TabPage
$tabBroken.Text = "Broken Rules"
$tabBroken.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabBroken)

$splitBroken = New-Object System.Windows.Forms.SplitContainer
$splitBroken.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitBroken.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitBroken.SplitterDistance = 300
$splitBroken.SplitterWidth = 6
$splitBroken.BackColor = $clrSepLine
$splitBroken.Panel1.BackColor = $clrPanelBg
$splitBroken.Panel2.BackColor = $clrPanelBg
$splitBroken.Panel1MinSize = 100
$splitBroken.Panel2MinSize = 80
$tabBroken.Controls.Add($splitBroken)

$gridBroken = New-ThemedGrid

$colBIssue   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colBIssue.HeaderText = "Issue Type";   $colBIssue.DataPropertyName = "IssueType";   $colBIssue.Width = 150
$colBSev     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colBSev.HeaderText = "Severity";       $colBSev.DataPropertyName = "Severity";      $colBSev.Width = 80
$colBCat     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colBCat.HeaderText = "Category";       $colBCat.DataPropertyName = "Category";      $colBCat.Width = 110
$colBFrom    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colBFrom.HeaderText = "From App";      $colBFrom.DataPropertyName = "FromApp";      $colBFrom.Width = 200
$colBTo      = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colBTo.HeaderText = "To App";          $colBTo.DataPropertyName = "ToApp";          $colBTo.Width = 200
$colBDesc    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colBDesc.HeaderText = "Description";   $colBDesc.DataPropertyName = "Description";  $colBDesc.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

$gridBroken.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colBIssue, $colBSev, $colBCat, $colBFrom, $colBTo, $colBDesc))
$splitBroken.Panel1.Controls.Add($gridBroken)

$dtBroken = New-Object System.Data.DataTable
[void]$dtBroken.Columns.Add("IssueType", [string])
[void]$dtBroken.Columns.Add("Severity", [string])
[void]$dtBroken.Columns.Add("Category", [string])
[void]$dtBroken.Columns.Add("FromApp", [string])
[void]$dtBroken.Columns.Add("ToApp", [string])
[void]$dtBroken.Columns.Add("Description", [string])
[void]$dtBroken.Columns.Add("Remediation", [string])
$gridBroken.DataSource = $dtBroken

$txtBrokenDetail = New-Object System.Windows.Forms.RichTextBox
$txtBrokenDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtBrokenDetail.ReadOnly = $true
$txtBrokenDetail.BackColor = $clrDetailBg
$txtBrokenDetail.ForeColor = $clrText
$txtBrokenDetail.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$txtBrokenDetail.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$splitBroken.Panel2.Controls.Add($txtBrokenDetail)

$gridBroken.Add_RowPrePaint({
    param($s, $e)
    try {
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $dtBroken.DefaultView.Count) {
            $rowView = $dtBroken.DefaultView[$e.RowIndex]
            $sev = [string]$rowView["Severity"]
            if ($sev -eq 'Error')   { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrErrText }
            elseif ($sev -eq 'Warning') { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrWarnText }
            elseif ($sev -eq 'Info')    { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrInfoText }
            else { $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrGridText }
        }
    } catch {}
})

$gridBroken.Add_SelectionChanged({
    if ($gridBroken.SelectedRows.Count -eq 0) { $txtBrokenDetail.Text = ''; return }
    $rowIdx = $gridBroken.SelectedRows[0].Index
    if ($rowIdx -lt 0 -or $rowIdx -ge $dtBroken.DefaultView.Count) { return }
    $row = $dtBroken.DefaultView[$rowIdx]

    $lines = @(
        "BROKEN RULE DETAILS",
        ("-" * 40),
        "",
        "Issue:       $($row['IssueType'])",
        "Severity:    $($row['Severity'])",
        "Category:    $($row['Category'])",
        "From App:    $($row['FromApp'])",
        "To App:      $($row['ToApp'])",
        "",
        "Description:",
        $row['Description'],
        "",
        "Remediation:",
        $row['Remediation']
    )

    $txtBrokenDetail.Text = $lines -join "`r`n"
})

# ===================== TAB 3: Tree View =====================

$tabTreeView = New-Object System.Windows.Forms.TabPage
$tabTreeView.Text = "Tree View"
$tabTreeView.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabTreeView)

$splitTree = New-Object System.Windows.Forms.SplitContainer
$splitTree.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitTree.Orientation = [System.Windows.Forms.Orientation]::Vertical
$splitTree.SplitterWidth = 6
$splitTree.BackColor = $clrSepLine
$splitTree.Panel1.BackColor = $clrPanelBg
$splitTree.Panel2.BackColor = $clrPanelBg
$splitTree.Panel1MinSize = 100
$splitTree.Panel2MinSize = 100
$tabTreeView.Controls.Add($splitTree)
$splitTree.SplitterDistance = [Math]::Max($splitTree.Panel1MinSize, [int]($splitTree.Width * 0.5))

$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Dock = [System.Windows.Forms.DockStyle]::Fill
$treeView.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$treeView.BackColor = $clrTreeBg
$treeView.ForeColor = $clrText
$treeView.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$treeView.FullRowSelect = $true
$treeView.HideSelection = $false
$treeView.ShowLines = $true
$treeView.ShowPlusMinus = $true
$treeView.ShowRootLines = $true
if ($script:Prefs.DarkMode) {
    $treeView.LineColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
}
Enable-DoubleBuffer -Control $treeView
$splitTree.Panel1.Controls.Add($treeView)

$txtTreeDetail = New-Object System.Windows.Forms.RichTextBox
$txtTreeDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtTreeDetail.ReadOnly = $true
$txtTreeDetail.BackColor = $clrDetailBg
$txtTreeDetail.ForeColor = $clrText
$txtTreeDetail.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$txtTreeDetail.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$splitTree.Panel2.Controls.Add($txtTreeDetail)

$treeView.Add_AfterSelect({
    param($s, $e)
    $node = $e.Node
    if (-not $node -or -not $node.Tag) { $txtTreeDetail.Text = ''; return }

    $ciid = [int]$node.Tag
    if (-not $script:AppLookup.ContainsKey($ciid)) {
        $txtTreeDetail.Text = "Application CI_ID: $ciid (not found in current scan)"
        return
    }

    $app = $script:AppLookup[$ciid]
    $lines = @(
        "APPLICATION DETAILS",
        ("-" * 40),
        "",
        "Name:          $($app.LocalizedDisplayName)",
        "Version:       $($app.SoftwareVersion)",
        "Manufacturer:  $($app.Manufacturer)",
        "CI_ID:         $($app.CI_ID)",
        "",
        "Enabled:       $($app.IsEnabled)",
        "Expired:       $($app.IsExpired)",
        "Is Superseded: $($app.IsSuperseded)",
        "Is Superseding: $($app.IsSuperseding)",
        "Has Content:   $($app.HasContent)",
        "",
        "Deploy Types:  $($app.NumberOfDeploymentTypes)",
        "Deployments:   $($app.NumberOfDeployments)",
        "",
        "Created:       $($app.DateCreated)",
        "Created By:    $($app.CreatedBy)",
        "Modified:      $($app.DateLastModified)",
        "Modified By:   $($app.LastModifiedBy)"
    )

    $txtTreeDetail.Text = $lines -join "`r`n"
})

# ---------------------------------------------------------------------------
# Finalize dock Z-order
# ---------------------------------------------------------------------------

$form.Controls.Add($menuStrip)
$menuStrip.SendToBack()

$pnlSep3.BringToFront()
$pnlFilter.BringToFront()
$pnlSep2.BringToFront()
$pnlCards.BringToFront()
$pnlSep1.BringToFront()
$pnlConnBar.BringToFront()
$pnlHeader.BringToFront()

$tabMain.BringToFront()

# ---------------------------------------------------------------------------
# Module-scoped data (populated by Scan)
# ---------------------------------------------------------------------------

$script:AppLookup          = @{}
$script:DTLookup           = @{}
$script:SupersedenceData   = @()
$script:DependencyData     = @()
$script:BrokenData         = @()
$script:ScanCounts         = $null
$script:LastScanTime       = $null

# ---------------------------------------------------------------------------
# Status bar update helper
# ---------------------------------------------------------------------------

function Update-StatusBar {
    $parts = @()

    if (Test-CMConnection) {
        $parts += "Connected to $($script:Prefs.SiteCode)"
    } else {
        $parts += "Disconnected"
    }

    if ($script:LastScanTime) {
        $parts += "Last scan: $($script:LastScanTime.ToString('HH:mm:ss'))"
    }

    $statusLabel.Text = $parts -join " | "

    # Row count for active tab
    $tabIdx = $tabMain.SelectedIndex
    $count = switch ($tabIdx) {
        0 { $dtSupersedence.DefaultView.Count }
        1 { $dtDependencies.DefaultView.Count }
        2 { $dtBroken.DefaultView.Count }
        3 { $treeView.GetNodeCount($true) }
        default { 0 }
    }
    $statusRowCount.Text = "$count rows"
}

# ---------------------------------------------------------------------------
# Filter logic
# ---------------------------------------------------------------------------

function Invoke-ApplyFilter {
    $filterText = $txtFilter.Text.Trim()
    $statusFilter = $cboStatus.SelectedItem

    # Build RowFilter for DataTables (tabs 0-2)
    $tabIdx = $tabMain.SelectedIndex

    switch ($tabIdx) {
        0 { # Supersedence
            $parts = @()
            if ($filterText) {
                $escaped = $filterText.Replace("'", "''")
                $parts += "(SupersedingApp LIKE '%$escaped%' OR SupersededApp LIKE '%$escaped%')"
            }
            if ($statusFilter -and $statusFilter -ne 'All') {
                switch ($statusFilter) {
                    'Healthy'        { $parts += "Status = 'Healthy'" }
                    'Broken/Warning' { $parts += "(Status <> 'Healthy')" }
                    'Error'          { $parts += "(Status = 'Orphaned' OR Status = 'Circular')" }
                }
            }
            $dtSupersedence.DefaultView.RowFilter = if ($parts.Count -gt 0) { $parts -join ' AND ' } else { '' }
        }
        1 { # Dependencies
            $parts = @()
            if ($filterText) {
                $escaped = $filterText.Replace("'", "''")
                $parts += "(ParentApp LIKE '%$escaped%' OR DependencyApp LIKE '%$escaped%')"
            }
            if ($statusFilter -and $statusFilter -ne 'All') {
                switch ($statusFilter) {
                    'Healthy'        { $parts += "Status = 'Healthy'" }
                    'Broken/Warning' { $parts += "(Status <> 'Healthy')" }
                    'Error'          { $parts += "(Status = 'Orphaned' OR Status = 'Missing Content')" }
                }
            }
            $dtDependencies.DefaultView.RowFilter = if ($parts.Count -gt 0) { $parts -join ' AND ' } else { '' }
        }
        2 { # Broken Rules
            $parts = @()
            if ($filterText) {
                $escaped = $filterText.Replace("'", "''")
                $parts += "(FromApp LIKE '%$escaped%' OR ToApp LIKE '%$escaped%' OR Description LIKE '%$escaped%')"
            }
            if ($statusFilter -and $statusFilter -ne 'All') {
                switch ($statusFilter) {
                    'Healthy'        { $parts += "1=0" }  # No healthy rows in broken tab
                    'Broken/Warning' { $parts += "Severity = 'Warning'" }
                    'Error'          { $parts += "Severity = 'Error'" }
                }
            }
            $dtBroken.DefaultView.RowFilter = if ($parts.Count -gt 0) { $parts -join ' AND ' } else { '' }
        }
    }

    Update-StatusBar
}

$txtFilter.Add_TextChanged({ Invoke-ApplyFilter })
$cboStatus.Add_SelectedIndexChanged({ Invoke-ApplyFilter })
$tabMain.Add_SelectedIndexChanged({ Invoke-ApplyFilter; Update-StatusBar })

# ---------------------------------------------------------------------------
# Tree population helper
# ---------------------------------------------------------------------------

function Set-TreeViewData {
    param(
        [object[]]$SupersedenceRoots,
        [object[]]$DependencyRoots
    )

    $treeView.BeginUpdate()
    $treeView.Nodes.Clear()

    # Supersedence Chains
    $nodeSuper = New-Object System.Windows.Forms.TreeNode("Supersedence Chains ($(@($SupersedenceRoots).Count))")
    $nodeSuper.ForeColor = $clrText

    function Add-SupersedenceNodes {
        param([System.Windows.Forms.TreeNode]$ParentNode, [object[]]$Children)
        foreach ($child in $Children) {
            $suffix = ''
            if ($child.Status -eq 'Expired' -or $child.Status -eq 'Expired Target') { $suffix = ' (EXPIRED)' }
            elseif ($child.Status -eq 'Disabled' -or $child.Status -eq 'Disabled Source') { $suffix = ' (DISABLED)' }

            $text = "$($child.Name) ($($child.Version))$suffix"
            $childNode = New-Object System.Windows.Forms.TreeNode($text)
            $childNode.Tag = $child.CIID

            if ($suffix -match 'EXPIRED')  { $childNode.ForeColor = $clrErrText }
            elseif ($suffix -match 'DISABLED') { $childNode.ForeColor = $clrWarnText }
            else { $childNode.ForeColor = $clrText }

            $ParentNode.Nodes.Add($childNode) | Out-Null

            if ($child.Children -and $child.Children.Count -gt 0) {
                Add-SupersedenceNodes -ParentNode $childNode -Children $child.Children
            }
        }
    }

    foreach ($root in $SupersedenceRoots) {
        $suffix = ''
        if ($root.Status -eq 'Expired') { $suffix = ' (EXPIRED)' }
        elseif ($root.Status -eq 'Disabled') { $suffix = ' (DISABLED)' }

        $text = "$($root.Name) ($($root.Version))$suffix"
        $rootNode = New-Object System.Windows.Forms.TreeNode($text)
        $rootNode.Tag = $root.CIID
        $rootNode.ForeColor = if ($suffix -match 'EXPIRED') { $clrErrText } elseif ($suffix -match 'DISABLED') { $clrWarnText } else { $clrText }
        $rootNode.NodeFont = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

        $nodeSuper.Nodes.Add($rootNode) | Out-Null

        if ($root.Children -and $root.Children.Count -gt 0) {
            Add-SupersedenceNodes -ParentNode $rootNode -Children $root.Children
        }
    }

    $treeView.Nodes.Add($nodeSuper) | Out-Null

    # Dependency Trees
    $nodeDeps = New-Object System.Windows.Forms.TreeNode("Dependency Trees ($(@($DependencyRoots).Count))")
    $nodeDeps.ForeColor = $clrText

    foreach ($root in $DependencyRoots) {
        $suffix = ''
        if ($root.Status -eq 'Expired') { $suffix = ' (EXPIRED)' }
        elseif ($root.Status -eq 'Disabled') { $suffix = ' (DISABLED)' }

        $text = "$($root.Name) ($($root.Version))$suffix"
        $rootNode = New-Object System.Windows.Forms.TreeNode($text)
        $rootNode.Tag = $root.CIID
        $rootNode.ForeColor = if ($suffix -match 'EXPIRED') { $clrErrText } elseif ($suffix -match 'DISABLED') { $clrWarnText } else { $clrText }
        $rootNode.NodeFont = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

        $nodeDeps.Nodes.Add($rootNode) | Out-Null

        if ($root.Children -and $root.Children.Count -gt 0) {
            foreach ($child in $root.Children) {
                $cSuffix = ''
                if ($child.Status -in 'Expired Target', 'Expired') { $cSuffix = ' (EXPIRED)' }
                elseif ($child.Status -in 'Disabled Target', 'Disabled') { $cSuffix = ' (DISABLED)' }
                elseif ($child.Status -eq 'Missing Content') { $cSuffix = ' (NO CONTENT)' }

                $cText = "$($child.Name) ($($child.Version)) [$($child.Type)]$cSuffix"
                $childNode = New-Object System.Windows.Forms.TreeNode($cText)
                $childNode.Tag = $child.CIID

                if ($cSuffix -match 'EXPIRED|NO CONTENT') { $childNode.ForeColor = $clrErrText }
                elseif ($cSuffix -match 'DISABLED') { $childNode.ForeColor = $clrWarnText }
                else { $childNode.ForeColor = $clrText }

                $rootNode.Nodes.Add($childNode) | Out-Null
            }
        }
    }

    $treeView.Nodes.Add($nodeDeps) | Out-Null

    # Expand root nodes
    $nodeSuper.Expand()
    $nodeDeps.Expand()

    $treeView.EndUpdate()
}

# ---------------------------------------------------------------------------
# Core operation: Scan Environment
# ---------------------------------------------------------------------------

function Invoke-ScanEnvironment {
    if (-not $script:Prefs.SiteCode -or -not $script:Prefs.SMSProvider) {
        [System.Windows.Forms.MessageBox]::Show(
            "Site Code and SMS Provider must be configured in File > Preferences.",
            "Configuration Required", "OK", "Warning") | Out-Null
        return
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnScan.Enabled = $false

    try {
        # Connect if needed
        if (-not (Test-CMConnection)) {
            Add-LogLine -TextBox $txtLog -Message "Connecting to $($script:Prefs.SiteCode)..."
            [System.Windows.Forms.Application]::DoEvents()

            $connected = Connect-CMSite -SiteCode $script:Prefs.SiteCode -SMSProvider $script:Prefs.SMSProvider
            if (-not $connected) {
                Add-LogLine -TextBox $txtLog -Message "ERROR: Failed to connect to CM site"
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to connect to site $($script:Prefs.SiteCode). Check log for details.",
                    "Connection Failed", "OK", "Error") | Out-Null
                return
            }
            Add-LogLine -TextBox $txtLog -Message "Connected to $($script:Prefs.SiteCode)"
        }

        # Clear existing data
        $dtSupersedence.Clear()
        $dtDependencies.Clear()
        $dtBroken.Clear()
        $treeView.Nodes.Clear()
        [System.Windows.Forms.Application]::DoEvents()

        # WMI Query 1: Applications
        Add-LogLine -TextBox $txtLog -Message "Querying all applications..."
        [System.Windows.Forms.Application]::DoEvents()
        $script:AppLookup = Get-AllApplicationSummary -SMSProvider $script:Prefs.SMSProvider -SiteCode $script:Prefs.SiteCode
        Add-LogLine -TextBox $txtLog -Message "Loaded $($script:AppLookup.Count) applications"
        [System.Windows.Forms.Application]::DoEvents()

        # WMI Query 2: Deployment Types
        Add-LogLine -TextBox $txtLog -Message "Querying deployment types..."
        [System.Windows.Forms.Application]::DoEvents()
        $script:DTLookup = Get-AllDeploymentTypeSummary -SMSProvider $script:Prefs.SMSProvider -SiteCode $script:Prefs.SiteCode
        Add-LogLine -TextBox $txtLog -Message "Loaded $($script:DTLookup.Count) deployment types"
        [System.Windows.Forms.Application]::DoEvents()

        # WMI Query 3: Relationships
        Add-LogLine -TextBox $txtLog -Message "Querying all relationships..."
        [System.Windows.Forms.Application]::DoEvents()
        $rawRelationships = Get-AllRelationships -SMSProvider $script:Prefs.SMSProvider -SiteCode $script:Prefs.SiteCode
        Add-LogLine -TextBox $txtLog -Message "Loaded $(@($rawRelationships).Count) raw relationship records"
        [System.Windows.Forms.Application]::DoEvents()

        # Resolve relationships
        Add-LogLine -TextBox $txtLog -Message "Resolving relationships..."
        [System.Windows.Forms.Application]::DoEvents()
        $resolved = Resolve-RelationshipData -RawRelationships $rawRelationships -AppLookup $script:AppLookup -DTLookup $script:DTLookup
        if (-not $resolved) { $resolved = @() }
        Add-LogLine -TextBox $txtLog -Message "Resolved $(@($resolved).Count) relevant relationships"
        [System.Windows.Forms.Application]::DoEvents()

        # Find supersedence chains
        Add-LogLine -TextBox $txtLog -Message "Analyzing supersedence chains..."
        [System.Windows.Forms.Application]::DoEvents()
        $script:SupersedenceData = @(Find-SupersedenceChains -ResolvedRelationships $resolved -AppLookup $script:AppLookup)
        Add-LogLine -TextBox $txtLog -Message "Found $($script:SupersedenceData.Count) supersedence relationships"
        [System.Windows.Forms.Application]::DoEvents()

        # Find dependency groups
        Add-LogLine -TextBox $txtLog -Message "Analyzing dependency groups..."
        [System.Windows.Forms.Application]::DoEvents()
        $script:DependencyData = @(Find-DependencyGroups -ResolvedRelationships $resolved -AppLookup $script:AppLookup)
        Add-LogLine -TextBox $txtLog -Message "Found $($script:DependencyData.Count) dependency relationships"
        [System.Windows.Forms.Application]::DoEvents()

        # Find broken rules
        Add-LogLine -TextBox $txtLog -Message "Detecting broken rules..."
        [System.Windows.Forms.Application]::DoEvents()
        $brokenSup = @(Find-BrokenSupersedence -SupersedenceData $script:SupersedenceData -ResolvedRelationships $resolved -AppLookup $script:AppLookup)
        $brokenDep = @(Find-BrokenDependencies -DependencyData $script:DependencyData -ResolvedRelationships $resolved -AppLookup $script:AppLookup)
        $undocumented = @(Find-UndocumentedRelationships -SupersedenceData $script:SupersedenceData -DependencyData $script:DependencyData -AppLookup $script:AppLookup)
        $script:BrokenData = @($brokenSup) + @($brokenDep) + @($undocumented)
        Add-LogLine -TextBox $txtLog -Message "Found $($script:BrokenData.Count) broken/info rules"
        [System.Windows.Forms.Application]::DoEvents()

        # Populate Supersedence DataTable
        $dtSupersedence.BeginLoadData()
        foreach ($rel in $script:SupersedenceData) {
            [void]$dtSupersedence.Rows.Add(
                $rel.SupersedingApp, $rel.SupersedingVersion, [string]$rel.SupersedingCIID,
                $rel.SupersededApp, $rel.SupersededVersion, [string]$rel.SupersededCIID,
                $rel.ChainDepth, $rel.Status
            )
        }
        $dtSupersedence.EndLoadData()

        # Populate Dependencies DataTable
        $dtDependencies.BeginLoadData()
        foreach ($rel in $script:DependencyData) {
            [void]$dtDependencies.Rows.Add(
                $rel.ParentApp, $rel.ParentVersion, [string]$rel.ParentCIID,
                $rel.DependencyApp, $rel.DependencyVersion, [string]$rel.DependencyCIID,
                $rel.DependencyType, $rel.Level, $rel.Status
            )
        }
        $dtDependencies.EndLoadData()

        # Populate Broken Rules DataTable
        $dtBroken.BeginLoadData()
        foreach ($rule in $script:BrokenData) {
            [void]$dtBroken.Rows.Add(
                $rule.IssueType, $rule.Severity, $rule.Category,
                $rule.FromApp, $rule.ToApp, $rule.Description, $rule.Remediation
            )
        }
        $dtBroken.EndLoadData()
        [System.Windows.Forms.Application]::DoEvents()

        # Build and populate trees
        Add-LogLine -TextBox $txtLog -Message "Building tree view..."
        [System.Windows.Forms.Application]::DoEvents()
        $supRoots = Build-SupersedenceTree -SupersedenceData $script:SupersedenceData -AppLookup $script:AppLookup
        $depRoots = Build-DependencyTree -DependencyData $script:DependencyData -AppLookup $script:AppLookup
        Set-TreeViewData -SupersedenceRoots $supRoots -DependencyRoots $depRoots

        # Update summary cards
        $script:ScanCounts = Get-ScanSummaryCounts -AppCount $script:AppLookup.Count `
            -SupersedenceData $script:SupersedenceData `
            -DependencyData $script:DependencyData `
            -BrokenRules $script:BrokenData

        Update-Card -Card $cardApps -ValueText "$($script:ScanCounts.AppCount) apps" -Severity 'info'
        Update-Card -Card $cardSupersede -ValueText "$($script:ScanCounts.SupersedenceTotal) rules" -Severity $(if ($script:ScanCounts.SupersedenceBroken -gt 0) { 'warn' } else { 'ok' })
        Update-Card -Card $cardDependency -ValueText "$($script:ScanCounts.DependencyTotal) rules" -Severity $(if ($script:ScanCounts.DependencyBroken -gt 0) { 'warn' } else { 'ok' })
        Update-Card -Card $cardBroken -ValueText "$($script:ScanCounts.BrokenRulesTotal) issues" -Severity $(if ($script:ScanCounts.BrokenErrors -gt 0) { 'critical' } elseif ($script:ScanCounts.BrokenWarnings -gt 0) { 'warn' } else { 'ok' })

        $script:LastScanTime = Get-Date
        Add-LogLine -TextBox $txtLog -Message "Scan complete. $($script:ScanCounts.AppCount) apps, $($script:ScanCounts.SupersedenceTotal) supersedence, $($script:ScanCounts.DependencyTotal) dependencies, $($script:ScanCounts.BrokenRulesTotal) broken rules."
    }
    catch {
        Add-LogLine -TextBox $txtLog -Message "ERROR: $($_.Exception.Message)"
        Write-Log "Scan failed: $_" -Level ERROR
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnScan.Enabled = $true
        Update-StatusBar
    }
}

$btnScan.Add_Click({ Invoke-ScanEnvironment })

# ---------------------------------------------------------------------------
# Export handlers
# ---------------------------------------------------------------------------

function Get-ActiveDataTable {
    $tabIdx = $tabMain.SelectedIndex
    switch ($tabIdx) {
        0 { return @{ DataTable = $dtSupersedence; Name = "Supersedence" } }
        1 { return @{ DataTable = $dtDependencies; Name = "Dependencies" } }
        2 { return @{ DataTable = $dtBroken; Name = "BrokenRules" } }
        3 { return $null }  # Tree view - no DataTable
        default { return $null }
    }
}

$btnExportCsv.Add_Click({
    $info = Get-ActiveDataTable
    if (-not $info) {
        [System.Windows.Forms.MessageBox]::Show("CSV export is not available for the Tree View tab.", "Export", "OK", "Information") | Out-Null
        return
    }
    if ($info.DataTable.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export. Run a scan first.", "Export", "OK", "Information") | Out-Null
        return
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV files (*.csv)|*.csv"
    $sfd.FileName = "Audit-$($info.Name)-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $sfd.InitialDirectory = Join-Path $PSScriptRoot "Reports"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-AuditCsv -DataTable $info.DataTable -OutputPath $sfd.FileName
        Add-LogLine -TextBox $txtLog -Message "Exported CSV: $($sfd.FileName)"
    }
    $sfd.Dispose()
})

$btnExportHtml.Add_Click({
    $info = Get-ActiveDataTable
    if (-not $info) {
        [System.Windows.Forms.MessageBox]::Show("HTML export is not available for the Tree View tab.", "Export", "OK", "Information") | Out-Null
        return
    }
    if ($info.DataTable.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export. Run a scan first.", "Export", "OK", "Information") | Out-Null
        return
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "HTML files (*.html)|*.html"
    $sfd.FileName = "Audit-$($info.Name)-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    $sfd.InitialDirectory = Join-Path $PSScriptRoot "Reports"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-AuditHtml -DataTable $info.DataTable -OutputPath $sfd.FileName -ReportTitle "Audit Report: $($info.Name)"
        Add-LogLine -TextBox $txtLog -Message "Exported HTML: $($sfd.FileName)"
    }
    $sfd.Dispose()
})

$btnCopySummary.Add_Click({
    if (-not $script:ScanCounts) {
        [System.Windows.Forms.MessageBox]::Show("No scan data available. Run a scan first.", "Copy Summary", "OK", "Information") | Out-Null
        return
    }

    $summary = New-AuditSummaryText -Counts $script:ScanCounts
    [System.Windows.Forms.Clipboard]::SetText($summary)
    Add-LogLine -TextBox $txtLog -Message "Summary copied to clipboard"
})

# ---------------------------------------------------------------------------
# Form events
# ---------------------------------------------------------------------------

$form.Add_FormClosing({
    Save-WindowState
    if (Test-CMConnection) { Disconnect-CMSite }
})

$form.Add_Shown({
    Restore-WindowState
    Update-StatusBar
    Add-LogLine -TextBox $txtLog -Message "Supersedence & Dependency Auditor ready. Configure Site/Provider in Preferences, then click Scan."
})

# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------

[System.Windows.Forms.Application]::Run($form)
