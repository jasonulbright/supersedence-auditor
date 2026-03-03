<#
.SYNOPSIS
    Core module for MECM Supersedence & Dependency Auditor.

.DESCRIPTION
    Import this module to get:
      - Structured logging (Initialize-Logging, Write-Log)
      - CM site connection management (Connect-CMSite, Disconnect-CMSite, Test-CMConnection)
      - WMI bulk discovery of applications, deployment types, and relationships
      - Relationship resolution and enrichment (CI_ID -> friendly names)
      - Supersedence chain and dependency group analysis
      - Broken rule detection (orphaned, circular, expired, disabled, missing content)
      - Tree building for hierarchical visualization
      - Export to CSV, HTML, and clipboard summary

.EXAMPLE
    Import-Module "$PSScriptRoot\Module\SupersedenceAuditorCommon.psd1" -Force
    Initialize-Logging -LogPath "C:\temp\audit.log"
    Connect-CMSite -SiteCode 'MCM' -SMSProvider 'sccm01.contoso.com'
#>

# ---------------------------------------------------------------------------
# Module-scoped state
# ---------------------------------------------------------------------------

$script:__SALogPath            = $null
$script:OriginalLocation       = $null
$script:ConnectedSiteCode      = $null
$script:ConnectedSMSProvider   = $null

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Initialize-Logging {
    param([string]$LogPath)

    $script:__SALogPath = $LogPath

    if ($LogPath) {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        $header = "[{0}] [INFO ] === Log initialized ===" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Set-Content -LiteralPath $LogPath -Value $header -Encoding UTF8
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, severity-tagged log message.
    #>
    param(
        [AllowEmptyString()]
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',

        [switch]$Quiet
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formatted = "[{0}] [{1,-5}] {2}" -f $timestamp, $Level, $Message

    if (-not $Quiet) {
        Write-Host $formatted
        if ($Level -eq 'ERROR') {
            $host.UI.WriteErrorLine($formatted)
        }
    }

    if ($script:__SALogPath) {
        Add-Content -LiteralPath $script:__SALogPath -Value $formatted -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# CM Connection
# ---------------------------------------------------------------------------

function Connect-CMSite {
    <#
    .SYNOPSIS
        Imports the ConfigurationManager module, creates a PSDrive, and sets location.
    .DESCRIPTION
        Returns $true on success, $false on failure.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$SMSProvider
    )

    $script:OriginalLocation = Get-Location

    # Import CM module if not already loaded
    if (-not (Get-Module ConfigurationManager -ErrorAction SilentlyContinue)) {
        $cmModulePath = $null
        if ($env:SMS_ADMIN_UI_PATH) {
            $cmModulePath = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
        }

        if (-not $cmModulePath -or -not (Test-Path -LiteralPath $cmModulePath)) {
            Write-Log "ConfigurationManager module not found. Ensure the CM console is installed." -Level ERROR
            return $false
        }

        try {
            Import-Module $cmModulePath -ErrorAction Stop
            Write-Log "Imported ConfigurationManager module"
        }
        catch {
            Write-Log "Failed to import ConfigurationManager module: $_" -Level ERROR
            return $false
        }
    }

    # Create PSDrive if needed
    if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        try {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SMSProvider -ErrorAction Stop | Out-Null
            Write-Log "Created PSDrive for site $SiteCode"
        }
        catch {
            Write-Log "Failed to create PSDrive for site $SiteCode : $_" -Level ERROR
            return $false
        }
    }

    try {
        Set-Location "${SiteCode}:" -ErrorAction Stop
        $site = Get-CMSite -SiteCode $SiteCode -ErrorAction Stop
        Write-Log "Connected to site $SiteCode ($($site.SiteName))"
        $script:ConnectedSiteCode    = $SiteCode
        $script:ConnectedSMSProvider = $SMSProvider
        return $true
    }
    catch {
        Write-Log "Failed to connect to site $SiteCode : $_" -Level ERROR
        return $false
    }
}

function Disconnect-CMSite {
    <#
    .SYNOPSIS
        Restores the original location before CM connection.
    #>
    if ($script:OriginalLocation) {
        try { Set-Location $script:OriginalLocation -ErrorAction SilentlyContinue } catch { }
    }
    $script:ConnectedSiteCode    = $null
    $script:ConnectedSMSProvider = $null
    Write-Log "Disconnected from CM site"
}

function Test-CMConnection {
    <#
    .SYNOPSIS
        Returns $true if currently connected to a CM site.
    #>
    if (-not $script:ConnectedSiteCode) { return $false }

    try {
        $drive = Get-PSDrive -Name $script:ConnectedSiteCode -PSProvider CMSite -ErrorAction Stop
        return ($null -ne $drive)
    }
    catch {
        return $false
    }
}

# ---------------------------------------------------------------------------
# WMI Bulk Discovery
# ---------------------------------------------------------------------------

function Get-AllApplicationSummary {
    <#
    .SYNOPSIS
        Bulk WMI query for all latest-revision applications.
    .DESCRIPTION
        Returns a hashtable keyed by CI_ID for O(1) lookup during relationship resolution.
    #>
    param(
        [Parameter(Mandatory)][string]$SMSProvider,
        [Parameter(Mandatory)][string]$SiteCode
    )

    Write-Log "Querying all applications (SMS_Application WHERE IsLatest = TRUE)..."

    $raw = Get-CimInstance -ComputerName $SMSProvider `
        -Namespace "root\SMS\site_$SiteCode" `
        -ClassName SMS_Application `
        -Filter "IsLatest = TRUE" `
        -ErrorAction Stop

    $lookup = @{}
    foreach ($app in $raw) {
        $lookup[$app.CI_ID] = [PSCustomObject]@{
            CI_ID                  = [uint32]$app.CI_ID
            LocalizedDisplayName   = [string]$app.LocalizedDisplayName
            SoftwareVersion        = [string]$app.SoftwareVersion
            Manufacturer           = [string]$app.Manufacturer
            IsEnabled              = [bool]$app.IsEnabled
            IsExpired              = [bool]$app.IsExpired
            IsSuperseded           = [bool]$app.IsSuperseded
            IsSuperseding          = [bool]$app.IsSuperseding
            HasContent             = [bool]$app.HasContent
            NumberOfDeploymentTypes = [uint32]$app.NumberOfDeploymentTypes
            NumberOfDeployments    = [uint32]$app.NumberOfDeployments
            DateCreated            = $app.DateCreated
            DateLastModified       = $app.DateLastModified
            CreatedBy              = [string]$app.CreatedBy
            LastModifiedBy         = [string]$app.LastModifiedBy
        }
    }

    Write-Log "Loaded $($lookup.Count) applications into lookup"
    return $lookup
}

function Get-AllDeploymentTypeSummary {
    <#
    .SYNOPSIS
        Bulk WMI query for all latest-revision deployment types.
    .DESCRIPTION
        Returns a hashtable keyed by CI_ID for O(1) lookup during relationship resolution.
    #>
    param(
        [Parameter(Mandatory)][string]$SMSProvider,
        [Parameter(Mandatory)][string]$SiteCode
    )

    Write-Log "Querying all deployment types..."

    $raw = Get-CimInstance -ComputerName $SMSProvider `
        -Namespace "root\SMS\site_$SiteCode" `
        -Query "SELECT CI_ID, LocalizedDisplayName, ModelName FROM SMS_ConfigurationItemLatestBaseClass WHERE CIType_ID = 21 AND IsLatest = TRUE" `
        -ErrorAction Stop

    $lookup = @{}
    foreach ($dt in $raw) {
        $lookup[$dt.CI_ID] = [PSCustomObject]@{
            CI_ID                = [uint32]$dt.CI_ID
            LocalizedDisplayName = [string]$dt.LocalizedDisplayName
            ModelName            = [string]$dt.ModelName
        }
    }

    Write-Log "Loaded $($lookup.Count) deployment types into lookup"
    return $lookup
}

function Get-AllRelationships {
    <#
    .SYNOPSIS
        Bulk WMI query for all application relationship records.
    .DESCRIPTION
        Returns raw SMS_AppRelation_Flat records. Single query captures all
        supersedence and dependency relationships in the environment.
    #>
    param(
        [Parameter(Mandatory)][string]$SMSProvider,
        [Parameter(Mandatory)][string]$SiteCode
    )

    Write-Log "Querying all relationships (SMS_AppRelation_Flat)..."

    $raw = Get-CimInstance -ComputerName $SMSProvider `
        -Namespace "root\SMS\site_$SiteCode" `
        -ClassName SMS_AppRelation_Flat `
        -ErrorAction Stop

    Write-Log "Loaded $(@($raw).Count) raw relationship records"
    return $raw
}

function Resolve-RelationshipData {
    <#
    .SYNOPSIS
        Joins raw SMS_AppRelation_Flat records with app/DT lookups for friendly names.
    .DESCRIPTION
        Filters to relevant RelationTypes (2, 4, 6, 10, 15) and enriches each
        record with application and deployment type names from the lookup hashtables.
    #>
    param(
        [Parameter(Mandatory)][object[]]$RawRelationships,
        [Parameter(Mandatory)][hashtable]$AppLookup,
        [Parameter(Mandatory)][hashtable]$DTLookup
    )

    $relationTypeNames = @{
        2  = 'Required'
        4  = 'Optional'
        6  = 'Superseded'
        10 = 'AppDependence'
        15 = 'ApplicationSuperSeded'
    }

    $relevantTypes = @(2, 4, 6, 10, 15)

    $results = foreach ($rel in $RawRelationships) {
        $relType = [int]$rel.RelationType
        if ($relType -notin $relevantTypes) { continue }

        $fromAppCIID = [uint32]$rel.FromApplicationCIID
        $toAppCIID   = [uint32]$rel.ToApplicationCIID
        $fromDTCIID  = [uint32]$rel.FromDeploymentTypeCIID
        $toDTCIID    = [uint32]$rel.ToDeploymentTypeCIID

        $fromApp = if ($AppLookup.ContainsKey($fromAppCIID)) { $AppLookup[$fromAppCIID] } else { $null }
        $toApp   = if ($AppLookup.ContainsKey($toAppCIID))   { $AppLookup[$toAppCIID] }   else { $null }
        $fromDT  = if ($DTLookup.ContainsKey($fromDTCIID))   { $DTLookup[$fromDTCIID] }   else { $null }
        $toDT    = if ($DTLookup.ContainsKey($toDTCIID))     { $DTLookup[$toDTCIID] }     else { $null }

        [PSCustomObject]@{
            FromAppCIID      = $fromAppCIID
            FromAppName      = if ($fromApp) { $fromApp.LocalizedDisplayName } else { "Unknown (CI_ID: $fromAppCIID)" }
            FromAppVersion   = if ($fromApp) { $fromApp.SoftwareVersion } else { '' }
            FromAppExists    = ($null -ne $fromApp)
            FromDTCIID       = $fromDTCIID
            FromDTName       = if ($fromDT) { $fromDT.LocalizedDisplayName } else { '' }
            ToAppCIID        = $toAppCIID
            ToAppName        = if ($toApp) { $toApp.LocalizedDisplayName } else { "Unknown (CI_ID: $toAppCIID)" }
            ToAppVersion     = if ($toApp) { $toApp.SoftwareVersion } else { '' }
            ToAppExists      = ($null -ne $toApp)
            ToDTCIID         = $toDTCIID
            ToDTName         = if ($toDT) { $toDT.LocalizedDisplayName } else { '' }
            RelationType     = $relType
            RelationTypeName = if ($relationTypeNames.ContainsKey($relType)) { $relationTypeNames[$relType] } else { "Type $relType" }
            Level            = [uint32]$rel.Level
        }
    }

    Write-Log "Resolved $(@($results).Count) relevant relationships"
    return $results
}

# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

function Find-SupersedenceChains {
    <#
    .SYNOPSIS
        Extracts supersedence relationships and computes chain depth.
    .DESCRIPTION
        Uses RelationType 6 (Superseded) and 15 (ApplicationSuperSeded).
        Returns enriched objects with chain depth and status.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ResolvedRelationships,
        [Parameter(Mandatory)][hashtable]$AppLookup
    )

    $supersedence = @($ResolvedRelationships | Where-Object { $_.RelationType -in 6, 15 })

    # Build adjacency for chain depth: superseding -> superseded
    $adj = @{}
    foreach ($rel in $supersedence) {
        $key = $rel.FromAppCIID
        if (-not $adj.ContainsKey($key)) { $adj[$key] = @() }
        $adj[$key] += $rel.ToAppCIID
    }

    # BFS from each superseding app to compute chain depth
    $depthCache = @{}
    function Get-ChainDepth {
        param([uint32]$AppCIID, [hashtable]$Adj, [hashtable]$Cache, [hashtable]$Visiting)
        if ($Cache.ContainsKey($AppCIID)) { return $Cache[$AppCIID] }
        if ($Visiting.ContainsKey($AppCIID)) { return -1 }  # circular
        $Visiting[$AppCIID] = $true
        $maxChild = 0
        if ($Adj.ContainsKey($AppCIID)) {
            foreach ($child in $Adj[$AppCIID]) {
                $d = Get-ChainDepth -AppCIID $child -Adj $Adj -Cache $Cache -Visiting $Visiting
                if ($d -eq -1) { $Visiting.Remove($AppCIID); return -1 }
                if ($d + 1 -gt $maxChild) { $maxChild = $d + 1 }
            }
        }
        $Visiting.Remove($AppCIID)
        $Cache[$AppCIID] = $maxChild
        return $maxChild
    }

    $results = foreach ($rel in $supersedence) {
        $depth = Get-ChainDepth -AppCIID $rel.FromAppCIID -Adj $adj -Cache $depthCache -Visiting @{}

        # Determine status
        $status = 'Healthy'
        if (-not $rel.FromAppExists -or -not $rel.ToAppExists) {
            $status = 'Orphaned'
        }
        elseif ($depth -eq -1) {
            $status = 'Circular'
        }
        else {
            $fromApp = $AppLookup[$rel.FromAppCIID]
            $toApp   = $AppLookup[$rel.ToAppCIID]
            if ($toApp -and $toApp.IsExpired) { $status = 'Expired Target' }
            elseif ($fromApp -and -not $fromApp.IsEnabled) { $status = 'Disabled Source' }
        }

        [PSCustomObject]@{
            SupersedingApp     = $rel.FromAppName
            SupersedingVersion = $rel.FromAppVersion
            SupersedingCIID    = $rel.FromAppCIID
            SupersededApp      = $rel.ToAppName
            SupersededVersion  = $rel.ToAppVersion
            SupersededCIID     = $rel.ToAppCIID
            ChainDepth         = if ($depth -ge 0) { $depth } else { 0 }
            Status             = $status
        }
    }

    Write-Log "Found $(@($results).Count) supersedence relationships"
    return $results
}

function Find-DependencyGroups {
    <#
    .SYNOPSIS
        Extracts dependency relationships and classifies by type.
    .DESCRIPTION
        Uses RelationType 2 (Required), 4 (Optional), 10 (AppDependence).
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ResolvedRelationships,
        [Parameter(Mandatory)][hashtable]$AppLookup
    )

    $dependencies = @($ResolvedRelationships | Where-Object { $_.RelationType -in 2, 4, 10 })

    $results = foreach ($rel in $dependencies) {
        $depType = switch ($rel.RelationType) {
            2  { 'Required' }
            4  { 'Optional' }
            10 { 'App Dependency' }
        }

        # Determine status
        $status = 'Healthy'
        if (-not $rel.FromAppExists -or -not $rel.ToAppExists) {
            $status = 'Orphaned'
        }
        else {
            $toApp = $AppLookup[$rel.ToAppCIID]
            if ($toApp) {
                if ($toApp.IsExpired) { $status = 'Expired Target' }
                elseif (-not $toApp.IsEnabled) { $status = 'Disabled Target' }
                elseif (-not $toApp.HasContent) { $status = 'Missing Content' }
            }
        }

        [PSCustomObject]@{
            ParentApp        = $rel.FromAppName
            ParentVersion    = $rel.FromAppVersion
            ParentCIID       = $rel.FromAppCIID
            DependencyApp    = $rel.ToAppName
            DependencyVersion = $rel.ToAppVersion
            DependencyCIID   = $rel.ToAppCIID
            DependencyType   = $depType
            Level            = $rel.Level
            Status           = $status
        }
    }

    Write-Log "Found $(@($results).Count) dependency relationships"
    return $results
}

function Find-BrokenSupersedence {
    <#
    .SYNOPSIS
        Detects broken supersedence rules.
    .DESCRIPTION
        Returns broken rule objects with issue type, severity, and remediation guidance.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SupersedenceData,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ResolvedRelationships,
        [Parameter(Mandatory)][hashtable]$AppLookup
    )

    $results = @()

    # Check each supersedence relationship for issues
    foreach ($rel in $SupersedenceData) {
        switch ($rel.Status) {
            'Orphaned' {
                $results += [PSCustomObject]@{
                    IssueType   = 'Orphaned Reference'
                    Severity    = 'Error'
                    Category    = 'Supersedence'
                    FromApp     = $rel.SupersedingApp
                    ToApp       = $rel.SupersededApp
                    Description = "Supersedence references an application that no longer exists or has been deleted."
                    Remediation = "Open the superseding application's properties in the console, navigate to the Supersedence tab, and remove the broken reference."
                }
            }
            'Circular' {
                $results += [PSCustomObject]@{
                    IssueType   = 'Circular Chain'
                    Severity    = 'Error'
                    Category    = 'Supersedence'
                    FromApp     = $rel.SupersedingApp
                    ToApp       = $rel.SupersededApp
                    Description = "Circular supersedence chain detected. App A supersedes B which eventually supersedes A."
                    Remediation = "Review the full supersedence chain and remove the relationship that creates the loop."
                }
            }
            'Expired Target' {
                $results += [PSCustomObject]@{
                    IssueType   = 'Expired Target'
                    Severity    = 'Warning'
                    Category    = 'Supersedence'
                    FromApp     = $rel.SupersedingApp
                    ToApp       = $rel.SupersededApp
                    Description = "Superseded application '$($rel.SupersededApp)' is expired/retired. The supersedence rule still exists but the target is no longer active."
                    Remediation = "Consider removing the supersedence relationship since the target is already retired, or leave it if the retired cleanup task will handle it."
                }
            }
            'Disabled Source' {
                $results += [PSCustomObject]@{
                    IssueType   = 'Disabled Source'
                    Severity    = 'Warning'
                    Category    = 'Supersedence'
                    FromApp     = $rel.SupersedingApp
                    ToApp       = $rel.SupersededApp
                    Description = "Superseding application '$($rel.SupersedingApp)' is disabled. The supersedence rule exists but the replacement app cannot be deployed."
                    Remediation = "Enable the superseding application or remove the supersedence relationship."
                }
            }
        }
    }

    # Detect circular chains via DFS on the full graph
    $adj = @{}
    $supersedence = @($ResolvedRelationships | Where-Object { $_.RelationType -in 6, 15 })
    foreach ($rel in $supersedence) {
        $key = $rel.FromAppCIID
        if (-not $adj.ContainsKey($key)) { $adj[$key] = @() }
        $adj[$key] += $rel.ToAppCIID
    }

    $visited = @{}
    $inStack = @{}
    $circularFound = @{}

    function Test-Circular {
        param([uint32]$Node)
        if ($inStack.ContainsKey($Node)) {
            if (-not $circularFound.ContainsKey($Node)) {
                $circularFound[$Node] = $true
            }
            return
        }
        if ($visited.ContainsKey($Node)) { return }
        $visited[$Node] = $true
        $inStack[$Node] = $true
        if ($adj.ContainsKey($Node)) {
            foreach ($child in $adj[$Node]) {
                Test-Circular -Node $child
            }
        }
        $inStack.Remove($Node)
    }

    foreach ($node in $adj.Keys) {
        Test-Circular -Node $node
    }

    # Add circular entries that weren't already caught by status
    foreach ($ciid in $circularFound.Keys) {
        $appName = if ($AppLookup.ContainsKey($ciid)) { $AppLookup[$ciid].LocalizedDisplayName } else { "Unknown (CI_ID: $ciid)" }
        $alreadyReported = $results | Where-Object { $_.IssueType -eq 'Circular Chain' -and ($_.FromApp -eq $appName -or $_.ToApp -eq $appName) }
        if (-not $alreadyReported) {
            $results += [PSCustomObject]@{
                IssueType   = 'Circular Chain'
                Severity    = 'Error'
                Category    = 'Supersedence'
                FromApp     = $appName
                ToApp       = '(cycle participant)'
                Description = "Application '$appName' is part of a circular supersedence chain."
                Remediation = "Trace the supersedence chain from this application to find and break the loop."
            }
        }
    }

    return $results
}

function Find-BrokenDependencies {
    <#
    .SYNOPSIS
        Detects broken dependency rules.
    .DESCRIPTION
        Returns broken rule objects with issue type, severity, and remediation guidance.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$DependencyData,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ResolvedRelationships,
        [Parameter(Mandatory)][hashtable]$AppLookup
    )

    $results = @()

    foreach ($rel in $DependencyData) {
        switch ($rel.Status) {
            'Orphaned' {
                $results += [PSCustomObject]@{
                    IssueType   = 'Orphaned Reference'
                    Severity    = 'Error'
                    Category    = 'Dependency'
                    FromApp     = $rel.ParentApp
                    ToApp       = $rel.DependencyApp
                    Description = "Dependency references an application that no longer exists or has been deleted."
                    Remediation = "Open the parent application's deployment type properties, navigate to the Dependencies tab, and remove the broken reference."
                }
            }
            'Expired Target' {
                $results += [PSCustomObject]@{
                    IssueType   = 'Expired Target'
                    Severity    = 'Warning'
                    Category    = 'Dependency'
                    FromApp     = $rel.ParentApp
                    ToApp       = $rel.DependencyApp
                    Description = "Dependency target '$($rel.DependencyApp)' is expired/retired. Deployments of '$($rel.ParentApp)' may fail if this dependency is required."
                    Remediation = "Update the dependency to point to a current version of the required application, or remove it if no longer needed."
                }
            }
            'Disabled Target' {
                $results += [PSCustomObject]@{
                    IssueType   = 'Disabled Target'
                    Severity    = 'Warning'
                    Category    = 'Dependency'
                    FromApp     = $rel.ParentApp
                    ToApp       = $rel.DependencyApp
                    Description = "Dependency target '$($rel.DependencyApp)' is disabled. It cannot be auto-installed as a dependency."
                    Remediation = "Enable the dependency application or update the dependency to point to an active version."
                }
            }
            'Missing Content' {
                $results += [PSCustomObject]@{
                    IssueType   = 'Missing Content'
                    Severity    = 'Error'
                    Category    = 'Dependency'
                    FromApp     = $rel.ParentApp
                    ToApp       = $rel.DependencyApp
                    Description = "Dependency target '$($rel.DependencyApp)' has no content. It cannot be installed as a dependency."
                    Remediation = "Distribute content for the dependency application or update the dependency reference."
                }
            }
        }
    }

    # Detect circular dependencies via DFS
    $adj = @{}
    $depRels = @($ResolvedRelationships | Where-Object { $_.RelationType -in 2, 4, 10 })
    foreach ($rel in $depRels) {
        $key = $rel.FromAppCIID
        if (-not $adj.ContainsKey($key)) { $adj[$key] = @() }
        $adj[$key] += $rel.ToAppCIID
    }

    $visited = @{}
    $inStack = @{}
    $circularFound = @{}

    function Test-DepCircular {
        param([uint32]$Node)
        if ($inStack.ContainsKey($Node)) {
            if (-not $circularFound.ContainsKey($Node)) {
                $circularFound[$Node] = $true
            }
            return
        }
        if ($visited.ContainsKey($Node)) { return }
        $visited[$Node] = $true
        $inStack[$Node] = $true
        if ($adj.ContainsKey($Node)) {
            foreach ($child in $adj[$Node]) {
                Test-DepCircular -Node $child
            }
        }
        $inStack.Remove($Node)
    }

    foreach ($node in $adj.Keys) {
        Test-DepCircular -Node $node
    }

    foreach ($ciid in $circularFound.Keys) {
        $appName = if ($AppLookup.ContainsKey($ciid)) { $AppLookup[$ciid].LocalizedDisplayName } else { "Unknown (CI_ID: $ciid)" }
        $results += [PSCustomObject]@{
            IssueType   = 'Circular Dependency'
            Severity    = 'Error'
            Category    = 'Dependency'
            FromApp     = $appName
            ToApp       = '(cycle participant)'
            Description = "Application '$appName' is part of a circular dependency chain. MECM enforces a max depth of 5."
            Remediation = "Trace the dependency chain from this application to find and break the loop."
        }
    }

    return $results
}

function Find-UndocumentedRelationships {
    <#
    .SYNOPSIS
        Finds apps with supersedence or dependencies but no description/comments.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SupersedenceData,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$DependencyData,
        [Parameter(Mandatory)][hashtable]$AppLookup
    )

    # Collect all unique app CI_IDs that participate in relationships
    $appCIIDs = @{}
    foreach ($rel in $SupersedenceData) {
        $appCIIDs[$rel.SupersedingCIID] = $true
        $appCIIDs[$rel.SupersededCIID]  = $true
    }
    foreach ($rel in $DependencyData) {
        $appCIIDs[$rel.ParentCIID]     = $true
        $appCIIDs[$rel.DependencyCIID] = $true
    }

    $results = @()
    foreach ($ciid in $appCIIDs.Keys) {
        if (-not $AppLookup.ContainsKey($ciid)) { continue }
        $app = $AppLookup[$ciid]

        # Check if the app has relationships but we can't easily check the
        # AdminUI description field from WMI. The SMS_Application class does
        # have LocalizedDescription - check if it's empty.
        # Note: We already have the app object but LocalizedDescription may
        # not be in our summary. We'll flag apps where the manufacturer field
        # is empty as a proxy for "poorly documented".
        # A more accurate check would require an additional WMI property.
    }

    # For now, return apps that have relationships and IsSuperseded or IsSuperseding
    # but are missing manufacturer info (common indicator of poor documentation)
    foreach ($ciid in $appCIIDs.Keys) {
        if (-not $AppLookup.ContainsKey($ciid)) { continue }
        $app = $AppLookup[$ciid]
        if ([string]::IsNullOrWhiteSpace($app.Manufacturer)) {
            $results += [PSCustomObject]@{
                IssueType   = 'Undocumented'
                Severity    = 'Info'
                Category    = 'Documentation'
                FromApp     = $app.LocalizedDisplayName
                ToApp       = ''
                Description = "Application '$($app.LocalizedDisplayName)' participates in supersedence or dependency relationships but has no Manufacturer set."
                Remediation = "Set the Manufacturer field in the application properties to improve traceability."
            }
        }
    }

    Write-Log "Found $(@($results).Count) undocumented relationship participants"
    return $results
}

function Get-ScanSummaryCounts {
    <#
    .SYNOPSIS
        Aggregates counts for the 4 summary cards.
    #>
    param(
        [int]$AppCount,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SupersedenceData,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$DependencyData,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$BrokenRules
    )

    $supBroken = @($SupersedenceData | Where-Object { $_.Status -ne 'Healthy' }).Count
    $depBroken = @($DependencyData | Where-Object { $_.Status -ne 'Healthy' }).Count

    return [PSCustomObject]@{
        AppCount            = $AppCount
        SupersedenceTotal   = @($SupersedenceData).Count
        SupersedenceBroken  = $supBroken
        DependencyTotal     = @($DependencyData).Count
        DependencyBroken    = $depBroken
        BrokenRulesTotal    = @($BrokenRules).Count
        BrokenErrors        = @($BrokenRules | Where-Object { $_.Severity -eq 'Error' }).Count
        BrokenWarnings      = @($BrokenRules | Where-Object { $_.Severity -eq 'Warning' }).Count
        BrokenInfo          = @($BrokenRules | Where-Object { $_.Severity -eq 'Info' }).Count
    }
}

# ---------------------------------------------------------------------------
# Tree Building
# ---------------------------------------------------------------------------

function Build-SupersedenceTree {
    <#
    .SYNOPSIS
        Builds a nested hierarchy for supersedence chain visualization.
    .DESCRIPTION
        Returns root nodes (apps that supersede others but are not themselves superseded).
        Each node has Children containing the apps it supersedes.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SupersedenceData,
        [Parameter(Mandatory)][hashtable]$AppLookup
    )

    # Build parent->children map (superseding -> superseded)
    $children = @{}
    $hasParent = @{}

    foreach ($rel in $SupersedenceData) {
        $parentKey = $rel.SupersedingCIID
        if (-not $children.ContainsKey($parentKey)) { $children[$parentKey] = @() }
        $children[$parentKey] += [PSCustomObject]@{
            CIID    = $rel.SupersededCIID
            Name    = $rel.SupersededApp
            Version = $rel.SupersededVersion
            Status  = $rel.Status
        }
        $hasParent[$rel.SupersededCIID] = $true
    }

    # Root nodes: superseding apps that are not themselves superseded
    $roots = @()
    foreach ($rel in $SupersedenceData) {
        $ciid = $rel.SupersedingCIID
        if (-not $hasParent.ContainsKey($ciid)) {
            # Only add each root once
            $existing = $roots | Where-Object { $_.CIID -eq $ciid }
            if (-not $existing) {
                $app = if ($AppLookup.ContainsKey($ciid)) { $AppLookup[$ciid] } else { $null }
                $roots += [PSCustomObject]@{
                    CIID     = $ciid
                    Name     = $rel.SupersedingApp
                    Version  = $rel.SupersedingVersion
                    Status   = if ($app -and $app.IsExpired) { 'Expired' } elseif ($app -and -not $app.IsEnabled) { 'Disabled' } else { 'Active' }
                    Children = if ($children.ContainsKey($ciid)) { $children[$ciid] } else { @() }
                }
            }
        }
    }

    # Recursively attach grandchildren
    function Attach-Children {
        param([object[]]$Nodes)
        foreach ($node in $Nodes) {
            if ($children.ContainsKey($node.CIID)) {
                $node | Add-Member -NotePropertyName Children -NotePropertyValue $children[$node.CIID] -Force
                Attach-Children -Nodes $node.Children
            }
            else {
                $node | Add-Member -NotePropertyName Children -NotePropertyValue @() -Force
            }
        }
    }

    Attach-Children -Nodes $roots
    return $roots
}

function Build-DependencyTree {
    <#
    .SYNOPSIS
        Builds a nested hierarchy for dependency visualization.
    .DESCRIPTION
        Returns root nodes (apps that have dependencies).
        Each node has Children containing the apps it depends on.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$DependencyData,
        [Parameter(Mandatory)][hashtable]$AppLookup
    )

    # Build parent->children map
    $children = @{}

    foreach ($rel in $DependencyData) {
        $parentKey = $rel.ParentCIID
        if (-not $children.ContainsKey($parentKey)) { $children[$parentKey] = @() }
        $children[$parentKey] += [PSCustomObject]@{
            CIID    = $rel.DependencyCIID
            Name    = $rel.DependencyApp
            Version = $rel.DependencyVersion
            Type    = $rel.DependencyType
            Status  = $rel.Status
        }
    }

    # Root nodes: apps that have dependencies
    $roots = @()
    foreach ($parentCIID in $children.Keys) {
        $app = if ($AppLookup.ContainsKey($parentCIID)) { $AppLookup[$parentCIID] } else { $null }
        $name = if ($app) { $app.LocalizedDisplayName } else { "Unknown (CI_ID: $parentCIID)" }
        $version = if ($app) { $app.SoftwareVersion } else { '' }

        $roots += [PSCustomObject]@{
            CIID     = $parentCIID
            Name     = $name
            Version  = $version
            Status   = if ($app -and $app.IsExpired) { 'Expired' } elseif ($app -and -not $app.IsEnabled) { 'Disabled' } else { 'Active' }
            Children = $children[$parentCIID]
        }
    }

    return $roots
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

function Export-AuditCsv {
    <#
    .SYNOPSIS
        Exports a DataTable to CSV.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $rows = @()
    foreach ($row in $DataTable.Rows) {
        $obj = [ordered]@{}
        foreach ($col in $DataTable.Columns) {
            $obj[$col.ColumnName] = $row[$col.ColumnName]
        }
        $rows += [PSCustomObject]$obj
    }

    $rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Log "Exported CSV to $OutputPath"
}

function Export-AuditHtml {
    <#
    .SYNOPSIS
        Exports a DataTable to a self-contained HTML report with color-coded severity.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$ReportTitle = 'Supersedence & Dependency Audit Report'
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $css = @(
        '<style>',
        'body { font-family: "Segoe UI", Arial, sans-serif; margin: 20px; background: #fafafa; }',
        'h1 { color: #0078D4; margin-bottom: 4px; }',
        '.summary { color: #666; margin-bottom: 12px; font-size: 0.9em; }',
        'table { border-collapse: collapse; width: 100%; margin-top: 12px; }',
        'th { background: #0078D4; color: #fff; padding: 8px 12px; text-align: left; }',
        'td { padding: 6px 12px; border-bottom: 1px solid #e0e0e0; }',
        'tr:nth-child(even) { background: #f5f5f5; }',
        '.error { color: #c00; font-weight: bold; }',
        '.warning { color: #b87800; }',
        '.info { color: #0078D4; }',
        '.healthy { color: #228b22; }',
        '</style>'
    ) -join "`r`n"

    $headerRow = ($DataTable.Columns | ForEach-Object { "<th>$($_.ColumnName)</th>" }) -join ''
    $bodyRows = foreach ($row in $DataTable.Rows) {
        $cells = foreach ($col in $DataTable.Columns) {
            $val = [string]$row[$col.ColumnName]
            $cssClass = ''
            if ($col.ColumnName -eq 'Severity' -or $col.ColumnName -eq 'Status') {
                switch -Regex ($val) {
                    '^Error$'          { $cssClass = ' class="error"' }
                    'Orphaned|Circular|Missing' { $cssClass = ' class="error"' }
                    '^Warning$'        { $cssClass = ' class="warning"' }
                    'Expired|Disabled' { $cssClass = ' class="warning"' }
                    '^Info$'           { $cssClass = ' class="info"' }
                    '^Healthy$'        { $cssClass = ' class="healthy"' }
                }
            }
            "<td$cssClass>$val</td>"
        }
        "<tr>$($cells -join '')</tr>"
    }

    $html = @(
        '<!DOCTYPE html>',
        '<html><head><meta charset="utf-8"><title>' + $ReportTitle + '</title>',
        $css,
        '</head><body>',
        "<h1>$ReportTitle</h1>",
        "<div class='summary'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Rows: $($DataTable.Rows.Count)</div>",
        "<table><thead><tr>$headerRow</tr></thead>",
        "<tbody>$($bodyRows -join "`r`n")</tbody></table>",
        '</body></html>'
    ) -join "`r`n"

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
    Write-Log "Exported HTML to $OutputPath"
}

function New-AuditSummaryText {
    <#
    .SYNOPSIS
        Returns a plain text summary of the audit for clipboard/log.
    #>
    param(
        [Parameter(Mandatory)][PSCustomObject]$Counts
    )

    $lines = @(
        "Supersedence & Dependency Audit Summary - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        ("-" * 60),
        "Applications:  $($Counts.AppCount) scanned",
        "Supersedence:  $($Counts.SupersedenceTotal) rules ($($Counts.SupersedenceBroken) broken)",
        "Dependencies:  $($Counts.DependencyTotal) rules ($($Counts.DependencyBroken) broken)",
        "Broken Rules:  $($Counts.BrokenRulesTotal) total ($($Counts.BrokenErrors) errors, $($Counts.BrokenWarnings) warnings, $($Counts.BrokenInfo) info)"
    )

    return ($lines -join "`r`n")
}
