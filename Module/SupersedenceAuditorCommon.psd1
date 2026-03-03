@{
    RootModule        = 'SupersedenceAuditorCommon.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'
    Author            = 'Jason Ulbright'
    Description       = 'Supersedence and dependency auditor for MECM applications - relationship discovery, broken rule detection, tree visualization.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging
        'Initialize-Logging'
        'Write-Log'

        # CM Connection
        'Connect-CMSite'
        'Disconnect-CMSite'
        'Test-CMConnection'

        # WMI Bulk Discovery
        'Get-AllApplicationSummary'
        'Get-AllDeploymentTypeSummary'
        'Get-AllRelationships'
        'Resolve-RelationshipData'

        # Analysis
        'Find-SupersedenceChains'
        'Find-DependencyGroups'
        'Find-BrokenSupersedence'
        'Find-BrokenDependencies'
        'Find-UndocumentedRelationships'
        'Get-ScanSummaryCounts'

        # Tree Building
        'Build-SupersedenceTree'
        'Build-DependencyTree'

        # Export
        'Export-AuditCsv'
        'Export-AuditHtml'
        'New-AuditSummaryText'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
