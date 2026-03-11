#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x tests for SupersedenceAuditorCommon shared module.

.DESCRIPTION
    Tests pure-logic functions: logging, relationship resolution, supersedence
    chain analysis, dependency group analysis, broken rule detection, circular
    detection, tree building, export, and summary text. Does NOT require MECM,
    WMI, or administrator elevation.

.EXAMPLE
    Invoke-Pester .\SupersedenceAuditorCommon.Tests.ps1
#>

BeforeAll {
    Import-Module "$PSScriptRoot\SupersedenceAuditorCommon.psd1" -Force -DisableNameChecking
}

# ============================================================================
# Write-Log / Initialize-Logging
# ============================================================================

Describe 'Write-Log' {
    It 'writes formatted message to log file' {
        $logFile = Join-Path $TestDrive 'test.log'
        Initialize-Logging -LogPath $logFile

        Write-Log 'Hello world' -Quiet

        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[INFO \] Hello world'
    }

    It 'tags WARN messages correctly' {
        $logFile = Join-Path $TestDrive 'warn.log'
        Initialize-Logging -LogPath $logFile

        Write-Log 'Something odd' -Level WARN -Quiet

        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[WARN \] Something odd'
    }

    It 'tags ERROR messages correctly' {
        $logFile = Join-Path $TestDrive 'error.log'
        Initialize-Logging -LogPath $logFile

        Write-Log 'Failure' -Level ERROR -Quiet

        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[ERROR\] Failure'
    }

    It 'accepts empty string message' {
        $logFile = Join-Path $TestDrive 'empty.log'
        Initialize-Logging -LogPath $logFile

        { Write-Log '' -Quiet } | Should -Not -Throw

        $lines = Get-Content -LiteralPath $logFile
        $lines.Count | Should -BeGreaterOrEqual 2
    }
}

Describe 'Initialize-Logging' {
    It 'creates log file with header line' {
        $logFile = Join-Path $TestDrive 'init.log'
        Initialize-Logging -LogPath $logFile

        Test-Path -LiteralPath $logFile | Should -BeTrue
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[INFO \] === Log initialized ==='
    }

    It 'creates parent directories if missing' {
        $logFile = Join-Path $TestDrive 'sub\dir\deep.log'
        Initialize-Logging -LogPath $logFile

        Test-Path -LiteralPath $logFile | Should -BeTrue
    }
}

# ============================================================================
# Note: Get-AllResolvedRelationships requires CM cmdlets (integration test).
# The downstream analysis functions are tested with mock resolved objects below.
# ============================================================================

# ============================================================================
# Find-SupersedenceChains
# ============================================================================

Describe 'Find-SupersedenceChains' {
    BeforeAll {
        $appLookup = @{
            100 = [PSCustomObject]@{ CI_ID = 100; LocalizedDisplayName = 'New App v3'; SoftwareVersion = '3.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $true; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 1; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
            200 = [PSCustomObject]@{ CI_ID = 200; LocalizedDisplayName = 'Old App v2'; SoftwareVersion = '2.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $true; IsSuperseding = $true; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 1; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
            300 = [PSCustomObject]@{ CI_ID = 300; LocalizedDisplayName = 'Ancient App v1'; SoftwareVersion = '1.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $true; IsSuperseded = $true; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 0; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
        }

        # Chain: v3 supersedes v2, v2 supersedes v1
        $resolved = @(
            [PSCustomObject]@{ FromAppCIID = 100; FromAppName = 'New App v3'; FromAppVersion = '3.0'; FromAppExists = $true; FromDTCIID = 0; FromDTName = ''; ToAppCIID = 200; ToAppName = 'Old App v2'; ToAppVersion = '2.0'; ToAppExists = $true; ToDTCIID = 0; ToDTName = ''; RelationType = 6; RelationTypeName = 'Superseded'; Level = 0 }
            [PSCustomObject]@{ FromAppCIID = 200; FromAppName = 'Old App v2'; FromAppVersion = '2.0'; FromAppExists = $true; FromDTCIID = 0; FromDTName = ''; ToAppCIID = 300; ToAppName = 'Ancient App v1'; ToAppVersion = '1.0'; ToAppExists = $true; ToDTCIID = 0; ToDTName = ''; RelationType = 6; RelationTypeName = 'Superseded'; Level = 0 }
        )

        $logFile = Join-Path $TestDrive 'sup.log'
        Initialize-Logging -LogPath $logFile
        $script:supData = @(Find-SupersedenceChains -ResolvedRelationships $resolved -AppLookup $appLookup)
    }

    It 'finds both supersedence relationships' {
        $script:supData.Count | Should -Be 2
    }

    It 'computes chain depth for the root' {
        $rootRel = $script:supData | Where-Object { $_.SupersedingApp -eq 'New App v3' }
        $rootRel.ChainDepth | Should -BeGreaterOrEqual 1
    }

    It 'marks expired target' {
        $ancientRel = $script:supData | Where-Object { $_.SupersededApp -eq 'Ancient App v1' }
        $ancientRel.Status | Should -Be 'Expired Target'
    }

    It 'marks healthy relationships correctly' {
        $healthyRel = $script:supData | Where-Object { $_.SupersededApp -eq 'Old App v2' }
        $healthyRel.Status | Should -Be 'Healthy'
    }
}

Describe 'Find-SupersedenceChains with disabled source' {
    It 'detects disabled source' {
        $appLookup = @{
            10 = [PSCustomObject]@{ CI_ID = 10; LocalizedDisplayName = 'Disabled New'; SoftwareVersion = '2.0'; Manufacturer = 'V'; IsEnabled = $false; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $true; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 0; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
            20 = [PSCustomObject]@{ CI_ID = 20; LocalizedDisplayName = 'Active Old'; SoftwareVersion = '1.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $true; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 1; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
        }

        $resolved = @(
            [PSCustomObject]@{ FromAppCIID = 10; FromAppName = 'Disabled New'; FromAppVersion = '2.0'; FromAppExists = $true; FromDTCIID = 0; FromDTName = ''; ToAppCIID = 20; ToAppName = 'Active Old'; ToAppVersion = '1.0'; ToAppExists = $true; ToDTCIID = 0; ToDTName = ''; RelationType = 6; RelationTypeName = 'Superseded'; Level = 0 }
        )

        $logFile = Join-Path $TestDrive 'dis.log'
        Initialize-Logging -LogPath $logFile
        $result = @(Find-SupersedenceChains -ResolvedRelationships $resolved -AppLookup $appLookup)
        $result[0].Status | Should -Be 'Disabled Source'
    }
}

Describe 'Find-SupersedenceChains with orphaned reference' {
    It 'detects orphaned when target app missing from lookup' {
        $appLookup = @{
            10 = [PSCustomObject]@{ CI_ID = 10; LocalizedDisplayName = 'Orphan Source'; SoftwareVersion = '2.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $true; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 1; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
        }

        $resolved = @(
            [PSCustomObject]@{ FromAppCIID = 10; FromAppName = 'Orphan Source'; FromAppVersion = '2.0'; FromAppExists = $true; FromDTCIID = 0; FromDTName = ''; ToAppCIID = 9999; ToAppName = 'Unknown (CI_ID: 9999)'; ToAppVersion = ''; ToAppExists = $false; ToDTCIID = 0; ToDTName = ''; RelationType = 6; RelationTypeName = 'Superseded'; Level = 0 }
        )

        $logFile = Join-Path $TestDrive 'orph.log'
        Initialize-Logging -LogPath $logFile
        $result = @(Find-SupersedenceChains -ResolvedRelationships $resolved -AppLookup $appLookup)
        $result[0].Status | Should -Be 'Orphaned'
    }
}

# ============================================================================
# Find-DependencyGroups
# ============================================================================

Describe 'Find-DependencyGroups' {
    BeforeAll {
        $appLookup = @{
            100 = [PSCustomObject]@{ CI_ID = 100; LocalizedDisplayName = 'Main App'; SoftwareVersion = '1.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 2; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
            200 = [PSCustomObject]@{ CI_ID = 200; LocalizedDisplayName = 'VC++ Runtime'; SoftwareVersion = '14.0'; Manufacturer = 'Microsoft'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 5; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
            300 = [PSCustomObject]@{ CI_ID = 300; LocalizedDisplayName = 'Expired Dep'; SoftwareVersion = '1.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $true; IsSuperseded = $false; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 0; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
            400 = [PSCustomObject]@{ CI_ID = 400; LocalizedDisplayName = 'No Content Dep'; SoftwareVersion = '1.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $false; HasContent = $false; NumberOfDeploymentTypes = 1; NumberOfDeployments = 0; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
            500 = [PSCustomObject]@{ CI_ID = 500; LocalizedDisplayName = 'Disabled Dep'; SoftwareVersion = '1.0'; Manufacturer = 'V'; IsEnabled = $false; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 0; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
        }

        $resolved = @(
            [PSCustomObject]@{ FromAppCIID = 100; FromAppName = 'Main App'; FromAppVersion = '1.0'; FromAppExists = $true; FromDTCIID = 0; FromDTName = ''; ToAppCIID = 200; ToAppName = 'VC++ Runtime'; ToAppVersion = '14.0'; ToAppExists = $true; ToDTCIID = 0; ToDTName = ''; RelationType = 2; RelationTypeName = 'Required'; Level = 0 }
            [PSCustomObject]@{ FromAppCIID = 100; FromAppName = 'Main App'; FromAppVersion = '1.0'; FromAppExists = $true; FromDTCIID = 0; FromDTName = ''; ToAppCIID = 300; ToAppName = 'Expired Dep'; ToAppVersion = '1.0'; ToAppExists = $true; ToDTCIID = 0; ToDTName = ''; RelationType = 4; RelationTypeName = 'Optional'; Level = 0 }
            [PSCustomObject]@{ FromAppCIID = 100; FromAppName = 'Main App'; FromAppVersion = '1.0'; FromAppExists = $true; FromDTCIID = 0; FromDTName = ''; ToAppCIID = 400; ToAppName = 'No Content Dep'; ToAppVersion = '1.0'; ToAppExists = $true; ToDTCIID = 0; ToDTName = ''; RelationType = 10; RelationTypeName = 'AppDependence'; Level = 0 }
            [PSCustomObject]@{ FromAppCIID = 100; FromAppName = 'Main App'; FromAppVersion = '1.0'; FromAppExists = $true; FromDTCIID = 0; FromDTName = ''; ToAppCIID = 500; ToAppName = 'Disabled Dep'; ToAppVersion = '1.0'; ToAppExists = $true; ToDTCIID = 0; ToDTName = ''; RelationType = 2; RelationTypeName = 'Required'; Level = 0 }
        )

        $logFile = Join-Path $TestDrive 'dep.log'
        Initialize-Logging -LogPath $logFile
        $script:depData = @(Find-DependencyGroups -ResolvedRelationships $resolved -AppLookup $appLookup)
    }

    It 'finds all 4 dependency relationships' {
        $script:depData.Count | Should -Be 4
    }

    It 'classifies Required type correctly' {
        $reqDeps = $script:depData | Where-Object { $_.DependencyApp -eq 'VC++ Runtime' }
        $reqDeps.DependencyType | Should -Be 'Required'
    }

    It 'classifies Optional type correctly' {
        $optDeps = $script:depData | Where-Object { $_.DependencyApp -eq 'Expired Dep' }
        $optDeps.DependencyType | Should -Be 'Optional'
    }

    It 'classifies App Dependency type correctly' {
        $appDeps = $script:depData | Where-Object { $_.DependencyApp -eq 'No Content Dep' }
        $appDeps.DependencyType | Should -Be 'App Dependency'
    }

    It 'marks healthy dependency correctly' {
        $healthy = $script:depData | Where-Object { $_.DependencyApp -eq 'VC++ Runtime' }
        $healthy.Status | Should -Be 'Healthy'
    }

    It 'marks expired target correctly' {
        $expired = $script:depData | Where-Object { $_.DependencyApp -eq 'Expired Dep' }
        $expired.Status | Should -Be 'Expired Target'
    }

    It 'marks missing content correctly' {
        $noContent = $script:depData | Where-Object { $_.DependencyApp -eq 'No Content Dep' }
        $noContent.Status | Should -Be 'Missing Content'
    }

    It 'marks disabled target correctly' {
        $disabled = $script:depData | Where-Object { $_.DependencyApp -eq 'Disabled Dep' }
        $disabled.Status | Should -Be 'Disabled Target'
    }
}

# ============================================================================
# Find-BrokenSupersedence
# ============================================================================

Describe 'Find-BrokenSupersedence' {
    BeforeAll {
        $appLookup = @{
            10 = [PSCustomObject]@{ CI_ID = 10; LocalizedDisplayName = 'Good New'; SoftwareVersion = '2.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $true; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 1; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
            20 = [PSCustomObject]@{ CI_ID = 20; LocalizedDisplayName = 'Expired Old'; SoftwareVersion = '1.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $true; IsSuperseded = $true; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 0; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
        }

        $supData = @(
            [PSCustomObject]@{ SupersedingApp = 'Good New'; SupersedingVersion = '2.0'; SupersedingCIID = 10; SupersededApp = 'Expired Old'; SupersededVersion = '1.0'; SupersededCIID = 20; ChainDepth = 1; Status = 'Expired Target' }
        )

        $resolved = @(
            [PSCustomObject]@{ FromAppCIID = 10; ToAppCIID = 20; RelationType = 6 }
        )

        $logFile = Join-Path $TestDrive 'bsup.log'
        Initialize-Logging -LogPath $logFile
        $script:brokenSup = @(Find-BrokenSupersedence -SupersedenceData $supData -ResolvedRelationships $resolved -AppLookup $appLookup)
    }

    It 'detects expired target as warning' {
        $script:brokenSup.Count | Should -BeGreaterOrEqual 1
        $expRule = $script:brokenSup | Where-Object { $_.IssueType -eq 'Expired Target' }
        $expRule | Should -Not -BeNullOrEmpty
        $expRule.Severity | Should -Be 'Warning'
        $expRule.Category | Should -Be 'Supersedence'
    }

    It 'includes remediation text' {
        $script:brokenSup[0].Remediation | Should -Not -BeNullOrEmpty
    }
}

# ============================================================================
# Find-BrokenDependencies
# ============================================================================

Describe 'Find-BrokenDependencies' {
    BeforeAll {
        $appLookup = @{
            100 = [PSCustomObject]@{ CI_ID = 100; LocalizedDisplayName = 'Parent'; SoftwareVersion = '1.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 1; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
        }

        $depData = @(
            [PSCustomObject]@{ ParentApp = 'Parent'; ParentVersion = '1.0'; ParentCIID = 100; DependencyApp = 'Unknown (CI_ID: 9999)'; DependencyVersion = ''; DependencyCIID = 9999; DependencyType = 'Required'; Level = 0; Status = 'Orphaned' }
            [PSCustomObject]@{ ParentApp = 'Parent'; ParentVersion = '1.0'; ParentCIID = 100; DependencyApp = 'Some Dep'; DependencyVersion = '1.0'; DependencyCIID = 200; DependencyType = 'Required'; Level = 0; Status = 'Missing Content' }
        )

        $resolved = @(
            [PSCustomObject]@{ FromAppCIID = 100; ToAppCIID = 9999; RelationType = 2 }
            [PSCustomObject]@{ FromAppCIID = 100; ToAppCIID = 200; RelationType = 2 }
        )

        $logFile = Join-Path $TestDrive 'bdep.log'
        Initialize-Logging -LogPath $logFile
        $script:brokenDep = @(Find-BrokenDependencies -DependencyData $depData -ResolvedRelationships $resolved -AppLookup $appLookup)
    }

    It 'detects orphaned reference as error' {
        $orphaned = $script:brokenDep | Where-Object { $_.IssueType -eq 'Orphaned Reference' }
        $orphaned | Should -Not -BeNullOrEmpty
        $orphaned.Severity | Should -Be 'Error'
        $orphaned.Category | Should -Be 'Dependency'
    }

    It 'detects missing content as error' {
        $missing = $script:brokenDep | Where-Object { $_.IssueType -eq 'Missing Content' }
        $missing | Should -Not -BeNullOrEmpty
        $missing.Severity | Should -Be 'Error'
    }
}

# ============================================================================
# Find-UndocumentedRelationships
# ============================================================================

Describe 'Find-UndocumentedRelationships' {
    It 'flags apps with empty Manufacturer' {
        $appLookup = @{
            10 = [PSCustomObject]@{ CI_ID = 10; LocalizedDisplayName = 'Documented App'; SoftwareVersion = '1.0'; Manufacturer = 'Vendor'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $true; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 1; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
            20 = [PSCustomObject]@{ CI_ID = 20; LocalizedDisplayName = 'Undocumented App'; SoftwareVersion = '1.0'; Manufacturer = ''; IsEnabled = $true; IsExpired = $false; IsSuperseded = $true; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 0; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
        }

        $supData = @(
            [PSCustomObject]@{ SupersedingApp = 'Documented App'; SupersedingVersion = '1.0'; SupersedingCIID = 10; SupersededApp = 'Undocumented App'; SupersededVersion = '1.0'; SupersededCIID = 20; ChainDepth = 1; Status = 'Healthy' }
        )

        $logFile = Join-Path $TestDrive 'undoc.log'
        Initialize-Logging -LogPath $logFile
        $result = @(Find-UndocumentedRelationships -SupersedenceData $supData -DependencyData @() -AppLookup $appLookup)

        $undoc = $result | Where-Object { $_.FromApp -eq 'Undocumented App' }
        $undoc | Should -Not -BeNullOrEmpty
        $undoc.IssueType | Should -Be 'Undocumented'
        $undoc.Severity | Should -Be 'Info'
    }

    It 'does not flag apps with Manufacturer set' {
        $appLookup = @{
            10 = [PSCustomObject]@{ CI_ID = 10; LocalizedDisplayName = 'Good App'; SoftwareVersion = '1.0'; Manufacturer = 'Vendor A'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 1; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
        }

        $depData = @(
            [PSCustomObject]@{ ParentApp = 'Good App'; ParentVersion = '1.0'; ParentCIID = 10; DependencyApp = 'Something'; DependencyVersion = '1.0'; DependencyCIID = 99; DependencyType = 'Required'; Level = 0; Status = 'Healthy' }
        )

        $logFile = Join-Path $TestDrive 'undoc2.log'
        Initialize-Logging -LogPath $logFile
        $result = @(Find-UndocumentedRelationships -SupersedenceData @() -DependencyData $depData -AppLookup $appLookup)

        $flagged = $result | Where-Object { $_.FromApp -eq 'Good App' }
        $flagged | Should -BeNullOrEmpty
    }
}

# ============================================================================
# Get-ScanSummaryCounts
# ============================================================================

Describe 'Get-ScanSummaryCounts' {
    It 'aggregates counts correctly' {
        $supData = @(
            [PSCustomObject]@{ Status = 'Healthy' }
            [PSCustomObject]@{ Status = 'Expired Target' }
            [PSCustomObject]@{ Status = 'Healthy' }
        )
        $depData = @(
            [PSCustomObject]@{ Status = 'Healthy' }
            [PSCustomObject]@{ Status = 'Missing Content' }
        )
        $brokenRules = @(
            [PSCustomObject]@{ Severity = 'Error' }
            [PSCustomObject]@{ Severity = 'Warning' }
            [PSCustomObject]@{ Severity = 'Info' }
        )

        $counts = Get-ScanSummaryCounts -AppCount 900 -SupersedenceData $supData -DependencyData $depData -BrokenRules $brokenRules

        $counts.AppCount | Should -Be 900
        $counts.SupersedenceTotal | Should -Be 3
        $counts.SupersedenceBroken | Should -Be 1
        $counts.DependencyTotal | Should -Be 2
        $counts.DependencyBroken | Should -Be 1
        $counts.BrokenRulesTotal | Should -Be 3
        $counts.BrokenErrors | Should -Be 1
        $counts.BrokenWarnings | Should -Be 1
        $counts.BrokenInfo | Should -Be 1
    }
}

# ============================================================================
# Build-SupersedenceTree
# ============================================================================

Describe 'Build-SupersedenceTree' {
    It 'builds root nodes for top-level superseding apps' {
        $appLookup = @{
            100 = [PSCustomObject]@{ CI_ID = 100; LocalizedDisplayName = 'App v3'; SoftwareVersion = '3.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $true; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 1; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
            200 = [PSCustomObject]@{ CI_ID = 200; LocalizedDisplayName = 'App v2'; SoftwareVersion = '2.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $true; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 0; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
        }

        $supData = @(
            [PSCustomObject]@{ SupersedingApp = 'App v3'; SupersedingVersion = '3.0'; SupersedingCIID = 100; SupersededApp = 'App v2'; SupersededVersion = '2.0'; SupersededCIID = 200; ChainDepth = 1; Status = 'Healthy' }
        )

        $logFile = Join-Path $TestDrive 'tree.log'
        Initialize-Logging -LogPath $logFile
        $roots = @(Build-SupersedenceTree -SupersedenceData $supData -AppLookup $appLookup)

        $roots.Count | Should -Be 1
        $roots[0].Name | Should -Be 'App v3'
        $roots[0].CIID | Should -Be 100
        $roots[0].Children.Count | Should -Be 1
        $roots[0].Children[0].Name | Should -Be 'App v2'
    }
}

# ============================================================================
# Build-DependencyTree
# ============================================================================

Describe 'Build-DependencyTree' {
    It 'builds root nodes for apps with dependencies' {
        $appLookup = @{
            100 = [PSCustomObject]@{ CI_ID = 100; LocalizedDisplayName = 'LOB App'; SoftwareVersion = '1.0'; Manufacturer = 'V'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 2; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
            200 = [PSCustomObject]@{ CI_ID = 200; LocalizedDisplayName = 'VC++ Runtime'; SoftwareVersion = '14.0'; Manufacturer = 'Microsoft'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 5; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
            300 = [PSCustomObject]@{ CI_ID = 300; LocalizedDisplayName = '.NET Runtime'; SoftwareVersion = '8.0'; Manufacturer = 'Microsoft'; IsEnabled = $true; IsExpired = $false; IsSuperseded = $false; IsSuperseding = $false; HasContent = $true; NumberOfDeploymentTypes = 1; NumberOfDeployments = 3; DateCreated = (Get-Date); DateLastModified = (Get-Date); CreatedBy = 'a'; LastModifiedBy = 'a' }
        }

        $depData = @(
            [PSCustomObject]@{ ParentApp = 'LOB App'; ParentVersion = '1.0'; ParentCIID = 100; DependencyApp = 'VC++ Runtime'; DependencyVersion = '14.0'; DependencyCIID = 200; DependencyType = 'Required'; Level = 0; Status = 'Healthy' }
            [PSCustomObject]@{ ParentApp = 'LOB App'; ParentVersion = '1.0'; ParentCIID = 100; DependencyApp = '.NET Runtime'; DependencyVersion = '8.0'; DependencyCIID = 300; DependencyType = 'Required'; Level = 0; Status = 'Healthy' }
        )

        $logFile = Join-Path $TestDrive 'dtree.log'
        Initialize-Logging -LogPath $logFile
        $roots = @(Build-DependencyTree -DependencyData $depData -AppLookup $appLookup)

        $roots.Count | Should -Be 1
        $roots[0].Name | Should -Be 'LOB App'
        $roots[0].Children.Count | Should -Be 2
        $childNames = $roots[0].Children | ForEach-Object { $_.Name }
        $childNames | Should -Contain 'VC++ Runtime'
        $childNames | Should -Contain '.NET Runtime'
    }
}

# ============================================================================
# Export-AuditCsv
# ============================================================================

Describe 'Export-AuditCsv' {
    It 'writes CSV file with correct columns and rows' {
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add("Name", [string])
        [void]$dt.Columns.Add("Status", [string])
        [void]$dt.Rows.Add("App Alpha", "Healthy")
        [void]$dt.Rows.Add("App Beta", "Expired")

        $csvPath = Join-Path $TestDrive 'export.csv'
        $logFile = Join-Path $TestDrive 'csv.log'
        Initialize-Logging -LogPath $logFile

        Export-AuditCsv -DataTable $dt -OutputPath $csvPath

        Test-Path -LiteralPath $csvPath | Should -BeTrue
        $rows = Import-Csv -LiteralPath $csvPath
        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be 'App Alpha'
        $rows[1].Status | Should -Be 'Expired'
    }

    It 'creates parent directories if missing' {
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add("Col1", [string])
        [void]$dt.Rows.Add("val")

        $csvPath = Join-Path $TestDrive 'deep\sub\export.csv'
        $logFile = Join-Path $TestDrive 'csv2.log'
        Initialize-Logging -LogPath $logFile

        Export-AuditCsv -DataTable $dt -OutputPath $csvPath
        Test-Path -LiteralPath $csvPath | Should -BeTrue
    }
}

# ============================================================================
# Export-AuditHtml
# ============================================================================

Describe 'Export-AuditHtml' {
    It 'writes valid HTML file with table' {
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add("App", [string])
        [void]$dt.Columns.Add("Severity", [string])
        [void]$dt.Rows.Add("Widget", "Error")
        [void]$dt.Rows.Add("Gadget", "Healthy")

        $htmlPath = Join-Path $TestDrive 'report.html'
        $logFile = Join-Path $TestDrive 'html.log'
        Initialize-Logging -LogPath $logFile

        Export-AuditHtml -DataTable $dt -OutputPath $htmlPath -ReportTitle 'Test Report'

        Test-Path -LiteralPath $htmlPath | Should -BeTrue
        $content = Get-Content -LiteralPath $htmlPath -Raw
        $content | Should -Match 'Test Report'
        $content | Should -Match '<title>'
        $content | Should -Match '<th>App</th>'
        $content | Should -Match '<th>Severity</th>'
        $content | Should -Match 'Widget'
        $content | Should -Match 'class="error"'
        $content | Should -Match 'class="healthy"'
    }
}

# ============================================================================
# New-AuditSummaryText
# ============================================================================

Describe 'New-AuditSummaryText' {
    It 'returns formatted summary string' {
        $counts = [PSCustomObject]@{
            AppCount           = 900
            SupersedenceTotal  = 45
            SupersedenceBroken = 3
            DependencyTotal    = 120
            DependencyBroken   = 7
            BrokenRulesTotal   = 12
            BrokenErrors       = 5
            BrokenWarnings     = 4
            BrokenInfo         = 3
        }

        $summary = New-AuditSummaryText -Counts $counts

        $summary | Should -Match '900 scanned'
        $summary | Should -Match '45 rules \(3 broken\)'
        $summary | Should -Match '120 rules \(7 broken\)'
        $summary | Should -Match '12 total \(5 errors, 4 warnings, 3 info\)'
    }
}

# ============================================================================
# Empty data edge cases
# ============================================================================

Describe 'Empty data handling' {
    BeforeAll {
        $logFile = Join-Path $TestDrive 'empty.log'
        Initialize-Logging -LogPath $logFile
    }

    It 'Find-SupersedenceChains handles no relationships' {
        $result = @(Find-SupersedenceChains -ResolvedRelationships @() -AppLookup @{})
        $result.Count | Should -Be 0
    }

    It 'Find-DependencyGroups handles no relationships' {
        $result = @(Find-DependencyGroups -ResolvedRelationships @() -AppLookup @{})
        $result.Count | Should -Be 0
    }

    It 'Find-BrokenSupersedence handles no data' {
        $result = @(Find-BrokenSupersedence -SupersedenceData @() -ResolvedRelationships @() -AppLookup @{})
        $result.Count | Should -Be 0
    }

    It 'Find-BrokenDependencies handles no data' {
        $result = @(Find-BrokenDependencies -DependencyData @() -ResolvedRelationships @() -AppLookup @{})
        $result.Count | Should -Be 0
    }

    It 'Build-SupersedenceTree handles no data' {
        $result = @(Build-SupersedenceTree -SupersedenceData @() -AppLookup @{})
        $result.Count | Should -Be 0
    }

    It 'Build-DependencyTree handles no data' {
        $result = @(Build-DependencyTree -DependencyData @() -AppLookup @{})
        $result.Count | Should -Be 0
    }

    It 'Get-ScanSummaryCounts handles all zeros' {
        $counts = Get-ScanSummaryCounts -AppCount 0 -SupersedenceData @() -DependencyData @() -BrokenRules @()
        $counts.AppCount | Should -Be 0
        $counts.SupersedenceTotal | Should -Be 0
        $counts.DependencyTotal | Should -Be 0
        $counts.BrokenRulesTotal | Should -Be 0
    }
}
