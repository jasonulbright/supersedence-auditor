# Supersedence & Dependency Auditor

A WinForms-based PowerShell GUI for auditing all application supersedence and dependency relationships in an MECM (Configuration Manager) environment. Discovers relationships via bulk WMI queries, detects broken rules, and visualizes hierarchies in a tree view.

![Supersedence & Dependency Auditor](screenshot.png)

## Requirements

- Windows 10/11
- PowerShell 5.1
- .NET Framework 4.8+
- Configuration Manager console installed (provides the ConfigurationManager PowerShell module)
- WMI access to the SMS Provider server

## Quick Start

```powershell
powershell -ExecutionPolicy Bypass -File start-supersedenceauditor.ps1
```

1. Open **File > Preferences** and set your Site Code and SMS Provider
2. Click **Scan Environment**
3. Review the four tabs: Supersedence, Dependencies, Broken Rules, Tree View

## Features

### Bulk WMI Scan

Three WMI queries load the entire relationship graph in seconds:

- All applications (`SMS_Application WHERE IsLatest = TRUE`)
- All deployment types (`SMS_ConfigurationItemLatestBaseClass WHERE CIType_ID = 21`)
- All relationships (`SMS_AppRelation_Flat`)

CI_IDs are resolved to friendly names via O(1) hashtable lookups.

### Supersedence Tab

Lists all supersedence relationships with:

- Superseding and superseded application names and versions
- Chain depth (how many levels deep the supersedence extends)
- Status (Healthy, Orphaned, Circular, Expired Target, Disabled Source)
- Detail panel with full app metadata on row selection

### Dependencies Tab

Lists all dependency relationships with:

- Parent and dependency application names and versions
- Dependency type (Required, Optional, App Dependency)
- Relationship level
- Status (Healthy, Orphaned, Expired Target, Disabled Target, Missing Content)
- Detail panel on row selection

### Broken Rules Tab

Unified view of all detected issues across both relationship types:

| Issue Type | Category | Severity | Description |
|---|---|---|---|
| Orphaned Reference | Both | Error | Relationship references a deleted application |
| Circular Chain | Supersedence | Error | Circular supersedence loop detected |
| Circular Dependency | Dependency | Error | Circular dependency loop detected |
| Expired Target | Both | Warning | Target application is retired/expired |
| Disabled Source | Supersedence | Warning | Superseding application is disabled |
| Disabled Target | Dependency | Warning | Dependency target is disabled |
| Missing Content | Dependency | Error | Dependency target has no distributed content |
| Undocumented | Both | Info | App has relationships but no Manufacturer set |

Each broken rule includes a remediation description in the detail panel.

### Tree View Tab

Hierarchical visualization with two root nodes:

- **Supersedence Chains** -- expand to see newest-to-oldest supersedence hierarchy
- **Dependency Trees** -- expand to see what each app depends on

Nodes are color-coded: red for expired, yellow for disabled, default for active. Click any node to see full application details.

### Filtering

- Text filter searches across app names and descriptions
- Status filter: All, Healthy, Broken/Warning, Error
- Filters apply to the currently active tab

### Export

- **CSV** -- export the active tab's data to CSV
- **HTML** -- styled report with color-coded severity columns
- **Copy Summary** -- plain text summary to clipboard (app count, rule counts, broken counts)

## Project Structure

```
supersedenceauditor/
├── start-supersedenceauditor.ps1          # WinForms GUI
├── Module/
│   ├── SupersedenceAuditorCommon.psd1     # Module manifest
│   ├── SupersedenceAuditorCommon.psm1     # Business logic (21 functions)
│   └── SupersedenceAuditorCommon.Tests.ps1 # Pester 5.x tests (46 tests)
├── Logs/                                   # Session logs
├── Reports/                                # Export output
├── SupersedenceAuditor.prefs.json         # User preferences
├── SupersedenceAuditor.windowstate.json   # Window state persistence
├── CHANGELOG.md
├── LICENSE
└── README.md
```

## Module Functions

### Logging
- `Initialize-Logging` -- create timestamped log file
- `Write-Log` -- severity-tagged log messages (INFO/WARN/ERROR)

### CM Connection
- `Connect-CMSite` -- import CM module, create PSDrive, verify connection
- `Disconnect-CMSite` -- restore original location
- `Test-CMConnection` -- check if connected

### WMI Discovery
- `Get-AllApplicationSummary` -- bulk load all apps into hashtable
- `Get-AllDeploymentTypeSummary` -- bulk load all DTs into hashtable
- `Get-AllRelationships` -- bulk load all SMS_AppRelation_Flat records
- `Resolve-RelationshipData` -- join CI_IDs with friendly names

### Analysis
- `Find-SupersedenceChains` -- extract and analyze supersedence pairs
- `Find-DependencyGroups` -- extract and classify dependencies
- `Find-BrokenSupersedence` -- detect broken supersedence rules
- `Find-BrokenDependencies` -- detect broken dependency rules
- `Find-UndocumentedRelationships` -- find poorly documented relationship participants
- `Get-ScanSummaryCounts` -- aggregate counts for summary cards

### Tree Building
- `Build-SupersedenceTree` -- nested hierarchy for supersedence chains
- `Build-DependencyTree` -- nested hierarchy for dependency trees

### Export
- `Export-AuditCsv` -- DataTable to CSV
- `Export-AuditHtml` -- DataTable to styled HTML report
- `New-AuditSummaryText` -- plain text summary for clipboard

## Tests

46 Pester 5.x tests covering all module functions. No MECM, WMI, or admin elevation required.

```powershell
cd Module
Invoke-Pester .\SupersedenceAuditorCommon.Tests.ps1
```

## License

See [LICENSE](LICENSE) file.

## Author

Jason Ulbright
