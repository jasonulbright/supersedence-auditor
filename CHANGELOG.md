# Changelog

All notable changes to the Supersedence & Dependency Auditor are documented in this file.

## [1.0.0] - 2026-03-03

### Added
- **WinForms GUI** (`start-supersedenceauditor.ps1`) for auditing MECM application relationships
  - Header panel with title and subtitle
  - Connection bar with Site Code, SMS Provider labels, and Scan Environment button
  - 4 summary cards: Applications Scanned, Supersedence Rules, Dependency Rules, Broken Rules
  - Text filter and status filter (All, Healthy, Broken/Warning, Error) across all tabs
  - Timestamped log console with scan progress
  - Status bar with connection state, last scan time, and row count

- **4 tabbed views**
  - **Supersedence** -- DataGridView with superseding/superseded app pairs, chain depth, status; detail panel with full app metadata
  - **Dependencies** -- DataGridView with parent/dependency app pairs, type (Required/Optional/App Dependency), level, status; detail panel
  - **Broken Rules** -- unified view of all broken supersedence and dependency rules with severity, category, description, and remediation guidance
  - **Tree View** -- TreeView control with two root nodes (Supersedence Chains, Dependency Trees); color-coded nodes for expired (red), disabled (yellow), and active (default) apps; click-to-detail panel with full app properties

- **Bulk WMI scan** (3 queries total)
  - `SMS_Application WHERE IsLatest = TRUE` for all applications
  - `SMS_ConfigurationItemLatestBaseClass WHERE CIType_ID = 21` for all deployment types
  - `SMS_AppRelation_Flat` for all relationship records
  - O(1) hashtable lookups for CI_ID-to-name resolution

- **Broken rule detection** (9 issue types)
  - Orphaned Reference (supersedence/dependency targets deleted app)
  - Circular Chain / Circular Dependency (DFS cycle detection)
  - Expired Target (superseded/dependency app retired)
  - Disabled Source / Disabled Target (app not enabled)
  - Missing Content (dependency target has no content)
  - Undocumented (app participates in relationships but has no Manufacturer set)

- **Core module** (`SupersedenceAuditorCommon.psm1`) with 21 exported functions
  - Structured logging (Initialize-Logging, Write-Log)
  - CM site connection management (Connect-CMSite, Disconnect-CMSite, Test-CMConnection)
  - WMI bulk discovery (Get-AllApplicationSummary, Get-AllDeploymentTypeSummary, Get-AllRelationships)
  - Relationship resolution (Resolve-RelationshipData)
  - Analysis (Find-SupersedenceChains, Find-DependencyGroups, Find-BrokenSupersedence, Find-BrokenDependencies, Find-UndocumentedRelationships, Get-ScanSummaryCounts)
  - Tree building (Build-SupersedenceTree, Build-DependencyTree)
  - Export (Export-AuditCsv, Export-AuditHtml, New-AuditSummaryText)

- **Export**: CSV, HTML (color-coded severity), clipboard summary

- **Dark/light theme** with 20+ color variables, custom DarkToolStripRenderer, owner-draw TabControl with ClearTypeGridFit

- **Window state persistence** (position, size, maximized, active tab, splitter distances)

- **Preferences dialog** (File > Preferences) with Dark Mode toggle, Site Code, SMS Provider

- **Menu bar** with File (Preferences, Exit), View (tab switchers), Help (About)

- Row color coding on all grids: red for errors (orphaned, circular, missing content), yellow for warnings (expired, disabled), green/default for healthy
