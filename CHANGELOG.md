# Changelog

All notable changes to the Supersedence & Dependency Auditor are documented in this file.

## [1.2.0] - 2026-03-24

### Changed
- **Replaced per-app CM cmdlet calls with in-memory SDMPackageXML parsing** -- eliminates thousands of sequential round-trips to the SMS Provider that caused 30+ minute hangs and 100% MP CPU on environments with 650+ applications
  - `Get-AllApplicationSummary` now uses `Get-CMApplication` (without `-Fast`) to capture `SDMPackageXML` on each app object
  - `Get-AllResolvedRelationships` parses the embedded XML to extract all supersedence and dependency relationships using XPath, replacing `Get-CMDeploymentType`, `Get-CMDeploymentTypeSupersedence`, `Get-CMDeploymentTypeDependencyGroup`, and `Get-CMDeploymentTypeDependency` per-app calls
  - Net effect: 1 bulk query + in-memory parsing instead of ~2,000+ provider round-trips

---

## [1.1.0] - 2026-03-11

### Changed
- **Replaced all direct WMI queries with supported ConfigurationManager PowerShell cmdlets** -- eliminates WS-Management/DCOM connectivity issues (0x80041001) by using the established CM PSDrive instead of raw `Get-CimInstance` calls
  - `Get-AllApplicationSummary` now uses `Get-CMApplication -Fast` (was `Get-CimInstance ... SMS_Application`)
  - New `Get-AllResolvedRelationships` uses `Get-CMDeploymentType`, `Get-CMDeploymentTypeSupersedence`, `Get-CMDeploymentTypeDependencyGroup`, and `Get-CMDeploymentTypeDependency` to discover all relationships
- Module export count reduced from 21 to 18 functions (3 WMI functions consolidated into 1 CM cmdlet function)

### Removed
- `Get-AllDeploymentTypeSummary` -- was bulk WMI query against `SMS_ConfigurationItemLatestBaseClass`; no longer needed
- `Get-AllRelationships` -- was bulk WMI query against `SMS_AppRelation_Flat`; replaced by CM cmdlet pipeline
- `Resolve-RelationshipData` -- was the WMI join/enrichment layer; now handled internally by `Get-AllResolvedRelationships`
- `Resolve-RelationshipData` Pester tests (function removed; downstream analysis tests remain unchanged)

---

## [1.0.2] - 2026-03-04

### Fixed
- Remaining `[uint32]` casts in UI detail panels (lines 1015, 1016, 1127, 1282) -- the v1.0.1 module fix changed hashtable keys to `[int]` but missed 4 lookup casts in the GUI, causing detail panels to silently show nothing on selection

---

## [1.0.1] - 2026-03-03

### Added
- `SupersedenceAuditorCommon.Tests.ps1` -- Pester 5.x test suite (46 tests); covers Write-Log, Initialize-Logging, Resolve-RelationshipData (type filtering, name resolution, orphaned CI_IDs), Find-SupersedenceChains (chain depth, expired target, disabled source, orphaned reference), Find-DependencyGroups (Required/Optional/AppDependence classification, all 4 broken statuses), Find-BrokenSupersedence, Find-BrokenDependencies, Find-UndocumentedRelationships, Get-ScanSummaryCounts, Build-SupersedenceTree, Build-DependencyTree, Export-AuditCsv, Export-AuditHtml, New-AuditSummaryText, and empty data edge cases; uses `$TestDrive` for all file I/O

### Fixed
- Hashtable key type mismatch in module -- changed `[uint32]` casts to `[int]` throughout; PowerShell hashtables treat `[int]100` and `[uint32]100` as different keys, causing CI_ID lookups to fail silently
- Tree View SplitContainer `Panel2MinSize` constraint error -- `SplitterDistance` now set after control is parented (after `Controls.Add`) so actual width is available for validation
- Suppressed unapproved verb warning on module import (`-DisableNameChecking`)

---

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

- **Core module** (`SupersedenceAuditorCommon.psm1`)
  - Structured logging (Initialize-Logging, Write-Log)
  - CM site connection management (Connect-CMSite, Disconnect-CMSite, Test-CMConnection)
  - Data discovery (Get-AllApplicationSummary, Get-AllResolvedRelationships)
  - Analysis (Find-SupersedenceChains, Find-DependencyGroups, Find-BrokenSupersedence, Find-BrokenDependencies, Find-UndocumentedRelationships, Get-ScanSummaryCounts)
  - Tree building (Build-SupersedenceTree, Build-DependencyTree)
  - Export (Export-AuditCsv, Export-AuditHtml, New-AuditSummaryText)

- **Export**: CSV, HTML (color-coded severity), clipboard summary

- **Dark/light theme** with 20+ color variables, custom DarkToolStripRenderer, owner-draw TabControl with ClearTypeGridFit

- **Window state persistence** (position, size, maximized, active tab, splitter distances)

- **Preferences dialog** (File > Preferences) with Dark Mode toggle, Site Code, SMS Provider

- **Menu bar** with File (Preferences, Exit), View (tab switchers), Help (About)

- Row color coding on all grids: red for errors (orphaned, circular, missing content), yellow for warnings (expired, disabled), green/default for healthy
