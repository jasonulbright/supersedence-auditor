# Changelog

All notable changes to Supersedence and Dependency Auditor are
documented in this file.

## [1.0.0] - 2026-05-02

Supersedence and Dependency Auditor maps every supersedence and
dependency relationship in an MECM environment from a single bulk
`Get-CMApplication` query plus an in-memory `SDMPackageXML` parse.
Detects orphans, circular chains, expired targets, disabled sources,
missing content, and undocumented apps. Read-only by design: only
`Get-CMSite` and `Get-CMApplication` are called against the SMS
Provider, never any `Set-` / `New-` / `Remove-` / `Add-` cmdlet.

Extract the zip and run `start-supersedenceauditor.ps1`.

### Features

- **Sidebar navigation** across four primary views (Supersedence,
  Dependencies, Broken Rules, Tree View) plus an Options button.
  Theme toggle bottom-docked on the sidebar.
- **Supersedence view** — DataGrid of superseding / superseded app
  pairs with chain depth and status. Detail panel renders full app
  metadata for the selected row.
- **Dependencies view** — parent / dependency pairs by type
  (Required / Optional / App Dependency), level, and status.
- **Broken Rules view** — unified view of every broken supersedence
  and dependency rule with severity, category, description, and
  remediation guidance. Nine issue types: Orphaned Reference,
  Circular Chain, Circular Dependency (DFS cycle detection),
  Expired Target, Disabled Source, Disabled Target, Missing Content,
  Undocumented.
- **Tree view** — TreeView with two root nodes (Supersedence Chains,
  Dependency Trees). Status via glyph on each node; click-to-detail
  side panel with full app properties.
- **Fast discovery** — one `Get-CMApplication` call per site
  retrieves all apps with their `SDMPackageXML`. Relationships
  resolved by XPath against the embedded XML, in memory, with no
  per-app `Get-CMDeploymentType*` round-trips. Scales to thousands
  of apps in seconds.
- **Filter** — text filter across app names plus status filter
  (All / Healthy / Broken or Warning / Error) on every view.
- **Export** — CSV and HTML (glyph-based severity, no color-as-
  status per brand) plus a plain-text clipboard summary.
- **MahApps Dark.Steel / Light.Blue themes** with live swap (no
  restart). Theme preference persisted across sessions.
- **Async progress overlay** during scan. ProgressRing + per-step
  status text driven by a background STA runspace and a
  DispatcherTimer; the UI stays responsive while the bulk query
  and SDMPackageXML parse run.
- **Title-bar drag fallback** — native `WM_NCHITTEST` HwndSource
  hook + managed `DragMove` for the main window and modal Options
  dialog so the title bar drags reliably under any host.
- **Crash handlers** on both the Dispatcher and the AppDomain
  write full stack traces to a per-launch crash file under `Logs/`.
- **Window state persistence** — size, position, maximized flag,
  and last-active view restored across launches.
- **Status conveyed via glyph, not color.** A leading column on
  every grid renders check / x / warn / ellipsis at theme
  foreground; tree nodes use the same glyph prefix. WCAG-friendly.

### Module

`SupersedenceAuditorCommon.psm1` with 18 exported functions covering
CM connection, data discovery (`Get-AllApplicationSummary`,
`Get-AllResolvedRelationships`), analysis (`Find-SupersedenceChains`,
`Find-DependencyGroups`, `Find-BrokenSupersedence`,
`Find-BrokenDependencies`, `Find-UndocumentedRelationships`,
`Get-ScanSummaryCounts`), tree building, and export.

### Stack

- PowerShell 5.1, .NET Framework 4.7.2+
- MahApps.Metro WPF (vendored DLLs in `Lib/`)
- ConfigurationManager PowerShell module (provided by the MECM
  Console install)
