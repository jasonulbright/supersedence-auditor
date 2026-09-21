# Releasing Supersedence and Dependency Auditor

## 1. Test

Run the full suite under Windows PowerShell 5.1 with Pester 5. The release stops if a test fails. Never run the suite under PowerShell 7.

```powershell
powershell -NoProfile -Command "Import-Module Pester -MinimumVersion 5.0; $r = Invoke-Pester -Path Module -PassThru -Output None; 'passed {0} failed {1}' -f $r.PassedCount, $r.FailedCount"
```

## 2. Version

The version is `YYYY.MM.DD.BBBB`: the release date, then a four-digit, zero-padded build number. The build number increases by 1 for each release and never resets. Build 0003 is the first release with this scheme; the 2 releases before it used `1.x.y` numbers. Keep the zero-padded text everywhere; `[version]` drops the leading zeros.

Set the same version in three places:

- `Module/SupersedenceAuditorCommon.psd1`, `ModuleVersion`
- `start-supersedenceauditor.ps1`, header line `Version    : <ver>`
- `CHANGELOG.md`, the top heading `## [<ver>] - <date>`

Add the new `CHANGELOG.md` entry at the top. Do not edit the entries of earlier releases.

## 3. Commit, tag, and package

Commit to `main`. Every commit to `main` is part of a release: the suite installer build refuses a component whose `main` is ahead of its latest tag. Tag the commit `v<ver>` with an annotated tag. Build the zip from the tag. The zip excludes the tests and this file.

```bash
git tag -a v<ver> -m v<ver>
git archive --format=zip -o SupersedenceAuditor-<ver>.zip v<ver> -- . ':(exclude)Tests' ':(exclude)*.Tests.ps1' ':(exclude)RELEASING.md'
sha256sum SupersedenceAuditor-<ver>.zip | sed 's/ \*/  /' > checksums.txt
```

Extract the zip to a temporary folder. Import `Module/SupersedenceAuditorCommon.psd1` under Windows PowerShell 5.1. The import must succeed.

## 4. Publish

Push `main` and the tag. Create the GitHub release with the title `v<ver>` and two assets: the zip and `checksums.txt`. The release must not be a draft. The release notes have a `##` headline with one concrete outcome metric, the `###` sections of the changelog entry, and the footer `Full changelog: CHANGELOG.md`.
