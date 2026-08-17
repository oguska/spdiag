# SharePoint Farm Diagnostic Report

`New-SPFarmReport.ps1` creates a single offline HTML diagnostic report for SharePoint 2013, 2016, 2019, and Subscription Edition farms.

## Usage

Run from a SharePoint farm server in Windows PowerShell 5.1 as a farm administrator:

```powershell
.\New-SPFarmReport.ps1 -OutputPath C:\Temp\SPFarmReport.html
```

By default, the script scans all report sections. Skip expensive sections only when needed:

```powershell
.\New-SPFarmReport.ps1 `
  -OutputPath C:\Temp\SPFarmReport.html `
  -SkipSiteCollections `
  -SkipFeatureInventory
```

Generate a Turkish report:

```powershell
.\New-SPFarmReport.ps1 -Language tr-TR -OutputPath C:\Temp\SPFarmReport_TR.html
```

Compare farm server builds to a known latest SharePoint build that you provide:

```powershell
.\New-SPFarmReport.ps1 `
  -LatestKnownSharePointBuild 16.0.10417.20018 `
  -LatestKnownSharePointUpdateName "Example CU/Security Update Name"
```

By default, the script tries to fetch the latest SharePoint update metadata from Microsoft when internet access is available. Disable this for isolated/offline farms:

```powershell
.\New-SPFarmReport.ps1 -SkipMicrosoftUpdateCheck
```

Microsoft update metadata is cached locally after a successful lookup. The cache stores previous CU/security update rows for the detected SharePoint product and is reused when Microsoft access is unavailable or skipped:

```powershell
.\New-SPFarmReport.ps1 `
  -UpdateCachePath C:\ProgramData\SPFarmReport\SharePointUpdatesCache.json `
  -UpdateCacheMaxAgeDays 30
```

Force a fresh Microsoft lookup even when the cache is still fresh:

```powershell
.\New-SPFarmReport.ps1 -ForceMicrosoftUpdateRefresh
```

## Notes

- The script loads `Microsoft.SharePoint.PowerShell` automatically when available.
- Each report section is collected independently, so a failed section is shown in the report instead of stopping the whole run.
- Passwords and credential secrets are not exported.
- IIS sections report only the local server where the script is executed.
- Full scan is the default. Use `-SkipSiteCollections`, `-SkipTimerJobs`, `-SkipHealthAnalyzer`, `-SkipSearchTopology`, `-SkipFeatureInventory`, or `-SkipIisDetails` to reduce runtime on very large farms.
- Older `-Include...` switches are still accepted but no longer required because these sections are included by default.
- `Local Server Health` checks local fixed drives only.
- `Farm Server Update Status` shows each server version, but detects the installed product and compares latest CU/security update status using the farm build from `Get-SPFarm.BuildVersion`.
- Use `-LatestKnownSharePointBuild` and `-LatestKnownSharePointUpdateName` to override or provide latest update details manually.
- Use `-SkipMicrosoftUpdateCheck` to avoid outbound Microsoft access attempts.
- Default update cache path is `C:\ProgramData\SPFarmReport\SharePointUpdatesCache.json`.
- The `Cached SharePoint Update History` section lists cached previous CU/security update rows for the detected SharePoint product.
- Supported report languages are `en-US` and `tr-TR`.
