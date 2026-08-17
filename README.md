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

## Notes

- The script loads `Microsoft.SharePoint.PowerShell` automatically when available.
- Each report section is collected independently, so a failed section is shown in the report instead of stopping the whole run.
- Passwords and credential secrets are not exported.
- IIS sections report only the local server where the script is executed.
- Full scan is the default. Use `-SkipSiteCollections`, `-SkipTimerJobs`, `-SkipHealthAnalyzer`, `-SkipSearchTopology`, `-SkipFeatureInventory`, or `-SkipIisDetails` to reduce runtime on very large farms.
- Older `-Include...` switches are still accepted but no longer required because these sections are included by default.
