# SharePoint Farm Diagnostic Report

`New-SPFarmReport.ps1` creates a single offline HTML diagnostic report for SharePoint farms.

Supported SharePoint versions:

- SharePoint 2013
- SharePoint 2016
- SharePoint 2019
- SharePoint Server Subscription Edition

The script is designed for Windows PowerShell 5.1 and the SharePoint Management Shell snap-in.

## Requirements

- Run on a SharePoint farm server.
- Run with an account that has SharePoint farm administration permissions.
- Windows PowerShell 5.1 is recommended.
- The `Microsoft.SharePoint.PowerShell` snap-in must be registered on the server.
- Some sections require local administrator or remote WMI/RPC permissions, especially IIS and remote hotfix inventory.
- Internet access is optional and only used for Microsoft SharePoint update metadata lookup.

## Quick Start

Generate a full report with default settings:

```powershell
.\New-SPFarmReport.ps1
```

Generate to a specific path:

```powershell
.\New-SPFarmReport.ps1 -OutputPath C:\Temp\SPFarmReport.html
```

Generate a Turkish report:

```powershell
.\New-SPFarmReport.ps1 -Language tr-TR -OutputPath C:\Temp\SPFarmReport_TR.html
```

Generate a faster report by skipping expensive sections:

```powershell
.\New-SPFarmReport.ps1 `
  -OutputPath C:\Temp\SPFarmReport.html `
  -SkipSiteCollections `
  -SkipTimerJobs `
  -SkipFeatureInventory `
  -SkipIisDetails
```

## Output

The output is a standalone `.html` file with embedded CSS and JavaScript.

HTML features:

- Offline viewing.
- Expand/collapse all sections.
- Dark/light theme toggle.
- Persistent theme preference through browser `localStorage`.
- Color-coded status cells.
- Mobile-friendly responsive layout.
- Print-friendly CSS.
- Turkish table headers when `-Language tr-TR` is used.

## Parameters

### Core

`-OutputPath`

Path of the generated HTML report. If omitted, a timestamped report is created in the current directory.

`-ReportTitle`

Custom report title. If omitted, the title is localized based on `-Language`.

`-Language`

Report language. Supported values:

- `en-US`
- `tr-TR`

### Skip Switches

By default, the script attempts to collect all sections. Use skip switches for very large farms or restricted environments.

Available skip switches:

- `-SkipSiteCollections`
- `-SkipTimerJobs`
- `-SkipHealthAnalyzer`
- `-SkipSearchTopology`
- `-SkipFeatureInventory`
- `-SkipIisDetails`
- `-SkipMicrosoftUpdateCheck`

Older `-Include...` switches are still accepted for compatibility, but they are no longer required because full scan is the default.

### Microsoft Update Lookup

`-SkipMicrosoftUpdateCheck`

Disables outbound Microsoft update metadata lookup. Cached data is still used when available.

`-ForceMicrosoftUpdateRefresh`

Forces a fresh Microsoft lookup even when cached update metadata is still fresh.

`-LatestKnownSharePointBuild`

Manually supplies the latest SharePoint build for comparison.

`-LatestKnownSharePointUpdateName`

Manually supplies the latest update name for reporting.

`-UpdateCachePath`

Path to the Microsoft SharePoint update metadata cache.

Default:

```powershell
C:\ProgramData\SPFarmReport\SharePointUpdatesCache.json
```

`-UpdateCacheMaxAgeDays`

Number of days cached Microsoft update metadata is considered fresh. Default is `30`.

## Update Cache Examples

Use the default cache path:

```powershell
.\New-SPFarmReport.ps1
```

Place the cache next to the script/report execution folder:

```powershell
.\New-SPFarmReport.ps1 -UpdateCachePath (Join-Path $PWD 'SharePointUpdatesCache.json')
```

Use a custom cache path:

```powershell
.\New-SPFarmReport.ps1 `
  -UpdateCachePath C:\Temp\SharePointUpdatesCache.json `
  -UpdateCacheMaxAgeDays 7
```

Run in an offline farm and reuse cache if available:

```powershell
.\New-SPFarmReport.ps1 -SkipMicrosoftUpdateCheck
```

Force a fresh lookup:

```powershell
.\New-SPFarmReport.ps1 -ForceMicrosoftUpdateRefresh
```

## Report Sections

The report includes these sections when available:

- Farm Version
- Farm Overview
- Servers
- Farm Server Update Status
- Installed Updates On Farm Servers
- Cached SharePoint Update History
- Services On Servers
- Web Applications
- Content Databases
- Local Server Health
- All SharePoint Databases
- Service Applications
- Service Application Proxies
- Managed Accounts
- Alternate Access Mappings
- Farm Solutions
- Blocked File Types
- Usage And Diagnostic Logging
- Outgoing Email
- Site Collections
- Timer Jobs
- Unhealthy Health Analyzer Findings
- Health Analyzer Rules
- Search Topology
- Installed Features
- IIS Application Pools
- IIS Sites

## Update And Patch Behavior

`Farm Server Update Status` uses `Get-SPFarm.BuildVersion` to detect the installed SharePoint product and compare the farm with the latest known SharePoint CU/security update.

The script attempts to fetch latest update metadata from Microsoft Learn:

```text
https://learn.microsoft.com/en-us/officeupdates/sharepoint-updates?view=officeupdates-raw
```

If Microsoft access succeeds:

- The latest build, KB number, release date, and update name are reported.
- All parsed update rows for the detected product are cached.

If Microsoft access fails:

- Cached update data is used when available.
- Manual values from `-LatestKnownSharePointBuild` and `-LatestKnownSharePointUpdateName` are used when provided.
- The lookup failure is shown in the report instead of stopping execution.

Servers with SharePoint role `Invalid`:

- Are excluded from the main `Servers` section.
- Are excluded from Microsoft latest-build evaluation.
- Are excluded from remote `Get-HotFix` checks.
- Are still visible in `Farm Server Update Status` as excluded rows when returned by `Get-SPServer`.

`Installed Updates On Farm Servers` lists all `Get-HotFix` entries for valid farm servers only.

## Health Analyzer Behavior

`Health Analyzer Rules` lists Health Analyzer rules and their configured severity when SharePoint exposes the cmdlets.

`Unhealthy Health Analyzer Findings` reads Central Administration Health Reports when available and includes:

- Title
- Category
- Severity
- Current status
- Explanation
- Remedy
- Possible solution
- Failing servers
- Failing services
- Modified date
- Rule ID

When a Health Analyzer item has a remedy value, that text is reused as the possible solution. If no remedy is available, the report includes fallback remediation guidance.

## Local Server Health

`Local Server Health` reports only local fixed disks on the server where the script runs.

It includes:

- Operating system
- OS version
- Last boot time
- PowerShell version
- Total memory
- Processor
- Local fixed drive size
- Free space
- Free space percentage
- Space status

Disk space status thresholds:

- `Good`: 20% or more free space
- `Warning`: below 20% free space
- `Critical`: below 10% free space

## Localization

Supported languages:

- `en-US`
- `tr-TR`

When `-Language tr-TR` is used:

- Report title is Turkish unless `-ReportTitle` is supplied.
- Main UI buttons are Turkish.
- Section titles and descriptions are Turkish.
- Table headers are Turkish.
- Original property names remain available as table header hover text.

The script is stored as UTF-8 with BOM so Turkish characters render correctly in Windows PowerShell 5.1.

## Security Notes

- Passwords and credential secrets are not exported.
- Managed account names are reported, but secret values are not read or displayed.
- Connection strings and credentials are not intentionally exported.
- The report may contain server names, account names, URLs, database names, patch levels, and farm configuration details. Treat generated reports as sensitive operational data.

## Performance Notes

- Full scan is the default.
- Site collection inventory can be slow on large farms.
- Timer job and feature inventory sections can be large.
- Remote `Get-HotFix` can be slow or fail if RPC/WMI/firewall permissions are restricted.
- IIS sections only report the local server where the script is executed.
- Every section is collected independently. If one section fails, the error is shown in the report and remaining sections continue.

## Troubleshooting

If Microsoft update lookup fails:

- Verify internet access from the SharePoint server.
- Re-run with `-ForceMicrosoftUpdateRefresh`.
- Use `-SkipMicrosoftUpdateCheck` for offline environments.
- Provide manual latest build details with `-LatestKnownSharePointBuild` and `-LatestKnownSharePointUpdateName`.

If Turkish characters look broken:

- Ensure the script file is saved as UTF-8 with BOM.
- Run in Windows PowerShell 5.1 or a console that supports UTF-8 output.

If remote installed updates are missing:

- Verify RPC/WMI firewall access.
- Verify the executing account has permission to query remote hotfixes.
- Check the `UpdateQueryStatus` column in the report.

If a section is empty:

- Confirm the corresponding `-Skip...` switch was not used.
- Confirm the SharePoint cmdlet is available in the installed SharePoint version.
- Confirm the executing account has enough permissions.

## Example Full Command

```powershell
.\New-SPFarmReport.ps1 `
  -Language tr-TR `
  -OutputPath C:\Temp\SPFarmReport_TR.html `
  -UpdateCachePath C:\ProgramData\SPFarmReport\SharePointUpdatesCache.json `
  -UpdateCacheMaxAgeDays 30
```
