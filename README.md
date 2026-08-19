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
- Some sections require local administrator or remote CIM/WMI/RPC permissions, especially server health, IIS, and remote hotfix inventory.
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

`-IncludeCentralAdminHealthReports`

Reads the Central Administration Health Reports list for detailed unhealthy Health Analyzer findings. This is disabled by default because some farms can take a long time to open or enumerate that list. Without this switch, the findings section uses fast Health Analyzer rule inventory fallback rows with remediation guidance.

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
- Central Administration Patch Status
- Installed Windows Updates On Farm Servers
- Cached SharePoint Update History
- Services On Servers
- Web Applications
- Content Databases
- Local Server Health / Server Health
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
- Current Search Application Crawl Jobs
- Installed Features
- IIS Application Pools
- IIS Sites

## Update And Patch Behavior

`Farm Server Update Status` uses farm-level `Get-SPProduct` server status as the preferred source for per-server SharePoint product action state. `InstallStatus` values such as `NoActionRequired` are reported directly because they show whether SharePoint expects another install/configuration action on that server.

When per-server product versions are available, the script uses the highest installed SharePoint product version per server for the `SharePointBuild` value. If farm-level `Get-SPProduct` only returns server action state, the farm build is used for the build column.

If `Get-SPProduct` cannot return product data for a server, the script falls back to Central Administration `Manage Patch Status`. It queries `/_admin/PatchStatus.aspx`, parses the Server/Product/Version/Install Status table, and uses those rows when available.

`SPServer.Version` and `SPServer.NeedsUpgrade` are not used as SharePoint product patch indicators because they can be stale or misleading for per-server CU status.

Patch health is evaluated from SharePoint product install status and the Microsoft/latest-known build comparison.

The installed SharePoint KB is resolved by matching `SharePointBuild` to the cached Microsoft SharePoint update history. The script also checks `Get-HotFix` for that KB and reports whether Windows Update/Quick Fix Engineering returned it for the server.

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
- Are excluded from `Farm Server Update Status`.
- Are excluded from `Central Administration Patch Status` comparisons.
- Are excluded from Microsoft latest-build evaluation.
- Are excluded from remote `Get-HotFix` checks.

`Installed Windows Updates On Farm Servers` lists all `Get-HotFix` entries for valid farm servers only. This is Windows hotfix inventory and is not the authoritative SharePoint product patch version source.

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

By default, this section uses fast Health Analyzer rule inventory fallback rows with remediation guidance and does not open the Central Administration Health Reports list. Use `-IncludeCentralAdminHealthReports` to read detailed Central Administration findings. When enabled, rows without actual finding data are filtered out, the script reads the 200 most recently modified report items, and it returns up to 50 finding rows. When a Health Analyzer item has a remedy value, that text is reused as the possible solution. If Central Administration does not expose populated finding fields, the section falls back to Health Analyzer rule inventory rows with fallback remediation guidance instead of returning an empty table.

## Server Health

`Local Server Health` is used for single-server farms. `Server Health` is used when the farm has more than one valid SharePoint server.

For multi-server farms, the script probes every valid farm server with CIM and returns an error row for any server that cannot be queried.

It includes:

- Server role
- Probe status
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
- Error details when a server probe fails

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
- Remote server health and `Get-HotFix` queries can be slow or fail if CIM/WMI/RPC/firewall permissions are restricted.
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

If remote server health or installed updates are missing:

- Verify CIM/WMI/RPC firewall access.
- Verify the executing account has permission to query remote servers.
- Check the `ProbeStatus`, `Error`, `Details`, and `UpdateQueryStatus` columns in the report.

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
