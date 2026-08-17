<#
.SYNOPSIS
    Creates an offline HTML diagnostic report for SharePoint 2013, 2016, 2019, and Subscription Edition farms.

.DESCRIPTION
    Run from a SharePoint farm server as an account with SharePoint farm administration rights.
    The script uses the Microsoft.SharePoint.PowerShell snap-in and continues reporting even when
    individual sections fail because of permissions, unavailable services, or version differences.

.EXAMPLE
    .\New-SPFarmReport.ps1 -OutputPath C:\Temp\SPFarmReport.html

.EXAMPLE
    .\New-SPFarmReport.ps1 -OutputPath C:\Temp\SPFarmReport.html -SkipSiteCollections -SkipFeatureInventory
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath ("SPFarmReport_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ReportTitle = 'SharePoint Farm Diagnostic Report',

    [Parameter()]
    [switch]$IncludeSiteCollections,

    [Parameter()]
    [switch]$IncludeTimerJobs,

    [Parameter()]
    [switch]$IncludeHealthAnalyzer,

    [Parameter()]
    [switch]$IncludeSearchTopology,

    [Parameter()]
    [switch]$IncludeFeatureInventory,

    [Parameter()]
    [switch]$IncludeIisDetails,

    [Parameter()]
    [switch]$SkipSiteCollections,

    [Parameter()]
    [switch]$SkipTimerJobs,

    [Parameter()]
    [switch]$SkipHealthAnalyzer,

    [Parameter()]
    [switch]$SkipSearchTopology,

    [Parameter()]
    [switch]$SkipFeatureInventory,

    [Parameter()]
    [switch]$SkipIisDetails
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:ReportSections = New-Object System.Collections.ArrayList
$script:SummaryItems = New-Object System.Collections.ArrayList

function Add-SummaryItem {
    param(
        [string]$Name,
        [string]$Value,
        [ValidateSet('Good', 'Warning', 'Critical', 'Unknown')]
        [string]$Status = 'Unknown'
    )

    [void]$script:SummaryItems.Add([pscustomobject]@{
        Name   = $Name
        Value  = $Value
        Status = $Status
    })
}

function ConvertTo-HtmlText {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    return [System.Web.HttpUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-DisplayValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss') }
    if ($Value -is [bool]) { return $Value.ToString() }
    if ($Value -is [System.Array]) { return ($Value -join ', ') }
    return [string]$Value
}

function Format-ByteSize {
    param([AllowNull()][object]$Bytes)

    if ($null -eq $Bytes) { return '' }

    $number = 0.0
    if (-not [double]::TryParse([string]$Bytes, [ref]$number)) { return [string]$Bytes }

    if ($number -ge 1TB) { return ('{0:N2} TB' -f ($number / 1TB)) }
    if ($number -ge 1GB) { return ('{0:N2} GB' -f ($number / 1GB)) }
    if ($number -ge 1MB) { return ('{0:N2} MB' -f ($number / 1MB)) }
    if ($number -ge 1KB) { return ('{0:N2} KB' -f ($number / 1KB)) }
    return ('{0:N0} B' -f $number)
}

function Get-ObjectValue {
    param(
        [AllowNull()][object]$InputObject,
        [string]$PropertyName
    )

    if ($null -eq $InputObject) { return $null }
    $member = $InputObject | Get-Member -Name $PropertyName -MemberType Properties -ErrorAction SilentlyContinue
    if ($null -eq $member) { return $null }
    return $InputObject.$PropertyName
}

function Test-ObjectProperty {
    param(
        [AllowNull()][object]$InputObject,
        [string]$PropertyName
    )

    if ($null -eq $InputObject) { return $false }
    return [bool]($InputObject | Get-Member -Name $PropertyName -MemberType Properties -ErrorAction SilentlyContinue)
}

function Get-NestedObjectValue {
    param(
        [AllowNull()][object]$InputObject,
        [string[]]$PropertyPath
    )

    $current = $InputObject
    foreach ($propertyName in $PropertyPath) {
        $current = Get-ObjectValue -InputObject $current -PropertyName $propertyName
        if ($null -eq $current) { return $null }
    }

    return $current
}

function Get-SPReportUserName {
    param([AllowNull()][object]$User)

    foreach ($propertyName in @('UserLogin', 'LoginName', 'Name', 'DisplayName')) {
        $value = Get-ObjectValue -InputObject $User -PropertyName $propertyName
        if ($value) { return $value }
    }

    return ''
}

function Get-SPReportConfigurationDatabase {
    $databases = @(Get-SPDatabase)
    $configDb = $databases | Where-Object { $_.GetType().Name -eq 'SPConfigurationDatabase' } | Select-Object -First 1
    if ($null -eq $configDb) {
        $configDb = $databases | Where-Object { (Get-ObjectValue $_ 'TypeName') -match 'Configuration' } | Select-Object -First 1
    }
    return $configDb
}

function Get-ObjectCount {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return 0 }
    return @($Value).Count
}

function Test-CommandAvailable {
    param([string]$Name)
    return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

function Invoke-SafeCollect {
    param(
        [string]$Name,
        [scriptblock]$ScriptBlock
    )

    Write-Progress -Activity 'Creating SharePoint farm report' -Status $Name
    try {
        $result = & $ScriptBlock
        if ($null -eq $result) { return @() }
        return @($result)
    }
    catch {
        return @([pscustomobject]@{
            Error   = $_.Exception.Message
            Details = $_.ScriptStackTrace
        })
    }
}

function Add-ReportSection {
    param(
        [string]$Title,
        [string]$Description,
        [object[]]$Data,
        [string[]]$Columns,
        [ValidateSet('Good', 'Warning', 'Critical', 'Unknown')]
        [string]$Status = 'Unknown'
    )

    [void]$script:ReportSections.Add([pscustomobject]@{
        Title       = $Title
        Description = $Description
        Data        = @($Data)
        Columns     = $Columns
        Status      = $Status
    })
}

function Get-SectionStatus {
    param([object[]]$Data)

    if (@($Data | Where-Object { Test-ObjectProperty -InputObject $_ -PropertyName 'Error' }).Count -gt 0) { return 'Warning' }
    return 'Good'
}

function Get-SPReportFarmVersion {
    $farm = Get-SPFarm
    $build = $farm.BuildVersion
    $major = [int]$build.Major
    $configDb = Get-SPReportConfigurationDatabase

    $product = switch ($major) {
        15 { 'SharePoint 2013' }
        16 { 'SharePoint 2016 / 2019 / Subscription Edition' }
        default { 'Unknown SharePoint version' }
    }

    return [pscustomobject]@{
        Product       = $product
        BuildVersion  = $build.ToString()
        FarmId        = $farm.Id
        Configuration = Get-ObjectValue -InputObject $configDb -PropertyName 'Name'
        Status        = $farm.Status
    }
}

function Get-SPReportFarmOverview {
    $farm = Get-SPFarm
    $configDb = Get-SPReportConfigurationDatabase
    $timerService = Get-ObjectValue -InputObject $farm -PropertyName 'TimerService'
    $centralAdministration = @(Get-SPWebApplication -IncludeCentralAdministration | Where-Object { Get-ObjectValue -InputObject $_ -PropertyName 'IsAdministrationWebApplication' } | Select-Object -First 1)

    [pscustomobject]@{
        FarmId                    = $farm.Id
        Status                    = $farm.Status
        BuildVersion              = $farm.BuildVersion
        ConfigurationDatabase     = Get-ObjectValue -InputObject $configDb -PropertyName 'Name'
        ConfigurationDatabaseSize = Format-ByteSize (Get-ObjectValue -InputObject $configDb -PropertyName 'DiskSizeRequired')
        TimerServiceAccount       = Get-NestedObjectValue -InputObject $timerService -PropertyPath @('ProcessIdentity', 'Username')
        CentralAdministration     = Get-ObjectValue -InputObject $centralAdministration[0] -PropertyName 'Url'
    }
}

function Get-SPReportServers {
    Get-SPServer | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{
            Name             = $_.Name
            Role             = Get-ObjectValue $_ 'Role'
            Address          = $_.Address
            Status           = $_.Status
            NeedsUpgrade     = $_.NeedsUpgrade
            Version          = $_.Version
            ServerRole       = Get-ObjectValue $_ 'ServerRole'
            CompliantWithMinRole = Get-ObjectValue $_ 'CompliantWithMinRole'
        }
    }
}

function Get-SPReportServicesOnServers {
    Get-SPServiceInstance | Sort-Object Server, TypeName | ForEach-Object {
        $server = Get-ObjectValue -InputObject $_ -PropertyName 'Server'
        [pscustomobject]@{
            Server        = Get-ObjectValue -InputObject $server -PropertyName 'Name'
            Service       = $_.TypeName
            Status        = $_.Status
            Id            = $_.Id
            ServiceType   = $_.GetType().Name
        }
    }
}

function Get-SPReportLocalServerHealth {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $computer = Get-CimInstance -ClassName Win32_ComputerSystem
    $processor = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
    $disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" | Sort-Object DeviceID

    foreach ($disk in $disks) {
        $freePercent = if ($disk.Size -gt 0) { [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2) } else { 0 }
        [pscustomobject]@{
            Server             = $env:COMPUTERNAME
            OperatingSystem    = $os.Caption
            OSVersion          = $os.Version
            LastBootTime       = $os.LastBootUpTime
            PowerShellVersion  = $PSVersionTable.PSVersion.ToString()
            TotalMemory        = Format-ByteSize $computer.TotalPhysicalMemory
            Processor          = $processor.Name
            Drive              = $disk.DeviceID
            DriveSize          = Format-ByteSize $disk.Size
            DriveFree          = Format-ByteSize $disk.FreeSpace
            DriveFreePercent   = ('{0:N2}%' -f $freePercent)
        }
    }
}

function Get-SPReportBlockedFileTypes {
    Get-SPWebApplication -IncludeCentralAdministration | Sort-Object Url | ForEach-Object {
        $webApplication = $_
        $webApplication.BlockedFileExtensions | Sort-Object | ForEach-Object {
            [pscustomobject]@{
                WebApplication = $webApplication.Url
                Extension      = $_
            }
        }
    }
}

function Get-SPReportWebApplications {
    Get-SPWebApplication -IncludeCentralAdministration | Sort-Object Url | ForEach-Object {
        $applicationPool = Get-ObjectValue -InputObject $_ -PropertyName 'ApplicationPool'
        $iisSettings = Get-ObjectValue -InputObject $_ -PropertyName 'IisSettings'
        $authProvider = ''
        if ($iisSettings) {
            $authProvider = @($iisSettings.Values | ForEach-Object {
                $providers = Get-ObjectValue -InputObject $_ -PropertyName 'ClaimsAuthenticationProviders'
                Get-ObjectValue -InputObject $providers -PropertyName 'DisplayName'
            }) -join ', '
        }

        [pscustomobject]@{
            Name                     = $_.Name
            Url                      = $_.Url
            ApplicationPool          = Get-ObjectValue -InputObject $applicationPool -PropertyName 'Name'
            ApplicationPoolAccount   = Get-ObjectValue -InputObject $applicationPool -PropertyName 'Username'
            ClaimsAuthentication     = Get-ObjectValue -InputObject $_ -PropertyName 'UseClaimsAuthentication'
            AllowAnonymous           = Get-ObjectValue -InputObject $_ -PropertyName 'AllowAnonymousAccess'
            AuthenticationProvider   = $authProvider
            ContentDatabases         = Get-ObjectCount (Get-ObjectValue -InputObject $_ -PropertyName 'ContentDatabases')
            MaximumFileSizeMB        = Get-ObjectValue -InputObject $_ -PropertyName 'MaximumFileSize'
            TimeZone                 = Get-NestedObjectValue -InputObject $_ -PropertyPath @('DefaultTimeZone', 'DisplayName')
            IsCentralAdministration  = Get-ObjectValue -InputObject $_ -PropertyName 'IsAdministrationWebApplication'
        }
    }
}

function Get-SPReportContentDatabases {
    Get-SPContentDatabase | Sort-Object WebApplication, Name | ForEach-Object {
        $webApplication = Get-ObjectValue -InputObject $_ -PropertyName 'WebApplication'
        [pscustomobject]@{
            Name              = $_.Name
            WebApplication    = Get-ObjectValue -InputObject $webApplication -PropertyName 'Url'
            Server            = $_.Server
            Status            = $_.Status
            CurrentSiteCount  = $_.CurrentSiteCount
            WarningSiteCount  = $_.WarningSiteCount
            MaximumSiteCount  = $_.MaximumSiteCount
            DiskSizeRequired  = Format-ByteSize $_.DiskSizeRequired
            NeedsUpgrade      = $_.NeedsUpgrade
            Id                = $_.Id
        }
    }
}

function Get-SPReportSiteCollections {
    Get-SPSite -Limit All | Sort-Object Url | ForEach-Object {
        try {
            $rootWeb = Get-ObjectValue -InputObject $_ -PropertyName 'RootWeb'
            $quota = Get-ObjectValue -InputObject $_ -PropertyName 'Quota'
            $usage = Get-ObjectValue -InputObject $_ -PropertyName 'Usage'
            [pscustomobject]@{
                Url                = $_.Url
                Owner              = Get-SPReportUserName (Get-ObjectValue -InputObject $_ -PropertyName 'Owner')
                SecondaryOwner     = Get-SPReportUserName (Get-ObjectValue -InputObject $_ -PropertyName 'SecondaryContact')
                ContentDatabase    = Get-NestedObjectValue -InputObject $_ -PropertyPath @('ContentDatabase', 'Name')
                Template           = Get-ObjectValue -InputObject $rootWeb -PropertyName 'WebTemplate'
                CompatibilityLevel = Get-ObjectValue -InputObject $_ -PropertyName 'CompatibilityLevel'
                StorageUsed        = Format-ByteSize (Get-ObjectValue -InputObject $usage -PropertyName 'Storage')
                StorageQuota       = Format-ByteSize (Get-ObjectValue -InputObject $quota -PropertyName 'StorageMaximumLevel')
                LastContentModifiedDate = Get-ObjectValue -InputObject $_ -PropertyName 'LastContentModifiedDate'
                LockState          = Get-ObjectValue -InputObject $_ -PropertyName 'LockState'
            }
        }
        finally {
            $_.Dispose()
        }
    }
}

function Get-SPReportServiceApplications {
    Get-SPServiceApplication | Sort-Object TypeName, DisplayName | ForEach-Object {
        $applicationPool = Get-ObjectValue -InputObject $_ -PropertyName 'ApplicationPool'
        [pscustomobject]@{
            Name        = $_.DisplayName
            TypeName    = $_.TypeName
            Status      = $_.Status
            Id          = $_.Id
            ApplicationPool = Get-ObjectValue -InputObject $applicationPool -PropertyName 'Name'
        }
    }
}

function Get-SPReportServiceApplicationProxies {
    Get-SPServiceApplicationProxy | Sort-Object TypeName, DisplayName | ForEach-Object {
        [pscustomobject]@{
            Name                = $_.DisplayName
            TypeName            = $_.TypeName
            Status              = $_.Status
            Id                  = $_.Id
            IsConnected         = Get-ObjectValue $_ 'IsConnected'
            ServiceApplication  = Get-ObjectValue $_ 'ServiceApplication'
        }
    }
}

function Get-SPReportManagedAccounts {
    Get-SPManagedAccount | Sort-Object UserName | ForEach-Object {
        [pscustomobject]@{
            UserName               = $_.UserName
            DisplayName            = Get-ObjectValue -InputObject $_ -PropertyName 'DisplayName'
            AutomaticChangeEnabled = Get-ObjectValue -InputObject $_ -PropertyName 'AutomaticChangeEnabled'
            DaysBeforeExpiryToWarn = Get-ObjectValue -InputObject $_ -PropertyName 'DaysBeforeExpiryToWarn'
            PasswordLastChanged    = Get-ObjectValue -InputObject $_ -PropertyName 'PasswordLastChanged'
        }
    }
}

function Get-SPReportAam {
    Get-SPAlternateUrl | Sort-Object WebApplication, Zone, PublicUrl | ForEach-Object {
        $webApplication = Get-ObjectValue -InputObject $_ -PropertyName 'WebApplication'
        [pscustomobject]@{
            WebApplication = Get-ObjectValue -InputObject $webApplication -PropertyName 'DisplayName'
            Zone           = Get-ObjectValue -InputObject $_ -PropertyName 'Zone'
            PublicUrl      = Get-ObjectValue -InputObject $_ -PropertyName 'PublicUrl'
            IncomingUrl    = Get-ObjectValue -InputObject $_ -PropertyName 'IncomingUrl'
            UriScheme      = Get-ObjectValue -InputObject $_ -PropertyName 'UriScheme'
        }
    }
}

function Get-SPReportSolutions {
    Get-SPSolution | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{
            Name             = $_.Name
            Deployed         = $_.Deployed
            ContainsGlobalAssembly = $_.ContainsGlobalAssembly
            ContainsCasPolicy = $_.ContainsCasPolicy
            DeploymentState  = $_.DeploymentState
            LastOperationResult = $_.LastOperationResult
            LastOperationEndTime = $_.LastOperationEndTime
        }
    }
}

function Get-SPReportFeatures {
    Get-SPFeature | Sort-Object Scope, DisplayName | ForEach-Object {
        [pscustomobject]@{
            DisplayName = $_.DisplayName
            Scope       = $_.Scope
            Id          = $_.Id
            CompatibilityLevel = $_.CompatibilityLevel
            Version     = $_.Version
        }
    }
}

function Get-SPReportTimerJobs {
    Get-SPTimerJob | Sort-Object DisplayName | ForEach-Object {
        $webApplication = Get-ObjectValue -InputObject $_ -PropertyName 'WebApplication'
        $server = Get-ObjectValue -InputObject $_ -PropertyName 'Server'
        $schedule = Get-ObjectValue -InputObject $_ -PropertyName 'Schedule'
        [pscustomobject]@{
            Name             = $_.DisplayName
            TypeName         = $_.TypeName
            Schedule         = Get-ObjectValue -InputObject $schedule -PropertyName 'Description'
            IsDisabled       = Get-ObjectValue -InputObject $_ -PropertyName 'IsDisabled'
            LastRunTime      = Get-ObjectValue -InputObject $_ -PropertyName 'LastRunTime'
            WebApplication   = Get-ObjectValue -InputObject $webApplication -PropertyName 'Url'
            Server           = Get-ObjectValue -InputObject $server -PropertyName 'Name'
        }
    }
}

function Get-SPReportHealthAnalyzer {
    if (-not (Test-CommandAvailable -Name Get-SPHealthAnalysisRule)) {
        return [pscustomobject]@{ Error = 'Get-SPHealthAnalysisRule is not available in this environment.'; Details = '' }
    }

    Get-SPHealthAnalysisRule | Sort-Object Category, Summary | ForEach-Object {
        [pscustomobject]@{
            Category     = Get-ObjectValue -InputObject $_ -PropertyName 'Category'
            Summary      = Get-ObjectValue -InputObject $_ -PropertyName 'Summary'
            Severity     = Get-ObjectValue -InputObject $_ -PropertyName 'Severity'
            Enabled      = Get-ObjectValue -InputObject $_ -PropertyName 'Enabled'
            Schedule     = Get-ObjectValue -InputObject $_ -PropertyName 'Schedule'
            RepairAutomatically = Get-ObjectValue -InputObject $_ -PropertyName 'RepairAutomatically'
        }
    }
}

function Get-SPReportSearchTopology {
    if (-not (Test-CommandAvailable -Name Get-SPEnterpriseSearchServiceApplication)) {
        return [pscustomobject]@{ Error = 'Search cmdlets are not available in this environment.'; Details = '' }
    }

    $apps = Get-SPEnterpriseSearchServiceApplication
    foreach ($app in $apps) {
        $topology = Get-SPEnterpriseSearchTopology -SearchApplication $app -Active
        Get-SPEnterpriseSearchComponent -SearchTopology $topology | Sort-Object ServerName, Name | ForEach-Object {
            [pscustomobject]@{
                SearchApplication = $app.Name
                ComponentName     = $_.Name
                ComponentType     = $_.GetType().Name
                ServerName        = $_.ServerName
                RootDirectory     = Get-ObjectValue $_ 'RootDirectory'
                IndexPartition    = Get-ObjectValue $_ 'IndexPartitionOrdinal'
            }
        }
    }
}

function Get-SPReportIisApplicationPools {
    Import-Module WebAdministration -ErrorAction Stop
    Get-ChildItem IIS:\AppPools | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{
            Name         = $_.Name
            State        = $_.State
            Runtime      = $_.managedRuntimeVersion
            PipelineMode = $_.managedPipelineMode
            IdentityType = $_.processModel.identityType
            UserName     = $_.processModel.userName
            Enable32Bit  = $_.enable32BitAppOnWin64
        }
    }
}

function Get-SPReportIisSites {
    Import-Module WebAdministration -ErrorAction Stop
    Get-ChildItem IIS:\Sites | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{
            Name         = $_.Name
            Id           = $_.Id
            State        = $_.State
            PhysicalPath = $_.physicalPath
            Bindings     = ($_.Bindings.Collection | ForEach-Object { $_.protocol + ' ' + $_.bindingInformation }) -join '; '
        }
    }
}

function Get-SPReportDatabases {
    Get-SPDatabase | Sort-Object TypeName, Name | ForEach-Object {
        [pscustomobject]@{
            Name             = $_.Name
            TypeName         = $_.TypeName
            Server           = $_.Server
            Status           = $_.Status
            DiskSizeRequired = Format-ByteSize $_.DiskSizeRequired
            NeedsUpgrade     = $_.NeedsUpgrade
            Id               = $_.Id
        }
    }
}

function Get-SPReportUsageAndLogging {
    $diagnostic = Get-SPDiagnosticConfig
    $usage = Get-SPUsageService

    [pscustomobject]@{
        LogLocation                  = Get-ObjectValue -InputObject $diagnostic -PropertyName 'LogLocation'
        LogDiskSpaceUsageGB          = Get-ObjectValue -InputObject $diagnostic -PropertyName 'LogDiskSpaceUsageGB'
        LogMaxDiskSpaceUsageEnabled  = Get-ObjectValue -InputObject $diagnostic -PropertyName 'LogMaxDiskSpaceUsageEnabled'
        DaysToKeepLogs               = Get-ObjectValue -InputObject $diagnostic -PropertyName 'DaysToKeepLogs'
        UsageServiceStatus           = Get-ObjectValue -InputObject $usage -PropertyName 'Status'
        UsageLogLocation             = Get-ObjectValue -InputObject $usage -PropertyName 'UsageLogLocation'
        UsageLogMaxSpaceGB           = Get-ObjectValue -InputObject $usage -PropertyName 'UsageLogMaxSpaceGB'
        UsageLogCutTime              = Get-ObjectValue -InputObject $usage -PropertyName 'UsageLogCutTime'
    }
}

function Get-SPReportOutgoingEmail {
    $webApp = Get-SPWebApplication -IncludeCentralAdministration | Select-Object -First 1
    $adminWebApp = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local
    $settings = Get-ObjectValue -InputObject $adminWebApp -PropertyName 'OutboundMailServiceInstance'
    $webAppMailService = Get-ObjectValue -InputObject $webApp -PropertyName 'OutboundMailServiceInstance'

    [pscustomobject]@{
        FarmOutboundMailService = Get-NestedObjectValue -InputObject $settings -PropertyPath @('Server', 'Address')
        SampleWebApplication    = Get-ObjectValue -InputObject $webApp -PropertyName 'Url'
        OutboundMailServer      = Get-NestedObjectValue -InputObject $webAppMailService -PropertyPath @('Server', 'Address')
        FromAddress             = Get-ObjectValue -InputObject $webApp -PropertyName 'OutboundMailSenderAddress'
        ReplyToAddress          = Get-ObjectValue -InputObject $webApp -PropertyName 'OutboundMailReplyToAddress'
    }
}

function Convert-DataToTableHtml {
    param(
        [object[]]$Data,
        [string[]]$Columns
    )

    if ($null -eq $Data -or (Get-ObjectCount $Data) -eq 0) {
        return '<p class="empty">No data returned.</p>'
    }

    $hasError = @($Data | Where-Object { Test-ObjectProperty -InputObject $_ -PropertyName 'Error' }).Count -gt 0
    if ($hasError) {
        $Columns = @('Error', 'Details')
    }
    elseif ($null -eq $Columns -or (Get-ObjectCount $Columns) -eq 0) {
        $Columns = @($Data[0] | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name)
    }

    $html = New-Object System.Text.StringBuilder
    [void]$html.AppendLine('<div class="table-wrap"><table>')
    [void]$html.AppendLine('<thead><tr>')
    foreach ($column in $Columns) {
        [void]$html.AppendLine(('<th>{0}</th>' -f (ConvertTo-HtmlText $column)))
    }
    [void]$html.AppendLine('</tr></thead><tbody>')

    foreach ($row in $Data) {
        [void]$html.AppendLine('<tr>')
        foreach ($column in $Columns) {
            $value = ConvertTo-DisplayValue (Get-ObjectValue -InputObject $row -PropertyName $column)
            $class = ''
            if ($column -match 'Status|State|Severity|Error|NeedsUpgrade|IsDisabled') {
                $class = ' class="status-cell"'
            }
            [void]$html.AppendLine(('<td{0}>{1}</td>' -f $class, (ConvertTo-HtmlText $value)))
        }
        [void]$html.AppendLine('</tr>')
    }

    [void]$html.AppendLine('</tbody></table></div>')
    return $html.ToString()
}

function Convert-ReportToHtml {
    param(
        [string]$Title,
        [object[]]$Sections,
        [object[]]$Summary
    )

    $generated = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $computer = $env:COMPUTERNAME
    $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

    $summaryHtml = New-Object System.Text.StringBuilder
    foreach ($item in $Summary) {
        [void]$summaryHtml.AppendLine(('<div class="card {0}"><div class="label">{1}</div><div class="value">{2}</div></div>' -f $item.Status.ToLowerInvariant(), (ConvertTo-HtmlText $item.Name), (ConvertTo-HtmlText $item.Value)))
    }

    $sectionsHtml = New-Object System.Text.StringBuilder
    foreach ($section in $Sections) {
        $count = Get-ObjectCount $section.Data
        [void]$sectionsHtml.AppendLine(('<section class="section {0}">' -f $section.Status.ToLowerInvariant()))
        [void]$sectionsHtml.AppendLine(('<button class="section-title" type="button"><span>{0}</span><span>{1} item(s)</span></button>' -f (ConvertTo-HtmlText $section.Title), $count))
        [void]$sectionsHtml.AppendLine('<div class="section-body">')
        if ($section.Description) {
            [void]$sectionsHtml.AppendLine(('<p>{0}</p>' -f (ConvertTo-HtmlText $section.Description)))
        }
        [void]$sectionsHtml.AppendLine((Convert-DataToTableHtml -Data $section.Data -Columns $section.Columns))
        [void]$sectionsHtml.AppendLine('</div></section>')
    }

@"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$(ConvertTo-HtmlText $Title)</title>
<style>
:root { --bg:#0f172a; --panel:#111827; --panel2:#1f2937; --text:#e5e7eb; --muted:#9ca3af; --good:#22c55e; --warn:#f59e0b; --bad:#ef4444; --unknown:#64748b; --line:#334155; }
* { box-sizing:border-box; }
body { margin:0; font-family:Segoe UI, Arial, sans-serif; background:#0b1120; color:var(--text); }
header { padding:28px 32px; background:linear-gradient(135deg,#172554,#111827 65%); border-bottom:1px solid var(--line); }
h1 { margin:0 0 8px 0; font-size:28px; }
.meta { color:var(--muted); display:flex; gap:18px; flex-wrap:wrap; font-size:13px; }
.container { padding:24px 32px 40px; }
.summary { display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:14px; margin-bottom:24px; }
.card { background:var(--panel); border:1px solid var(--line); border-left-width:5px; padding:14px; border-radius:12px; min-height:82px; }
.card.good { border-left-color:var(--good); } .card.warning { border-left-color:var(--warn); } .card.critical { border-left-color:var(--bad); } .card.unknown { border-left-color:var(--unknown); }
.label { color:var(--muted); font-size:12px; text-transform:uppercase; letter-spacing:.06em; }
.value { font-size:20px; margin-top:8px; word-break:break-word; }
.toolbar { display:flex; gap:10px; margin-bottom:14px; flex-wrap:wrap; }
.toolbar button, .section-title { cursor:pointer; background:var(--panel2); color:var(--text); border:1px solid var(--line); border-radius:10px; padding:10px 13px; }
.section { margin:12px 0; border:1px solid var(--line); border-radius:12px; overflow:hidden; background:var(--panel); }
.section.good { border-left:5px solid var(--good); } .section.warning { border-left:5px solid var(--warn); } .section.critical { border-left:5px solid var(--bad); } .section.unknown { border-left:5px solid var(--unknown); }
.section-title { width:100%; display:flex; justify-content:space-between; align-items:center; border:0; border-radius:0; font-size:16px; font-weight:600; text-align:left; }
.section-body { padding:16px; display:block; }
.section.collapsed .section-body { display:none; }
.section-body p { color:var(--muted); margin-top:0; }
.table-wrap { overflow:auto; border:1px solid var(--line); border-radius:10px; }
table { width:100%; border-collapse:collapse; min-width:760px; }
th, td { padding:9px 10px; border-bottom:1px solid var(--line); text-align:left; vertical-align:top; font-size:13px; }
th { position:sticky; top:0; background:#020617; color:#cbd5e1; z-index:1; }
tr:nth-child(even) td { background:rgba(255,255,255,.025); }
.status-cell { font-weight:600; }
.empty { color:var(--muted); font-style:italic; }
@media (max-width:700px) { header,.container { padding-left:16px; padding-right:16px; } h1 { font-size:22px; } .value { font-size:17px; } }
@media print { body { background:white; color:black; } header,.card,.section,.section-title { background:white; color:black; } .toolbar { display:none; } .section-body { display:block !important; } }
</style>
</head>
<body>
<header>
<h1>$(ConvertTo-HtmlText $Title)</h1>
<div class="meta"><span>Generated: $(ConvertTo-HtmlText $generated)</span><span>Server: $(ConvertTo-HtmlText $computer)</span><span>User: $(ConvertTo-HtmlText $user)</span></div>
</header>
<main class="container">
<div class="summary">
$($summaryHtml.ToString())
</div>
<div class="toolbar"><button type="button" onclick="setAll(false)">Expand all</button><button type="button" onclick="setAll(true)">Collapse all</button></div>
$($sectionsHtml.ToString())
</main>
<script>
function setAll(c){document.querySelectorAll('.section').forEach(function(s){s.classList.toggle('collapsed',c);});}
document.querySelectorAll('.section-title').forEach(function(b){b.addEventListener('click',function(){b.parentElement.classList.toggle('collapsed');});});
document.querySelectorAll('.status-cell').forEach(function(td){var v=td.textContent.toLowerCase(); if(v.indexOf('online')>=0||v.indexOf('succeeded')>=0||v==='false'||v.indexOf('good')>=0){td.style.color='var(--good)';} if(v.indexOf('warning')>=0||v.indexOf('failed')>=0||v.indexOf('error')>=0||v==='true'){td.style.color='var(--warn)';} if(v.indexOf('critical')>=0||v.indexOf('offline')>=0){td.style.color='var(--bad)';}});
</script>
</body>
</html>
"@
}

function Add-SectionFromCollector {
    param(
        [string]$Title,
        [string]$Description,
        [scriptblock]$Collector,
        [string[]]$Columns
    )

    $data = Invoke-SafeCollect -Name $Title -ScriptBlock $Collector
    Add-ReportSection -Title $Title -Description $Description -Data $data -Columns $Columns -Status (Get-SectionStatus -Data $data)
}

try {
    if (-not (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
        if (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -Registered -ErrorAction SilentlyContinue) {
            Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
        }
        else {
            throw 'Microsoft.SharePoint.PowerShell snap-in is not registered. Run this on a SharePoint farm server.'
        }
    }

    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Web')

    $version = Invoke-SafeCollect -Name 'Farm version' -ScriptBlock { Get-SPReportFarmVersion }
    $farmOverview = Invoke-SafeCollect -Name 'Farm overview' -ScriptBlock { Get-SPReportFarmOverview }
    $servers = Invoke-SafeCollect -Name 'Servers' -ScriptBlock { Get-SPReportServers }
    $webApplications = Invoke-SafeCollect -Name 'Web applications' -ScriptBlock { Get-SPReportWebApplications }
    $contentDatabases = Invoke-SafeCollect -Name 'Content databases' -ScriptBlock { Get-SPReportContentDatabases }
    $servicesOnServers = Invoke-SafeCollect -Name 'Services on servers' -ScriptBlock { Get-SPReportServicesOnServers }

    $versionValue = if (Test-ObjectProperty -InputObject $version[0] -PropertyName 'BuildVersion') { $version[0].BuildVersion } else { 'Unknown' }
    $productValue = if (Test-ObjectProperty -InputObject $version[0] -PropertyName 'Product') { $version[0].Product } else { 'Unknown' }
    Add-SummaryItem -Name 'Product' -Value $productValue -Status 'Unknown'
    Add-SummaryItem -Name 'Build' -Value $versionValue -Status 'Unknown'
    Add-SummaryItem -Name 'Servers' -Value (@($servers | Where-Object { -not (Test-ObjectProperty -InputObject $_ -PropertyName 'Error') }).Count) -Status (Get-SectionStatus -Data $servers)
    Add-SummaryItem -Name 'Web Applications' -Value (@($webApplications | Where-Object { -not (Test-ObjectProperty -InputObject $_ -PropertyName 'Error') }).Count) -Status (Get-SectionStatus -Data $webApplications)
    Add-SummaryItem -Name 'Content Databases' -Value (@($contentDatabases | Where-Object { -not (Test-ObjectProperty -InputObject $_ -PropertyName 'Error') }).Count) -Status (Get-SectionStatus -Data $contentDatabases)

    Add-ReportSection -Title 'Farm Version' -Description 'Detected farm version and configuration database.' -Data $version -Columns @('Product', 'BuildVersion', 'FarmId', 'Configuration', 'Status') -Status (Get-SectionStatus -Data $version)
    Add-ReportSection -Title 'Farm Overview' -Description 'Core farm identity, status, timer service identity, and Central Administration URL.' -Data $farmOverview -Columns @('FarmId', 'Status', 'BuildVersion', 'ConfigurationDatabase', 'ConfigurationDatabaseSize', 'TimerServiceAccount', 'CentralAdministration') -Status (Get-SectionStatus -Data $farmOverview)
    Add-ReportSection -Title 'Servers' -Description 'Servers joined to the farm and version/role indicators.' -Data $servers -Columns @('Name', 'Role', 'ServerRole', 'CompliantWithMinRole', 'Address', 'Status', 'NeedsUpgrade', 'Version') -Status (Get-SectionStatus -Data $servers)
    Add-ReportSection -Title 'Services On Servers' -Description 'SharePoint service instances and their status on farm servers.' -Data $servicesOnServers -Columns @('Server', 'Service', 'Status', 'ServiceType', 'Id') -Status (Get-SectionStatus -Data $servicesOnServers)
    Add-ReportSection -Title 'Web Applications' -Description 'Web application configuration, authentication, application pools, and content database counts.' -Data $webApplications -Columns @('Name', 'Url', 'ApplicationPool', 'ApplicationPoolAccount', 'ClaimsAuthentication', 'AllowAnonymous', 'AuthenticationProvider', 'ContentDatabases', 'MaximumFileSizeMB', 'TimeZone', 'IsCentralAdministration') -Status (Get-SectionStatus -Data $webApplications)
    Add-ReportSection -Title 'Content Databases' -Description 'Content database sizing, site counts, and upgrade indicators.' -Data $contentDatabases -Columns @('Name', 'WebApplication', 'Server', 'Status', 'CurrentSiteCount', 'WarningSiteCount', 'MaximumSiteCount', 'DiskSizeRequired', 'NeedsUpgrade', 'Id') -Status (Get-SectionStatus -Data $contentDatabases)

    Add-SectionFromCollector -Title 'Local Server Health' -Description 'Operating system, memory, processor, PowerShell version, and fixed disk free space on the server where the script ran.' -Collector { Get-SPReportLocalServerHealth } -Columns @('Server', 'OperatingSystem', 'OSVersion', 'LastBootTime', 'PowerShellVersion', 'TotalMemory', 'Processor', 'Drive', 'DriveSize', 'DriveFree', 'DriveFreePercent')
    Add-SectionFromCollector -Title 'All SharePoint Databases' -Description 'All SharePoint databases registered in the configuration database.' -Collector { Get-SPReportDatabases } -Columns @('Name', 'TypeName', 'Server', 'Status', 'DiskSizeRequired', 'NeedsUpgrade', 'Id')
    Add-SectionFromCollector -Title 'Service Applications' -Description 'Service applications provisioned in the farm.' -Collector { Get-SPReportServiceApplications } -Columns @('Name', 'TypeName', 'Status', 'ApplicationPool', 'Id')
    Add-SectionFromCollector -Title 'Service Application Proxies' -Description 'Service application proxies and connection status when exposed by the object model.' -Collector { Get-SPReportServiceApplicationProxies } -Columns @('Name', 'TypeName', 'Status', 'IsConnected', 'ServiceApplication', 'Id')
    Add-SectionFromCollector -Title 'Managed Accounts' -Description 'Managed accounts. Password values are never exported.' -Collector { Get-SPReportManagedAccounts } -Columns @('UserName', 'DisplayName', 'AutomaticChangeEnabled', 'DaysBeforeExpiryToWarn', 'PasswordLastChanged')
    Add-SectionFromCollector -Title 'Alternate Access Mappings' -Description 'Public and internal URLs configured for SharePoint zones.' -Collector { Get-SPReportAam } -Columns @('WebApplication', 'Zone', 'PublicUrl', 'IncomingUrl', 'UriScheme')
    Add-SectionFromCollector -Title 'Farm Solutions' -Description 'Farm solution deployment status and last deployment result.' -Collector { Get-SPReportSolutions } -Columns @('Name', 'Deployed', 'ContainsGlobalAssembly', 'ContainsCasPolicy', 'DeploymentState', 'LastOperationResult', 'LastOperationEndTime')
    Add-SectionFromCollector -Title 'Blocked File Types' -Description 'Blocked file extensions configured per web application.' -Collector { Get-SPReportBlockedFileTypes } -Columns @('WebApplication', 'Extension')
    Add-SectionFromCollector -Title 'Usage And Diagnostic Logging' -Description 'ULS diagnostic and usage logging configuration.' -Collector { Get-SPReportUsageAndLogging } -Columns @('LogLocation', 'LogDiskSpaceUsageGB', 'LogMaxDiskSpaceUsageEnabled', 'DaysToKeepLogs', 'UsageServiceStatus', 'UsageLogLocation', 'UsageLogMaxSpaceGB', 'UsageLogCutTime')
    Add-SectionFromCollector -Title 'Outgoing Email' -Description 'Farm/web application outgoing mail indicators. Values may differ per web application.' -Collector { Get-SPReportOutgoingEmail } -Columns @('FarmOutboundMailService', 'SampleWebApplication', 'OutboundMailServer', 'FromAddress', 'ReplyToAddress')

    if (-not $SkipSiteCollections) {
        Add-SectionFromCollector -Title 'Site Collections' -Description 'Site collection inventory. This can be slow on large farms.' -Collector { Get-SPReportSiteCollections } -Columns @('Url', 'Owner', 'SecondaryOwner', 'ContentDatabase', 'Template', 'CompatibilityLevel', 'StorageUsed', 'StorageQuota', 'LastContentModifiedDate', 'LockState')
    }
    else {
        Add-ReportSection -Title 'Site Collections' -Description 'Skipped by -SkipSiteCollections.' -Data @() -Columns @('Url', 'Owner', 'ContentDatabase') -Status 'Unknown'
    }

    if (-not $SkipTimerJobs) {
        Add-SectionFromCollector -Title 'Timer Jobs' -Description 'Timer job schedule and disablement status.' -Collector { Get-SPReportTimerJobs } -Columns @('Name', 'TypeName', 'Schedule', 'IsDisabled', 'LastRunTime', 'WebApplication', 'Server')
    }
    else {
        Add-ReportSection -Title 'Timer Jobs' -Description 'Skipped by -SkipTimerJobs.' -Data @() -Columns @('Name', 'Schedule', 'IsDisabled') -Status 'Unknown'
    }

    if (-not $SkipHealthAnalyzer) {
        Add-SectionFromCollector -Title 'Health Analyzer Rules' -Description 'Health Analyzer rules and configured severity.' -Collector { Get-SPReportHealthAnalyzer } -Columns @('Category', 'Summary', 'Severity', 'Enabled', 'Schedule', 'RepairAutomatically')
    }
    else {
        Add-ReportSection -Title 'Health Analyzer Rules' -Description 'Skipped by -SkipHealthAnalyzer.' -Data @() -Columns @('Category', 'Summary', 'Severity') -Status 'Unknown'
    }

    if (-not $SkipSearchTopology) {
        Add-SectionFromCollector -Title 'Search Topology' -Description 'Active enterprise search topology components.' -Collector { Get-SPReportSearchTopology } -Columns @('SearchApplication', 'ComponentName', 'ComponentType', 'ServerName', 'RootDirectory', 'IndexPartition')
    }
    else {
        Add-ReportSection -Title 'Search Topology' -Description 'Skipped by -SkipSearchTopology.' -Data @() -Columns @('SearchApplication', 'ComponentName', 'ComponentType', 'ServerName') -Status 'Unknown'
    }

    if (-not $SkipFeatureInventory) {
        Add-SectionFromCollector -Title 'Installed Features' -Description 'Installed SharePoint feature definitions. This is a definition inventory, not activation state.' -Collector { Get-SPReportFeatures } -Columns @('DisplayName', 'Scope', 'Id', 'CompatibilityLevel', 'Version')
    }
    else {
        Add-ReportSection -Title 'Installed Features' -Description 'Skipped by -SkipFeatureInventory.' -Data @() -Columns @('DisplayName', 'Scope', 'Id') -Status 'Unknown'
    }

    if (-not $SkipIisDetails) {
        Add-SectionFromCollector -Title 'IIS Application Pools' -Description 'Local IIS application pools on the server where the script ran.' -Collector { Get-SPReportIisApplicationPools } -Columns @('Name', 'State', 'Runtime', 'PipelineMode', 'IdentityType', 'UserName', 'Enable32Bit')
        Add-SectionFromCollector -Title 'IIS Sites' -Description 'Local IIS sites and bindings on the server where the script ran.' -Collector { Get-SPReportIisSites } -Columns @('Name', 'Id', 'State', 'PhysicalPath', 'Bindings')
    }
    else {
        Add-ReportSection -Title 'IIS Details' -Description 'Skipped by -SkipIisDetails.' -Data @() -Columns @('Name', 'State') -Status 'Unknown'
    }

    $html = Convert-ReportToHtml -Title $ReportTitle -Sections @($script:ReportSections) -Summary @($script:SummaryItems)
    $outputDirectory = Split-Path -Path $OutputPath -Parent
    if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.Encoding]::UTF8)
    Write-Progress -Activity 'Creating SharePoint farm report' -Completed
    Write-Host ('Report created: {0}' -f $OutputPath)
}
catch {
    $line = $_.InvocationInfo.ScriptLineNumber
    $command = $_.InvocationInfo.Line
    Write-Error ('{0} Line {1}: {2}' -f $_.Exception.Message, $line, $command)
    exit 1
}
