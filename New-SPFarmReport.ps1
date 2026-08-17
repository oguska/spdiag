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
    [switch]$SkipIisDetails,

    [Parameter()]
    [ValidateSet('en-US', 'tr-TR')]
    [string]$Language = 'en-US',

    [Parameter()]
    [string]$LatestKnownSharePointBuild,

    [Parameter()]
    [string]$LatestKnownSharePointUpdateName,

    [Parameter()]
    [switch]$SkipMicrosoftUpdateCheck,

    [Parameter()]
    [string]$UpdateCachePath = (Join-Path -Path $env:ProgramData -ChildPath 'SPFarmReport\SharePointUpdatesCache.json'),

    [Parameter()]
    [ValidateRange(1, 365)]
    [int]$UpdateCacheMaxAgeDays = 30,

    [Parameter()]
    [switch]$ForceMicrosoftUpdateRefresh
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:ReportSections = New-Object System.Collections.ArrayList
$script:SummaryItems = New-Object System.Collections.ArrayList

$script:Translations = @{
    'en-US' = @{
        ReportTitle = 'SharePoint Farm Diagnostic Report'
        Generated = 'Generated'
        Server = 'Server'
        User = 'User'
        ExpandAll = 'Expand all'
        CollapseAll = 'Collapse all'
        Items = 'item(s)'
        NoData = 'No data returned.'
        Product = 'Product'
        Build = 'Build'
        Servers = 'Servers'
        WebApplications = 'Web Applications'
        ContentDatabases = 'Content Databases'
        FarmVersionTitle = 'Farm Version'
        FarmVersionDescription = 'Detected farm version and configuration database.'
        FarmOverviewTitle = 'Farm Overview'
        FarmOverviewDescription = 'Core farm identity, status, timer service identity, and Central Administration URL.'
        ServersTitle = 'Servers'
        ServersDescription = 'Servers joined to the farm and version/role indicators.'
        FarmServerUpdatesTitle = 'Farm Server Update Status'
        FarmServerUpdatesDescription = 'SharePoint build, upgrade state, and latest installed Windows update indicators per farm server. If Microsoft access is available, the latest SharePoint update is checked automatically.'
        UpdateCacheHistoryTitle = 'Cached SharePoint Update History'
        UpdateCacheHistoryDescription = 'Previously cached Microsoft SharePoint CU/security update rows for the detected product. This cache is used when Microsoft access is unavailable or skipped.'
        ServicesTitle = 'Services On Servers'
        ServicesDescription = 'SharePoint service instances and their status on farm servers.'
        WebAppsTitle = 'Web Applications'
        WebAppsDescription = 'Web application configuration, authentication, application pools, and content database counts.'
        ContentDbTitle = 'Content Databases'
        ContentDbDescription = 'Content database sizing, site counts, and upgrade indicators.'
        LocalHealthTitle = 'Local Server Health'
        LocalHealthDescription = 'Operating system, memory, processor, PowerShell version, and all attached logical disk space analysis on the server where the script ran.'
    }
    'tr-TR' = @{
        ReportTitle = 'SharePoint Farm Tanılama Raporu'
        Generated = 'Oluşturulma'
        Server = 'Sunucu'
        User = 'Kullanıcı'
        ExpandAll = 'Tümünü genişlet'
        CollapseAll = 'Tümünü daralt'
        Items = 'öğe'
        NoData = 'Veri dönmedi.'
        Product = 'Ürün'
        Build = 'Derleme'
        Servers = 'Sunucular'
        WebApplications = 'Web Uygulamaları'
        ContentDatabases = 'İçerik Veritabanları'
        FarmVersionTitle = 'Farm Sürümü'
        FarmVersionDescription = 'Algılanan farm sürümü ve yapılandırma veritabanı.'
        FarmOverviewTitle = 'Farm Özeti'
        FarmOverviewDescription = 'Temel farm kimliği, durumu, zamanlayıcı servis hesabı ve Central Administration URL bilgisi.'
        ServersTitle = 'Sunucular'
        ServersDescription = 'Farma bağlı sunucular ve sürüm/rol göstergeleri.'
        FarmServerUpdatesTitle = 'Farm Sunucu Güncelleme Durumu'
        FarmServerUpdatesDescription = 'Farm sunucuları için SharePoint derlemesi, yükseltme durumu ve en son yüklü Windows güncelleme göstergeleri. Microsoft erişimi varsa en güncel SharePoint güncellemesi otomatik kontrol edilir.'
        UpdateCacheHistoryTitle = 'Önbelleğe Alınmış SharePoint Güncelleme Geçmişi'
        UpdateCacheHistoryDescription = 'Algılanan ürün için daha önce önbelleğe alınmış Microsoft SharePoint CU/güvenlik güncelleme satırları. Microsoft erişimi yoksa veya atlanırsa bu önbellek kullanılır.'
        ServicesTitle = 'Sunuculardaki Servisler'
        ServicesDescription = 'Farm sunucularındaki SharePoint servis örnekleri ve durumları.'
        WebAppsTitle = 'Web Uygulamaları'
        WebAppsDescription = 'Web uygulaması yapılandırması, kimlik doğrulama, uygulama havuzları ve içerik veritabanı sayıları.'
        ContentDbTitle = 'İçerik Veritabanları'
        ContentDbDescription = 'İçerik veritabanı boyutları, site sayıları ve yükseltme göstergeleri.'
        LocalHealthTitle = 'Yerel Sunucu Sağlığı'
        LocalHealthDescription = 'Betiğin çalıştığı sunucuda işletim sistemi, bellek, işlemci, PowerShell sürümü ve tüm bağlı mantıksal disk alanı analizi.'
    }
}

function Get-ReportText {
    param(
        [string]$Key,
        [string]$Default = $Key
    )

    if ($script:Translations.ContainsKey($Language) -and $script:Translations[$Language].ContainsKey($Key)) {
        return $script:Translations[$Language][$Key]
    }

    if ($script:Translations['en-US'].ContainsKey($Key)) {
        return $script:Translations['en-US'][$Key]
    }

    return $Default
}

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

function ConvertTo-VersionNumber {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return 0 }
    $digits = ([string]$Value) -replace '[^0-9]', ''
    if (-not $digits) { return 0 }

    $number = 0L
    if ([long]::TryParse($digits, [ref]$number)) { return $number }
    return 0
}

function Get-VersionParts {
    param([AllowNull()][object]$Value)

    $parts = @([string]$Value -split '\.' | ForEach-Object {
        $number = 0
        if ([int]::TryParse($_, [ref]$number)) { $number } else { 0 }
    })

    while ($parts.Count -lt 4) { $parts += 0 }
    return $parts
}

function Get-SPProductUpdateSectionName {
    param([AllowNull()][object]$BuildVersion)

    $parts = Get-VersionParts $BuildVersion
    if ($parts[0] -eq 15) { return 'SharePoint 2013' }
    if ($parts[0] -ne 16) { return '' }

    if ($parts[2] -ge 14000) { return 'SharePoint Server Subscription Edition' }
    if ($parts[2] -ge 10000) { return 'SharePoint 2019' }
    return 'SharePoint 2016'
}

function Get-LatestVersionFromText {
    param([string]$Text)

    $versions = @([regex]::Matches($Text, '\d+\.\d+\.\d+\.\d+') | ForEach-Object { $_.Value })
    if ($versions.Count -eq 0) { return '' }
    return ($versions | Sort-Object { ConvertTo-VersionNumber $_ } -Descending | Select-Object -First 1)
}

function Read-SPReportUpdateCache {
    if (-not $UpdateCachePath) { return $null }
    if (-not (Test-Path -LiteralPath $UpdateCachePath)) { return $null }

    try {
        $json = [System.IO.File]::ReadAllText($UpdateCachePath, [System.Text.Encoding]::UTF8)
        if (-not $json) { return $null }
        return ($json | ConvertFrom-Json)
    }
    catch {
        return $null
    }
}

function Write-SPReportUpdateCache {
    param(
        [string]$Product,
        [object[]]$Updates,
        [string]$Source
    )

    if (-not $UpdateCachePath -or -not $Product -or (Get-ObjectCount $Updates) -eq 0) { return }

    try {
        $cache = Read-SPReportUpdateCache
        $products = @{}
        if ($cache -and (Test-ObjectProperty -InputObject $cache -PropertyName 'Products')) {
            foreach ($property in $cache.Products.PSObject.Properties) {
                $products[$property.Name] = $property.Value
            }
        }

        $products[$Product] = [pscustomobject]@{
            Product       = $Product
            Source        = $Source
            FetchedAt     = (Get-Date).ToString('o')
            UpdateCount   = (Get-ObjectCount $Updates)
            Updates       = @($Updates)
        }

        $cacheObject = [pscustomobject]@{
            SchemaVersion = 1
            CreatedBy     = 'New-SPFarmReport.ps1'
            UpdatedAt     = (Get-Date).ToString('o')
            Products      = [pscustomobject]$products
        }

        $cacheDirectory = Split-Path -Path $UpdateCachePath -Parent
        if ($cacheDirectory -and -not (Test-Path -LiteralPath $cacheDirectory)) {
            New-Item -Path $cacheDirectory -ItemType Directory -Force | Out-Null
        }

        [System.IO.File]::WriteAllText($UpdateCachePath, ($cacheObject | ConvertTo-Json -Depth 8), [System.Text.Encoding]::UTF8)
    }
    catch {
        Write-Verbose ('Failed to write update cache: {0}' -f $_.Exception.Message)
    }
}

function Get-SPReportCachedLatestUpdate {
    param([string]$Product)

    $cache = Read-SPReportUpdateCache
    if (-not $cache -or -not (Test-ObjectProperty -InputObject $cache -PropertyName 'Products')) { return $null }
    $productCache = Get-ObjectValue -InputObject $cache.Products -PropertyName $Product
    if (-not $productCache) { return $null }

    $updates = @(Get-ObjectValue -InputObject $productCache -PropertyName 'Updates')
    if ($updates.Count -eq 0) { return $null }

    $fetchedAtValue = Get-ObjectValue -InputObject $productCache -PropertyName 'FetchedAt'
    $fetchedAt = [datetime]::MinValue
    [void][datetime]::TryParse([string]$fetchedAtValue, [ref]$fetchedAt)
    $ageDays = if ($fetchedAt -gt [datetime]::MinValue) { [math]::Round(((Get-Date) - $fetchedAt).TotalDays, 1) } else { '' }
    $ageStatus = if ($ageDays -ne '' -and $ageDays -gt $UpdateCacheMaxAgeDays) { 'Stale' } else { 'Fresh' }
    $latest = $updates | Sort-Object { ConvertTo-VersionNumber (Get-ObjectValue -InputObject $_ -PropertyName 'LatestBuild') } -Descending | Select-Object -First 1

    return [pscustomobject]@{
        Product        = $Product
        LatestBuild    = Get-ObjectValue -InputObject $latest -PropertyName 'LatestBuild'
        UpdateName     = Get-ObjectValue -InputObject $latest -PropertyName 'UpdateName'
        KB             = Get-ObjectValue -InputObject $latest -PropertyName 'KB'
        ReleaseDate    = Get-ObjectValue -InputObject $latest -PropertyName 'ReleaseDate'
        Source         = ('Cache: {0}' -f $UpdateCachePath)
        LookupStatus   = ('Using cached Microsoft update data ({0}, age {1} day(s))' -f $ageStatus, $ageDays)
        CacheLastRefresh = $fetchedAtValue
        CachedUpdateCount = Get-ObjectValue -InputObject $productCache -PropertyName 'UpdateCount'
    }
}

function Get-SPReportUpdatesFromMicrosoftContent {
    param(
        [string]$Content,
        [string]$Product
    )

    $escapedProduct = [regex]::Escape($Product)
    $sectionMatch = [regex]::Match($Content, "(?is)##\s+$escapedProduct\s+update history(.*?)(\r?\n##\s+|$)")
    if (-not $sectionMatch.Success) { return @() }

    $updates = New-Object System.Collections.ArrayList
    $lines = @($sectionMatch.Groups[1].Value -split "\r?\n")
    foreach ($line in $lines) {
        if ($line -notmatch '^\|') { continue }
        if ($line -match '^\|\s*-') { continue }
        if ($line -match 'Package Name|KB Number|Version|Release Date') { continue }

        $columns = @($line.Trim('|') -split '\|')
        if ($columns.Count -lt 4) { continue }

        $updateName = (($columns[0] -replace '<[^>]+>', '') -replace '\s+', ' ').Trim()
        $kb = (@([regex]::Matches($columns[1], 'KB\s*\d+') | ForEach-Object { ($_.Value -replace '\s+', ' ') }) -join ', ')
        $latestBuild = Get-LatestVersionFromText $columns[2]
        $releaseDate = (($columns[3] -replace '<[^>]+>', '') -replace '\s+', ' ').Trim()

        if ($latestBuild) {
            [void]$updates.Add([pscustomobject]@{
                Product     = $Product
                LatestBuild = $latestBuild
                UpdateName  = $updateName
                KB          = $kb
                ReleaseDate = $releaseDate
            })
        }
    }

    return @($updates)
}

function Get-SPReportMicrosoftLatestSharePointUpdate {
    param([AllowNull()][object]$InstalledBuild)

    $product = Get-SPProductUpdateSectionName $InstalledBuild

    if ($SkipMicrosoftUpdateCheck) {
        $cached = Get-SPReportCachedLatestUpdate -Product $product
        if ($cached) { return $cached }

        return [pscustomobject]@{
            Product        = $product
            LatestBuild    = $LatestKnownSharePointBuild
            UpdateName     = $LatestKnownSharePointUpdateName
            KB             = ''
            ReleaseDate    = ''
            Source         = 'Manual/Skipped Microsoft check'
            LookupStatus   = 'Skipped by -SkipMicrosoftUpdateCheck'
            CacheLastRefresh = ''
            CachedUpdateCount = 0
        }
    }

    if (-not $product) {
        return [pscustomobject]@{
            Product        = ''
            LatestBuild    = $LatestKnownSharePointBuild
            UpdateName     = $LatestKnownSharePointUpdateName
            KB             = ''
            ReleaseDate    = ''
            Source         = 'Manual/Fallback'
            LookupStatus   = 'Could not infer SharePoint product from build.'
            CacheLastRefresh = ''
            CachedUpdateCount = 0
        }
    }

    $cachedBeforeRefresh = Get-SPReportCachedLatestUpdate -Product $product
    if ($cachedBeforeRefresh -and -not $ForceMicrosoftUpdateRefresh) {
        $cacheStatus = Get-ObjectValue -InputObject $cachedBeforeRefresh -PropertyName 'LookupStatus'
        if ($cacheStatus -match 'Fresh') { return $cachedBeforeRefresh }
    }

    $sources = @(
        'https://raw.githubusercontent.com/MicrosoftDocs/OfficeDocs-OfficeUpdates-pr/live/OfficeUpdates/sharepoint-updates.md',
        'https://learn.microsoft.com/en-us/officeupdates/sharepoint-updates'
    )
    $lastError = ''

    foreach ($source in $sources) {
        try {
            $response = Invoke-WebRequest -Uri $source -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop
            $content = [string]$response.Content
            $updates = @(Get-SPReportUpdatesFromMicrosoftContent -Content $content -Product $product)
            if ($updates.Count -eq 0) { continue }

            Write-SPReportUpdateCache -Product $product -Updates $updates -Source $source
            $latest = $updates | Sort-Object { ConvertTo-VersionNumber (Get-ObjectValue -InputObject $_ -PropertyName 'LatestBuild') } -Descending | Select-Object -First 1
            return [pscustomobject]@{
                Product        = $product
                LatestBuild    = Get-ObjectValue -InputObject $latest -PropertyName 'LatestBuild'
                UpdateName     = Get-ObjectValue -InputObject $latest -PropertyName 'UpdateName'
                KB             = Get-ObjectValue -InputObject $latest -PropertyName 'KB'
                ReleaseDate    = Get-ObjectValue -InputObject $latest -PropertyName 'ReleaseDate'
                Source         = $source
                LookupStatus   = ('Success; cached {0} update row(s)' -f $updates.Count)
                CacheLastRefresh = (Get-Date).ToString('o')
                CachedUpdateCount = $updates.Count
            }
        }
        catch {
            $lastError = $_.Exception.Message
        }
    }

    $cachedAfterFailure = Get-SPReportCachedLatestUpdate -Product $product
    if ($cachedAfterFailure) {
        $cachedAfterFailure.LookupStatus = ('Microsoft lookup failed; {0}' -f (Get-ObjectValue -InputObject $cachedAfterFailure -PropertyName 'LookupStatus'))
        return $cachedAfterFailure
    }

    return [pscustomobject]@{
        Product        = $product
        LatestBuild    = $LatestKnownSharePointBuild
        UpdateName     = $LatestKnownSharePointUpdateName
        KB             = ''
        ReleaseDate    = ''
        Source         = 'Manual/Fallback'
        LookupStatus   = if ($lastError) { $lastError } else { 'Microsoft update information could not be parsed.' }
        CacheLastRefresh = ''
        CachedUpdateCount = 0
    }
}

function Get-DiskTypeName {
    param([AllowNull()][object]$DriveType)

    switch ([int]$DriveType) {
        0 { 'Unknown' }
        1 { 'No Root Directory' }
        2 { 'Removable Disk' }
        3 { 'Local Disk' }
        4 { 'Network Drive' }
        5 { 'Compact Disc' }
        6 { 'RAM Disk' }
        default { 'Unknown' }
    }
}

function Get-DiskSpaceStatus {
    param([double]$FreePercent)

    if ($FreePercent -lt 10) { return 'Critical' }
    if ($FreePercent -lt 20) { return 'Warning' }
    return 'Good'
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
    $disks = Get-CimInstance -ClassName Win32_LogicalDisk | Sort-Object DeviceID

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
            DriveType          = Get-DiskTypeName $disk.DriveType
            VolumeName         = $disk.VolumeName
            FileSystem         = $disk.FileSystem
            DriveSize          = Format-ByteSize $disk.Size
            DriveFree          = Format-ByteSize $disk.FreeSpace
            DriveFreePercent   = ('{0:N2}%' -f $freePercent)
            SpaceStatus        = Get-DiskSpaceStatus $freePercent
        }
    }
}

function Get-SPReportFarmServerUpdateStatus {
    $servers = @(Get-SPServer | Sort-Object Name)
    $farmMaxVersionNumber = 0
    foreach ($server in $servers) {
        $versionNumber = ConvertTo-VersionNumber (Get-ObjectValue -InputObject $server -PropertyName 'Version')
        if ($versionNumber -gt $farmMaxVersionNumber) { $farmMaxVersionNumber = $versionNumber }
    }

    $farmMaxVersion = ($servers | Sort-Object { ConvertTo-VersionNumber (Get-ObjectValue -InputObject $_ -PropertyName 'Version') } -Descending | Select-Object -First 1 | ForEach-Object { Get-ObjectValue -InputObject $_ -PropertyName 'Version' })
    $microsoftLatest = Get-SPReportMicrosoftLatestSharePointUpdate -InstalledBuild $farmMaxVersion
    $effectiveLatestBuild = if ($LatestKnownSharePointBuild) { $LatestKnownSharePointBuild } else { Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'LatestBuild' }
    $effectiveLatestUpdateName = if ($LatestKnownSharePointUpdateName) { $LatestKnownSharePointUpdateName } else { Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'UpdateName' }
    $latestKnownVersionNumber = ConvertTo-VersionNumber $effectiveLatestBuild

    foreach ($server in $servers) {
        $serverName = Get-ObjectValue -InputObject $server -PropertyName 'Name'
        $serverVersion = Get-ObjectValue -InputObject $server -PropertyName 'Version'
        $serverVersionNumber = ConvertTo-VersionNumber $serverVersion
        $needsUpgrade = Get-ObjectValue -InputObject $server -PropertyName 'NeedsUpgrade'
        $farmRelativeStatus = 'Current farm maximum'
        $latestKnownStatus = 'Not evaluated'
        $latestUpdate = ''
        $latestSecurityUpdate = ''
        $updateQueryStatus = 'Not queried'

        if ($serverVersionNumber -lt $farmMaxVersionNumber) { $farmRelativeStatus = 'Behind farm maximum' }
        if ($needsUpgrade) { $farmRelativeStatus = 'Needs SharePoint upgrade' }

        if ($latestKnownVersionNumber -gt 0) {
            if ($serverVersionNumber -ge $latestKnownVersionNumber) { $latestKnownStatus = 'At or above latest known build' }
            else { $latestKnownStatus = 'Below latest known build' }
        }

        try {
            $hotfixes = @(Get-HotFix -ComputerName $serverName -ErrorAction Stop | Sort-Object InstalledOn -Descending)
            $latest = $hotfixes | Select-Object -First 1
            $latestSecurity = $hotfixes | Where-Object { $_.Description -match 'Security' } | Select-Object -First 1
            if ($latest) { $latestUpdate = ('{0} {1} {2}' -f $latest.HotFixID, $latest.Description, $latest.InstalledOn) }
            if ($latestSecurity) { $latestSecurityUpdate = ('{0} {1} {2}' -f $latestSecurity.HotFixID, $latestSecurity.Description, $latestSecurity.InstalledOn) }
            $updateQueryStatus = 'Success'
        }
        catch {
            $updateQueryStatus = $_.Exception.Message
        }

        [pscustomobject]@{
            Server                  = $serverName
            Role                    = Get-ObjectValue -InputObject $server -PropertyName 'Role'
            Status                  = Get-ObjectValue -InputObject $server -PropertyName 'Status'
            SharePointBuild         = $serverVersion
            NeedsUpgrade            = $needsUpgrade
            FarmRelativeStatus      = $farmRelativeStatus
            LatestKnownBuild        = $effectiveLatestBuild
            LatestKnownUpdateName   = $effectiveLatestUpdateName
            LatestKnownKB           = Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'KB'
            LatestKnownReleaseDate  = Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'ReleaseDate'
            LatestKnownStatus       = $latestKnownStatus
            MicrosoftLookupProduct  = Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'Product'
            MicrosoftLookupStatus   = Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'LookupStatus'
            MicrosoftLookupSource   = Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'Source'
            CacheLastRefresh        = Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'CacheLastRefresh'
            CachedUpdateCount       = Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'CachedUpdateCount'
            LatestInstalledUpdate   = $latestUpdate
            LatestSecurityUpdate    = $latestSecurityUpdate
            UpdateQueryStatus       = $updateQueryStatus
        }
    }
}

function Get-SPReportCachedSharePointUpdateHistory {
    $servers = @(Get-SPServer | Sort-Object Name)
    $farmMaxVersion = ($servers | Sort-Object { ConvertTo-VersionNumber (Get-ObjectValue -InputObject $_ -PropertyName 'Version') } -Descending | Select-Object -First 1 | ForEach-Object { Get-ObjectValue -InputObject $_ -PropertyName 'Version' })
    $product = Get-SPProductUpdateSectionName $farmMaxVersion
    if (-not $product) { return @() }

    $cache = Read-SPReportUpdateCache
    if (-not $cache -or -not (Test-ObjectProperty -InputObject $cache -PropertyName 'Products')) { return @() }
    $productCache = Get-ObjectValue -InputObject $cache.Products -PropertyName $product
    if (-not $productCache) { return @() }

    $source = Get-ObjectValue -InputObject $productCache -PropertyName 'Source'
    $fetchedAt = Get-ObjectValue -InputObject $productCache -PropertyName 'FetchedAt'
    @(Get-ObjectValue -InputObject $productCache -PropertyName 'Updates') | Sort-Object { ConvertTo-VersionNumber (Get-ObjectValue -InputObject $_ -PropertyName 'LatestBuild') } -Descending | ForEach-Object {
        [pscustomobject]@{
            Product       = Get-ObjectValue -InputObject $_ -PropertyName 'Product'
            LatestBuild   = Get-ObjectValue -InputObject $_ -PropertyName 'LatestBuild'
            UpdateName    = Get-ObjectValue -InputObject $_ -PropertyName 'UpdateName'
            KB            = Get-ObjectValue -InputObject $_ -PropertyName 'KB'
            ReleaseDate   = Get-ObjectValue -InputObject $_ -PropertyName 'ReleaseDate'
            CacheFetchedAt = $fetchedAt
            CacheSource   = $source
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
        return ('<p class="empty">{0}</p>' -f (ConvertTo-HtmlText (Get-ReportText -Key 'NoData')))
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
        [void]$sectionsHtml.AppendLine(('<button class="section-title" type="button"><span>{0}</span><span>{1} {2}</span></button>' -f (ConvertTo-HtmlText $section.Title), $count, (ConvertTo-HtmlText (Get-ReportText -Key 'Items'))))
        [void]$sectionsHtml.AppendLine('<div class="section-body">')
        if ($section.Description) {
            [void]$sectionsHtml.AppendLine(('<p>{0}</p>' -f (ConvertTo-HtmlText $section.Description)))
        }
        [void]$sectionsHtml.AppendLine((Convert-DataToTableHtml -Data $section.Data -Columns $section.Columns))
        [void]$sectionsHtml.AppendLine('</div></section>')
    }

@"
<!doctype html>
<html lang="$Language">
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
<div class="meta"><span>$(ConvertTo-HtmlText (Get-ReportText -Key 'Generated')): $(ConvertTo-HtmlText $generated)</span><span>$(ConvertTo-HtmlText (Get-ReportText -Key 'Server')): $(ConvertTo-HtmlText $computer)</span><span>$(ConvertTo-HtmlText (Get-ReportText -Key 'User')): $(ConvertTo-HtmlText $user)</span></div>
</header>
<main class="container">
<div class="summary">
$($summaryHtml.ToString())
</div>
<div class="toolbar"><button type="button" onclick="setAll(false)">$(ConvertTo-HtmlText (Get-ReportText -Key 'ExpandAll'))</button><button type="button" onclick="setAll(true)">$(ConvertTo-HtmlText (Get-ReportText -Key 'CollapseAll'))</button></div>
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
    if (-not $PSBoundParameters.ContainsKey('ReportTitle')) {
        $ReportTitle = Get-ReportText -Key 'ReportTitle'
    }

    $version = Invoke-SafeCollect -Name 'Farm version' -ScriptBlock { Get-SPReportFarmVersion }
    $farmOverview = Invoke-SafeCollect -Name 'Farm overview' -ScriptBlock { Get-SPReportFarmOverview }
    $servers = Invoke-SafeCollect -Name 'Servers' -ScriptBlock { Get-SPReportServers }
    $webApplications = Invoke-SafeCollect -Name 'Web applications' -ScriptBlock { Get-SPReportWebApplications }
    $contentDatabases = Invoke-SafeCollect -Name 'Content databases' -ScriptBlock { Get-SPReportContentDatabases }
    $servicesOnServers = Invoke-SafeCollect -Name 'Services on servers' -ScriptBlock { Get-SPReportServicesOnServers }
    $farmServerUpdateStatus = Invoke-SafeCollect -Name 'Farm server update status' -ScriptBlock { Get-SPReportFarmServerUpdateStatus }
    $cachedUpdateHistory = Invoke-SafeCollect -Name 'Cached SharePoint update history' -ScriptBlock { Get-SPReportCachedSharePointUpdateHistory }

    $versionValue = if (Test-ObjectProperty -InputObject $version[0] -PropertyName 'BuildVersion') { $version[0].BuildVersion } else { 'Unknown' }
    $productValue = if (Test-ObjectProperty -InputObject $version[0] -PropertyName 'Product') { $version[0].Product } else { 'Unknown' }
    Add-SummaryItem -Name (Get-ReportText -Key 'Product') -Value $productValue -Status 'Unknown'
    Add-SummaryItem -Name (Get-ReportText -Key 'Build') -Value $versionValue -Status 'Unknown'
    Add-SummaryItem -Name (Get-ReportText -Key 'Servers') -Value (@($servers | Where-Object { -not (Test-ObjectProperty -InputObject $_ -PropertyName 'Error') }).Count) -Status (Get-SectionStatus -Data $servers)
    Add-SummaryItem -Name (Get-ReportText -Key 'WebApplications') -Value (@($webApplications | Where-Object { -not (Test-ObjectProperty -InputObject $_ -PropertyName 'Error') }).Count) -Status (Get-SectionStatus -Data $webApplications)
    Add-SummaryItem -Name (Get-ReportText -Key 'ContentDatabases') -Value (@($contentDatabases | Where-Object { -not (Test-ObjectProperty -InputObject $_ -PropertyName 'Error') }).Count) -Status (Get-SectionStatus -Data $contentDatabases)

    Add-ReportSection -Title (Get-ReportText -Key 'FarmVersionTitle') -Description (Get-ReportText -Key 'FarmVersionDescription') -Data $version -Columns @('Product', 'BuildVersion', 'FarmId', 'Configuration', 'Status') -Status (Get-SectionStatus -Data $version)
    Add-ReportSection -Title (Get-ReportText -Key 'FarmOverviewTitle') -Description (Get-ReportText -Key 'FarmOverviewDescription') -Data $farmOverview -Columns @('FarmId', 'Status', 'BuildVersion', 'ConfigurationDatabase', 'ConfigurationDatabaseSize', 'TimerServiceAccount', 'CentralAdministration') -Status (Get-SectionStatus -Data $farmOverview)
    Add-ReportSection -Title (Get-ReportText -Key 'ServersTitle') -Description (Get-ReportText -Key 'ServersDescription') -Data $servers -Columns @('Name', 'Role', 'ServerRole', 'CompliantWithMinRole', 'Address', 'Status', 'NeedsUpgrade', 'Version') -Status (Get-SectionStatus -Data $servers)
    Add-ReportSection -Title (Get-ReportText -Key 'FarmServerUpdatesTitle') -Description (Get-ReportText -Key 'FarmServerUpdatesDescription') -Data $farmServerUpdateStatus -Columns @('Server', 'Role', 'Status', 'SharePointBuild', 'NeedsUpgrade', 'FarmRelativeStatus', 'LatestKnownBuild', 'LatestKnownUpdateName', 'LatestKnownKB', 'LatestKnownReleaseDate', 'LatestKnownStatus', 'MicrosoftLookupProduct', 'MicrosoftLookupStatus', 'MicrosoftLookupSource', 'CacheLastRefresh', 'CachedUpdateCount', 'LatestInstalledUpdate', 'LatestSecurityUpdate', 'UpdateQueryStatus') -Status (Get-SectionStatus -Data $farmServerUpdateStatus)
    Add-ReportSection -Title (Get-ReportText -Key 'UpdateCacheHistoryTitle') -Description (Get-ReportText -Key 'UpdateCacheHistoryDescription') -Data $cachedUpdateHistory -Columns @('Product', 'LatestBuild', 'UpdateName', 'KB', 'ReleaseDate', 'CacheFetchedAt', 'CacheSource') -Status (Get-SectionStatus -Data $cachedUpdateHistory)
    Add-ReportSection -Title (Get-ReportText -Key 'ServicesTitle') -Description (Get-ReportText -Key 'ServicesDescription') -Data $servicesOnServers -Columns @('Server', 'Service', 'Status', 'ServiceType', 'Id') -Status (Get-SectionStatus -Data $servicesOnServers)
    Add-ReportSection -Title (Get-ReportText -Key 'WebAppsTitle') -Description (Get-ReportText -Key 'WebAppsDescription') -Data $webApplications -Columns @('Name', 'Url', 'ApplicationPool', 'ApplicationPoolAccount', 'ClaimsAuthentication', 'AllowAnonymous', 'AuthenticationProvider', 'ContentDatabases', 'MaximumFileSizeMB', 'TimeZone', 'IsCentralAdministration') -Status (Get-SectionStatus -Data $webApplications)
    Add-ReportSection -Title (Get-ReportText -Key 'ContentDbTitle') -Description (Get-ReportText -Key 'ContentDbDescription') -Data $contentDatabases -Columns @('Name', 'WebApplication', 'Server', 'Status', 'CurrentSiteCount', 'WarningSiteCount', 'MaximumSiteCount', 'DiskSizeRequired', 'NeedsUpgrade', 'Id') -Status (Get-SectionStatus -Data $contentDatabases)

    Add-SectionFromCollector -Title (Get-ReportText -Key 'LocalHealthTitle') -Description (Get-ReportText -Key 'LocalHealthDescription') -Collector { Get-SPReportLocalServerHealth } -Columns @('Server', 'OperatingSystem', 'OSVersion', 'LastBootTime', 'PowerShellVersion', 'TotalMemory', 'Processor', 'Drive', 'DriveType', 'VolumeName', 'FileSystem', 'DriveSize', 'DriveFree', 'DriveFreePercent', 'SpaceStatus')
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
