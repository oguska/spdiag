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
    [switch]$IncludeCentralAdminHealthReports,

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
        ToggleTheme = 'Toggle theme'
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
        ServersDescription = 'Servers joined to the farm and role/configuration indicators.'
        FarmServerUpdatesTitle = 'Farm Server Update Status'
        FarmServerUpdatesDescription = 'SharePoint product/update build and latest installed Windows update indicators per farm server. Get-SPProduct is used first; Central Administration Patch Status is used as fallback.'
        CentralAdminPatchStatusTitle = 'Central Administration Patch Status'
        CentralAdminPatchStatusDescription = 'Product patch status from Central Administration Manage Patch Status.'
        InstalledUpdatesTitle = 'Installed Windows Updates On Farm Servers'
        InstalledUpdatesDescription = 'All Windows hotfix/update entries returned by Get-HotFix for valid farm servers. Servers with role Invalid are excluded.'
        UpdateCacheHistoryTitle = 'Cached SharePoint Update History'
        UpdateCacheHistoryDescription = 'Previously cached Microsoft SharePoint CU/security update rows for the detected product. This cache is used when Microsoft access is unavailable or skipped.'
        ServicesTitle = 'Services On Servers'
        ServicesDescription = 'SharePoint service instances and their status on farm servers.'
        WebAppsTitle = 'Web Applications'
        WebAppsDescription = 'Web application configuration, authentication, application pools, and content database counts.'
        ContentDbTitle = 'Content Databases'
        ContentDbDescription = 'Content database sizing, site counts, and upgrade indicators.'
        LocalHealthTitle = 'Local Server Health'
        LocalHealthDescription = 'Operating system, memory, processor, PowerShell version, and local fixed disk space analysis on the server where the script ran.'
        ServerHealthTitle = 'Server Health'
        ServerHealthDescription = 'Operating system, memory, processor, PowerShell version, and fixed disk space analysis for all valid farm servers.'
        AllDatabasesTitle = 'All SharePoint Databases'
        AllDatabasesDescription = 'All SharePoint databases registered in the configuration database.'
        ServiceApplicationsTitle = 'Service Applications'
        ServiceApplicationsDescription = 'Service applications provisioned in the farm.'
        ServiceApplicationProxiesTitle = 'Service Application Proxies'
        ServiceApplicationProxiesDescription = 'Service application proxies and connection status when exposed by the object model.'
        ManagedAccountsTitle = 'Managed Accounts'
        ManagedAccountsDescription = 'Managed accounts. Password values are never exported.'
        AamTitle = 'Alternate Access Mappings'
        AamDescription = 'Public and internal URLs configured for SharePoint zones.'
        FarmSolutionsTitle = 'Farm Solutions'
        FarmSolutionsDescription = 'Farm solution deployment status and last deployment result.'
        BlockedFileTypesTitle = 'Blocked File Types'
        BlockedFileTypesDescription = 'Blocked file extensions configured per web application.'
        UsageLoggingTitle = 'Usage And Diagnostic Logging'
        UsageLoggingDescription = 'ULS diagnostic and usage logging configuration.'
        OutgoingEmailTitle = 'Outgoing Email'
        OutgoingEmailDescription = 'Farm/web application outgoing mail indicators. Values may differ per web application.'
        SiteCollectionsTitle = 'Site Collections'
        SiteCollectionsDescription = 'Site collection inventory. This can be slow on large farms.'
        TimerJobsTitle = 'Timer Jobs'
        TimerJobsDescription = 'Timer job schedule and disablement status.'
        HealthAnalyzerTitle = 'Health Analyzer Rules'
        HealthAnalyzerDescription = 'Health Analyzer rules and configured severity.'
        HealthAnalyzerFindingsTitle = 'Unhealthy Health Analyzer Findings'
        HealthAnalyzerFindingsDescription = 'Current Health Analyzer findings from Central Administration with explanation, remedy, and possible solution guidance.'
        HealthAnalyzerFallbackSolution = 'Review the rule in Central Administration, validate affected services/servers, then run the rule again after remediation.'
        SearchTopologyTitle = 'Search Topology'
        SearchTopologyDescription = 'Active enterprise search topology components.'
        SearchCrawlJobsTitle = 'Current Search Application Crawl Jobs'
        SearchCrawlJobsDescription = 'Current crawl status for Search Service Application content sources.'
        InstalledFeaturesTitle = 'Installed Features'
        InstalledFeaturesDescription = 'Installed SharePoint feature definitions. This is a definition inventory, not activation state.'
        IisAppPoolsTitle = 'IIS Application Pools'
        IisAppPoolsDescription = 'Local IIS application pools on the server where the script ran.'
        IisSitesTitle = 'IIS Sites'
        IisSitesDescription = 'Local IIS sites and bindings on the server where the script ran.'
        IisDetailsTitle = 'IIS Details'
        SkippedBy = 'Skipped by'
    }
    'tr-TR' = @{
        ReportTitle = 'SharePoint Farm Tanılama Raporu'
        Generated = 'Oluşturulma'
        Server = 'Sunucu'
        User = 'Kullanıcı'
        ExpandAll = 'Tümünü genişlet'
        CollapseAll = 'Tümünü daralt'
        ToggleTheme = 'Tema değiştir'
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
        ServersDescription = 'Farma bağlı sunucular ve rol/yapılandırma göstergeleri.'
        FarmServerUpdatesTitle = 'Farm Sunucu Güncelleme Durumu'
        FarmServerUpdatesDescription = 'Farm sunucuları için SharePoint ürün/güncelleme derlemesi ve en son yüklü Windows güncelleme göstergeleri. Önce Get-SPProduct kullanılır; gerekirse Central Administration Patch Status yedek kaynak olarak kullanılır.'
        CentralAdminPatchStatusTitle = 'Central Administration Yama Durumu'
        CentralAdminPatchStatusDescription = 'Central Administration Manage Patch Status üzerinden ürün yama durumu.'
        InstalledUpdatesTitle = 'Farm Sunucularındaki Yüklü Windows Güncellemeleri'
        InstalledUpdatesDescription = 'Geçerli farm sunucuları için Get-HotFix tarafından döndürülen tüm Windows hotfix/güncelleme kayıtları. Rolü Invalid olan sunucular hariç tutulur.'
        UpdateCacheHistoryTitle = 'Önbelleğe Alınmış SharePoint Güncelleme Geçmişi'
        UpdateCacheHistoryDescription = 'Algılanan ürün için daha önce önbelleğe alınmış Microsoft SharePoint CU/güvenlik güncelleme satırları. Microsoft erişimi yoksa veya atlanırsa bu önbellek kullanılır.'
        ServicesTitle = 'Sunuculardaki Servisler'
        ServicesDescription = 'Farm sunucularındaki SharePoint servis örnekleri ve durumları.'
        WebAppsTitle = 'Web Uygulamaları'
        WebAppsDescription = 'Web uygulaması yapılandırması, kimlik doğrulama, uygulama havuzları ve içerik veritabanı sayıları.'
        ContentDbTitle = 'İçerik Veritabanları'
        ContentDbDescription = 'İçerik veritabanı boyutları, site sayıları ve yükseltme göstergeleri.'
        LocalHealthTitle = 'Yerel Sunucu Sağlığı'
        LocalHealthDescription = 'Betiğin çalıştığı sunucuda işletim sistemi, bellek, işlemci, PowerShell sürümü ve yerel sabit disk alanı analizi.'
        ServerHealthTitle = 'Sunucu Sağlığı'
        ServerHealthDescription = 'Tüm geçerli farm sunucuları için işletim sistemi, bellek, işlemci, PowerShell sürümü ve sabit disk alanı analizi.'
        AllDatabasesTitle = 'Tüm SharePoint Veritabanları'
        AllDatabasesDescription = 'Yapılandırma veritabanına kayıtlı tüm SharePoint veritabanları.'
        ServiceApplicationsTitle = 'Servis Uygulamaları'
        ServiceApplicationsDescription = 'Farm üzerinde sağlanan servis uygulamaları.'
        ServiceApplicationProxiesTitle = 'Servis Uygulaması Proxyleri'
        ServiceApplicationProxiesDescription = 'Servis uygulaması proxyleri ve nesne modeli tarafından sağlanıyorsa bağlantı durumu.'
        ManagedAccountsTitle = 'Yönetilen Hesaplar'
        ManagedAccountsDescription = 'Yönetilen hesaplar. Parola değerleri hiçbir zaman dışa aktarılmaz.'
        AamTitle = 'Alternatif Erişim Eşlemeleri'
        AamDescription = 'SharePoint bölgeleri için yapılandırılmış genel ve iç URL adresleri.'
        FarmSolutionsTitle = 'Farm Çözümleri'
        FarmSolutionsDescription = 'Farm çözüm dağıtım durumu ve son dağıtım sonucu.'
        BlockedFileTypesTitle = 'Engellenen Dosya Türleri'
        BlockedFileTypesDescription = 'Web uygulaması başına yapılandırılmış engellenen dosya uzantıları.'
        UsageLoggingTitle = 'Kullanım ve Tanılama Günlüğü'
        UsageLoggingDescription = 'ULS tanılama ve kullanım günlüğü yapılandırması.'
        OutgoingEmailTitle = 'Giden E-posta'
        OutgoingEmailDescription = 'Farm/web uygulaması giden posta göstergeleri. Değerler web uygulamasına göre değişebilir.'
        SiteCollectionsTitle = 'Site Koleksiyonları'
        SiteCollectionsDescription = 'Site koleksiyonu envanteri. Büyük farmlarda yavaş olabilir.'
        TimerJobsTitle = 'Zamanlayıcı İşleri'
        TimerJobsDescription = 'Zamanlayıcı işi zamanlaması ve devre dışı bırakma durumu.'
        HealthAnalyzerTitle = 'Sağlık Çözümleyici Kuralları'
        HealthAnalyzerDescription = 'Sağlık Çözümleyici kuralları ve yapılandırılmış önem derecesi.'
        HealthAnalyzerFindingsTitle = 'Sağlıksız Sağlık Çözümleyici Bulguları'
        HealthAnalyzerFindingsDescription = 'Central Administration üzerinden alınan geçerli Sağlık Çözümleyici bulguları; açıklama, çözüm ve olası çözüm önerileri ile birlikte.'
        HealthAnalyzerFallbackSolution = 'Kuralı Central Administration üzerinden inceleyin, etkilenen servisleri/sunucuları doğrulayın ve düzeltme sonrası kuralı yeniden çalıştırın.'
        SearchTopologyTitle = 'Arama Topolojisi'
        SearchTopologyDescription = 'Etkin kurumsal arama topolojisi bileşenleri.'
        SearchCrawlJobsTitle = 'Geçerli Arama Uygulaması Tarama İşleri'
        SearchCrawlJobsDescription = 'Arama Servis Uygulaması içerik kaynakları için geçerli tarama durumu.'
        InstalledFeaturesTitle = 'Yüklü Özellikler'
        InstalledFeaturesDescription = 'Yüklü SharePoint özellik tanımları. Bu bir tanım envanteridir, etkinleştirme durumu değildir.'
        IisAppPoolsTitle = 'IIS Uygulama Havuzları'
        IisAppPoolsDescription = 'Betiğin çalıştığı sunucudaki yerel IIS uygulama havuzları.'
        IisSitesTitle = 'IIS Siteleri'
        IisSitesDescription = 'Betiğin çalıştığı sunucudaki yerel IIS siteleri ve bağlamaları.'
        IisDetailsTitle = 'IIS Ayrıntıları'
        SkippedBy = 'Atlandı'
    }
}

$script:ColumnTranslations = @{
    'tr-TR' = @{
        Product = 'Ürün'
        BuildVersion = 'Derleme Sürümü'
        FarmId = 'Farm Kimliği'
        Configuration = 'Yapılandırma'
        Status = 'Durum'
        ConfigurationDatabase = 'Yapılandırma Veritabanı'
        ConfigurationDatabaseSize = 'Yapılandırma Veritabanı Boyutu'
        TimerServiceAccount = 'Zamanlayıcı Servis Hesabı'
        CentralAdministration = 'Merkezi Yönetim'
        Name = 'Ad'
        Role = 'Rol'
        ServerRole = 'Sunucu Rolü'
        CompliantWithMinRole = 'MinRole Uyumlu'
        Address = 'Adres'
        NeedsUpgrade = 'Yükseltme Gerekli'
        Version = 'Sürüm'
        Server = 'Sunucu'
        FarmBuild = 'Farm Derlemesi'
        SharePointBuild = 'SharePoint Derlemesi'
        InstalledSharePointUpdate = 'Yüklü SharePoint Güncellemesi'
        InstalledSharePointKB = 'Yüklü SharePoint KB'
        InstalledSharePointReleaseDate = 'Yüklü SharePoint Yayın Tarihi'
        WindowsSharePointUpdate = 'Windows Update SharePoint Kaydı'
        ProductInstallStatus = 'Ürün Kurulum Durumu'
        PatchStatusSource = 'Yama Durumu Kaynağı'
        InstallStatus = 'Kurulum Durumu'
        Source = 'Kaynak'
        FarmRelativeStatus = 'Farma Göre Durum'
        LatestKnownBuild = 'Bilinen En Güncel Derleme'
        LatestKnownUpdateName = 'Bilinen En Güncel Güncelleme'
        LatestKnownKB = 'Bilinen En Güncel KB'
        LatestKnownReleaseDate = 'Bilinen En Güncel Yayın Tarihi'
        LatestKnownStatus = 'Bilinen Güncellik Durumu'
        MicrosoftLookupProduct = 'Microsoft Arama Ürünü'
        MicrosoftLookupStatus = 'Microsoft Arama Durumu'
        MicrosoftLookupSource = 'Microsoft Arama Kaynağı'
        CacheLastRefresh = 'Önbellek Son Yenileme'
        CachedUpdateCount = 'Önbellekteki Güncelleme Sayısı'
        LatestWindowsUpdate = 'Son Windows Güncellemesi'
        LatestWindowsSecurityUpdate = 'Son Windows Güvenlik Güncellemesi'
        LatestInstalledUpdate = 'Son Yüklü Güncelleme'
        LatestSecurityUpdate = 'Son Güvenlik Güncellemesi'
        UpdateQueryStatus = 'Güncelleme Sorgu Durumu'
        HotFixId = 'Hotfix Kimliği'
        Description = 'Açıklama'
        InstalledBy = 'Yükleyen'
        InstalledOn = 'Yüklenme Tarihi'
        Caption = 'Başlık'
        UpdateType = 'Güncelleme Türü'
        LatestBuild = 'En Güncel Derleme'
        UpdateName = 'Güncelleme Adı'
        KB = 'KB'
        ReleaseDate = 'Yayın Tarihi'
        CacheFetchedAt = 'Önbelleğe Alınma Zamanı'
        CacheSource = 'Önbellek Kaynağı'
        Service = 'Servis'
        ServiceType = 'Servis Türü'
        Id = 'Kimlik'
        Url = 'URL'
        ApplicationPool = 'Uygulama Havuzu'
        ApplicationPoolAccount = 'Uygulama Havuzu Hesabı'
        ClaimsAuthentication = 'Claims Kimlik Doğrulama'
        AllowAnonymous = 'Anonim Erişim'
        AuthenticationProvider = 'Kimlik Doğrulama Sağlayıcısı'
        ContentDatabases = 'İçerik Veritabanları'
        MaximumFileSizeMB = 'Maksimum Dosya Boyutu (MB)'
        TimeZone = 'Saat Dilimi'
        IsCentralAdministration = 'Merkezi Yönetim mi'
        WebApplication = 'Web Uygulaması'
        CurrentSiteCount = 'Geçerli Site Sayısı'
        WarningSiteCount = 'Uyarı Site Sayısı'
        MaximumSiteCount = 'Maksimum Site Sayısı'
        DiskSizeRequired = 'Gerekli Disk Alanı'
        OperatingSystem = 'İşletim Sistemi'
        OSVersion = 'İS Sürümü'
        LastBootTime = 'Son Açılış Zamanı'
        PowerShellVersion = 'PowerShell Sürümü'
        ProbeStatus = 'Sorgu Durumu'
        TotalMemory = 'Toplam Bellek'
        Processor = 'İşlemci'
        Drive = 'Sürücü'
        DriveType = 'Sürücü Türü'
        VolumeName = 'Birim Adı'
        FileSystem = 'Dosya Sistemi'
        DriveSize = 'Sürücü Boyutu'
        DriveFree = 'Boş Alan'
        DriveFreePercent = 'Boş Alan Yüzdesi'
        SpaceStatus = 'Alan Durumu'
        TypeName = 'Tür Adı'
        IsConnected = 'Bağlı mı'
        ServiceApplication = 'Servis Uygulaması'
        UserName = 'Kullanıcı Adı'
        DisplayName = 'Görünen Ad'
        AutomaticChangeEnabled = 'Otomatik Değişim Etkin'
        DaysBeforeExpiryToWarn = 'Süre Dolmadan Uyarı Günü'
        PasswordLastChanged = 'Parola Son Değişim'
        Zone = 'Bölge'
        PublicUrl = 'Genel URL'
        IncomingUrl = 'Gelen URL'
        UriScheme = 'URI Şeması'
        Deployed = 'Dağıtıldı'
        ContainsGlobalAssembly = 'Global Assembly İçerir'
        ContainsCasPolicy = 'CAS İlkesi İçerir'
        DeploymentState = 'Dağıtım Durumu'
        LastOperationResult = 'Son İşlem Sonucu'
        LastOperationEndTime = 'Son İşlem Bitiş Zamanı'
        Extension = 'Uzantı'
        LogLocation = 'Günlük Konumu'
        LogDiskSpaceUsageGB = 'Günlük Disk Kullanımı (GB)'
        LogMaxDiskSpaceUsageEnabled = 'Maksimum Günlük Alanı Etkin'
        DaysToKeepLogs = 'Günlük Saklama Günü'
        UsageServiceStatus = 'Kullanım Servisi Durumu'
        UsageLogLocation = 'Kullanım Günlüğü Konumu'
        UsageLogMaxSpaceGB = 'Kullanım Günlüğü Maksimum Alanı (GB)'
        UsageLogCutTime = 'Kullanım Günlüğü Kesim Zamanı'
        FarmOutboundMailService = 'Farm Giden Posta Servisi'
        SampleWebApplication = 'Örnek Web Uygulaması'
        OutboundMailServer = 'Giden Posta Sunucusu'
        FromAddress = 'Gönderen Adresi'
        ReplyToAddress = 'Yanıt Adresi'
        Owner = 'Sahip'
        SecondaryOwner = 'İkincil Sahip'
        ContentDatabase = 'İçerik Veritabanı'
        Template = 'Şablon'
        CompatibilityLevel = 'Uyumluluk Seviyesi'
        StorageUsed = 'Kullanılan Depolama'
        StorageQuota = 'Depolama Kotası'
        LastContentModifiedDate = 'Son İçerik Değişim Tarihi'
        LockState = 'Kilit Durumu'
        Schedule = 'Zamanlama'
        IsDisabled = 'Devre Dışı mı'
        LastRunTime = 'Son Çalışma Zamanı'
        Category = 'Kategori'
        Summary = 'Özet'
        Severity = 'Önem Derecesi'
        Enabled = 'Etkin'
        RepairAutomatically = 'Otomatik Onar'
        Title = 'Başlık'
        CurrentStatus = 'Geçerli Durum'
        Explanation = 'Açıklama'
        Remedy = 'Çözüm'
        PossibleSolution = 'Olası Çözüm'
        Modified = 'Değiştirilme'
        FailingServers = 'Etkilenen Sunucular'
        FailingServices = 'Etkilenen Servisler'
        RuleId = 'Kural Kimliği'
        SearchApplication = 'Arama Uygulaması'
        ContentSource = 'İçerik Kaynağı'
        CrawlStatus = 'Tarama Durumu'
        CrawlStarted = 'Tarama Başlangıcı'
        CrawlCompleted = 'Tarama Bitişi'
        CrawlDuration = 'Tarama Süresi'
        IncrementalCrawlSchedule = 'Artımlı Tarama Zamanlaması'
        FullCrawlSchedule = 'Tam Tarama Zamanlaması'
        StartAddresses = 'Başlangıç Adresleri'
        SuccessCount = 'Başarılı Sayısı'
        ErrorCount = 'Hata Sayısı'
        DeleteCount = 'Silinen Sayısı'
        ComponentName = 'Bileşen Adı'
        ComponentType = 'Bileşen Türü'
        ServerName = 'Sunucu Adı'
        RootDirectory = 'Kök Dizin'
        IndexPartition = 'Dizin Bölümü'
        Scope = 'Kapsam'
        Runtime = 'Çalışma Zamanı'
        PipelineMode = 'Pipeline Modu'
        IdentityType = 'Kimlik Türü'
        Enable32Bit = '32-bit Etkin'
        PhysicalPath = 'Fiziksel Yol'
        Bindings = 'Bağlamalar'
        Error = 'Hata'
        Details = 'Ayrıntılar'
    }
}

function Get-ColumnHeaderText {
    param([string]$ColumnName)

    if ($script:ColumnTranslations.ContainsKey($Language) -and $script:ColumnTranslations[$Language].ContainsKey($ColumnName)) {
        return $script:ColumnTranslations[$Language][$ColumnName]
    }

    return $ColumnName
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

function ConvertFrom-HtmlFragmentText {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    $text = $text -replace '(?i)<br\s*/?>', ' '
    $text = $text -replace '<[^>]+>', ' '
    $text = [System.Web.HttpUtility]::HtmlDecode($text)
    return (($text -replace '\s+', ' ').Trim())
}

function ConvertFrom-PatchStatusHtml {
    param(
        [AllowNull()][string]$Html,
        [string]$Source
    )

    if (-not $Html) { return @() }

    $rowPattern = '(?is)<tr\b[^>]*>\s*<td\b[^>]*>.*?</td>\s*<td\b[^>]*>\s*(?<server>[A-Za-z0-9_.-]+)\s*</td>\s*<td\b[^>]*>(?<product>.*?)</td>\s*<td\b[^>]*>\s*(?<version>(?:&nbsp;|\d+(?:\.\d+){1,3})?)\s*</td>\s*<td\b[^>]*>\s*(?<status>(?:(?!</td>).)*)\s*</td>\s*</tr>'
    $rows = New-Object System.Collections.ArrayList
    foreach ($match in [regex]::Matches($Html, $rowPattern)) {
        $version = ConvertFrom-HtmlFragmentText $match.Groups['version'].Value
        if ($version -eq [char]160) { $version = '' }

        [void]$rows.Add([pscustomobject]@{
            Server        = ConvertFrom-HtmlFragmentText $match.Groups['server'].Value
            Product       = ConvertFrom-HtmlFragmentText $match.Groups['product'].Value
            Version       = $version
            InstallStatus = ConvertFrom-HtmlFragmentText $match.Groups['status'].Value
            Source        = $Source
        })
    }

    return @($rows)
}

function ConvertFrom-SPProductRecord {
    param(
        [object]$Product,
        [string]$ServerName,
        [string]$Source
    )

    $productName = Get-ObjectValue -InputObject $Product -PropertyName 'DisplayName'
    if (-not $productName) { $productName = Get-ObjectValue -InputObject $Product -PropertyName 'Name' }
    if (-not $productName) { $productName = Get-ObjectValue -InputObject $Product -PropertyName 'ProductName' }
    if (-not $productName) { $productName = Get-ObjectValue -InputObject $Product -PropertyName 'PatchableUnitDisplayName' }
    if (-not $productName) { $productName = ConvertFrom-HtmlFragmentText $Product.ToString() }

    $version = Get-ObjectValue -InputObject $Product -PropertyName 'Version'
    if (-not $version) { $version = Get-ObjectValue -InputObject $Product -PropertyName 'ProductVersion' }
    if (-not $version) { $version = Get-ObjectValue -InputObject $Product -PropertyName 'BuildVersion' }
    if (-not $version) { $version = Get-ObjectValue -InputObject $Product -PropertyName 'PatchVersion' }
    if (-not $version) { $version = Get-LatestVersionFromText $Product.ToString() }

    $installStatus = Get-ObjectValue -InputObject $Product -PropertyName 'InstallStatus'
    if (-not $installStatus) { $installStatus = Get-ObjectValue -InputObject $Product -PropertyName 'Status' }
    if (-not $installStatus) { $installStatus = if ($version) { 'Installed' } else { 'Unknown' } }

    [pscustomobject]@{
        Server        = $ServerName
        Product       = $productName
        Version       = $version
        InstallStatus = $installStatus
        Source        = $Source
    }
}

function ConvertFrom-SPProductServerRecord {
    param(
        [object]$ServerProduct,
        [string]$Source
    )

    $serverName = Get-ObjectValue -InputObject $ServerProduct -PropertyName 'ServerName'
    if (-not $serverName) { $serverName = Get-ObjectValue -InputObject $ServerProduct -PropertyName 'Name' }
    if (-not $serverName) { $serverName = Get-ObjectValue -InputObject $ServerProduct -PropertyName 'Server' }

    $products = @(Get-ObjectValue -InputObject $ServerProduct -PropertyName 'Products') | Where-Object { $_ }
    $productName = if ($products.Count -gt 0) { ($products | ForEach-Object { ConvertFrom-HtmlFragmentText $_.ToString() }) -join ', ' } else { '' }
    $version = Get-ObjectValue -InputObject $ServerProduct -PropertyName 'Version'
    if (-not $version) { $version = Get-ObjectValue -InputObject $ServerProduct -PropertyName 'ProductVersion' }
    if (-not $version) { $version = Get-LatestVersionFromText $ServerProduct.ToString() }

    [pscustomobject]@{
        Server        = $serverName
        Product       = $productName
        Version       = $version
        InstallStatus = Get-ObjectValue -InputObject $ServerProduct -PropertyName 'InstallStatus'
        Source        = $Source
    }
}

function Test-SPReportInstalledStatus {
    param([AllowNull()][object]$Status)

    $text = ([string]$Status).Trim()
    return ($text -eq 'Installed' -or $text -eq 'Online' -or $text -eq 'Success' -or $text -eq 'NoActionRequired')
}

function Get-SPReportProductPatchStatus {
    $cached = Get-Variable -Name 'ProductPatchStatusRows' -Scope Script -ErrorAction SilentlyContinue
    if ($cached) { return @($cached.Value) }

    if (-not (Test-CommandAvailable -Name 'Get-SPProduct')) {
        $result = @([pscustomobject]@{ Error = 'Get-SPProduct is not available in this environment.'; Details = '' })
        Set-Variable -Name 'ProductPatchStatusRows' -Scope Script -Value $result
        return $result
    }

    $rows = New-Object System.Collections.ArrayList
    $servers = @(Get-SPServer | Where-Object { (Get-ObjectValue -InputObject $_ -PropertyName 'Role') -ne 'Invalid' } | Sort-Object Name)
    try {
        foreach ($farmProduct in @(Get-SPProduct -ErrorAction Stop)) {
            foreach ($serverProduct in @(Get-ObjectValue -InputObject $farmProduct -PropertyName 'Servers')) {
                if (-not $serverProduct) { continue }
                $row = ConvertFrom-SPProductServerRecord -ServerProduct $serverProduct -Source 'Get-SPProduct farm Servers'
                $rowServer = Get-ObjectValue -InputObject $row -PropertyName 'Server'
                if (-not $rowServer) { continue }
                [void]$rows.Add($row)
            }
        }
    }
    catch {
        [void]$rows.Add([pscustomobject]@{
            Server  = ''
            Error   = 'Get-SPProduct farm server status could not be queried.'
            Details = ('Get-SPProduct: {0}' -f $_.Exception.Message)
        })
    }

    $farmServerRows = @($rows | Where-Object { -not ((Test-ObjectProperty -InputObject $_ -PropertyName 'Error') -and (Get-ObjectValue -InputObject $_ -PropertyName 'Error')) })
    foreach ($server in $servers) {
        $serverName = Get-ObjectValue -InputObject $server -PropertyName 'Name'
        $serverShortName = ($serverName -split '\.')[0]
        $existingRows = @($farmServerRows | Where-Object {
            $productServer = Get-ObjectValue -InputObject $_ -PropertyName 'Server'
            $productServer -eq $serverName -or $productServer -eq $serverShortName
        })
        if ($existingRows.Count -gt 0) { continue }

        $source = 'Get-SPProduct -Server'
        try {
            $products = @(Get-SPProduct -Server $serverName -ErrorAction Stop)
            if ($products.Count -eq 0) {
                [void]$rows.Add([pscustomobject]@{
                    Server        = $serverName
                    Product       = ''
                    Version       = ''
                    InstallStatus = 'No products returned'
                    Source        = $source
                })
                continue
            }

            foreach ($product in $products) {
                [void]$rows.Add((ConvertFrom-SPProductRecord -Product $product -ServerName $serverName -Source $source))
            }
        }
        catch {
            [void]$rows.Add([pscustomobject]@{
                Server  = $serverName
                Error   = 'Get-SPProduct could not be queried.'
                Details = ('{0}: {1}' -f $source, $_.Exception.Message)
            })
        }
    }

    $result = @($rows)
    Set-Variable -Name 'ProductPatchStatusRows' -Scope Script -Value $result
    return $result
}

function Get-SPReportProductPatchFallbackRows {
    param([string]$Reason)

    @(Get-SPReportProductPatchStatus) | Where-Object { -not ((Test-ObjectProperty -InputObject $_ -PropertyName 'Error') -and (Get-ObjectValue -InputObject $_ -PropertyName 'Error')) } | ForEach-Object {
        [pscustomobject]@{
            Server        = Get-ObjectValue -InputObject $_ -PropertyName 'Server'
            Product       = Get-ObjectValue -InputObject $_ -PropertyName 'Product'
            Version       = Get-ObjectValue -InputObject $_ -PropertyName 'Version'
            InstallStatus = Get-ObjectValue -InputObject $_ -PropertyName 'InstallStatus'
            Source        = ('Get-SPProduct fallback; {0}' -f $Reason)
        }
    }
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

function Get-SPReportUpdateForBuild {
    param(
        [string]$Product,
        [AllowNull()][object]$BuildVersion
    )

    if (-not $Product -or -not $BuildVersion) { return $null }
    $buildNumber = ConvertTo-VersionNumber $BuildVersion
    if ($buildNumber -le 0) { return $null }

    $cache = Read-SPReportUpdateCache
    if (-not $cache -or -not (Test-ObjectProperty -InputObject $cache -PropertyName 'Products')) { return $null }
    $productCache = Get-ObjectValue -InputObject $cache.Products -PropertyName $Product
    if (-not $productCache) { return $null }

    $updates = @(Get-ObjectValue -InputObject $productCache -PropertyName 'Updates')
    if ($updates.Count -eq 0) { return $null }

    $match = $updates | Where-Object { (ConvertTo-VersionNumber (Get-ObjectValue -InputObject $_ -PropertyName 'LatestBuild')) -eq $buildNumber } | Select-Object -First 1
    if (-not $match) { return $null }

    [pscustomobject]@{
        UpdateName  = Get-ObjectValue -InputObject $match -PropertyName 'UpdateName'
        KB          = Get-ObjectValue -InputObject $match -PropertyName 'KB'
        ReleaseDate = Get-ObjectValue -InputObject $match -PropertyName 'ReleaseDate'
    }
}

function Get-SPReportUpdatesFromMicrosoftContent {
    param(
        [string]$Content,
        [string]$Product
    )

    $escapedProduct = [regex]::Escape($Product)
    $sectionMatch = [regex]::Match($Content, "(?is)##\s+$escapedProduct\s+update history(.*?)(\r?\n##\s+|$)")
    if (-not $sectionMatch.Success) {
        $sectionMatch = [regex]::Match($Content, "(?is)<h2[^>]*>\s*$escapedProduct\s+update history\s*</h2>(.*?)(<h2[^>]*>|$)")
    }
    if (-not $sectionMatch.Success) { return @() }

    $updates = New-Object System.Collections.ArrayList
    $section = $sectionMatch.Groups[1].Value

    if ($section -match '(?is)<tr') {
        foreach ($row in [regex]::Matches($section, '(?is)<tr[^>]*>(.*?)</tr>')) {
            $cells = @([regex]::Matches($row.Groups[1].Value, '(?is)<td[^>]*>(.*?)</td>') | ForEach-Object { $_.Groups[1].Value })
            if ($cells.Count -lt 4) { continue }

            $updateName = ConvertFrom-HtmlFragmentText $cells[0]
            $kb = (@([regex]::Matches($cells[1], 'KB\s*\d+') | ForEach-Object { ($_.Value -replace '\s+', ' ') }) -join ', ')
            $latestBuild = Get-LatestVersionFromText (ConvertFrom-HtmlFragmentText $cells[2])
            $releaseDate = ConvertFrom-HtmlFragmentText $cells[3]

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

    $lines = @($section -split "\r?\n")
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
        'https://learn.microsoft.com/en-us/officeupdates/sharepoint-updates?view=officeupdates-raw',
        'https://learn.microsoft.com/en-us/officeupdates/sharepoint-updates'
    )
    $lastError = ''

    foreach ($source in $sources) {
        try {
            $response = Invoke-WebRequest -Uri $source -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop
            $content = [string]$response.Content
            $updates = @(Get-SPReportUpdatesFromMicrosoftContent -Content $content -Product $product)
            if ($updates.Count -eq 0) {
                $lastError = ('No update rows parsed from {0}' -f $source)
                continue
            }

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

function Get-SPListItemFieldText {
    param(
        [AllowNull()][object]$Item,
        [string[]]$FieldNames
    )

    if ($null -eq $Item) { return '' }

    foreach ($fieldName in $FieldNames) {
        try {
            $field = $null
            if ($Item.Fields.ContainsField($fieldName)) {
                $field = $Item.Fields.GetField($fieldName)
            }
            elseif ($Item.Fields.ContainsFieldWithStaticName($fieldName)) {
                $field = $Item.Fields.GetFieldByInternalName($fieldName)
            }

            if ($field) {
                $value = $Item[$field.InternalName]
                if ($null -ne $value -and [string]$value -ne '') { return (ConvertFrom-HtmlFragmentText $value) }
            }
        }
        catch { }

        try {
            $value = $Item[$fieldName]
            if ($null -ne $value -and [string]$value -ne '') { return (ConvertFrom-HtmlFragmentText $value) }
        }
        catch { }
    }

    return ''
}

function Get-SPListItemFieldTextByPattern {
    param(
        [AllowNull()][object]$Item,
        [string[]]$Patterns
    )

    if ($null -eq $Item) { return '' }

    foreach ($field in $Item.Fields) {
        $names = @(
            Get-ObjectValue -InputObject $field -PropertyName 'Title'
            Get-ObjectValue -InputObject $field -PropertyName 'InternalName'
            Get-ObjectValue -InputObject $field -PropertyName 'StaticName'
        ) | Where-Object { $_ }

        foreach ($pattern in $Patterns) {
            if (($names -join ' ') -match $pattern) {
                try {
                    $value = $Item[$field.InternalName]
                    if ($null -ne $value -and [string]$value -ne '') { return (ConvertFrom-HtmlFragmentText $value) }
                }
                catch { }
            }
        }
    }

    return ''
}

function Get-SPReportHealthFallbackRows {
    if (-not (Test-CommandAvailable -Name Get-SPHealthAnalysisRule)) { return @() }

    Get-SPHealthAnalysisRule | Sort-Object Category, Summary | ForEach-Object {
        $summary = Get-ObjectValue -InputObject $_ -PropertyName 'Summary'
        if (-not $summary) { $summary = Get-ObjectValue -InputObject $_ -PropertyName 'Title' }

        [pscustomobject]@{
            Title            = $summary
            Category         = Get-ObjectValue -InputObject $_ -PropertyName 'Category'
            Severity         = Get-ObjectValue -InputObject $_ -PropertyName 'Severity'
            CurrentStatus    = 'Rule inventory fallback'
            Explanation      = 'Central Administration Health Reports did not expose populated finding fields for this farm.'
            Remedy           = ''
            PossibleSolution = Get-ReportText -Key 'HealthAnalyzerFallbackSolution'
            FailingServers   = ''
            FailingServices  = ''
            Modified         = ''
            RuleId           = Get-ObjectValue -InputObject $_ -PropertyName 'Id'
        }
    }
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

function Get-SPReportCimInstance {
    param(
        [string]$ComputerName,
        [string]$ClassName,
        [string]$Filter
    )

    $localNames = @($env:COMPUTERNAME, ([System.Net.Dns]::GetHostName())) | Where-Object { $_ } | ForEach-Object { $_.ToUpperInvariant() }
    $target = ($ComputerName -split '\.')[0].ToUpperInvariant()
    $parameters = @{ ClassName = $ClassName; ErrorAction = 'Stop' }
    if ($Filter) { $parameters.Filter = $Filter }
    if ($localNames -notcontains $target) { $parameters.ComputerName = $ComputerName }

    Get-CimInstance @parameters
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

    if (@($Data | Where-Object { (Test-ObjectProperty -InputObject $_ -PropertyName 'Error') -and (Get-ObjectValue -InputObject $_ -PropertyName 'Error') }).Count -gt 0) { return 'Warning' }
    if (@($Data | Where-Object { (Test-ObjectProperty -InputObject $_ -PropertyName 'InstallStatus') -and (Get-ObjectValue -InputObject $_ -PropertyName 'InstallStatus') -ne 'Installed' }).Count -gt 0) { return 'Warning' }
    if (@($Data | Where-Object { (Test-ObjectProperty -InputObject $_ -PropertyName 'ProductInstallStatus') -and (Get-ObjectValue -InputObject $_ -PropertyName 'ProductInstallStatus') -eq 'Issues detected' }).Count -gt 0) { return 'Warning' }
    if (@($Data | Where-Object { (Test-ObjectProperty -InputObject $_ -PropertyName 'LatestKnownStatus') -and (Get-ObjectValue -InputObject $_ -PropertyName 'LatestKnownStatus') -eq 'Below latest known build' }).Count -gt 0) { return 'Warning' }
    return 'Good'
}

function Get-SPReportFarmVersion {
    $farm = Get-SPFarm
    $build = $farm.BuildVersion
    $configDb = Get-SPReportConfigurationDatabase
    $product = Get-SPProductUpdateSectionName $build.ToString()
    if (-not $product) { $product = 'Unknown SharePoint version' }

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
    $farmBuild = (Get-SPFarm).BuildVersion.ToString()
    Get-SPServer | Where-Object { (Get-ObjectValue -InputObject $_ -PropertyName 'Role') -ne 'Invalid' } | Sort-Object Name | ForEach-Object {
        [pscustomobject]@{
            Name             = $_.Name
            Role             = Get-ObjectValue $_ 'Role'
            Address          = $_.Address
            Status           = $_.Status
            SharePointBuild = $farmBuild
            ServerRole       = Get-ObjectValue $_ 'ServerRole'
            CompliantWithMinRole = Get-ObjectValue $_ 'CompliantWithMinRole'
        }
    }
}

function Get-SPReportCentralAdminPatchStatus {
    $cached = Get-Variable -Name 'CentralAdminPatchStatusRows' -Scope Script -ErrorAction SilentlyContinue
    if ($cached) { return @($cached.Value) }

    $adminWebApplication = Get-SPWebApplication -IncludeCentralAdministration | Where-Object { Get-ObjectValue -InputObject $_ -PropertyName 'IsAdministrationWebApplication' } | Select-Object -First 1
    if (-not $adminWebApplication) {
        $result = @([pscustomobject]@{ Error = 'Central Administration web application could not be found.'; Details = '' })
        Set-Variable -Name 'CentralAdminPatchStatusRows' -Scope Script -Value $result
        return $result
    }

    $adminUrl = (Get-ObjectValue -InputObject $adminWebApplication -PropertyName 'Url').TrimEnd('/')
    $patchStatusUrl = '{0}/_admin/PatchStatus.aspx' -f $adminUrl

    try {
        $response = Invoke-WebRequest -Uri $patchStatusUrl -UseDefaultCredentials -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
        $rows = @(ConvertFrom-PatchStatusHtml -Html $response.Content -Source $patchStatusUrl)
        if ($rows.Count -eq 0) {
            $fallbackRows = @(Get-SPReportProductPatchFallbackRows -Reason ('Central Administration Patch Status returned no parseable product rows from {0}' -f $patchStatusUrl))
            $rows = if ($fallbackRows.Count -gt 0) { $fallbackRows } else { @([pscustomobject]@{ Error = 'Central Administration Patch Status returned no parseable product rows.'; Details = $patchStatusUrl }) }
        }
        Set-Variable -Name 'CentralAdminPatchStatusRows' -Scope Script -Value $rows
        return $rows
    }
    catch {
        $details = ('{0}: {1}' -f $patchStatusUrl, $_.Exception.Message)
        $fallbackRows = @(Get-SPReportProductPatchFallbackRows -Reason ('Central Administration Patch Status could not be queried. {0}' -f $details))
        $result = if ($fallbackRows.Count -gt 0) { $fallbackRows } else { @([pscustomobject]@{ Error = 'Central Administration Patch Status could not be queried.'; Details = $details }) }
        Set-Variable -Name 'CentralAdminPatchStatusRows' -Scope Script -Value $result
        return $result
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
    $servers = @(Get-SPServer | Where-Object { (Get-ObjectValue -InputObject $_ -PropertyName 'Role') -ne 'Invalid' } | Sort-Object Name)

    foreach ($server in $servers) {
        $serverName = Get-ObjectValue -InputObject $server -PropertyName 'Name'
        $serverRole = Get-ObjectValue -InputObject $server -PropertyName 'Role'
        try {
            $os = Get-SPReportCimInstance -ComputerName $serverName -ClassName 'Win32_OperatingSystem'
            $computer = Get-SPReportCimInstance -ComputerName $serverName -ClassName 'Win32_ComputerSystem'
            $processor = Get-SPReportCimInstance -ComputerName $serverName -ClassName 'Win32_Processor' | Select-Object -First 1
            $disks = @(Get-SPReportCimInstance -ComputerName $serverName -ClassName 'Win32_LogicalDisk' -Filter 'DriveType=3' | Sort-Object DeviceID)

            if ($disks.Count -eq 0) {
                [pscustomobject]@{
                    Server            = $serverName
                    Role              = $serverRole
                    ProbeStatus       = 'No fixed disks returned'
                    OperatingSystem   = Get-ObjectValue -InputObject $os -PropertyName 'Caption'
                    OSVersion         = Get-ObjectValue -InputObject $os -PropertyName 'Version'
                    LastBootTime      = Get-ObjectValue -InputObject $os -PropertyName 'LastBootUpTime'
                    PowerShellVersion = if (($serverName -split '\.')[0] -eq $env:COMPUTERNAME) { $PSVersionTable.PSVersion.ToString() } else { '' }
                    TotalMemory       = Format-ByteSize (Get-ObjectValue -InputObject $computer -PropertyName 'TotalPhysicalMemory')
                    Processor         = Get-ObjectValue -InputObject $processor -PropertyName 'Name'
                    Drive             = ''
                    DriveType         = ''
                    VolumeName        = ''
                    FileSystem        = ''
                    DriveSize         = ''
                    DriveFree         = ''
                    DriveFreePercent  = ''
                    SpaceStatus       = 'Unknown'
                    Error             = ''
                    Details           = ''
                }
                continue
            }

            foreach ($disk in $disks) {
                $freePercent = if ($disk.Size -gt 0) { [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2) } else { 0 }
                [pscustomobject]@{
                    Server            = $serverName
                    Role              = $serverRole
                    ProbeStatus       = 'Success'
                    OperatingSystem   = $os.Caption
                    OSVersion         = $os.Version
                    LastBootTime      = $os.LastBootUpTime
                    PowerShellVersion = if (($serverName -split '\.')[0] -eq $env:COMPUTERNAME) { $PSVersionTable.PSVersion.ToString() } else { '' }
                    TotalMemory       = Format-ByteSize $computer.TotalPhysicalMemory
                    Processor         = $processor.Name
                    Drive             = $disk.DeviceID
                    DriveType         = Get-DiskTypeName $disk.DriveType
                    VolumeName        = $disk.VolumeName
                    FileSystem        = $disk.FileSystem
                    DriveSize         = Format-ByteSize $disk.Size
                    DriveFree         = Format-ByteSize $disk.FreeSpace
                    DriveFreePercent  = ('{0:N2}%' -f $freePercent)
                    SpaceStatus       = Get-DiskSpaceStatus $freePercent
                    Error             = ''
                    Details           = ''
                }
            }
        }
        catch {
            [pscustomobject]@{
                Server            = $serverName
                Role              = $serverRole
                ProbeStatus       = 'Failed'
                OperatingSystem   = ''
                OSVersion         = ''
                LastBootTime      = ''
                PowerShellVersion = ''
                TotalMemory       = ''
                Processor         = ''
                Drive             = ''
                DriveType         = ''
                VolumeName        = ''
                FileSystem        = ''
                DriveSize         = ''
                DriveFree         = ''
                DriveFreePercent  = ''
                SpaceStatus       = 'Unknown'
                Error             = $_.Exception.Message
                Details           = $_.ScriptStackTrace
            }
        }
    }
}

function Get-SPReportFarmServerUpdateStatus {
    $farmBuild = (Get-SPFarm).BuildVersion.ToString()
    $activeFarmServers = @(Get-SPServer | Where-Object { (Get-ObjectValue -InputObject $_ -PropertyName 'Role') -ne 'Invalid' } | Sort-Object Name)
    $productPatchRows = @(Get-SPReportProductPatchStatus)
    $productPatchRowsWithoutErrors = @($productPatchRows | Where-Object { -not ((Test-ObjectProperty -InputObject $_ -PropertyName 'Error') -and (Get-ObjectValue -InputObject $_ -PropertyName 'Error')) })
    $patchStatusRows = @(Get-SPReportCentralAdminPatchStatus)
    $patchStatusError = $patchStatusRows | Where-Object { (Test-ObjectProperty -InputObject $_ -PropertyName 'Error') -and (Get-ObjectValue -InputObject $_ -PropertyName 'Error') } | Select-Object -First 1
    $patchStatusProductRows = @($patchStatusRows | Where-Object { -not ((Test-ObjectProperty -InputObject $_ -PropertyName 'Error') -and (Get-ObjectValue -InputObject $_ -PropertyName 'Error')) })

    $microsoftLatest = Get-SPReportMicrosoftLatestSharePointUpdate -InstalledBuild $farmBuild
    $effectiveLatestBuild = if ($LatestKnownSharePointBuild) { $LatestKnownSharePointBuild } else { Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'LatestBuild' }
    $effectiveLatestUpdateName = if ($LatestKnownSharePointUpdateName) { $LatestKnownSharePointUpdateName } else { Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'UpdateName' }
    $latestKnownVersionNumber = ConvertTo-VersionNumber $effectiveLatestBuild

    foreach ($server in $activeFarmServers) {
        $serverName = Get-ObjectValue -InputObject $server -PropertyName 'Name'
        $serverShortName = ($serverName -split '\.')[0]
        $serverRole = Get-ObjectValue -InputObject $server -PropertyName 'Role'
        $serverPatchRows = @($patchStatusProductRows | Where-Object {
            $patchServer = Get-ObjectValue -InputObject $_ -PropertyName 'Server'
            $patchServer -eq $serverName -or $patchServer -eq $serverShortName
        })
        $serverProductRows = @($productPatchRowsWithoutErrors | Where-Object {
            $productServer = Get-ObjectValue -InputObject $_ -PropertyName 'Server'
            $productServer -eq $serverName -or $productServer -eq $serverShortName
        })
        $serverVersionRows = if ($serverProductRows.Count -gt 0) { $serverProductRows } else { $serverPatchRows }
        $installedPatchRows = @($serverVersionRows | Where-Object { (Test-SPReportInstalledStatus (Get-ObjectValue -InputObject $_ -PropertyName 'InstallStatus')) -and (Get-ObjectValue -InputObject $_ -PropertyName 'Version') })
        $latestServerPatchRow = $installedPatchRows | Sort-Object { ConvertTo-VersionNumber (Get-ObjectValue -InputObject $_ -PropertyName 'Version') } -Descending | Select-Object -First 1
        $serverBuild = Get-ObjectValue -InputObject $latestServerPatchRow -PropertyName 'Version'
        if (-not $serverBuild) { $serverBuild = $farmBuild }
        $serverBuildVersionNumber = ConvertTo-VersionNumber $serverBuild
        $serverProductName = Get-SPProductUpdateSectionName $serverBuild
        $installedSharePointUpdate = Get-SPReportUpdateForBuild -Product $serverProductName -BuildVersion $serverBuild
        if (-not $installedSharePointUpdate -and $effectiveLatestBuild -and (ConvertTo-VersionNumber $serverBuild) -eq (ConvertTo-VersionNumber $effectiveLatestBuild)) {
            $installedSharePointUpdate = [pscustomobject]@{
                UpdateName  = $effectiveLatestUpdateName
                KB          = Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'KB'
                ReleaseDate = Get-ObjectValue -InputObject $microsoftLatest -PropertyName 'ReleaseDate'
            }
        }

        $productInstallStatus = 'Not available'
        $patchStatusSource = 'Farm build fallback'
        $productStatusIsHealthy = $false
        if ($serverProductRows.Count -gt 0) {
            $notInstalled = @($serverProductRows | Where-Object { -not (Test-SPReportInstalledStatus (Get-ObjectValue -InputObject $_ -PropertyName 'InstallStatus')) })
            $productStatuses = @($serverProductRows | ForEach-Object { Get-ObjectValue -InputObject $_ -PropertyName 'InstallStatus' } | Where-Object { $_ } | Select-Object -Unique)
            $productInstallStatus = if ($notInstalled.Count -gt 0) { 'Issues detected' } elseif ($productStatuses.Count -gt 0) { $productStatuses -join ', ' } else { 'NoActionRequired' }
            $productStatusIsHealthy = ($notInstalled.Count -eq 0)
            $patchStatusSource = 'Get-SPProduct -Server'
        }
        elseif ($serverPatchRows.Count -gt 0) {
            $notInstalled = @($serverPatchRows | Where-Object { -not (Test-SPReportInstalledStatus (Get-ObjectValue -InputObject $_ -PropertyName 'InstallStatus')) })
            $productStatuses = @($serverPatchRows | ForEach-Object { Get-ObjectValue -InputObject $_ -PropertyName 'InstallStatus' } | Where-Object { $_ } | Select-Object -Unique)
            $productInstallStatus = if ($notInstalled.Count -gt 0) { 'Issues detected' } elseif ($productStatuses.Count -gt 0) { $productStatuses -join ', ' } else { 'Installed' }
            $productStatusIsHealthy = ($notInstalled.Count -eq 0)
            $patchStatusSource = 'Central Administration Patch Status'
        }
        elseif ($patchStatusError) {
            $patchStatusSource = Get-ObjectValue -InputObject $patchStatusError -PropertyName 'Details'
        }

        $farmRelativeStatus = 'Patch status not available; farm build used'
        $latestKnownStatus = 'Not evaluated'
        $latestUpdate = ''
        $latestSecurityUpdate = ''
        $windowsSharePointUpdate = ''
        $updateQueryStatus = 'Not queried'

        if ($serverProductRows.Count -gt 0 -or $serverPatchRows.Count -gt 0) {
            if ($productStatusIsHealthy) { $farmRelativeStatus = 'No SharePoint patch action required' }
            else { $farmRelativeStatus = 'Patch install issue detected' }
        }

        if ($latestKnownVersionNumber -gt 0) {
            if ($serverBuildVersionNumber -ge $latestKnownVersionNumber) { $latestKnownStatus = 'At or above latest known build' }
            else { $latestKnownStatus = 'Below latest known build' }
        }

        try {
            $hotfixes = @(Get-HotFix -ComputerName $serverName -ErrorAction Stop | Sort-Object InstalledOn -Descending)
            $latest = $hotfixes | Select-Object -First 1
            $latestSecurity = $hotfixes | Where-Object { $_.Description -match 'Security' } | Select-Object -First 1
            if ($latest) { $latestUpdate = ('{0} {1} {2}' -f $latest.HotFixID, $latest.Description, $latest.InstalledOn) }
            if ($latestSecurity) { $latestSecurityUpdate = ('{0} {1} {2}' -f $latestSecurity.HotFixID, $latestSecurity.Description, $latestSecurity.InstalledOn) }
            $installedSharePointKbText = [string](Get-ObjectValue -InputObject $installedSharePointUpdate -PropertyName 'KB')
            $installedSharePointKbs = @([regex]::Matches($installedSharePointKbText, 'KB\s*\d+') | ForEach-Object { ($_.Value -replace '\s+', '') })
            if ($installedSharePointKbs.Count -gt 0) {
                $sharePointHotfix = $hotfixes | Where-Object { $installedSharePointKbs -contains (($_.HotFixID -replace '\s+', '').ToUpperInvariant()) } | Select-Object -First 1
                if ($sharePointHotfix) { $windowsSharePointUpdate = ('{0} {1} {2}' -f $sharePointHotfix.HotFixID, $sharePointHotfix.Description, $sharePointHotfix.InstalledOn) }
                else { $windowsSharePointUpdate = ('{0} not returned by Get-HotFix' -f ($installedSharePointKbs -join ', ')) }
            }
            $updateQueryStatus = 'Success'
        }
        catch {
            $updateQueryStatus = $_.Exception.Message
        }

        [pscustomobject]@{
            Server                  = $serverName
            Role                    = $serverRole
            Status                  = Get-ObjectValue -InputObject $server -PropertyName 'Status'
            FarmBuild               = $farmBuild
            SharePointBuild         = $serverBuild
            InstalledSharePointUpdate = Get-ObjectValue -InputObject $installedSharePointUpdate -PropertyName 'UpdateName'
            InstalledSharePointKB   = Get-ObjectValue -InputObject $installedSharePointUpdate -PropertyName 'KB'
            InstalledSharePointReleaseDate = Get-ObjectValue -InputObject $installedSharePointUpdate -PropertyName 'ReleaseDate'
            ProductInstallStatus    = $productInstallStatus
            PatchStatusSource       = $patchStatusSource
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
            LatestWindowsUpdate     = $latestUpdate
            LatestWindowsSecurityUpdate = $latestSecurityUpdate
            WindowsSharePointUpdate = $windowsSharePointUpdate
            UpdateQueryStatus       = $updateQueryStatus
        }
    }
}

function Get-SPReportInstalledUpdates {
    $servers = @(Get-SPServer | Where-Object { (Get-ObjectValue -InputObject $_ -PropertyName 'Role') -ne 'Invalid' } | Sort-Object Name)
    foreach ($server in $servers) {
        $serverName = Get-ObjectValue -InputObject $server -PropertyName 'Name'
        try {
            Get-HotFix -ComputerName $serverName -ErrorAction Stop | Sort-Object -Property InstalledOn, HotFixID -Descending | ForEach-Object {
                [pscustomobject]@{
                    Server      = $serverName
                    Role        = Get-ObjectValue -InputObject $server -PropertyName 'Role'
                    HotFixId    = $_.HotFixID
                    Description = $_.Description
                    InstalledBy = $_.InstalledBy
                    InstalledOn = $_.InstalledOn
                    Caption     = $_.Caption
                    UpdateType  = if ($_.Description -match 'Security') { 'Security' } else { 'Update' }
                }
            }
        }
        catch {
            [pscustomobject]@{
                Server      = $serverName
                Role        = Get-ObjectValue -InputObject $server -PropertyName 'Role'
                HotFixId    = ''
                Description = ''
                InstalledBy = ''
                InstalledOn = ''
                Caption     = ''
                UpdateType  = ('Error: {0}' -f $_.Exception.Message)
            }
        }
    }
}

function Get-SPReportCachedSharePointUpdateHistory {
    $farmBuild = (Get-SPFarm).BuildVersion.ToString()
    $product = Get-SPProductUpdateSectionName $farmBuild
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

function Get-SPReportHealthAnalyzerFindings {
    if (-not $IncludeCentralAdminHealthReports) {
        return @(Get-SPReportHealthFallbackRows)
    }

    $adminWebApplication = Get-SPWebApplication -IncludeCentralAdministration | Where-Object { Get-ObjectValue -InputObject $_ -PropertyName 'IsAdministrationWebApplication' } | Select-Object -First 1
    if (-not $adminWebApplication) {
        return [pscustomobject]@{ Error = 'Central Administration web application could not be found.'; Details = '' }
    }

    $site = $null
    $web = $null
    try {
        $site = Get-SPSite (Get-ObjectValue -InputObject $adminWebApplication -PropertyName 'Url')
        $web = $site.OpenWeb()
        $list = $web.Lists.TryGetList('Health Reports')
        if (-not $list) { $list = $web.Lists.TryGetList('HealthReports') }
        if (-not $list) {
            foreach ($candidate in $web.Lists) {
                $rootFolder = Get-ObjectValue -InputObject $candidate -PropertyName 'RootFolder'
                $rootName = Get-ObjectValue -InputObject $rootFolder -PropertyName 'Name'
                if ($candidate.Title -match 'Health|Sağlık' -or $rootName -eq 'HealthReports') {
                    $list = $candidate
                    break
                }
            }
        }

        if (-not $list) {
            return [pscustomobject]@{ Error = 'Central Administration Health Reports list could not be found.'; Details = '' }
        }

        $query = New-Object Microsoft.SharePoint.SPQuery
        $query.RowLimit = 200
        $query.Query = '<OrderBy><FieldRef Name="Modified" Ascending="FALSE" /></OrderBy>'

        $items = $list.GetItems($query)
        $findings = New-Object System.Collections.ArrayList
        foreach ($item in $items) {
            $title = Get-SPListItemFieldText -Item $item -FieldNames @('Title', 'Name')
            $category = Get-SPListItemFieldText -Item $item -FieldNames @('Category', 'HealthReportCategory')
            if (-not $category) { $category = Get-SPListItemFieldTextByPattern -Item $item -Patterns @('(?i)category|kategori') }

            $severity = Get-SPListItemFieldText -Item $item -FieldNames @('Severity', 'HealthReportSeverity')
            if (-not $severity) { $severity = Get-SPListItemFieldTextByPattern -Item $item -Patterns @('(?i)severity|önem|onem') }

            $currentStatus = Get-SPListItemFieldText -Item $item -FieldNames @('Status', 'HealthReportStatus')
            if (-not $currentStatus) { $currentStatus = Get-SPListItemFieldTextByPattern -Item $item -Patterns @('(?i)status|durum') }

            $explanation = Get-SPListItemFieldText -Item $item -FieldNames @('Explanation', 'HealthReportExplanation')
            if (-not $explanation) { $explanation = Get-SPListItemFieldTextByPattern -Item $item -Patterns @('(?i)explanation|açıklama|aciklama') }

            $remedy = Get-SPListItemFieldText -Item $item -FieldNames @('Remedy', 'HealthReportRemedy')
            if (-not $remedy) { $remedy = Get-SPListItemFieldTextByPattern -Item $item -Patterns @('(?i)remedy|solution|çözüm|cozum|repair') }

            $failingServers = Get-SPListItemFieldText -Item $item -FieldNames @('Failing Servers', 'FailingServers', 'Failing_x0020_Servers')
            if (-not $failingServers) { $failingServers = Get-SPListItemFieldTextByPattern -Item $item -Patterns @('(?i)failing.*server|server.*fail|sunucu') }

            $failingServices = Get-SPListItemFieldText -Item $item -FieldNames @('Failing Services', 'FailingServices', 'Failing_x0020_Services')
            if (-not $failingServices) { $failingServices = Get-SPListItemFieldTextByPattern -Item $item -Patterns @('(?i)failing.*service|service.*fail|servis') }

            $ruleId = Get-SPListItemFieldText -Item $item -FieldNames @('RuleId', 'Rule ID', 'HealthReportRuleId')
            if (-not $ruleId) { $ruleId = Get-SPListItemFieldTextByPattern -Item $item -Patterns @('(?i)rule.*id|id.*rule|kural') }

            $hasFindingData = $severity -or $currentStatus -or $explanation -or $remedy -or $failingServers -or $failingServices -or $ruleId
            if (-not $hasFindingData) { continue }

            $possibleSolution = if ($remedy) { $remedy } else { Get-ReportText -Key 'HealthAnalyzerFallbackSolution' }

            [void]$findings.Add([pscustomobject]@{
                Title            = $title
                Category         = $category
                Severity         = $severity
                CurrentStatus    = $currentStatus
                Explanation      = $explanation
                Remedy           = $remedy
                PossibleSolution = $possibleSolution
                FailingServers   = $failingServers
                FailingServices  = $failingServices
                Modified         = Get-SPListItemFieldText -Item $item -FieldNames @('Modified')
                RuleId           = $ruleId
            })

            if ($findings.Count -ge 50) { break }
        }

        if ($findings.Count -gt 0) { return @($findings) }
        return @(Get-SPReportHealthFallbackRows)
    }
    finally {
        if ($web) { $web.Dispose() }
        if ($site) { $site.Dispose() }
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

function Get-SPReportSearchCrawlJobs {
    if (-not (Test-CommandAvailable -Name Get-SPEnterpriseSearchServiceApplication)) {
        return [pscustomobject]@{ Error = 'Search cmdlets are not available in this environment.'; Details = '' }
    }
    if (-not (Test-CommandAvailable -Name Get-SPEnterpriseSearchCrawlContentSource)) {
        return [pscustomobject]@{ Error = 'Get-SPEnterpriseSearchCrawlContentSource is not available in this environment.'; Details = '' }
    }

    $apps = Get-SPEnterpriseSearchServiceApplication
    foreach ($app in $apps) {
        Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $app | Sort-Object Name | ForEach-Object {
            $started = Get-ObjectValue -InputObject $_ -PropertyName 'CrawlStarted'
            $completed = Get-ObjectValue -InputObject $_ -PropertyName 'CrawlCompleted'
            $duration = ''
            if ($started -and $completed) {
                try { $duration = ([datetime]$completed - [datetime]$started).ToString() } catch { $duration = '' }
            }

            [pscustomobject]@{
                SearchApplication        = Get-ObjectValue -InputObject $app -PropertyName 'Name'
                ContentSource            = Get-ObjectValue -InputObject $_ -PropertyName 'Name'
                Type                     = Get-ObjectValue -InputObject $_ -PropertyName 'Type'
                CrawlStatus              = Get-ObjectValue -InputObject $_ -PropertyName 'CrawlStatus'
                CrawlStarted             = $started
                CrawlCompleted           = $completed
                CrawlDuration            = $duration
                IncrementalCrawlSchedule = Get-ObjectValue -InputObject $_ -PropertyName 'IncrementalCrawlSchedule'
                FullCrawlSchedule        = Get-ObjectValue -InputObject $_ -PropertyName 'FullCrawlSchedule'
                StartAddresses           = Get-ObjectValue -InputObject $_ -PropertyName 'StartAddresses'
                SuccessCount             = Get-ObjectValue -InputObject $_ -PropertyName 'SuccessCount'
                ErrorCount               = Get-ObjectValue -InputObject $_ -PropertyName 'ErrorCount'
                DeleteCount              = Get-ObjectValue -InputObject $_ -PropertyName 'DeleteCount'
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

    $hasError = @($Data | Where-Object { (Test-ObjectProperty -InputObject $_ -PropertyName 'Error') -and (Get-ObjectValue -InputObject $_ -PropertyName 'Error') }).Count -gt 0
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
        [void]$html.AppendLine(('<th title="{0}">{1}</th>' -f (ConvertTo-HtmlText $column), (ConvertTo-HtmlText (Get-ColumnHeaderText -ColumnName $column))))
    }
    [void]$html.AppendLine('</tr></thead><tbody>')

    foreach ($row in $Data) {
        [void]$html.AppendLine('<tr>')
        foreach ($column in $Columns) {
            $value = ConvertTo-DisplayValue (Get-ObjectValue -InputObject $row -PropertyName $column)
            $class = ''
            if ($column -match 'Status|State|Severity|Error|NeedsUpgrade|UpgradeRequired|IsDisabled') {
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
:root { --bg:#0b1120; --header1:#172554; --header2:#111827; --panel:#111827; --panel2:#1f2937; --tableHead:#020617; --stripe:rgba(255,255,255,.025); --text:#e5e7eb; --muted:#9ca3af; --good:#22c55e; --warn:#f59e0b; --bad:#ef4444; --unknown:#64748b; --line:#334155; }
body.light { --bg:#f8fafc; --header1:#dbeafe; --header2:#ffffff; --panel:#ffffff; --panel2:#eff6ff; --tableHead:#e0f2fe; --stripe:rgba(15,23,42,.035); --text:#0f172a; --muted:#475569; --good:#15803d; --warn:#b45309; --bad:#b91c1c; --unknown:#64748b; --line:#cbd5e1; }
* { box-sizing:border-box; }
body { margin:0; font-family:Segoe UI, Arial, sans-serif; background:var(--bg); color:var(--text); }
header { padding:28px 32px; background:linear-gradient(135deg,var(--header1),var(--header2) 65%); border-bottom:1px solid var(--line); }
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
th { position:sticky; top:0; background:var(--tableHead); color:var(--text); z-index:1; }
tr:nth-child(even) td { background:var(--stripe); }
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
<div class="toolbar"><button type="button" onclick="setAll(false)">$(ConvertTo-HtmlText (Get-ReportText -Key 'ExpandAll'))</button><button type="button" onclick="setAll(true)">$(ConvertTo-HtmlText (Get-ReportText -Key 'CollapseAll'))</button><button type="button" onclick="toggleTheme()">$(ConvertTo-HtmlText (Get-ReportText -Key 'ToggleTheme'))</button></div>
$($sectionsHtml.ToString())
</main>
<script>
function setAll(c){document.querySelectorAll('.section').forEach(function(s){s.classList.toggle('collapsed',c);});}
function applyTheme(t){document.body.classList.toggle('light',t==='light'); try{localStorage.setItem('spReportTheme',t);}catch(e){}}
function toggleTheme(){applyTheme(document.body.classList.contains('light')?'dark':'light');}
try{applyTheme(localStorage.getItem('spReportTheme')||'dark');}catch(e){applyTheme('dark');}
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
    $centralAdminPatchStatus = Invoke-SafeCollect -Name 'Central Administration patch status' -ScriptBlock { Get-SPReportCentralAdminPatchStatus }
    $farmServerUpdateStatus = Invoke-SafeCollect -Name 'Farm server update status' -ScriptBlock { Get-SPReportFarmServerUpdateStatus }
    $installedUpdates = Invoke-SafeCollect -Name 'Installed updates on farm servers' -ScriptBlock { Get-SPReportInstalledUpdates }
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
    Add-ReportSection -Title (Get-ReportText -Key 'ServersTitle') -Description (Get-ReportText -Key 'ServersDescription') -Data $servers -Columns @('Name', 'Role', 'ServerRole', 'CompliantWithMinRole', 'Address', 'Status', 'SharePointBuild') -Status (Get-SectionStatus -Data $servers)
    Add-ReportSection -Title (Get-ReportText -Key 'FarmServerUpdatesTitle') -Description (Get-ReportText -Key 'FarmServerUpdatesDescription') -Data $farmServerUpdateStatus -Columns @('Server', 'Role', 'Status', 'FarmBuild', 'SharePointBuild', 'InstalledSharePointUpdate', 'InstalledSharePointKB', 'InstalledSharePointReleaseDate', 'ProductInstallStatus', 'PatchStatusSource', 'FarmRelativeStatus', 'LatestKnownBuild', 'LatestKnownUpdateName', 'LatestKnownKB', 'LatestKnownReleaseDate', 'LatestKnownStatus', 'MicrosoftLookupProduct', 'MicrosoftLookupStatus', 'MicrosoftLookupSource', 'CacheLastRefresh', 'CachedUpdateCount', 'WindowsSharePointUpdate', 'LatestWindowsUpdate', 'LatestWindowsSecurityUpdate', 'UpdateQueryStatus') -Status (Get-SectionStatus -Data $farmServerUpdateStatus)
    Add-ReportSection -Title (Get-ReportText -Key 'CentralAdminPatchStatusTitle') -Description (Get-ReportText -Key 'CentralAdminPatchStatusDescription') -Data $centralAdminPatchStatus -Columns @('Server', 'Product', 'Version', 'InstallStatus', 'Source', 'Error', 'Details') -Status (Get-SectionStatus -Data $centralAdminPatchStatus)
    Add-ReportSection -Title (Get-ReportText -Key 'InstalledUpdatesTitle') -Description (Get-ReportText -Key 'InstalledUpdatesDescription') -Data $installedUpdates -Columns @('Server', 'Role', 'HotFixId', 'Description', 'InstalledBy', 'InstalledOn', 'Caption', 'UpdateType') -Status (Get-SectionStatus -Data $installedUpdates)
    Add-ReportSection -Title (Get-ReportText -Key 'UpdateCacheHistoryTitle') -Description (Get-ReportText -Key 'UpdateCacheHistoryDescription') -Data $cachedUpdateHistory -Columns @('Product', 'LatestBuild', 'UpdateName', 'KB', 'ReleaseDate', 'CacheFetchedAt', 'CacheSource') -Status (Get-SectionStatus -Data $cachedUpdateHistory)
    Add-ReportSection -Title (Get-ReportText -Key 'ServicesTitle') -Description (Get-ReportText -Key 'ServicesDescription') -Data $servicesOnServers -Columns @('Server', 'Service', 'Status', 'ServiceType', 'Id') -Status (Get-SectionStatus -Data $servicesOnServers)
    Add-ReportSection -Title (Get-ReportText -Key 'WebAppsTitle') -Description (Get-ReportText -Key 'WebAppsDescription') -Data $webApplications -Columns @('Name', 'Url', 'ApplicationPool', 'ApplicationPoolAccount', 'ClaimsAuthentication', 'AllowAnonymous', 'AuthenticationProvider', 'ContentDatabases', 'MaximumFileSizeMB', 'TimeZone', 'IsCentralAdministration') -Status (Get-SectionStatus -Data $webApplications)
    Add-ReportSection -Title (Get-ReportText -Key 'ContentDbTitle') -Description (Get-ReportText -Key 'ContentDbDescription') -Data $contentDatabases -Columns @('Name', 'WebApplication', 'Server', 'Status', 'CurrentSiteCount', 'WarningSiteCount', 'MaximumSiteCount', 'DiskSizeRequired', 'NeedsUpgrade', 'Id') -Status (Get-SectionStatus -Data $contentDatabases)

    $healthTitleKey = if (@($servers | Where-Object { -not (Test-ObjectProperty -InputObject $_ -PropertyName 'Error') }).Count -gt 1) { 'ServerHealthTitle' } else { 'LocalHealthTitle' }
    $healthDescriptionKey = if ($healthTitleKey -eq 'ServerHealthTitle') { 'ServerHealthDescription' } else { 'LocalHealthDescription' }
    Add-SectionFromCollector -Title (Get-ReportText -Key $healthTitleKey) -Description (Get-ReportText -Key $healthDescriptionKey) -Collector { Get-SPReportLocalServerHealth } -Columns @('Server', 'Role', 'ProbeStatus', 'OperatingSystem', 'OSVersion', 'LastBootTime', 'PowerShellVersion', 'TotalMemory', 'Processor', 'Drive', 'DriveType', 'VolumeName', 'FileSystem', 'DriveSize', 'DriveFree', 'DriveFreePercent', 'SpaceStatus', 'Error', 'Details')
    Add-SectionFromCollector -Title (Get-ReportText -Key 'AllDatabasesTitle') -Description (Get-ReportText -Key 'AllDatabasesDescription') -Collector { Get-SPReportDatabases } -Columns @('Name', 'TypeName', 'Server', 'Status', 'DiskSizeRequired', 'NeedsUpgrade', 'Id')
    Add-SectionFromCollector -Title (Get-ReportText -Key 'ServiceApplicationsTitle') -Description (Get-ReportText -Key 'ServiceApplicationsDescription') -Collector { Get-SPReportServiceApplications } -Columns @('Name', 'TypeName', 'Status', 'ApplicationPool', 'Id')
    Add-SectionFromCollector -Title (Get-ReportText -Key 'ServiceApplicationProxiesTitle') -Description (Get-ReportText -Key 'ServiceApplicationProxiesDescription') -Collector { Get-SPReportServiceApplicationProxies } -Columns @('Name', 'TypeName', 'Status', 'IsConnected', 'ServiceApplication', 'Id')
    Add-SectionFromCollector -Title (Get-ReportText -Key 'ManagedAccountsTitle') -Description (Get-ReportText -Key 'ManagedAccountsDescription') -Collector { Get-SPReportManagedAccounts } -Columns @('UserName', 'DisplayName', 'AutomaticChangeEnabled', 'DaysBeforeExpiryToWarn', 'PasswordLastChanged')
    Add-SectionFromCollector -Title (Get-ReportText -Key 'AamTitle') -Description (Get-ReportText -Key 'AamDescription') -Collector { Get-SPReportAam } -Columns @('WebApplication', 'Zone', 'PublicUrl', 'IncomingUrl', 'UriScheme')
    Add-SectionFromCollector -Title (Get-ReportText -Key 'FarmSolutionsTitle') -Description (Get-ReportText -Key 'FarmSolutionsDescription') -Collector { Get-SPReportSolutions } -Columns @('Name', 'Deployed', 'ContainsGlobalAssembly', 'ContainsCasPolicy', 'DeploymentState', 'LastOperationResult', 'LastOperationEndTime')
    Add-SectionFromCollector -Title (Get-ReportText -Key 'BlockedFileTypesTitle') -Description (Get-ReportText -Key 'BlockedFileTypesDescription') -Collector { Get-SPReportBlockedFileTypes } -Columns @('WebApplication', 'Extension')
    Add-SectionFromCollector -Title (Get-ReportText -Key 'UsageLoggingTitle') -Description (Get-ReportText -Key 'UsageLoggingDescription') -Collector { Get-SPReportUsageAndLogging } -Columns @('LogLocation', 'LogDiskSpaceUsageGB', 'LogMaxDiskSpaceUsageEnabled', 'DaysToKeepLogs', 'UsageServiceStatus', 'UsageLogLocation', 'UsageLogMaxSpaceGB', 'UsageLogCutTime')
    Add-SectionFromCollector -Title (Get-ReportText -Key 'OutgoingEmailTitle') -Description (Get-ReportText -Key 'OutgoingEmailDescription') -Collector { Get-SPReportOutgoingEmail } -Columns @('FarmOutboundMailService', 'SampleWebApplication', 'OutboundMailServer', 'FromAddress', 'ReplyToAddress')

    if (-not $SkipSiteCollections) {
        Add-SectionFromCollector -Title (Get-ReportText -Key 'SiteCollectionsTitle') -Description (Get-ReportText -Key 'SiteCollectionsDescription') -Collector { Get-SPReportSiteCollections } -Columns @('Url', 'Owner', 'SecondaryOwner', 'ContentDatabase', 'Template', 'CompatibilityLevel', 'StorageUsed', 'StorageQuota', 'LastContentModifiedDate', 'LockState')
    }
    else {
        Add-ReportSection -Title (Get-ReportText -Key 'SiteCollectionsTitle') -Description ('{0} -SkipSiteCollections.' -f (Get-ReportText -Key 'SkippedBy')) -Data @() -Columns @('Url', 'Owner', 'ContentDatabase') -Status 'Unknown'
    }

    if (-not $SkipTimerJobs) {
        Add-SectionFromCollector -Title (Get-ReportText -Key 'TimerJobsTitle') -Description (Get-ReportText -Key 'TimerJobsDescription') -Collector { Get-SPReportTimerJobs } -Columns @('Name', 'TypeName', 'Schedule', 'IsDisabled', 'LastRunTime', 'WebApplication', 'Server')
    }
    else {
        Add-ReportSection -Title (Get-ReportText -Key 'TimerJobsTitle') -Description ('{0} -SkipTimerJobs.' -f (Get-ReportText -Key 'SkippedBy')) -Data @() -Columns @('Name', 'Schedule', 'IsDisabled') -Status 'Unknown'
    }

    if (-not $SkipHealthAnalyzer) {
        Add-SectionFromCollector -Title (Get-ReportText -Key 'HealthAnalyzerFindingsTitle') -Description (Get-ReportText -Key 'HealthAnalyzerFindingsDescription') -Collector { Get-SPReportHealthAnalyzerFindings } -Columns @('Title', 'Category', 'Severity', 'CurrentStatus', 'Explanation', 'Remedy', 'PossibleSolution', 'FailingServers', 'FailingServices', 'Modified', 'RuleId')
        Add-SectionFromCollector -Title (Get-ReportText -Key 'HealthAnalyzerTitle') -Description (Get-ReportText -Key 'HealthAnalyzerDescription') -Collector { Get-SPReportHealthAnalyzer } -Columns @('Category', 'Summary', 'Severity', 'Enabled', 'Schedule', 'RepairAutomatically')
    }
    else {
        Add-ReportSection -Title (Get-ReportText -Key 'HealthAnalyzerTitle') -Description ('{0} -SkipHealthAnalyzer.' -f (Get-ReportText -Key 'SkippedBy')) -Data @() -Columns @('Category', 'Summary', 'Severity') -Status 'Unknown'
    }

    if (-not $SkipSearchTopology) {
        Add-SectionFromCollector -Title (Get-ReportText -Key 'SearchTopologyTitle') -Description (Get-ReportText -Key 'SearchTopologyDescription') -Collector { Get-SPReportSearchTopology } -Columns @('SearchApplication', 'ComponentName', 'ComponentType', 'ServerName', 'RootDirectory', 'IndexPartition')
        Add-SectionFromCollector -Title (Get-ReportText -Key 'SearchCrawlJobsTitle') -Description (Get-ReportText -Key 'SearchCrawlJobsDescription') -Collector { Get-SPReportSearchCrawlJobs } -Columns @('SearchApplication', 'ContentSource', 'Type', 'CrawlStatus', 'CrawlStarted', 'CrawlCompleted', 'CrawlDuration', 'IncrementalCrawlSchedule', 'FullCrawlSchedule', 'StartAddresses', 'SuccessCount', 'ErrorCount', 'DeleteCount')
    }
    else {
        Add-ReportSection -Title (Get-ReportText -Key 'SearchTopologyTitle') -Description ('{0} -SkipSearchTopology.' -f (Get-ReportText -Key 'SkippedBy')) -Data @() -Columns @('SearchApplication', 'ComponentName', 'ComponentType', 'ServerName') -Status 'Unknown'
        Add-ReportSection -Title (Get-ReportText -Key 'SearchCrawlJobsTitle') -Description ('{0} -SkipSearchTopology.' -f (Get-ReportText -Key 'SkippedBy')) -Data @() -Columns @('SearchApplication', 'ContentSource', 'CrawlStatus') -Status 'Unknown'
    }

    if (-not $SkipFeatureInventory) {
        Add-SectionFromCollector -Title (Get-ReportText -Key 'InstalledFeaturesTitle') -Description (Get-ReportText -Key 'InstalledFeaturesDescription') -Collector { Get-SPReportFeatures } -Columns @('DisplayName', 'Scope', 'Id', 'CompatibilityLevel', 'Version')
    }
    else {
        Add-ReportSection -Title (Get-ReportText -Key 'InstalledFeaturesTitle') -Description ('{0} -SkipFeatureInventory.' -f (Get-ReportText -Key 'SkippedBy')) -Data @() -Columns @('DisplayName', 'Scope', 'Id') -Status 'Unknown'
    }

    if (-not $SkipIisDetails) {
        Add-SectionFromCollector -Title (Get-ReportText -Key 'IisAppPoolsTitle') -Description (Get-ReportText -Key 'IisAppPoolsDescription') -Collector { Get-SPReportIisApplicationPools } -Columns @('Name', 'State', 'Runtime', 'PipelineMode', 'IdentityType', 'UserName', 'Enable32Bit')
        Add-SectionFromCollector -Title (Get-ReportText -Key 'IisSitesTitle') -Description (Get-ReportText -Key 'IisSitesDescription') -Collector { Get-SPReportIisSites } -Columns @('Name', 'Id', 'State', 'PhysicalPath', 'Bindings')
    }
    else {
        Add-ReportSection -Title (Get-ReportText -Key 'IisDetailsTitle') -Description ('{0} -SkipIisDetails.' -f (Get-ReportText -Key 'SkippedBy')) -Data @() -Columns @('Name', 'State') -Status 'Unknown'
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
