
<#
.SYNOPSIS
    Master Microsoft 365 Tenant Configuration Export Script

.DESCRIPTION
    Read-only collector for Microsoft 365 tenant governance review. Exports configuration from:
      - Microsoft Graph / Entra ID
      - Exchange Online and Defender for Office 365 policies
      - SharePoint Online
      - Microsoft Teams
      - Purview / Compliance
      - Intune / Endpoint Manager through Microsoft Graph
      - Power Platform
      - Optional Microsoft365DSC baseline export

.NOTES
    Author: Generated for tenant settings extraction governance review
    Recommended PowerShell: Windows PowerShell 5.1 or PowerShell 7+
    Run as: Admin PowerShell
    Authentication: Interactive delegated admin login. For full coverage, use a Global Admin or workload-specific admin with sufficient read permissions.
    Output: Timestamped folder containing JSON, CSV, TXT logs and transcript.

.EXAMPLE
    .\M365-FullTenantExport.ps1 -TenantName contoso.onmicrosoft.com -SPOAdminUrl https://contoso-admin.sharepoint.com

.EXAMPLE
    .\M365-FullTenantExport.ps1 -TenantName contoso.onmicrosoft.com -SPOAdminUrl https://contoso-admin.sharepoint.com -IncludeMicrosoft365DSC
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$TenantName,

    [Parameter(Mandatory=$false)]
    [string]$SPOAdminUrl,

    [Parameter(Mandatory=$false)]
    [string]$OutputRoot = "$env:USERPROFILE\Desktop\M365TenantExport",

    [Parameter(Mandatory=$false)]
    [switch]$SkipModuleInstall,

    [Parameter(Mandatory=$false)]
    [switch]$IncludeMicrosoft365DSC,

    [Parameter(Mandatory=$false)]
    [switch]$SkipIntuneBeta
)

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$BasePath = Join-Path $OutputRoot "TenantExport_$timestamp"
$LogPath = Join-Path $BasePath '_Logs'
$Summary = New-Object System.Collections.Generic.List[object]

function New-Folder {
    param([string]$Path)
    if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
}

function Write-Status {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -Path (Join-Path $LogPath 'run.log') -Value $line -Encoding UTF8
}

function Add-Summary {
    param([string]$Workload, [string]$Name, [string]$Status, [string]$Detail, [string]$Path)
    $Summary.Add([pscustomobject]@{
        Time     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Workload = $Workload
        Name     = $Name
        Status   = $Status
        Detail   = $Detail
        Path     = $Path
    }) | Out-Null
}

function Ensure-Module {
    param([string]$Name, [string]$MinimumVersion)
    if (Get-Module -ListAvailable -Name $Name) {
        Write-Status "Module available: $Name"
        return
    }
    if ($SkipModuleInstall) {
        Write-Status "Module missing and install skipped: $Name" 'WARN'
        return
    }
    try {
        Write-Status "Installing module: $Name"
        if ($MinimumVersion) {
            Install-Module -Name $Name -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber
        } else {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
        }
    } catch {
        Write-Status "Failed to install module $Name. $($_.Exception.Message)" 'ERROR'
    }
}

function ConvertTo-SafeFileName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'Unnamed' }
    return ($Name -replace '[\\/:*?"<>|]', '_' -replace '\s+', '_')
}

function Export-Data {
    param(
        [Parameter(Mandatory=$true)]$Data,
        [Parameter(Mandatory=$true)][string]$Folder,
        [Parameter(Mandatory=$true)][string]$Name,
        [string]$Workload = 'General'
    )
    New-Folder $Folder
    $safe = ConvertTo-SafeFileName $Name
    $jsonPath = Join-Path $Folder "$safe.json"
    $csvPath  = Join-Path $Folder "$safe.csv"
    try {
        $Data | ConvertTo-Json -Depth 100 | Out-File -FilePath $jsonPath -Encoding UTF8
        try {
            $Data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        } catch {
            # Some nested objects do not flatten well to CSV; JSON remains the source of truth.
        }
        $count = 0
        try { $count = @($Data).Count } catch { $count = 1 }
        Add-Summary -Workload $Workload -Name $Name -Status 'Success' -Detail "Exported $count object(s)" -Path $jsonPath
        Write-Status "Exported $Workload - $Name ($count object(s))"
    } catch {
        Add-Summary -Workload $Workload -Name $Name -Status 'Failed' -Detail $_.Exception.Message -Path $jsonPath
        Write-Status "Failed export $Workload - $Name. $($_.Exception.Message)" 'ERROR'
    }
}

function Invoke-ExportCommand {
    param(
        [string]$Workload,
        [string]$Name,
        [string]$Folder,
        [scriptblock]$Command
    )
    try {
        Write-Status "Collecting $Workload - $Name"
        $data = & $Command
        if ($null -eq $data) {
            Add-Summary -Workload $Workload -Name $Name -Status 'NoData' -Detail 'Command returned no data' -Path ''
            Write-Status "No data for $Workload - $Name" 'WARN'
        } else {
            Export-Data -Data $data -Folder $Folder -Name $Name -Workload $Workload
        }
    } catch {
        Add-Summary -Workload $Workload -Name $Name -Status 'Failed' -Detail $_.Exception.Message -Path ''
        Write-Status "Failed $Workload - $Name. $($_.Exception.Message)" 'ERROR'
    }
}

function Invoke-IfCmdletExists {
    param(
        [string]$Cmdlet,
        [string]$Workload,
        [string]$Name,
        [string]$Folder,
        [scriptblock]$Command
    )
    if (Get-Command $Cmdlet -ErrorAction SilentlyContinue) {
        Invoke-ExportCommand -Workload $Workload -Name $Name -Folder $Folder -Command $Command
    } else {
        Add-Summary -Workload $Workload -Name $Name -Status 'Skipped' -Detail "Cmdlet not available: $Cmdlet" -Path ''
        Write-Status "Skipped $Workload - $Name. Cmdlet not available: $Cmdlet" 'WARN'
    }
}

function Invoke-GraphExport {
    param(
        [string]$Workload,
        [string]$Name,
        [string]$Folder,
        [string]$Uri,
        [switch]$Beta
    )
    Invoke-ExportCommand -Workload $Workload -Name $Name -Folder $Folder -Command {
        $items = New-Object System.Collections.Generic.List[object]
        $next = $Uri
        do {
            $result = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
            if ($result.value) { foreach ($v in $result.value) { $items.Add($v) | Out-Null } }
            else { $items.Add($result) | Out-Null }
            $next = $result.'@odata.nextLink'
        } while ($next)
        $items
    }
}

New-Folder $BasePath
New-Folder $LogPath
Start-Transcript -Path (Join-Path $LogPath 'transcript.txt') -Force | Out-Null
Write-Status "Starting Microsoft 365 tenant export. Output folder: $BasePath"

# Module preparation
Ensure-Module Microsoft.Graph
Ensure-Module ExchangeOnlineManagement
Ensure-Module Microsoft.Online.SharePoint.PowerShell
Ensure-Module MicrosoftTeams
Ensure-Module Microsoft.PowerApps.Administration.PowerShell
Ensure-Module Microsoft.PowerApps.PowerShell
if ($IncludeMicrosoft365DSC) { Ensure-Module Microsoft365DSC }

# Import modules if available
foreach ($module in @('Microsoft.Graph','ExchangeOnlineManagement','Microsoft.Online.SharePoint.PowerShell','MicrosoftTeams','Microsoft.PowerApps.Administration.PowerShell','Microsoft.PowerApps.PowerShell')) {
    try { Import-Module $module -ErrorAction SilentlyContinue } catch { }
}

# Create workload folders
$Folders = @{}
foreach ($f in @('EntraID','ExchangeOnline','SharePointOnline','Teams','PurviewCompliance','DefenderOffice365','Intune','PowerPlatform','Microsoft365DSC','General','_Logs')) {
    $Folders[$f] = Join-Path $BasePath $f
    New-Folder $Folders[$f]
}

# Connection scopes
$GraphScopes = @(
    'Organization.Read.All','Directory.Read.All','Domain.Read.All','Policy.Read.All','RoleManagement.Read.Directory',
    'Application.Read.All','User.Read.All','Group.Read.All','AuditLog.Read.All','IdentityProvider.Read.All',
    'EntitlementManagement.Read.All','PrivilegedAccess.Read.AzureAD','DeviceManagementConfiguration.Read.All',
    'DeviceManagementApps.Read.All','DeviceManagementManagedDevices.Read.All','SecurityEvents.Read.All',
    'ThreatHunting.Read.All','Reports.Read.All','MultiTenantOrganization.Read.All'
)

try {
    Write-Status 'Connecting to Microsoft Graph'
    if ($TenantName) { Connect-MgGraph -TenantId $TenantName -Scopes $GraphScopes -NoWelcome }
    else { Connect-MgGraph -Scopes $GraphScopes -NoWelcome }
    $ctx = Get-MgContext
    Export-Data -Data $ctx -Folder $Folders['General'] -Name 'GraphContext' -Workload 'General'
} catch { Write-Status "Graph connection failed. $($_.Exception.Message)" 'ERROR' }

# General tenant / licensing / domains
Invoke-IfCmdletExists Get-MgOrganization 'General' 'Organization' $Folders['General'] { Get-MgOrganization -Property * }
Invoke-IfCmdletExists Get-MgSubscribedSku 'General' 'SubscribedSKUs' $Folders['General'] { Get-MgSubscribedSku -All }
Invoke-IfCmdletExists Get-MgDomain 'General' 'Domains' $Folders['General'] { Get-MgDomain -All }
Invoke-IfCmdletExists Get-MgDirectoryRole 'General' 'DirectoryRoles' $Folders['General'] { Get-MgDirectoryRole -All }
Invoke-IfCmdletExists Get-MgDirectoryRoleTemplate 'General' 'DirectoryRoleTemplates' $Folders['General'] { Get-MgDirectoryRoleTemplate -All }

# Entra ID via Graph cmdlets
Invoke-IfCmdletExists Get-MgPolicyAuthorizationPolicy 'EntraID' 'AuthorizationPolicy' $Folders['EntraID'] { Get-MgPolicyAuthorizationPolicy }
Invoke-IfCmdletExists Get-MgPolicyAuthenticationMethodPolicy 'EntraID' 'AuthenticationMethodPolicy' $Folders['EntraID'] { Get-MgPolicyAuthenticationMethodPolicy }
Invoke-IfCmdletExists Get-MgIdentityConditionalAccessPolicy 'EntraID' 'ConditionalAccessPolicies' $Folders['EntraID'] { Get-MgIdentityConditionalAccessPolicy -All }
Invoke-IfCmdletExists Get-MgIdentityConditionalAccessNamedLocation 'EntraID' 'ConditionalAccessNamedLocations' $Folders['EntraID'] { Get-MgIdentityConditionalAccessNamedLocation -All }
Invoke-IfCmdletExists Get-MgPolicyCrossTenantAccessPolicy 'EntraID' 'CrossTenantAccessPolicy' $Folders['EntraID'] { Get-MgPolicyCrossTenantAccessPolicy }
Invoke-IfCmdletExists Get-MgPolicyCrossTenantAccessPolicyDefault 'EntraID' 'CrossTenantAccessPolicyDefault' $Folders['EntraID'] { Get-MgPolicyCrossTenantAccessPolicyDefault }
Invoke-IfCmdletExists Get-MgPolicyCrossTenantAccessPolicyPartner 'EntraID' 'CrossTenantAccessPolicyPartners' $Folders['EntraID'] { Get-MgPolicyCrossTenantAccessPolicyPartner -All }
Invoke-IfCmdletExists Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy 'EntraID' 'SecurityDefaults' $Folders['EntraID'] { Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy }
Invoke-IfCmdletExists Get-MgDirectoryAdministrativeUnit 'EntraID' 'AdministrativeUnits' $Folders['EntraID'] { Get-MgDirectoryAdministrativeUnit -All }
Invoke-IfCmdletExists Get-MgServicePrincipal 'EntraID' 'ServicePrincipals' $Folders['EntraID'] { Get-MgServicePrincipal -All -Property Id,AppId,DisplayName,AccountEnabled,AppOwnerOrganizationId,PublisherName,SignInAudience,ServicePrincipalType,Tags,CreatedDateTime }
Invoke-IfCmdletExists Get-MgApplication 'EntraID' 'Applications' $Folders['EntraID'] { Get-MgApplication -All -Property Id,AppId,DisplayName,SignInAudience,PublisherDomain,CreatedDateTime,RequiredResourceAccess,Web,Spa,PublicClient,Api }
#Invoke-IfCmdletExists Get-MgGroup 'EntraID' 'Groups' $Folders['EntraID'] { Get-MgGroup -All -Property Id,DisplayName,Mail,MailEnabled,SecurityEnabled,GroupTypes,Visibility,CreatedDateTime }
#Invoke-IfCmdletExists Get-MgUser 'EntraID' 'Users_BasicInventory' $Folders['EntraID'] { Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled,UserType,CreatedDateTime,Department,JobTitle,OnPremisesSyncEnabled }
#Invoke-IfCmdletExists Get-MgDevice 'EntraID' 'Devices' $Folders['EntraID'] { Get-MgDevice -All -Property Id,DisplayName,AccountEnabled,OperatingSystem,OperatingSystemVersion,TrustType,ApproximateLastSignInDateTime,IsCompliant,IsManaged }

# Additional Entra / Graph REST exports
if (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
    Invoke-GraphExport 'EntraID' 'DirectorySettings' $Folders['EntraID'] '/v1.0/settings'
    Invoke-GraphExport 'EntraID' 'OrganizationBranding' $Folders['EntraID'] '/v1.0/organization?$expand=branding'
    Invoke-GraphExport 'EntraID' 'IdentityProviders' $Folders['EntraID'] '/v1.0/identity/identityProviders'
    Invoke-GraphExport 'EntraID' 'AccessPackageCatalogs' $Folders['EntraID'] '/v1.0/identityGovernance/entitlementManagement/catalogs'
    Invoke-GraphExport 'EntraID' 'AccessPackages' $Folders['EntraID'] '/v1.0/identityGovernance/entitlementManagement/accessPackages'
    Invoke-GraphExport 'EntraID' 'TermsOfUse' $Folders['EntraID'] '/v1.0/identityGovernance/termsOfUse/agreements'
    Invoke-GraphExport 'EntraID' 'MultiTenantOrganization' $Folders['EntraID'] '/v1.0/tenantRelationships/multiTenantOrganization'
}

# Exchange Online / Defender for Office 365 / Purview connection
try {
    Write-Status 'Connecting to Exchange Online'
    Connect-ExchangeOnline -ShowBanner:$false
} catch { Write-Status "Exchange Online connection failed. $($_.Exception.Message)" 'ERROR' }

# Exchange Online core
$exoFolder = $Folders['ExchangeOnline']
$exoCmds = @(
    'Get-OrganizationConfig','Get-AcceptedDomain','Get-RemoteDomain','Get-MailboxPlan','Get-RoleAssignmentPolicy','Get-ManagementRoleAssignment',
    'Get-TransportConfig','Get-TransportRule','Get-HostedOutboundSpamFilterPolicy','Get-InboundConnector','Get-OutboundConnector',
    'Get-OrganizationRelationship','Get-SharingPolicy','Get-OwaMailboxPolicy','Get-MobileDeviceMailboxPolicy','Get-ActiveSyncOrganizationSettings',
    'Get-AddressBookPolicy','Get-AddressList','Get-GlobalAddressList','Get-EmailAddressPolicy','Get-DistributionGroup',
    'Get-UnifiedGroup','Get-CASMailboxPlan' 
    #,'Get-MailboxAutoReplyConfiguration'
)
foreach ($cmd in $exoCmds) {
    Invoke-IfCmdletExists $cmd 'ExchangeOnline' $cmd $exoFolder ([scriptblock]::Create("$cmd | Select-Object *"))
}

# Defender for Office 365 policies available through EXO session
$defFolder = $Folders['DefenderOffice365']
$defCmds = @(
    'Get-AntiPhishPolicy','Get-AntiPhishRule','Get-MalwareFilterPolicy','Get-MalwareFilterRule','Get-HostedContentFilterPolicy','Get-HostedContentFilterRule',
    'Get-SafeAttachmentPolicy','Get-SafeAttachmentRule','Get-SafeLinksPolicy','Get-SafeLinksRule','Get-AtpPolicyForO365','Get-QuarantinePolicy',
    'Get-TenantAllowBlockListItems','Get-EmailTenantSettings','Get-ReportSubmissionPolicy','Get-ExternalInOutlook'
)
foreach ($cmd in $defCmds) {
    Invoke-IfCmdletExists $cmd 'DefenderOffice365' $cmd $defFolder ([scriptblock]::Create("$cmd | Select-Object *"))
}

# Purview / Compliance cmdlets through ExchangeOnlineManagement IPPS session if available
try {
    Write-Status 'Connecting to Purview compliance PowerShell'
    Connect-IPPSSession -ShowBanner:$false
} catch { Write-Status "Purview compliance connection failed or unavailable. $($_.Exception.Message)" 'WARN' }

$purvFolder = $Folders['PurviewCompliance']
$purvCmds = @(
    'Get-DlpCompliancePolicy','Get-DlpComplianceRule','Get-RetentionCompliancePolicy','Get-RetentionComplianceRule','Get-Label','Get-LabelPolicy',
    'Get-DataClassificationConfig','Get-ComplianceTag','Get-CaseHoldPolicy','Get-DeviceConditionalAccessPolicy','Get-InformationBarrierPolicy',
    'Get-InsiderRiskPolicy','Get-SupervisoryReviewPolicyV2','Get-CommunicationCompliancePolicy'
)
foreach ($cmd in $purvCmds) {
    Invoke-IfCmdletExists $cmd 'PurviewCompliance' $cmd $purvFolder ([scriptblock]::Create("$cmd | Select-Object *"))
}

# SharePoint Online
if ($SPOAdminUrl) {
    try {
        Write-Status "Connecting to SharePoint Online: $SPOAdminUrl"
        Connect-SPOService -Url $SPOAdminUrl
    } catch { Write-Status "SharePoint Online connection failed. $($_.Exception.Message)" 'ERROR' }

    $spoFolder = $Folders['SharePointOnline']
    Invoke-IfCmdletExists Get-SPOTenant 'SharePointOnline' 'TenantSettings' $spoFolder { Get-SPOTenant }
    Invoke-IfCmdletExists Get-SPOSite 'SharePointOnline' 'Sites_All' $spoFolder { Get-SPOSite -Limit All }
    Invoke-IfCmdletExists Get-SPOSite 'SharePointOnline' 'Sites_Sharing' $spoFolder { Get-SPOSite -Limit All | Select-Object Url,Owner,Template,StorageQuota,SharingCapability,DisableCompanyWideSharingLinks,DefaultSharingLinkType,DefaultLinkPermission,ConditionalAccessPolicy }
    Invoke-IfCmdletExists Get-SPODeletedSite 'SharePointOnline' 'DeletedSites' $spoFolder { Get-SPODeletedSite -Limit All }
    Invoke-IfCmdletExists Get-SPOTenantCdnEnabled 'SharePointOnline' 'TenantCdnPublicEnabled' $spoFolder { [pscustomobject]@{ PublicCdnEnabled = Get-SPOTenantCdnEnabled -CdnType Public } }
    Invoke-IfCmdletExists Get-SPOTenantCdnEnabled 'SharePointOnline' 'TenantCdnPrivateEnabled' $spoFolder { [pscustomobject]@{ PrivateCdnEnabled = Get-SPOTenantCdnEnabled -CdnType Private } }
    #Invoke-IfCmdletExists Get-SPOAppErrors 'SharePointOnline' 'AppErrors' $spoFolder { Get-SPOAppErrors }
    Invoke-IfCmdletExists Get-SPOOrgAssetsLibrary 'SharePointOnline' 'OrgAssetsLibraries' $spoFolder { Get-SPOOrgAssetsLibrary }
} else {
    Write-Status 'SPOAdminUrl not provided. SharePoint Online export skipped.' 'WARN'
    Add-Summary -Workload 'SharePointOnline' -Name 'Connection' -Status 'Skipped' -Detail 'SPOAdminUrl not provided' -Path ''
}

# Microsoft Teams
try {
    Write-Status 'Connecting to Microsoft Teams'
    Connect-MicrosoftTeams | Out-Null
} catch { Write-Status "Teams connection failed. $($_.Exception.Message)" 'ERROR' }

$teamsFolder = $Folders['Teams']
$teamsCmds = @(
    'Get-CsTenant','Get-CsTeamsMessagingPolicy','Get-CsTeamsMeetingPolicy','Get-CsTeamsCallingPolicy','Get-CsTeamsEventsPolicy',
    'Get-CsTeamsAppPermissionPolicy','Get-CsTeamsAppSetupPolicy','Get-CsTeamsChannelsPolicy','Get-CsTeamsFilesPolicy',
    'Get-CsTeamsGuestMessagingConfiguration','Get-CsTeamsGuestMeetingConfiguration','Get-CsTeamsGuestCallingConfiguration',
    'Get-CsTenantFederationConfiguration','Get-CsTenantMigrationConfiguration','Get-CsExternalAccessPolicy','Get-CsTeamsUpgradePolicy',
    #'Get-Team','Get-TeamChannel',
    'Get-TeamsApp'
)
foreach ($cmd in $teamsCmds) {
    if ($cmd -eq 'Get-TeamChannel') {
        Invoke-IfCmdletExists Get-TeamChannel 'Teams' 'TeamChannels' $teamsFolder {
            $teams = Get-Team
            foreach ($t in $teams) { Get-TeamChannel -GroupId $t.GroupId | Select-Object *, @{n='TeamDisplayName';e={$t.DisplayName}}, @{n='GroupId';e={$t.GroupId}} }
        }
    } else {
        Invoke-IfCmdletExists $cmd 'Teams' $cmd $teamsFolder ([scriptblock]::Create("$cmd | Select-Object *"))
    }
}

# Intune / Endpoint Manager via Graph
if (-not $SkipIntuneBeta -and (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
    $intuneFolder = $Folders['Intune']
    $intuneEndpoints = @{
        'DeviceCompliancePolicies' = '/beta/deviceManagement/deviceCompliancePolicies'
        'DeviceConfigurations' = '/beta/deviceManagement/deviceConfigurations'
        'ConfigurationPolicies_SettingsCatalog' = '/beta/deviceManagement/configurationPolicies'
        'GroupPolicyConfigurations' = '/beta/deviceManagement/groupPolicyConfigurations'
        'DeviceManagementScripts' = '/beta/deviceManagement/deviceManagementScripts'
        'DetectedApps' = '/beta/deviceManagement/detectedApps'
        'MobileApps' = '/beta/deviceAppManagement/mobileApps'
        'AppProtectionPolicies_iOS' = '/beta/deviceAppManagement/iosManagedAppProtections'
        'AppProtectionPolicies_Android' = '/beta/deviceAppManagement/androidManagedAppProtections'
        'DeviceEnrollmentConfigurations' = '/beta/deviceManagement/deviceEnrollmentConfigurations'
        'ManagedDevices' = '/beta/deviceManagement/managedDevices'
        'WindowsAutopilotDeploymentProfiles' = '/beta/deviceManagement/windowsAutopilotDeploymentProfiles'
        'TermsAndConditions' = '/beta/deviceManagement/termsAndConditions'
        'IntuneRBACRoleDefinitions' = '/beta/deviceManagement/roleDefinitions'
        'IntuneRBACRoleAssignments' = '/beta/deviceManagement/roleAssignments'
    }
    foreach ($key in $intuneEndpoints.Keys) { Invoke-GraphExport 'Intune' $key $intuneFolder $intuneEndpoints[$key] }
}

# Power Platform
try {
    Write-Status 'Connecting to Power Platform admin endpoint'
    Add-PowerAppsAccount | Out-Null
} catch { Write-Status "Power Platform connection failed. $($_.Exception.Message)" 'ERROR' }

$ppFolder = $Folders['PowerPlatform']
$ppCmds = @(
    'Get-TenantSettings','Get-AdminPowerAppEnvironment','Get-AdminPowerApp','Get-AdminFlow','Get-AdminPowerAppConnector',
    'Get-AdminDlpPolicy','Get-AdminPowerAppConnection','Get-AdminPowerAppConnectorRoleAssignment','Get-AdminPowerAppRoleAssignment'
)
foreach ($cmd in $ppCmds) {
    Invoke-IfCmdletExists $cmd 'PowerPlatform' $cmd $ppFolder ([scriptblock]::Create("$cmd | Select-Object *"))
}

# Optional Microsoft365DSC baseline snapshot
if ($IncludeMicrosoft365DSC) {
    $dscFolder = $Folders['Microsoft365DSC']
    Invoke-IfCmdletExists Export-M365DSCConfiguration 'Microsoft365DSC' 'M365DSC_Export_WebUI' $dscFolder {
        Write-Status 'Launching Microsoft365DSC Web UI. Select desired workloads and export location when prompted.'
        Export-M365DSCConfiguration -LaunchWebUI
    }
}

# Final summary
try {
    $summaryCsv = Join-Path $BasePath 'ExportSummary.csv'
    $summaryJson = Join-Path $BasePath 'ExportSummary.json'
    $Summary | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8
    $Summary | ConvertTo-Json -Depth 20 | Out-File -FilePath $summaryJson -Encoding UTF8

    $readme = @"
Microsoft 365 Tenant Export Summary
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Output Folder: $BasePath

Folders:
- General: tenant, licensing, domains, directory roles
- EntraID: identity, CA, auth methods, apps, service principals, cross-tenant settings
- ExchangeOnline: organization, mail flow, policies, connector and mailbox plan configuration
- DefenderOffice365: anti-phish, anti-spam, anti-malware, Safe Links, Safe Attachments, tenant allow/block
- PurviewCompliance: DLP, retention, sensitivity labels and compliance policies
- SharePointOnline: tenant settings, sites and sharing settings
- Teams: Teams tenant and policy settings
- Intune: device management, compliance, configuration, apps and enrollment settings
- PowerPlatform: tenant settings, environments, apps, flows, connectors and DLP policies
- Microsoft365DSC: optional DSC baseline export when enabled

Important:
- This script is read-only except for module installation and login/session creation.
- Some exports may be skipped if the signed-in admin lacks permissions or if a cmdlet is not available in the installed module version.
- JSON output is the source of truth for nested settings. CSV is included for quick review where possible.
"@
    $readme | Out-File -FilePath (Join-Path $BasePath 'README.txt') -Encoding UTF8
    Write-Status "Completed. Summary: $summaryCsv"
} finally {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    try { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue } catch { }
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
    Stop-Transcript | Out-Null
}
