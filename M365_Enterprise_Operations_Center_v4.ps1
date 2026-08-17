<#
.SYNOPSIS
    Microsoft 365 Enterprise Operations Center - WPF GUI Dashboard
.VERSION
    4.0
.AUTHOR
    Generated for Md. Raquibul Hoque
.DESCRIPTION
    Read-only Microsoft 365 operations dashboard with fixed PowerShell syntax, stronger Exchange Online connection handling,
    Secure Score direct Graph API fallback, license summary, MFA registration status, mail flow failures, Intune device compliance,
    tenant health score, service health, sign-in/risky user/domain status, and export options.

.NOTES
    Run PowerShell as Administrator the first time if modules need installation.
    This script is read-only. It does not modify tenant settings.

    Microsoft Graph scopes requested:
      ServiceHealth.Read.All
      Directory.Read.All
      User.Read.All
      AuditLog.Read.All
      SecurityEvents.Read.All
      IdentityRiskyUser.Read.All
      Reports.Read.All
      UserAuthenticationMethod.Read.All
      DeviceManagementManagedDevices.Read.All
      DeviceManagementConfiguration.Read.All
      Organization.Read.All
      LicenseAssignment.Read.All

    Exchange Online is optional. If Exchange connection fails, Mail Flow will show N/A and the rest of the dashboard will continue.
#>

#region Global Settings
$ErrorActionPreference = 'Continue'
$Script:DashboardVersion = '4.0 Fixed Enterprise Operations Center'
$Script:ConnectedGraph = $false
$Script:ConnectedExchange = $false
$Script:TenantInfo = [ordered]@{}
$Script:ReportFolder = Join-Path $env:USERPROFILE 'Documents\M365EnterpriseOperationsReports'
if (-not (Test-Path $Script:ReportFolder)) { New-Item -ItemType Directory -Path $Script:ReportFolder -Force | Out-Null }

$Script:HealthItems = @()
$Script:Messages = @()
$Script:Licenses = @()
$Script:Users = @()
$Script:RiskyUsers = @()
$Script:FailedSignIns = @()
$Script:Domains = @()
$Script:ManagedDevices = @()
$Script:SecureScorePercent = $null
$Script:SecureScoreText = 'N/A'

$Script:UsersSummary = @{ Total=0; Enabled=0; Disabled=0; Licensed=0; Guest=0 }
$Script:LicenseSummary = @{ Total=0; Used=0; Available=0; ExhaustedSkus=0 }
$Script:MfaSummary = @{ WithMfa=0; WithoutMfa=0; Checked=0; Percent=0 }
$Script:DeviceSummary = @{ Total=0; Compliant=0; NonCompliant=0; Unknown=0; CompliancePercent=0 }
$Script:MailFlowSummary = @{ Failed24h=0; Checked=$false; Error='N/A' }
$Script:TenantHealthScore = @{ Score=0; Rating='N/A'; Details='' }

$Script:Issues = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Recommendations = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:LicenseDetails = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:MfaDetails = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:DeviceDetails = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:MailFlowDetails = New-Object System.Collections.ObjectModel.ObservableCollection[object]
#endregion Global Settings

#region Helper Functions
function Write-DashboardLog {
    param([string]$Message,[string]$Level='INFO')
    $line = "[{0}][{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    if ($Global:txtLog) {
        try {
            $Global:txtLog.Dispatcher.Invoke([action]{
                $Global:txtLog.AppendText($line + [Environment]::NewLine)
                $Global:txtLog.ScrollToEnd()
            })
        } catch {}
    }
    Write-Host $line
}

function Test-IsAdmin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    } catch { return $false }
}

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    try {
        if (-not (Get-Module -ListAvailable -Name $Name)) {
            Write-DashboardLog "Installing module $Name ..." 'WARN'
            Install-Module $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        Import-Module $Name -Force -ErrorAction Stop
        Write-DashboardLog "Loaded module $Name"
        return $true
    } catch {
        Write-DashboardLog "Failed to load/install module $Name. $($_.Exception.Message)" 'ERROR'
        return $false
    }
}

function Safe-Run {
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock,[string]$Name='Operation',[object]$Default=@())
    try { return & $ScriptBlock }
    catch {
        Write-DashboardLog "$Name failed: $($_.Exception.Message)" 'ERROR'
        return $Default
    }
}

function Add-Issue {
    param([string]$Area,[string]$Status,[string]$Severity,[string]$Title,[string]$Details,[string]$Recommendation)
    $Script:Issues.Add([pscustomobject]@{
        Area=$Area; Status=$Status; Severity=$Severity; Title=$Title; Details=$Details; Recommendation=$Recommendation; Time=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }) | Out-Null
}

function Add-Recommendation {
    param([string]$Area,[string]$Recommendation,[string]$Priority='Medium')
    $Script:Recommendations.Add([pscustomobject]@{ Area=$Area; Priority=$Priority; Recommendation=$Recommendation }) | Out-Null
}

function Set-TextSafe {
    param([object]$Control,[string]$Text)
    if ($Control) {
        try { $Control.Dispatcher.Invoke([action]{ $Control.Text = $Text }) } catch {}
    }
}

function Set-CardStatus {
    param([string]$Name,[string]$Text,[string]$Status)
    $color = switch ($Status) {
        'Good' { '#22C55E' }
        'Warning' { '#F59E0B' }
        'Critical' { '#EF4444' }
        'Info' { '#3B82F6' }
        default { '#64748B' }
    }
    $label = Get-Variable -Name "lbl$Name" -Scope Global -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Value
    $border = Get-Variable -Name "card$Name" -Scope Global -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Value
    if ($label) { try { $label.Dispatcher.Invoke([action]{ $label.Text = $Text }) } catch {} }
    if ($border) { try { $border.Dispatcher.Invoke([action]{ $border.Background = $color }) } catch {} }
}

function Get-StatusByCount {
    param([int]$Critical,[int]$Warning)
    if ($Critical -gt 0) { return 'Critical' }
    if ($Warning -gt 0) { return 'Warning' }
    return 'Good'
}

function Get-PercentStatus {
    param([double]$Percent,[double]$GoodAt=90,[double]$WarnAt=70)
    if ($Percent -ge $GoodAt) { return 'Good' }
    if ($Percent -ge $WarnAt) { return 'Warning' }
    return 'Critical'
}

function ConvertTo-HtmlSafe {
    param([object]$Text)
    if ($null -eq $Text) { return '' }
    try { Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue } catch {}
    return [System.Web.HttpUtility]::HtmlEncode([string]$Text)
}
#endregion Helper Functions

#region Connection Functions
function Connect-M365Dashboard {
    Write-DashboardLog 'Starting Microsoft 365 Graph connection...'

    $modules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Identity.SignIns',
        'Microsoft.Graph.Reports',
        'Microsoft.Graph.Security',
        'Microsoft.Graph.Devices.ServiceAnnouncement',
        'Microsoft.Graph.DeviceManagement'
    )
    foreach ($m in $modules) { [void](Ensure-Module $m) }

    $scopes = @(
        'ServiceHealth.Read.All',
        'Directory.Read.All',
        'User.Read.All',
        'AuditLog.Read.All',
        'SecurityEvents.Read.All',
        'IdentityRiskyUser.Read.All',
        'Reports.Read.All',
        'UserAuthenticationMethod.Read.All',
        'DeviceManagementManagedDevices.Read.All',
        'DeviceManagementConfiguration.Read.All',
        'Organization.Read.All',
        'LicenseAssignment.Read.All'
    )

    try {
        Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop | Out-Null
        $ctx = Get-MgContext
        $Script:ConnectedGraph = $true
        $Script:TenantInfo.Account = $ctx.Account
        $Script:TenantInfo.TenantId = $ctx.TenantId
        Set-TextSafe $Global:txtTenant "Graph connected: $($ctx.Account) | Tenant: $($ctx.TenantId)"
        Write-DashboardLog "Connected to Microsoft Graph as $($ctx.Account)"
    } catch {
        $Script:ConnectedGraph = $false
        Set-TextSafe $Global:txtTenant 'Graph not connected'
        Write-DashboardLog "Graph connection failed: $($_.Exception.Message)" 'ERROR'
    }

    # Optional Exchange Online connection. The rest of the dashboard continues if this fails.
    $Script:ConnectedExchange = $false
    if (Ensure-Module 'ExchangeOnlineManagement') {
        try {
            Write-DashboardLog 'Connecting to Exchange Online for mail flow statistics...'
            # Do not pass UserPrincipalName here. This avoids the null reference issue some tenants/modules hit.
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
            $Script:ConnectedExchange = $true
            Write-DashboardLog 'Connected to Exchange Online.'
        } catch {
            $Script:ConnectedExchange = $false
            Write-DashboardLog "Exchange Online connection failed. Mail flow section will show N/A. $($_.Exception.Message)" 'WARN'
        }
    }
}
#endregion Connection Functions

#region Data Collection Functions
function Get-ServiceHealthData {
    Write-DashboardLog 'Collecting Service Health incidents...'
    $Script:HealthItems = Safe-Run -Name 'Service health issues' -Default @() -ScriptBlock {
        Get-MgServiceAnnouncementIssue -All | Sort-Object LastModifiedDateTime -Descending
    }
    foreach ($i in $Script:HealthItems | Select-Object -First 100) {
        $sev = if ($i.Classification -match 'Incident') {'Critical'} elseif ($i.Status -match 'ServiceDegradation|Investigating') {'High'} else {'Medium'}
        Add-Issue -Area 'Service Health' -Status $i.Status -Severity $sev -Title $i.Title -Details "$($i.Service) | $($i.FeatureGroup) | $($i.LastModifiedDateTime)" -Recommendation 'Review Microsoft 365 admin center Service health details and notify impacted users if required.'
    }
    $active = @($Script:HealthItems | Where-Object { $_.Status -notmatch 'serviceRestored|resolved|postIncidentReviewPublished' })
    if ($active.Count -gt 0) { Add-Recommendation -Area 'Service Health' -Recommendation 'Review active incidents/advisories and communicate impact to users/customers.' -Priority 'High' }
    return $active.Count
}

function Get-MessageCenterData {
    Write-DashboardLog 'Collecting Message Center announcements...'
    $Script:Messages = Safe-Run -Name 'Message center' -Default @() -ScriptBlock {
        Get-MgServiceAnnouncementMessage -All | Sort-Object LastModifiedDateTime -Descending
    }
    foreach ($m in $Script:Messages | Select-Object -First 50) {
        $cat = if ($m.IsMajorChange) {'High'} else {'Info'}
        Add-Issue -Area 'Message Center' -Status 'Announcement' -Severity $cat -Title $m.Title -Details "Category: $($m.Category) | Last Modified: $($m.LastModifiedDateTime)" -Recommendation 'Check whether this change affects tenant configuration, users, or customer communication.'
    }
    return @($Script:Messages | Where-Object {$_.IsMajorChange}).Count
}

function Get-LicenseData {
    Write-DashboardLog 'Collecting license summary and per-SKU details...'
    $Script:LicenseDetails.Clear()
    $Script:Licenses = Safe-Run -Name 'Subscribed SKUs' -Default @() -ScriptBlock { Get-MgSubscribedSku -All }

    $totalAll = 0; $usedAll = 0; $availableAll = 0; $exhausted = 0
    foreach ($sku in $Script:Licenses) {
        $total = [int]$sku.PrepaidUnits.Enabled
        $used = [int]$sku.ConsumedUnits
        $available = $total - $used
        if ($available -lt 0) { $available = 0 }
        $pct = if ($total -gt 0) { [math]::Round(($used / $total) * 100, 2) } else { 0 }
        $status = if ($available -le 0) {'Critical'} elseif ($pct -ge 90) {'Warning'} else {'Good'}
        $sev = if ($available -le 0) {'High'} elseif ($pct -ge 90) {'Medium'} else {'Info'}
        if ($available -le 0) { $exhausted++ }
        $totalAll += $total; $usedAll += $used; $availableAll += $available
        $Script:LicenseDetails.Add([pscustomobject]@{
            LicenseName=$sku.SkuPartNumber; Total=$total; Assigned=$used; Available=$available; UsagePercent="$pct%"; Status=$status
        }) | Out-Null
        Add-Issue -Area 'Licenses' -Status "$used/$total used" -Severity $sev -Title $sku.SkuPartNumber -Details "Available: $available | Usage: $pct%" -Recommendation 'Review unused or exhausted licenses and plan procurement or cleanup if needed.'
    }
    $Script:LicenseSummary = @{ Total=$totalAll; Used=$usedAll; Available=$availableAll; ExhaustedSkus=$exhausted }
    if ($exhausted -gt 0) { Add-Recommendation -Area 'Licenses' -Recommendation 'One or more license SKUs have no available capacity. Review assignments and procurement.' -Priority 'High' }
    return $exhausted
}

function Get-UsersSummaryData {
    Write-DashboardLog 'Collecting user summary...'
    $Script:Users = Safe-Run -Name 'Users' -Default @() -ScriptBlock {
        Get-MgUser -All -Property Id,UserPrincipalName,AccountEnabled,DisplayName,UserType,AssignedLicenses
    }
    $enabled = @($Script:Users | Where-Object {$_.AccountEnabled}).Count
    $disabled = @($Script:Users | Where-Object {-not $_.AccountEnabled}).Count
    $licensed = @($Script:Users | Where-Object {$_.AssignedLicenses.Count -gt 0}).Count
    $guest = @($Script:Users | Where-Object {$_.UserType -eq 'Guest'}).Count
    $Script:UsersSummary = @{ Total=$Script:Users.Count; Enabled=$enabled; Disabled=$disabled; Licensed=$licensed; Guest=$guest }
    Add-Issue -Area 'Users' -Status 'Summary' -Severity 'Info' -Title 'User account summary' -Details "Total: $($Script:Users.Count) | Enabled: $enabled | Disabled: $disabled | Licensed: $licensed | Guest: $guest" -Recommendation 'Review inactive users, disabled accounts with licenses, and guest accounts periodically.'
    return $Script:UsersSummary
}

function Get-UsersWithoutMfaData {
    Write-DashboardLog 'Checking MFA registration methods for enabled member users...'
    $Script:MfaDetails.Clear()
    if (-not $Script:Users -or $Script:Users.Count -eq 0) { [void](Get-UsersSummaryData) }

    $enabledUsers = @($Script:Users | Where-Object { $_.AccountEnabled -eq $true -and $_.UserType -ne 'Guest' })
    $withoutMfa = 0; $withMfa = 0; $checked = 0

    foreach ($u in $enabledUsers) {
        $checked++
        $methods = Safe-Run -Name "Authentication methods for $($u.UserPrincipalName)" -Default @() -ScriptBlock {
            Get-MgUserAuthenticationMethod -UserId $u.Id
        }
        $mfaMethods = @($methods | Where-Object {
            $odata = $_.AdditionalProperties['@odata.type']
            $odata -and $odata -notmatch 'passwordAuthenticationMethod'
        })
        if ($mfaMethods.Count -gt 0) {
            $withMfa++
            $methodTypes = ($mfaMethods | ForEach-Object { $_.AdditionalProperties['@odata.type'] -replace '#microsoft.graph.','' }) -join ', '
            $Script:MfaDetails.Add([pscustomobject]@{ DisplayName=$u.DisplayName; UserPrincipalName=$u.UserPrincipalName; AccountEnabled=$u.AccountEnabled; MfaStatus='Registered'; Methods=$methodTypes }) | Out-Null
        } else {
            $withoutMfa++
            $Script:MfaDetails.Add([pscustomobject]@{ DisplayName=$u.DisplayName; UserPrincipalName=$u.UserPrincipalName; AccountEnabled=$u.AccountEnabled; MfaStatus='Not Registered'; Methods='None or password only' }) | Out-Null
        }
    }
    $pct = if ($checked -gt 0) { [math]::Round(($withMfa / $checked) * 100, 2) } else { 0 }
    $Script:MfaSummary = @{ WithMfa=$withMfa; WithoutMfa=$withoutMfa; Checked=$checked; Percent=$pct }
    if ($withoutMfa -gt 0) {
        Add-Issue -Area 'MFA' -Status "$withoutMfa without MFA" -Severity 'High' -Title 'Users without MFA registration method' -Details "Checked: $checked | Registered: $withMfa | Without MFA: $withoutMfa | Coverage: $pct%" -Recommendation 'Require MFA registration using Conditional Access and authentication methods registration campaign.'
        Add-Recommendation -Area 'MFA' -Recommendation 'Prioritize MFA registration for all enabled member users, especially admins and high-risk users.' -Priority 'High'
    } else {
        Add-Issue -Area 'MFA' -Status 'Healthy' -Severity 'Info' -Title 'MFA registration status' -Details "Checked: $checked | Registered: $withMfa | Coverage: $pct%" -Recommendation 'Continue periodic MFA registration review.'
    }
    return $withoutMfa
}

function Get-SecurityScoreData {
    Write-DashboardLog 'Collecting Secure Score using direct Microsoft Graph API...'
    $Script:SecureScorePercent = $null
    $Script:SecureScoreText = 'N/A'
    $scoreStatus = 'Warning'

    $response = Safe-Run -Name 'Secure Score direct Graph API' -Default $null -ScriptBlock {
        Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/security/secureScores?$top=1'
    }

    try {
        $scoreObj = $null
        if ($response -is [System.Collections.IDictionary]) {
            if ($response.ContainsKey('value')) { $scoreObj = @($response['value'] | Select-Object -First 1)[0] }
        } elseif ($response -and $response.value) {
            $scoreObj = @($response.value | Select-Object -First 1)[0]
        }

        if ($null -ne $scoreObj) {
            $currentScore = $null; $maxScore = $null; $createdDate = $null
            if ($scoreObj -is [System.Collections.IDictionary]) {
                $currentScore = $scoreObj['currentScore']; $maxScore = $scoreObj['maxScore']; $createdDate = $scoreObj['createdDateTime']
            } else {
                $currentScore = $scoreObj.currentScore; $maxScore = $scoreObj.maxScore; $createdDate = $scoreObj.createdDateTime
            }

            if ($null -ne $currentScore -and $null -ne $maxScore -and [double]$maxScore -gt 0) {
                $pct = [math]::Round(([double]$currentScore / [double]$maxScore) * 100, 2)
                $Script:SecureScorePercent = $pct
                $Script:SecureScoreText = "$currentScore / $maxScore ($pct%)"
                $scoreStatus = if ($pct -ge 70) {'Good'} elseif ($pct -ge 50) {'Warning'} else {'Critical'}
                Add-Issue -Area 'Secure Score' -Status 'Collected' -Severity 'Info' -Title 'Microsoft Secure Score' -Details "$($Script:SecureScoreText) | Date: $createdDate" -Recommendation 'Review Secure Score improvement actions in Microsoft Defender portal.'
                Add-Recommendation -Area 'Secure Score' -Recommendation 'Review Secure Score improvement actions such as MFA, legacy authentication blocking, admin role hardening, and device compliance.' -Priority 'Medium'
            } else {
                Add-Issue -Area 'Secure Score' -Status 'Value empty' -Severity 'Medium' -Title 'Secure Score returned without score values' -Details 'currentScore or maxScore was empty in Graph response.' -Recommendation 'Confirm SecurityEvents.Read.All permission, admin consent, licensing, and Microsoft Defender security portal access.'
            }
        } else {
            Add-Issue -Area 'Secure Score' -Status 'No data' -Severity 'Medium' -Title 'Secure Score not returned' -Details 'Graph response did not contain secureScores value.' -Recommendation 'Confirm SecurityEvents.Read.All permission, admin consent, licensing, and security portal access.'
        }
    } catch {
        Write-DashboardLog "Secure Score parsing failed: $($_.Exception.Message)" 'ERROR'
        Add-Issue -Area 'Secure Score' -Status 'Error' -Severity 'High' -Title 'Secure Score parse error' -Details $_.Exception.Message -Recommendation 'Check Graph permissions and retry after updating Microsoft.Graph modules.'
    }

    Set-CardStatus 'SecureScore' $Script:SecureScoreText $scoreStatus
    return $Script:SecureScoreText
}

function Get-RiskyUsersData {
    Write-DashboardLog 'Collecting risky users...'
    $Script:RiskyUsers = Safe-Run -Name 'Risky users' -Default @() -ScriptBlock { Get-MgRiskyUser -All | Sort-Object RiskLastUpdatedDateTime -Descending }
    foreach ($r in $Script:RiskyUsers | Select-Object -First 100) {
        $sev = if ($r.RiskLevel -match 'high') {'Critical'} elseif ($r.RiskLevel -match 'medium') {'High'} else {'Medium'}
        Add-Issue -Area 'Risky Users' -Status $r.RiskState -Severity $sev -Title $r.UserPrincipalName -Details "Risk: $($r.RiskLevel) | Last Updated: $($r.RiskLastUpdatedDateTime)" -Recommendation 'Investigate sign-ins, reset password if required, and require MFA re-registration if needed.'
    }
    if ($Script:RiskyUsers.Count -gt 0) { Add-Recommendation -Area 'Identity Protection' -Recommendation 'Investigate risky users and remediate high-risk identities first.' -Priority 'High' }
    return $Script:RiskyUsers.Count
}

function Get-FailedSignInsData {
    Write-DashboardLog 'Collecting failed sign-ins from last 24 hours...'
    $start = (Get-Date).AddDays(-1).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $filter = "createdDateTime ge $start and status/errorCode ne 0"
    $Script:FailedSignIns = Safe-Run -Name 'Failed sign-ins' -Default @() -ScriptBlock { Get-MgAuditLogSignIn -Filter $filter -Top 100 }
    foreach ($f in $Script:FailedSignIns | Select-Object -First 100) {
        Add-Issue -Area 'Failed Sign-ins' -Status 'Failed' -Severity 'Medium' -Title $f.UserPrincipalName -Details "App: $($f.AppDisplayName) | IP: $($f.IpAddress) | Error: $($f.Status.ErrorCode)" -Recommendation 'Check failure reason, user location, client app, and Conditional Access impact.'
    }
    return $Script:FailedSignIns.Count
}

function Get-DomainData {
    Write-DashboardLog 'Collecting verified domains...'
    $Script:Domains = Safe-Run -Name 'Domains' -Default @() -ScriptBlock { Get-MgDomain -All }
    foreach ($d in $Script:Domains) {
        $sev = if ($d.IsVerified) {'Info'} else {'High'}
        Add-Issue -Area 'Domains' -Status $(if($d.IsVerified){'Verified'}else{'Not Verified'}) -Severity $sev -Title $d.Id -Details "Default: $($d.IsDefault) | Auth Type: $($d.AuthenticationType)" -Recommendation 'Verify DNS records for unverified domains and review federation or managed authentication status.'
    }
    $notVerified = @($Script:Domains | Where-Object {-not $_.IsVerified}).Count
    if ($notVerified -gt 0) { Add-Recommendation -Area 'Domains' -Recommendation 'One or more domains are not verified. Review DNS records.' -Priority 'High' }
    return $notVerified
}

function Get-MailFlowData {
    Write-DashboardLog 'Collecting Exchange Online mail flow failures for last 24 hours...'
    $Script:MailFlowDetails.Clear()
    $Script:MailFlowSummary = @{ Failed24h=0; Checked=$false; Error='N/A' }

    if (-not $Script:ConnectedExchange) {
        $msg = 'Exchange Online is not connected. Mail flow section is skipped.'
        Write-DashboardLog $msg 'WARN'
        Add-Issue -Area 'Mail Flow' -Status 'N/A' -Severity 'Medium' -Title 'Mail flow failures unavailable' -Details $msg -Recommendation 'Connect to Exchange Online with an account that can run message trace.'
        $Script:MailFlowSummary.Error = $msg
        return 0
    }

    $start = (Get-Date).AddDays(-1)
    $end = Get-Date
    $traces = Safe-Run -Name 'Message trace last 24 hours' -Default @() -ScriptBlock {
        Get-MessageTrace -StartDate $start -EndDate $end -PageSize 5000
    }
    $failed = @($traces | Where-Object { $_.Status -and $_.Status -notmatch 'Delivered|Expanded|Resolved' })
    foreach ($m in $failed | Select-Object -First 200) {
        $Script:MailFlowDetails.Add([pscustomobject]@{
            Received=$m.Received; Sender=$m.SenderAddress; Recipient=$m.RecipientAddress; Subject=$m.Subject; Status=$m.Status; MessageId=$m.MessageId
        }) | Out-Null
    }
    $Script:MailFlowSummary = @{ Failed24h=$failed.Count; Checked=$true; Error='' }
    if ($failed.Count -gt 0) {
        Add-Issue -Area 'Mail Flow' -Status "$($failed.Count) failed/non-delivered" -Severity 'High' -Title 'Mail flow failures in last 24 hours' -Details "Non-delivered, failed, or unresolved traces found: $($failed.Count)" -Recommendation 'Review message trace details, recipient errors, transport rules, and quarantine if applicable.'
        Add-Recommendation -Area 'Exchange Online' -Recommendation 'Investigate mail flow failures from the last 24 hours and validate affected senders/recipients.' -Priority 'High'
    } else {
        Add-Issue -Area 'Mail Flow' -Status 'Healthy' -Severity 'Info' -Title 'Mail flow failures in last 24 hours' -Details 'No failed/non-delivered message trace entries found in retrieved data.' -Recommendation 'Continue periodic mail flow monitoring.'
    }
    return $failed.Count
}

function Get-DeviceComplianceData {
    Write-DashboardLog 'Collecting Intune managed device compliance summary...'
    $Script:DeviceDetails.Clear()
    $Script:ManagedDevices = Safe-Run -Name 'Managed devices' -Default @() -ScriptBlock { Get-MgDeviceManagementManagedDevice -All }

    $total = @($Script:ManagedDevices).Count
    $compliant = @($Script:ManagedDevices | Where-Object { $_.ComplianceState -eq 'compliant' }).Count
    $nonCompliant = @($Script:ManagedDevices | Where-Object { $_.ComplianceState -eq 'noncompliant' }).Count
    $unknown = $total - $compliant - $nonCompliant
    if ($unknown -lt 0) { $unknown = 0 }
    $pct = if ($total -gt 0) { [math]::Round(($compliant / $total) * 100, 2) } else { 0 }

    foreach ($d in $Script:ManagedDevices | Select-Object -First 500) {
        $Script:DeviceDetails.Add([pscustomobject]@{
            DeviceName=$d.DeviceName; UserPrincipalName=$d.UserPrincipalName; OperatingSystem=$d.OperatingSystem; ComplianceState=$d.ComplianceState; LastSyncDateTime=$d.LastSyncDateTime
        }) | Out-Null
    }

    $Script:DeviceSummary = @{ Total=$total; Compliant=$compliant; NonCompliant=$nonCompliant; Unknown=$unknown; CompliancePercent=$pct }
    if ($nonCompliant -gt 0) {
        Add-Issue -Area 'Intune Devices' -Status "$nonCompliant non-compliant" -Severity 'High' -Title 'Device compliance' -Details "Total: $total | Compliant: $compliant | Non-compliant: $nonCompliant | Unknown: $unknown | Compliance: $pct%" -Recommendation 'Review non-compliant devices in Intune and identify policy or enrollment issues.'
        Add-Recommendation -Area 'Intune' -Recommendation 'Review non-compliant managed devices and remediate compliance policy failures.' -Priority 'High'
    } elseif ($total -gt 0) {
        Add-Issue -Area 'Intune Devices' -Status 'Healthy' -Severity 'Info' -Title 'Device compliance' -Details "Total: $total | Compliant: $compliant | Compliance: $pct%" -Recommendation 'Continue periodic device compliance review.'
    } else {
        Add-Issue -Area 'Intune Devices' -Status 'No data' -Severity 'Medium' -Title 'No managed devices returned' -Details 'No Intune managed devices were returned or permission/license may be missing.' -Recommendation 'Confirm Intune licensing and DeviceManagementManagedDevices.Read.All permission.'
    }
    return $nonCompliant
}

function Get-TenantHealthScore {
    Write-DashboardLog 'Calculating tenant health score...'
    $securePct = if ($null -ne $Script:SecureScorePercent) { [double]$Script:SecureScorePercent } else { 50 }
    $mfaPct = if ($Script:MfaSummary.Checked -gt 0) { [double]$Script:MfaSummary.Percent } else { 50 }
    $devicePct = if ($Script:DeviceSummary.Total -gt 0) { [double]$Script:DeviceSummary.CompliancePercent } else { 50 }

    $activeIncidentCount = @($Script:HealthItems | Where-Object { $_.Status -notmatch 'serviceRestored|resolved|postIncidentReviewPublished' }).Count
    $incidentPenalty = [math]::Min(100, $activeIncidentCount * 10)
    $riskPenalty = [math]::Min(100, $Script:RiskyUsers.Count * 2)
    $mailPenalty = [math]::Min(100, $Script:MailFlowSummary.Failed24h * 2)

    $score = [math]::Round(
        ($securePct * 0.30) +
        ($mfaPct * 0.25) +
        ((100 - $incidentPenalty) * 0.15) +
        ((100 - $riskPenalty) * 0.10) +
        ($devicePct * 0.10) +
        ((100 - $mailPenalty) * 0.10), 2)

    if ($score -lt 0) { $score = 0 }
    if ($score -gt 100) { $score = 100 }

    $rating = if ($score -ge 95) {'Excellent'} elseif ($score -ge 80) {'Good'} elseif ($score -ge 60) {'Needs Attention'} else {'Critical'}
    $status = if ($score -ge 80) {'Good'} elseif ($score -ge 60) {'Warning'} else {'Critical'}
    $details = "SecureScore=$securePct%, MFA=$mfaPct%, DeviceCompliance=$devicePct%, ActiveIncidents=$activeIncidentCount, IncidentPenalty=$incidentPenalty, RiskPenalty=$riskPenalty, MailPenalty=$mailPenalty"
    $Script:TenantHealthScore = @{ Score=$score; Rating=$rating; Details=$details }
    Add-Issue -Area 'Tenant Health Score' -Status $rating -Severity 'Info' -Title 'Calculated tenant health score' -Details "$score% | $details" -Recommendation 'Use the score as an operational indicator and review the underlying findings for remediation.'
    Set-CardStatus 'TenantHealth' "$score% $rating" $status
    return $score
}
#endregion Data Collection Functions

#region Refresh And Export Functions
function Refresh-Dashboard {
    if (-not $Script:ConnectedGraph) {
        [System.Windows.MessageBox]::Show('Please select Connect first.','Not connected','OK','Warning') | Out-Null
        return
    }

    $Script:Issues.Clear(); $Script:Recommendations.Clear(); $Script:LicenseDetails.Clear(); $Script:MfaDetails.Clear(); $Script:DeviceDetails.Clear(); $Script:MailFlowDetails.Clear()
    Write-DashboardLog 'Refreshing dashboard data...'
    Set-TextSafe $Global:txtLastRefresh 'Refreshing...'

    $activeHealth = Get-ServiceHealthData
    $majorMessages = Get-MessageCenterData
    $licenseIssues = Get-LicenseData
    $usersSummary = Get-UsersSummaryData
    $withoutMfa = Get-UsersWithoutMfaData
    [void](Get-SecurityScoreData)
    $risky = Get-RiskyUsersData
    $failedSignIns = Get-FailedSignInsData
    $domainIssues = Get-DomainData
    $mailFailures = Get-MailFlowData
    [void](Get-DeviceComplianceData)
    [void](Get-TenantHealthScore)

    $mfaStatusColor = Get-PercentStatus -Percent ([double]$Script:MfaSummary.Percent) -GoodAt 90 -WarnAt 70
    $deviceStatusColor = Get-PercentStatus -Percent ([double]$Script:DeviceSummary.CompliancePercent) -GoodAt 90 -WarnAt 70
    $mailStatusColor = if ($Script:MailFlowSummary.Checked -eq $false) { 'Warning' } else { Get-StatusByCount -Critical $mailFailures -Warning 0 }

    Set-CardStatus 'TotalLicenses' "$($Script:LicenseSummary.Total) total" 'Info'
    Set-CardStatus 'AvailableLicenses' "$($Script:LicenseSummary.Available) available" (Get-StatusByCount -Critical 0 -Warning $licenseIssues)
    Set-CardStatus 'LicensedUsers' "$($usersSummary.Licensed) licensed" 'Good'
    Set-CardStatus 'UsersWithoutMfa' "$withoutMfa users" (Get-StatusByCount -Critical $withoutMfa -Warning 0)
    Set-CardStatus 'MfaStatus' "$($Script:MfaSummary.Percent)% registered" $mfaStatusColor
    Set-CardStatus 'MailFlow' "$mailFailures failures" $mailStatusColor
    Set-CardStatus 'DeviceCompliance' "$($Script:DeviceSummary.CompliancePercent)% compliant" $deviceStatusColor
    Set-CardStatus 'ServiceHealth' "$activeHealth active" (Get-StatusByCount -Critical $activeHealth -Warning 0)
    Set-CardStatus 'RiskyUsers' "$risky risky" (Get-StatusByCount -Critical $risky -Warning 0)
    Set-CardStatus 'SignIns' "$failedSignIns failed" (Get-StatusByCount -Critical 0 -Warning $failedSignIns)
    Set-CardStatus 'Domains' "$domainIssues issue" (Get-StatusByCount -Critical 0 -Warning $domainIssues)
    Set-CardStatus 'Users' "$($usersSummary.Total) users" 'Good'
    Set-CardStatus 'MessageCenter' "$majorMessages major" (Get-StatusByCount -Critical 0 -Warning $majorMessages)

    Set-TextSafe $Global:txtLastRefresh "Last refresh: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
    Write-DashboardLog 'Dashboard refresh completed.'
}

function Export-DashboardHtml {
    $file = Join-Path $Script:ReportFolder ("M365EnterpriseOperations_v4_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $cards = @"
<div class='card'><b>Tenant Health Score</b><h2>$($Script:TenantHealthScore.Score)% $($Script:TenantHealthScore.Rating)</h2></div>
<div class='card'><b>Secure Score</b><h2>$($Script:SecureScoreText)</h2></div>
<div class='card'><b>Total Licenses</b><h2>$($Script:LicenseSummary.Total)</h2></div>
<div class='card'><b>Available Licenses</b><h2>$($Script:LicenseSummary.Available)</h2></div>
<div class='card'><b>Users Without MFA</b><h2>$($Script:MfaSummary.WithoutMfa)</h2></div>
<div class='card'><b>MFA Coverage</b><h2>$($Script:MfaSummary.Percent)%</h2></div>
<div class='card'><b>Mail Failures 24h</b><h2>$($Script:MailFlowSummary.Failed24h)</h2></div>
<div class='card'><b>Device Compliance</b><h2>$($Script:DeviceSummary.CompliancePercent)%</h2></div>
"@
    $issueRows = foreach ($i in $Script:Issues) { "<tr><td>$(ConvertTo-HtmlSafe $i.Area)</td><td>$(ConvertTo-HtmlSafe $i.Status)</td><td>$(ConvertTo-HtmlSafe $i.Severity)</td><td>$(ConvertTo-HtmlSafe $i.Title)</td><td>$(ConvertTo-HtmlSafe $i.Details)</td><td>$(ConvertTo-HtmlSafe $i.Recommendation)</td></tr>" }
    $licenseRows = foreach ($l in $Script:LicenseDetails) { "<tr><td>$(ConvertTo-HtmlSafe $l.LicenseName)</td><td>$($l.Total)</td><td>$($l.Assigned)</td><td>$($l.Available)</td><td>$($l.UsagePercent)</td><td>$($l.Status)</td></tr>" }
    $mfaRows = foreach ($m in $Script:MfaDetails | Where-Object {$_.MfaStatus -eq 'Not Registered'}) { "<tr><td>$(ConvertTo-HtmlSafe $m.DisplayName)</td><td>$(ConvertTo-HtmlSafe $m.UserPrincipalName)</td><td>$(ConvertTo-HtmlSafe $m.MfaStatus)</td><td>$(ConvertTo-HtmlSafe $m.Methods)</td></tr>" }
    $deviceRows = foreach ($d in $Script:DeviceDetails) { "<tr><td>$(ConvertTo-HtmlSafe $d.DeviceName)</td><td>$(ConvertTo-HtmlSafe $d.UserPrincipalName)</td><td>$(ConvertTo-HtmlSafe $d.OperatingSystem)</td><td>$(ConvertTo-HtmlSafe $d.ComplianceState)</td><td>$(ConvertTo-HtmlSafe $d.LastSyncDateTime)</td></tr>" }
    $mailRows = foreach ($m in $Script:MailFlowDetails) { "<tr><td>$(ConvertTo-HtmlSafe $m.Received)</td><td>$(ConvertTo-HtmlSafe $m.Sender)</td><td>$(ConvertTo-HtmlSafe $m.Recipient)</td><td>$(ConvertTo-HtmlSafe $m.Subject)</td><td>$(ConvertTo-HtmlSafe $m.Status)</td></tr>" }
    $recRows = foreach ($r in $Script:Recommendations) { "<tr><td>$(ConvertTo-HtmlSafe $r.Area)</td><td>$(ConvertTo-HtmlSafe $r.Priority)</td><td>$(ConvertTo-HtmlSafe $r.Recommendation)</td></tr>" }

    $html = @"
<!DOCTYPE html><html><head><meta charset='utf-8'><title>Microsoft 365 Enterprise Operations Report</title>
<style>body{font-family:Segoe UI,Arial;background:#f8fafc;color:#0f172a;margin:24px}.header{background:#0f172a;color:white;padding:24px;border-radius:14px}.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:18px 0}.card{background:white;border-radius:12px;padding:16px;box-shadow:0 2px 8px #ddd}table{width:100%;border-collapse:collapse;background:white;margin-top:15px}th{background:#1e293b;color:white;text-align:left}td,th{padding:10px;border-bottom:1px solid #e2e8f0}.section{margin-top:30px}</style></head><body>
<div class='header'><h1>Microsoft 365 Enterprise Operations Report</h1><p>Generated: $(Get-Date) | Account: $($Script:TenantInfo.Account) | Tenant: $($Script:TenantInfo.TenantId)</p></div>
<h2>Executive Summary</h2><div class='cards'>$cards</div>
<div class='section'><h2>License Details</h2><table><tr><th>License</th><th>Total</th><th>Assigned</th><th>Available</th><th>Usage %</th><th>Status</th></tr>$($licenseRows -join "`n")</table></div>
<div class='section'><h2>Users Without MFA</h2><table><tr><th>Display Name</th><th>User Principal Name</th><th>MFA Status</th><th>Methods</th></tr>$($mfaRows -join "`n")</table></div>
<div class='section'><h2>Mail Flow Failures 24h</h2><table><tr><th>Received</th><th>Sender</th><th>Recipient</th><th>Subject</th><th>Status</th></tr>$($mailRows -join "`n")</table></div>
<div class='section'><h2>Device Compliance</h2><table><tr><th>Device</th><th>User</th><th>OS</th><th>Compliance</th><th>Last Sync</th></tr>$($deviceRows -join "`n")</table></div>
<div class='section'><h2>Recommendations</h2><table><tr><th>Area</th><th>Priority</th><th>Recommendation</th></tr>$($recRows -join "`n")</table></div>
<div class='section'><h2>All Findings</h2><table><tr><th>Area</th><th>Status</th><th>Severity</th><th>Title</th><th>Details</th><th>Recommendation</th></tr>$($issueRows -join "`n")</table></div>
<p>Generated by Microsoft 365 Enterprise Operations Center v$Script:DashboardVersion.</p></body></html>
"@
    $html | Out-File -FilePath $file -Encoding UTF8
    Write-DashboardLog "HTML report exported: $file"
    [System.Windows.MessageBox]::Show("Report exported:`n$file",'Export Complete','OK','Information') | Out-Null
}

function Export-DashboardCsv {
    $base = Join-Path $Script:ReportFolder ("M365EnterpriseOperations_v4_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $Script:Issues | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$base`_Findings.csv"
    $Script:LicenseDetails | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$base`_Licenses.csv"
    $Script:MfaDetails | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$base`_MFA.csv"
    $Script:MailFlowDetails | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$base`_MailFlow.csv"
    $Script:DeviceDetails | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$base`_Devices.csv"
    $Script:Recommendations | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$base`_Recommendations.csv"
    Write-DashboardLog "CSV reports exported with base name: $base"
    [System.Windows.MessageBox]::Show("CSV reports exported with base name:`n$base",'Export Complete','OK','Information') | Out-Null
}
#endregion Refresh And Export Functions

#region WPF UI
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Microsoft 365 Enterprise Operations Center v4" Height="930" Width="1500" WindowStartupLocation="CenterScreen" Background="#F1F5F9">
<Grid Margin="14">
<Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="210"/></Grid.RowDefinitions>
<Border Grid.Row="0" Background="#0F172A" CornerRadius="16" Padding="18" Margin="0,0,0,12"><DockPanel>
<StackPanel DockPanel.Dock="Left"><TextBlock Text="Microsoft 365 Enterprise Operations Center" Foreground="White" FontSize="28" FontWeight="Bold"/><TextBlock Text="Version 4.0 - Fixed syntax, Exchange connection handling, Secure Score parser, License, MFA, Mail Flow, Intune" Foreground="#93C5FD" FontSize="12" Margin="0,2,0,0"/><TextBlock Name="txtTenant" Text="Not connected" Foreground="#CBD5E1" FontSize="13" Margin="0,6,0,0"/><TextBlock Name="txtLastRefresh" Text="Last refresh: Never" Foreground="#CBD5E1" FontSize="13" Margin="0,2,0,0"/></StackPanel>
<StackPanel Orientation="Horizontal" DockPanel.Dock="Right" HorizontalAlignment="Right"><Button Name="btnConnect" Content="Connect" Width="110" Height="38" Margin="6" Background="#2563EB" Foreground="White" FontWeight="SemiBold"/><Button Name="btnRefresh" Content="Refresh" Width="110" Height="38" Margin="6" Background="#16A34A" Foreground="White" FontWeight="SemiBold"/><Button Name="btnExportHtml" Content="Export HTML" Width="120" Height="38" Margin="6" Background="#7C3AED" Foreground="White" FontWeight="SemiBold"/><Button Name="btnExportCsv" Content="Export CSV" Width="110" Height="38" Margin="6" Background="#475569" Foreground="White" FontWeight="SemiBold"/></StackPanel>
</DockPanel></Border>

<UniformGrid Grid.Row="1" Columns="6" Rows="3" Margin="0,0,0,12">
<Border Name="cardTenantHealth" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Tenant Health Score" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblTenantHealth" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardSecureScore" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Secure Score" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblSecureScore" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardTotalLicenses" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Total Licenses" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblTotalLicenses" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardAvailableLicenses" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Available Licenses" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblAvailableLicenses" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardLicensedUsers" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Licensed Users" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblLicensedUsers" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardUsers" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Total Users" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblUsers" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardUsersWithoutMfa" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Users Without MFA" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblUsersWithoutMfa" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardMfaStatus" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="MFA Status" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblMfaStatus" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardMailFlow" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Mail Flow Failures 24h" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblMailFlow" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardDeviceCompliance" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Device Compliance" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblDeviceCompliance" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardServiceHealth" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Service Health" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblServiceHealth" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardRiskyUsers" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Risky Users" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblRiskyUsers" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardSignIns" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Failed Sign-ins" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblSignIns" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardDomains" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Domains" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblDomains" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardMessageCenter" Background="#64748B" CornerRadius="14" Padding="14" Margin="6"><StackPanel><TextBlock Text="Message Center" Foreground="White" FontWeight="Bold"/><TextBlock Name="lblMessageCenter" Text="N/A" Foreground="White" FontSize="21" FontWeight="Bold"/></StackPanel></Border>
</UniformGrid>

<TabControl Grid.Row="2" Background="White">
<TabItem Header="All Findings"><DataGrid Name="gridIssues" AutoGenerateColumns="False" IsReadOnly="True" AlternatingRowBackground="#F8FAFC" GridLinesVisibility="Horizontal"><DataGrid.Columns><DataGridTextColumn Header="Area" Binding="{Binding Area}" Width="145"/><DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="150"/><DataGridTextColumn Header="Severity" Binding="{Binding Severity}" Width="100"/><DataGridTextColumn Header="Title" Binding="{Binding Title}" Width="300"/><DataGridTextColumn Header="Details" Binding="{Binding Details}" Width="420"/><DataGridTextColumn Header="Recommendation" Binding="{Binding Recommendation}" Width="*"/></DataGrid.Columns></DataGrid></TabItem>
<TabItem Header="License Details"><DataGrid Name="gridLicenseDetails" AutoGenerateColumns="False" IsReadOnly="True" AlternatingRowBackground="#F8FAFC"><DataGrid.Columns><DataGridTextColumn Header="License Name" Binding="{Binding LicenseName}" Width="300"/><DataGridTextColumn Header="Total" Binding="{Binding Total}" Width="100"/><DataGridTextColumn Header="Assigned" Binding="{Binding Assigned}" Width="100"/><DataGridTextColumn Header="Available" Binding="{Binding Available}" Width="100"/><DataGridTextColumn Header="Usage %" Binding="{Binding UsagePercent}" Width="100"/><DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="120"/></DataGrid.Columns></DataGrid></TabItem>
<TabItem Header="MFA Status"><DataGrid Name="gridMfaDetails" AutoGenerateColumns="False" IsReadOnly="True" AlternatingRowBackground="#F8FAFC"><DataGrid.Columns><DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="220"/><DataGridTextColumn Header="User Principal Name" Binding="{Binding UserPrincipalName}" Width="320"/><DataGridTextColumn Header="Account Enabled" Binding="{Binding AccountEnabled}" Width="130"/><DataGridTextColumn Header="MFA Status" Binding="{Binding MfaStatus}" Width="140"/><DataGridTextColumn Header="Methods" Binding="{Binding Methods}" Width="*"/></DataGrid.Columns></DataGrid></TabItem>
<TabItem Header="Mail Flow 24h"><DataGrid Name="gridMailFlow" AutoGenerateColumns="False" IsReadOnly="True" AlternatingRowBackground="#F8FAFC"><DataGrid.Columns><DataGridTextColumn Header="Received" Binding="{Binding Received}" Width="170"/><DataGridTextColumn Header="Sender" Binding="{Binding Sender}" Width="240"/><DataGridTextColumn Header="Recipient" Binding="{Binding Recipient}" Width="240"/><DataGridTextColumn Header="Subject" Binding="{Binding Subject}" Width="360"/><DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="140"/><DataGridTextColumn Header="Message ID" Binding="{Binding MessageId}" Width="*"/></DataGrid.Columns></DataGrid></TabItem>
<TabItem Header="Device Compliance"><DataGrid Name="gridDeviceDetails" AutoGenerateColumns="False" IsReadOnly="True" AlternatingRowBackground="#F8FAFC"><DataGrid.Columns><DataGridTextColumn Header="Device Name" Binding="{Binding DeviceName}" Width="220"/><DataGridTextColumn Header="User" Binding="{Binding UserPrincipalName}" Width="280"/><DataGridTextColumn Header="OS" Binding="{Binding OperatingSystem}" Width="130"/><DataGridTextColumn Header="Compliance" Binding="{Binding ComplianceState}" Width="130"/><DataGridTextColumn Header="Last Sync" Binding="{Binding LastSyncDateTime}" Width="*"/></DataGrid.Columns></DataGrid></TabItem>
<TabItem Header="Recommendations"><DataGrid Name="gridRecommendations" AutoGenerateColumns="False" IsReadOnly="True" AlternatingRowBackground="#F8FAFC"><DataGrid.Columns><DataGridTextColumn Header="Area" Binding="{Binding Area}" Width="180"/><DataGridTextColumn Header="Priority" Binding="{Binding Priority}" Width="120"/><DataGridTextColumn Header="Recommendation" Binding="{Binding Recommendation}" Width="*"/></DataGrid.Columns></DataGrid></TabItem>
</TabControl>

<Border Grid.Row="3" Background="#020617" CornerRadius="12" Padding="12" Margin="0,12,0,0"><DockPanel><TextBlock DockPanel.Dock="Top" Text="Activity Log" Foreground="White" FontSize="14" FontWeight="Bold" Margin="0,0,0,6"/><TextBox Name="txtLog" Background="#020617" Foreground="#C4B5FD" BorderBrush="#334155" FontFamily="Consolas" FontSize="12" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" IsReadOnly="True"/></DockPanel></Border>
</Grid></Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$Global:Window = [Windows.Markup.XamlReader]::Load($reader)
$xaml.SelectNodes('//*[@Name]') | ForEach-Object {
    $name = $_.Name
    Set-Variable -Name $name -Value $Global:Window.FindName($name) -Scope Global
}

$Global:gridIssues.ItemsSource = $Script:Issues
$Global:gridRecommendations.ItemsSource = $Script:Recommendations
$Global:gridLicenseDetails.ItemsSource = $Script:LicenseDetails
$Global:gridMfaDetails.ItemsSource = $Script:MfaDetails
$Global:gridMailFlow.ItemsSource = $Script:MailFlowDetails
$Global:gridDeviceDetails.ItemsSource = $Script:DeviceDetails

$Global:btnConnect.Add_Click({ Connect-M365Dashboard })
$Global:btnRefresh.Add_Click({ Refresh-Dashboard })
$Global:btnExportHtml.Add_Click({ Export-DashboardHtml })
$Global:btnExportCsv.Add_Click({ Export-DashboardCsv })

Write-DashboardLog "Dashboard v$Script:DashboardVersion loaded. Select Connect, then Refresh."
if (-not (Test-IsAdmin)) { Write-DashboardLog 'PowerShell is not running as Administrator. Module installation may still work with CurrentUser scope.' 'WARN' }
[void]$Global:Window.ShowDialog()
#endregion WPF UI
