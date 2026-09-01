<#
.SYNOPSIS
 Microsoft 365 Enterprise Operations Center v5.2
.DESCRIPTION
 Read-only WPF dashboard. Every launch creates a fresh process-scoped Microsoft Graph session,
 validates an active Global Administrator assignment, connects Exchange Online separately,
 and then opens the dashboard.
#>

$ErrorActionPreference = 'Continue'
$Script:Version = '5.2'
$Script:GraphConnected = $false
$Script:ExchangeConnected = $false
$Script:AdminUPN = $null
$Script:TenantId = $null
$Script:ReportFolder = Join-Path $env:USERPROFILE 'Documents\M365EnterpriseOperationsReports'
New-Item -ItemType Directory -Path $Script:ReportFolder -Force | Out-Null

$Script:Issues = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Recommendations = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Licenses = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Mfa = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Devices = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Mail = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:RawUsers=@(); $Script:RawHealth=@(); $Script:RawRisk=@()
$Script:Summary=[ordered]@{TenantHealth='N/A';SecureScore='N/A';TotalLicenses=0;AvailableLicenses=0;LicensedUsers=0;TotalUsers=0;WithoutMfa=0;MfaPercent=0;MailFailures='N/A';DevicePercent=0;ActiveHealth=0;RiskyUsers=0;FailedSignIns=0;DomainIssues=0;MajorChanges=0}

function Write-Log {
 param([string]$Message,[ValidateSet('INFO','WARN','ERROR')][string]$Level='INFO')
 $line='[{0}][{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
 Write-Host $line
 if($Global:txtLog){try{$Global:txtLog.AppendText($line+[Environment]::NewLine);$Global:txtLog.ScrollToEnd()}catch{}}
}
function Ensure-Module {
 param([Parameter(Mandatory)][string]$Name)
 try{
  if(-not(Get-Module -ListAvailable $Name)){Write-Log "Installing $Name" WARN;Install-Module $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop}
  Import-Module $Name -Force -ErrorAction Stop
  Write-Log "Loaded $Name"
  return $true
 }catch{Write-Log "Unable to load $Name. $($_.Exception.Message)" ERROR;return $false}
}
function Invoke-Safe {
 param([scriptblock]$Action,[string]$Name,[object]$Default=@())
 try{return & $Action}catch{Write-Log "$Name failed: $($_.Exception.Message)" ERROR;return $Default}
}
function Add-Issue {
 param($Area,$Status,$Severity,$Title,$Details,$Recommendation)
 $Script:Issues.Add([pscustomobject]@{Area=$Area;Status=$Status;Severity=$Severity;Title=$Title;Details=$Details;Recommendation=$Recommendation;Time=Get-Date})|Out-Null
}
function Add-Rec {param($Area,$Priority,$Recommendation)$Script:Recommendations.Add([pscustomobject]@{Area=$Area;Priority=$Priority;Recommendation=$Recommendation})|Out-Null}
function HtmlEncode([object]$Value){Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue;[System.Web.HttpUtility]::HtmlEncode([string]$Value)}

function Disconnect-All {
 try{Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue}catch{}
 try{Disconnect-MgGraph -ErrorAction SilentlyContinue|Out-Null}catch{}
 $Script:GraphConnected=$false;$Script:ExchangeConnected=$false
}
function Test-GlobalAdmin {
 try{
  $me=Invoke-MgGraphRequest GET 'https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName'
  $uri="https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$filter=principalId eq '$($me.id)'&`$expand=roleDefinition"
  $result=Invoke-MgGraphRequest GET $uri
  return [bool](@($result.value|Where-Object{$_.roleDefinition.templateId -eq '62e90394-69f5-4237-9190-012177145e10'}).Count)
 }catch{Write-Log "Global Administrator validation failed: $($_.Exception.Message)" ERROR;return $false}
}
function Connect-FreshSession {
 Disconnect-All
 $mods=@('Microsoft.Graph.Authentication','Microsoft.Graph.Users','Microsoft.Graph.Identity.DirectoryManagement','Microsoft.Graph.Identity.SignIns','Microsoft.Graph.Reports','Microsoft.Graph.Security','Microsoft.Graph.Devices.ServiceAnnouncement','Microsoft.Graph.DeviceManagement')
 foreach($m in $mods){if(-not(Ensure-Module $m)){throw "Required module unavailable: $m"}}
 $scopes=@('ServiceHealth.Read.All','Directory.Read.All','RoleManagement.Read.Directory','User.Read.All','AuditLog.Read.All','SecurityEvents.Read.All','IdentityRiskyUser.Read.All','Reports.Read.All','UserAuthenticationMethod.Read.All','DeviceManagementManagedDevices.Read.All','Organization.Read.All','LicenseAssignment.Read.All')
 Write-Host ''
 Write-Host 'NEW MICROSOFT GRAPH SESSION' -ForegroundColor Cyan
 Write-Host 'Complete the device-code sign-in shown below using the required Global Administrator.' -ForegroundColor Yellow
 Connect-MgGraph -Scopes $scopes -ContextScope Process -UseDeviceAuthentication -NoWelcome -ErrorAction Stop | Out-Host
 $ctx=Get-MgContext
 if(-not$ctx.Account){throw 'Microsoft Graph returned no authenticated account.'}
 $Script:AdminUPN=$ctx.Account;$Script:TenantId=$ctx.TenantId
 if(-not(Test-GlobalAdmin)){Disconnect-MgGraph|Out-Null;throw "The selected account $($ctx.Account) does not have an active Global Administrator assignment. Activate PIM first if applicable."}
 $Script:GraphConnected=$true
 Write-Log "Graph connected as active Global Administrator: $Script:AdminUPN"
 if(Ensure-Module ExchangeOnlineManagement){
  try{
   $p=@{UserPrincipalName=$Script:AdminUPN;ShowBanner=$false;ErrorAction='Stop'}
   if((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')){$p.DisableWAM=$true}
   Write-Host '';Write-Host 'EXCHANGE ONLINE SESSION' -ForegroundColor Cyan
   Connect-ExchangeOnline @p | Out-Host
   $Script:ExchangeConnected=$true
   Write-Log "Exchange Online connected as $Script:AdminUPN"
  }catch{Write-Log "Exchange Online unavailable. Mail Flow will show N/A. $($_.Exception.Message)" WARN}
 }
}

function Collect-Data {
 $Script:Issues.Clear();$Script:Recommendations.Clear();$Script:Licenses.Clear();$Script:Mfa.Clear();$Script:Devices.Clear();$Script:Mail.Clear()
 Write-Log 'Collecting users...'
 $Script:RawUsers=Invoke-Safe {Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled,UserType,AssignedLicenses} 'Users'
 $Script:Summary.TotalUsers=@($Script:RawUsers).Count
 $Script:Summary.LicensedUsers=@($Script:RawUsers|Where-Object{$_.AssignedLicenses.Count -gt 0}).Count
 Add-Issue Users Summary Info 'User summary' "Total: $($Script:Summary.TotalUsers); Licensed: $($Script:Summary.LicensedUsers)" 'Review disabled licensed accounts and guests.'

 Write-Log 'Collecting licenses...'
 $skus=Invoke-Safe {Get-MgSubscribedSku -All} 'Licenses'
 $total=0;$avail=0
 foreach($s in $skus){$t=[int]$s.PrepaidUnits.Enabled;$u=[int]$s.ConsumedUnits;$a=[math]::Max(0,$t-$u);$pct=if($t){[math]::Round($u/$t*100,2)}else{0};$status=if($a-le0){'Critical'}elseif($pct-ge90){'Warning'}else{'Good'};$total+=$t;$avail+=$a;$Script:Licenses.Add([pscustomobject]@{License=$s.SkuPartNumber;Total=$t;Assigned=$u;Available=$a;Usage="$pct%";Status=$status})|Out-Null;if($status-ne'Good'){Add-Issue Licenses "$u/$t" $(if($status-eq'Critical'){'High'}else{'Medium'}) $s.SkuPartNumber "Available: $a; Usage: $pct%" 'Review assignments and procurement.'}}
 $Script:Summary.TotalLicenses=$total;$Script:Summary.AvailableLicenses=$avail

 Write-Log 'Collecting MFA registration report...'
 $registration=Invoke-Safe {Get-MgReportAuthenticationMethodUserRegistrationDetail -All} 'MFA registration report'
 if(@($registration).Count){
  foreach($r in $registration){$registered=[bool]$r.IsMfaRegistered;$Script:Mfa.Add([pscustomobject]@{UserPrincipalName=$r.UserPrincipalName;MfaRegistered=$registered;Methods=($r.MethodsRegistered -join ', ');Admin=$r.IsAdmin})|Out-Null}
  $members=@($registration);$with=@($members|Where-Object IsMfaRegistered).Count;$without=$members.Count-$with;$pct=if($members.Count){[math]::Round($with/$members.Count*100,2)}else{0}
 }else{$with=0;$without=0;$pct=0}
 $Script:Summary.WithoutMfa=$without;$Script:Summary.MfaPercent=$pct
 if($without){Add-Issue MFA "$without not registered" High 'MFA registration gap' "Coverage: $pct%" 'Require registration through Conditional Access and registration campaign.';Add-Rec MFA High 'Prioritize administrators and enabled member users without MFA registration.'}

 Write-Log 'Collecting Secure Score...'
 $score=Invoke-Safe {Invoke-MgGraphRequest GET 'https://graph.microsoft.com/v1.0/security/secureScores?$top=1'} 'Secure Score' $null
 $o=@($score.value)[0]
 if($o -and [double]$o.maxScore -gt 0){$sp=[math]::Round([double]$o.currentScore/[double]$o.maxScore*100,2);$Script:Summary.SecureScore="$($o.currentScore) / $($o.maxScore) ($sp%)"}else{$sp=50;$Script:Summary.SecureScore='N/A'}

 Write-Log 'Collecting service health and Message Center...'
 $Script:RawHealth=Invoke-Safe {Get-MgServiceAnnouncementIssue -All} 'Service health'
 $active=@($Script:RawHealth|Where-Object{$_.Status -notmatch 'serviceRestored|resolved|postIncidentReviewPublished'})
 $Script:Summary.ActiveHealth=$active.Count
 foreach($i in $active|Select-Object -First 100){Add-Issue 'Service Health' $i.Status High $i.Title "$($i.Service); $($i.LastModifiedDateTime)" 'Review the Microsoft 365 service health advisory.'}
 $messages=Invoke-Safe {Get-MgServiceAnnouncementMessage -All} 'Message Center'
 $Script:Summary.MajorChanges=@($messages|Where-Object IsMajorChange).Count

 Write-Log 'Collecting risky users and failed sign-ins...'
 $Script:RawRisk=Invoke-Safe {Get-MgRiskyUser -All} 'Risky users'
 $Script:Summary.RiskyUsers=@($Script:RawRisk).Count
 foreach($r in $Script:RawRisk|Select-Object -First 100){Add-Issue 'Risky Users' $r.RiskState $(if($r.RiskLevel-eq'high'){'Critical'}else{'High'}) $r.UserPrincipalName "Risk: $($r.RiskLevel)" 'Investigate and remediate the identity.'}
 $start=(Get-Date).AddDays(-1).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
 $failed=Invoke-Safe {Get-MgAuditLogSignIn -Filter "createdDateTime ge $start and status/errorCode ne 0" -Top 100} 'Failed sign-ins'
 $Script:Summary.FailedSignIns=@($failed).Count
 foreach($f in $failed){Add-Issue 'Failed Sign-ins' Failed Medium $f.UserPrincipalName "App: $($f.AppDisplayName); IP: $($f.IpAddress); Code: $($f.Status.ErrorCode)" 'Review the failure reason and Conditional Access result.'}

 Write-Log 'Collecting domains...'
 $domains=Invoke-Safe {Get-MgDomain -All} 'Domains';$Script:Summary.DomainIssues=@($domains|Where-Object{-not$_.IsVerified}).Count
 foreach($d in $domains|Where-Object{-not$_.IsVerified}){Add-Issue Domains 'Not verified' High $d.Id "Authentication: $($d.AuthenticationType)" 'Correct DNS verification records.'}

 Write-Log 'Collecting Intune devices...'
 $managed=Invoke-Safe {Get-MgDeviceManagementManagedDevice -All} 'Managed devices'
 foreach($d in $managed){$Script:Devices.Add([pscustomobject]@{Device=$d.DeviceName;User=$d.UserPrincipalName;OS=$d.OperatingSystem;Compliance=$d.ComplianceState;LastSync=$d.LastSyncDateTime})|Out-Null}
 $dt=@($managed).Count;$dc=@($managed|Where-Object ComplianceState -eq compliant).Count;$dn=@($managed|Where-Object ComplianceState -eq noncompliant).Count;$dp=if($dt){[math]::Round($dc/$dt*100,2)}else{0};$Script:Summary.DevicePercent=$dp
 if($dn){Add-Issue 'Intune Devices' "$dn non-compliant" High 'Device compliance' "Total: $dt; Compliant: $dc; Compliance: $dp%" 'Remediate non-compliant managed devices.'}

 Write-Log 'Collecting Exchange Online message trace...'
 if($Script:ExchangeConnected){
  $traces=Invoke-Safe {if(Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue){Get-MessageTraceV2 -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -ResultSize 5000}else{Get-MessageTrace -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -PageSize 5000}} 'Message trace'
  $mf=@($traces|Where-Object{$_.Status -and $_.Status -notmatch 'Delivered|Expanded|Resolved'})
  foreach($m in $mf|Select-Object -First 500){$Script:Mail.Add([pscustomobject]@{Received=$m.Received;Sender=$m.SenderAddress;Recipient=$m.RecipientAddress;Subject=$m.Subject;Status=$m.Status;TraceId=$(if($m.MessageTraceId){$m.MessageTraceId}else{$m.MessageId})})|Out-Null}
  $Script:Summary.MailFailures=$mf.Count
  if($mf.Count){Add-Issue 'Mail Flow' "$($mf.Count) failures" High 'Mail flow failures during last 24 hours' 'Failed, pending, filtered, or unresolved messages were returned.' 'Review message trace details and transport controls.'}
 }else{$Script:Summary.MailFailures='N/A'}

 $incidentPenalty=[math]::Min(100,$active.Count*10);$riskPenalty=[math]::Min(100,$Script:Summary.RiskyUsers*2);$mailNumber=if($Script:Summary.MailFailures -is [int]){$Script:Summary.MailFailures}else{0};$mailPenalty=[math]::Min(100,$mailNumber*2)
 $health=[math]::Round($sp*.30+$pct*.25+(100-$incidentPenalty)*.15+(100-$riskPenalty)*.10+$dp*.10+(100-$mailPenalty)*.10,2)
 $health=[math]::Max(0,[math]::Min(100,$health));$rating=if($health-ge95){'Excellent'}elseif($health-ge80){'Good'}elseif($health-ge60){'Needs Attention'}else{'Critical'}
 $Script:Summary.TenantHealth="$health% $rating"
 Add-Issue 'Tenant Health' $rating Info 'Calculated operational health score' $Script:Summary.TenantHealth 'Use the score as an operational indicator and review underlying findings.'
 Write-Log 'Data collection completed.'
}

function CardColor([string]$Name){
 switch($Name){'TenantHealth'{if($Script:Summary.TenantHealth-match'Excellent|Good'){'#16A34A'}elseif($Script:Summary.TenantHealth-match'Needs'){'#F59E0B'}else{'#EF4444'}}'WithoutMfa'{if($Script:Summary.WithoutMfa){'#EF4444'}else{'#16A34A'}}'MailFailures'{if($Script:Summary.MailFailures-eq'N/A'){'#64748B'}elseif([int]$Script:Summary.MailFailures){'#EF4444'}else{'#16A34A'}}default{'#2563EB'}}
}
function Update-Cards {
 $map=@{TenantHealth='TenantHealth';SecureScore='SecureScore';TotalLicenses='TotalLicenses';AvailableLicenses='AvailableLicenses';LicensedUsers='LicensedUsers';TotalUsers='TotalUsers';WithoutMfa='WithoutMfa';MfaPercent='MfaPercent';MailFailures='MailFailures';DevicePercent='DevicePercent';ActiveHealth='ActiveHealth';RiskyUsers='RiskyUsers';FailedSignIns='FailedSignIns';DomainIssues='DomainIssues';MajorChanges='MajorChanges'}
 foreach($key in $map.Keys){$label=Get-Variable "lbl$key" -Scope Global -ValueOnly -ErrorAction SilentlyContinue;$card=Get-Variable "card$key" -Scope Global -ValueOnly -ErrorAction SilentlyContinue;$v=$Script:Summary[$map[$key]];if($key-in @('MfaPercent','DevicePercent')){$v="$v%"};if($label){$label.Text=[string]$v};if($card){$card.Background=CardColor $key}}
 $Global:txtLastRefresh.Text="Last refresh: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}
function Export-CsvReports {
 $b=Join-Path $Script:ReportFolder ("M365Operations_v5.2_{0}" -f (Get-Date -Format yyyyMMdd_HHmmss))
 $Script:Issues|Export-Csv "$b`_Findings.csv" -NoTypeInformation -Encoding utf8
 $Script:Licenses|Export-Csv "$b`_Licenses.csv" -NoTypeInformation -Encoding utf8
 $Script:Mfa|Export-Csv "$b`_MFA.csv" -NoTypeInformation -Encoding utf8
 $Script:Devices|Export-Csv "$b`_Devices.csv" -NoTypeInformation -Encoding utf8
 $Script:Mail|Export-Csv "$b`_MailFlow.csv" -NoTypeInformation -Encoding utf8
 $Script:Recommendations|Export-Csv "$b`_Recommendations.csv" -NoTypeInformation -Encoding utf8
 [System.Windows.MessageBox]::Show("CSV reports exported:`n$b",'Export complete')|Out-Null
}
function Export-HtmlReport {
 $f=Join-Path $Script:ReportFolder ("M365Operations_v5.2_{0}.html" -f (Get-Date -Format yyyyMMdd_HHmmss))
 $rows=foreach($i in $Script:Issues){"<tr><td>$(HtmlEncode $i.Area)</td><td>$(HtmlEncode $i.Status)</td><td>$(HtmlEncode $i.Severity)</td><td>$(HtmlEncode $i.Title)</td><td>$(HtmlEncode $i.Details)</td><td>$(HtmlEncode $i.Recommendation)</td></tr>"}
 $html=@"
<!doctype html><html><head><meta charset='utf-8'><title>M365 Operations Report</title><style>body{font-family:Segoe UI;background:#f8fafc;margin:24px;color:#0f172a}.head{background:#0f172a;color:white;padding:20px;border-radius:12px}.cards{display:flex;gap:12px;flex-wrap:wrap;margin:16px 0}.card{background:white;padding:14px;border-radius:10px;min-width:190px}table{width:100%;border-collapse:collapse;background:white}th{background:#1e293b;color:white}td,th{padding:9px;border-bottom:1px solid #e2e8f0;text-align:left}</style></head><body><div class='head'><h1>Microsoft 365 Enterprise Operations Report</h1><p>Generated: $(Get-Date) | Admin: $(HtmlEncode $Script:AdminUPN) | Tenant: $(HtmlEncode $Script:TenantId)</p></div><div class='cards'><div class='card'><b>Tenant Health</b><h2>$($Script:Summary.TenantHealth)</h2></div><div class='card'><b>Secure Score</b><h2>$($Script:Summary.SecureScore)</h2></div><div class='card'><b>MFA Coverage</b><h2>$($Script:Summary.MfaPercent)%</h2></div><div class='card'><b>Device Compliance</b><h2>$($Script:Summary.DevicePercent)%</h2></div></div><h2>Findings</h2><table><tr><th>Area</th><th>Status</th><th>Severity</th><th>Title</th><th>Details</th><th>Recommendation</th></tr>$($rows -join "`n")</table></body></html>
"@
 $html|Set-Content $f -Encoding utf8
 [System.Windows.MessageBox]::Show("HTML report exported:`n$f",'Export complete')|Out-Null
}

try{Connect-FreshSession}catch{Write-Host "`nConnection failed: $($_.Exception.Message)" -ForegroundColor Red;Disconnect-All;Read-Host 'Press Enter to close';exit 1}

Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Xaml
[xml]$xaml=@"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Microsoft 365 Enterprise Operations Center v5.2" Height="920" Width="1500" WindowStartupLocation="CenterScreen" Background="#F1F5F9">
<Grid Margin="14"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="170"/></Grid.RowDefinitions>
<Border Grid.Row="0" Background="#0F172A" CornerRadius="14" Padding="16" Margin="0,0,0,10"><DockPanel><StackPanel><TextBlock Text="Microsoft 365 Enterprise Operations Center" Foreground="White" FontSize="27" FontWeight="Bold"/><TextBlock Name="txtTenant" Foreground="#93C5FD"/><TextBlock Name="txtLastRefresh" Text="Last refresh: Never" Foreground="#CBD5E1"/></StackPanel><StackPanel DockPanel.Dock="Right" Orientation="Horizontal"><Button Name="btnNewSession" Content="New Session" Width="115" Height="36" Margin="5"/><Button Name="btnRefresh" Content="Refresh" Width="100" Height="36" Margin="5"/><Button Name="btnHtml" Content="Export HTML" Width="110" Height="36" Margin="5"/><Button Name="btnCsv" Content="Export CSV" Width="105" Height="36" Margin="5"/></StackPanel></DockPanel></Border>
<UniformGrid Grid.Row="1" Columns="5" Rows="3" Margin="0,0,0,10">
<Border Name="cardTenantHealth" Background="#64748B" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Tenant Health" Foreground="White"/><TextBlock Name="lblTenantHealth" Text="N/A" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardSecureScore" Background="#2563EB" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Secure Score" Foreground="White"/><TextBlock Name="lblSecureScore" Text="N/A" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardTotalLicenses" Background="#2563EB" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Total Licenses" Foreground="White"/><TextBlock Name="lblTotalLicenses" Text="0" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardAvailableLicenses" Background="#2563EB" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Available Licenses" Foreground="White"/><TextBlock Name="lblAvailableLicenses" Text="0" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardLicensedUsers" Background="#2563EB" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Licensed Users" Foreground="White"/><TextBlock Name="lblLicensedUsers" Text="0" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardTotalUsers" Background="#2563EB" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Total Users" Foreground="White"/><TextBlock Name="lblTotalUsers" Text="0" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardWithoutMfa" Background="#64748B" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Without MFA" Foreground="White"/><TextBlock Name="lblWithoutMfa" Text="0" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardMfaPercent" Background="#2563EB" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="MFA Coverage" Foreground="White"/><TextBlock Name="lblMfaPercent" Text="0%" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardMailFailures" Background="#64748B" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Mail Failures 24h" Foreground="White"/><TextBlock Name="lblMailFailures" Text="N/A" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardDevicePercent" Background="#2563EB" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Device Compliance" Foreground="White"/><TextBlock Name="lblDevicePercent" Text="0%" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardActiveHealth" Background="#2563EB" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Active Health Issues" Foreground="White"/><TextBlock Name="lblActiveHealth" Text="0" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardRiskyUsers" Background="#2563EB" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Risky Users" Foreground="White"/><TextBlock Name="lblRiskyUsers" Text="0" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardFailedSignIns" Background="#2563EB" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Failed Sign-ins" Foreground="White"/><TextBlock Name="lblFailedSignIns" Text="0" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardDomainIssues" Background="#2563EB" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Domain Issues" Foreground="White"/><TextBlock Name="lblDomainIssues" Text="0" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
<Border Name="cardMajorChanges" Background="#2563EB" CornerRadius="10" Margin="4" Padding="10"><StackPanel><TextBlock Text="Major Changes" Foreground="White"/><TextBlock Name="lblMajorChanges" Text="0" Foreground="White" FontSize="19" FontWeight="Bold"/></StackPanel></Border>
</UniformGrid>
<TabControl Grid.Row="2"><TabItem Header="Findings"><DataGrid Name="gridIssues" IsReadOnly="True" AutoGenerateColumns="True"/></TabItem><TabItem Header="Licenses"><DataGrid Name="gridLicenses" IsReadOnly="True" AutoGenerateColumns="True"/></TabItem><TabItem Header="MFA"><DataGrid Name="gridMfa" IsReadOnly="True" AutoGenerateColumns="True"/></TabItem><TabItem Header="Mail Flow"><DataGrid Name="gridMail" IsReadOnly="True" AutoGenerateColumns="True"/></TabItem><TabItem Header="Devices"><DataGrid Name="gridDevices" IsReadOnly="True" AutoGenerateColumns="True"/></TabItem><TabItem Header="Recommendations"><DataGrid Name="gridRecommendations" IsReadOnly="True" AutoGenerateColumns="True"/></TabItem></TabControl>
<Border Grid.Row="3" Background="#020617" CornerRadius="10" Padding="10" Margin="0,10,0,0"><DockPanel><TextBlock DockPanel.Dock="Top" Text="Activity Log" Foreground="White" FontWeight="Bold"/><TextBox Name="txtLog" Background="#020617" Foreground="#C4B5FD" FontFamily="Consolas" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/></DockPanel></Border>
</Grid></Window>
"@
$reader=New-Object System.Xml.XmlNodeReader $xaml
$Global:Window=[Windows.Markup.XamlReader]::Load($reader)
$xaml.SelectNodes('//*[@Name]')|ForEach-Object{$n=$_.Name;Set-Variable -Name $n -Scope Global -Value $Global:Window.FindName($n)}
$Global:txtTenant.Text="Global Administrator: $Script:AdminUPN | Tenant: $Script:TenantId | Graph context: Process"
$Global:gridIssues.ItemsSource=$Script:Issues;$Global:gridLicenses.ItemsSource=$Script:Licenses;$Global:gridMfa.ItemsSource=$Script:Mfa;$Global:gridMail.ItemsSource=$Script:Mail;$Global:gridDevices.ItemsSource=$Script:Devices;$Global:gridRecommendations.ItemsSource=$Script:Recommendations
$Global:btnRefresh.Add_Click({$Global:btnRefresh.IsEnabled=$false;$Global:Window.Cursor='Wait';try{Collect-Data;Update-Cards}catch{Write-Log $_.Exception.Message ERROR}finally{$Global:Window.Cursor=$null;$Global:btnRefresh.IsEnabled=$true}})
$Global:btnNewSession.Add_Click({$Global:Window.Hide();try{Connect-FreshSession;$Global:txtTenant.Text="Global Administrator: $Script:AdminUPN | Tenant: $Script:TenantId | Graph context: Process"}catch{[System.Windows.MessageBox]::Show($_.Exception.Message,'Connection failed')|Out-Null}finally{$Global:Window.Show();$Global:Window.Activate()}})
$Global:btnCsv.Add_Click({Export-CsvReports});$Global:btnHtml.Add_Click({Export-HtmlReport})
$Global:Window.Add_Closing({Disconnect-All})
Write-Log "Dashboard v$Script:Version opened successfully. Select Refresh to collect tenant data."
[void]$Global:Window.ShowDialog()
