
<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Microsoft 365 Enterprise Command Center
VERSION     : 3.2 REST App-Only Dashboard
PURPOSE     : Technician console using app-only certificate auth and direct
              Microsoft Graph REST calls.
REQUIRES    : C:\Reporting\M365EOC_AppOnly_Config.json from Setup-M365EOC-AppOnly.ps1
==============================================================================
#>

$ErrorActionPreference = 'Continue'
$Script:Version = '3.2'
$Script:ReportPath = 'C:\Reporting'
$Script:ConfigPath = Join-Path $Script:ReportPath 'M365EOC_AppOnly_Config.json'
$Script:GraphConnected = $false

# ---------------- Collections ----------------
$Script:Findings = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Licenses = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Users = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Roles = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:CAPolicies = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:CAExclusions = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:NamedLocations = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Apps = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Privileged = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Enhanced = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:Devices = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:MailFlow = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:SecurityGroups = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:DistributionGroups = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:M365Groups = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:EmptyGroups = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$Script:RawUsers = @()
$Script:RawRoleAssignments = @()
$Script:RawCAExclusions = @()

$Script:MetricKeys = @(
    'Tenant','TotalUsers','ActiveUsers','DisabledUsers','LicensedUsers','GuestUsers',
    'TotalLicenses','AssignedLicenses','AvailableLicenses','DirectoryRoles','GlobalAdmins',
    'CAPolicies','EnabledPolicies','ReportOnlyPolicies','DisabledPolicies','CAExclusions','NamedLocations',
    'EnterpriseApps','SSOApps','PrivilegedAccounts','SecurityGroups','DistributionGroups','M365Groups','EmptyGroups',
    'Domains','SecureScore','MfaStatus','UsersWithoutMfa','RiskyUsers','FailedSignIns','ServiceHealth','MessageCenter','DeviceCompliance','MailFlowFailures'
)
$Script:Metrics = [ordered]@{}
foreach($k in $Script:MetricKeys){ $Script:Metrics[$k] = 'N/A' }

$Script:CardLabels = [ordered]@{
    Tenant='Tenant'; TotalUsers='Total Users'; ActiveUsers='Active Users'; DisabledUsers='Disabled Users'; LicensedUsers='Licensed Users'; GuestUsers='Guest Users'
    TotalLicenses='Total Licenses'; AssignedLicenses='Assigned Licenses'; AvailableLicenses='Available Licenses'; DirectoryRoles='Directory Roles'; GlobalAdmins='Global Admins'
    CAPolicies='CA Policies'; EnabledPolicies='Enabled Policies'; ReportOnlyPolicies='Report Only Policies'; DisabledPolicies='Disabled Policies'; CAExclusions='CA Exclusions'; NamedLocations='Named Locations'
    EnterpriseApps='Enterprise Apps'; SSOApps='SSO Apps'; PrivilegedAccounts='Privileged Accounts'; SecurityGroups='Security Groups'; DistributionGroups='Distribution Groups'; M365Groups='Microsoft 365 Groups'; EmptyGroups='Empty Groups'
    Domains='Domains'; SecureScore='Secure Score'; MfaStatus='MFA Status'; UsersWithoutMfa='Users Without MFA'; RiskyUsers='Risky Users'; FailedSignIns='Failed Sign-ins'; ServiceHealth='Service Health'; MessageCenter='Message Center'; DeviceCompliance='Device Compliance'; MailFlowFailures='Mail Flow Failures 24h'
}

function Write-Log {
    param([string]$Message,[string]$Level='INFO')
    $line = '[{0}][{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Write-Host $line
    if($Script:txtLog){ try { $Script:txtLog.AppendText($line + [Environment]::NewLine); $Script:txtLog.ScrollToEnd() } catch {} }
}
function Ensure-AuthModule {
    $name='Microsoft.Graph.Authentication'
    $m=Get-Module -ListAvailable -Name $name | Sort-Object Version -Descending | Select-Object -First 1
    if(-not $m){ Write-Log "Installing $name..." 'WARN'; Install-Module $name -Scope CurrentUser -Force -AllowClobber }
    Import-Module $name -Force -ErrorAction Stop
}
function Ensure-ImportExcel {
    try {
        if(-not (Get-Module -ListAvailable -Name ImportExcel)){ Write-Log 'Installing ImportExcel module for workbook export...' 'WARN'; Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber }
        Import-Module ImportExcel -Force -ErrorAction Stop
        return $true
    } catch { Write-Log "ImportExcel unavailable. $($_.Exception.Message)" 'ERROR'; return $false }
}
function Add-Finding { param($Area,$Status,$Severity,$Title,$Details,$Recommendation)
    $Script:Findings.Add([pscustomobject]@{Area=$Area;Status=$Status;Severity=$Severity;Title=$Title;Details=$Details;Recommendation=$Recommendation;Time=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')}) | Out-Null
}
function Add-Enhanced { param($Area,$Metric,$Value,$Status,$Details)
    $Script:Enhanced.Add([pscustomobject]@{Area=$Area;Metric=$Metric;Value=$Value;Status=$Status;Details=$Details}) | Out-Null
}
function ConvertTo-FlatString { param($Value)
    if($null -eq $Value){return ''}; if($Value -is [string]){return $Value}
    try { if($Value -is [System.Collections.IEnumerable]){ return (($Value | ForEach-Object { $_.ToString() }) -join '; ') } } catch {}
    try { return ($Value | ConvertTo-Json -Compress -Depth 8) } catch { return [string]$Value }
}
function Clear-Data {
    foreach($c in @($Script:Findings,$Script:Licenses,$Script:Users,$Script:Roles,$Script:CAPolicies,$Script:CAExclusions,$Script:NamedLocations,$Script:Apps,$Script:Privileged,$Script:Enhanced,$Script:Devices,$Script:MailFlow,$Script:SecurityGroups,$Script:DistributionGroups,$Script:M365Groups,$Script:EmptyGroups)){ $c.Clear() }
    foreach($k in $Script:MetricKeys){ $Script:Metrics[$k]='N/A' }
    $Script:RawUsers=@(); $Script:RawRoleAssignments=@(); $Script:RawCAExclusions=@()
}
function Invoke-GraphGetPage { param([string]$Uri,[hashtable]$Headers)
    if($Headers){ Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers $Headers -ErrorAction Stop } else { Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop }
}
function Get-GraphAll { param([string]$Uri,[hashtable]$Headers)
    $out=@(); $next=$Uri
    while($next){
        $r=Invoke-GraphGetPage -Uri $next -Headers $Headers
        if($r.value){ $out += @($r.value) } elseif($r){ $out += @($r) }
        if($r.'@odata.nextLink'){$next=$r.'@odata.nextLink'}else{$next=$null}
    }
    return @($out)
}
function Try-GraphAll { param($Area,$Metric,$Uri,[hashtable]$Headers)
    try { return @(Get-GraphAll -Uri $Uri -Headers $Headers) } catch { Add-Enhanced $Area $Metric 'Permission Required' 'Blocked' $_.Exception.Message; return @() }
}
function Get-GraphCount { param([string]$Uri)
    try { return [int](Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers @{ConsistencyLevel='eventual'} -ErrorAction Stop) } catch { return $null }
}

function Connect-GraphAppOnly {
    if(-not(Test-Path $Script:ConfigPath)){ [System.Windows.MessageBox]::Show("Missing $($Script:ConfigPath). Run setup first.",'Missing config','OK','Error')|Out-Null; return }
    try{
        Ensure-AuthModule
        $cfg=Get-Content $Script:ConfigPath -Raw | ConvertFrom-Json
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Connect-MgGraph -TenantId $cfg.TenantId -ClientId $cfg.ClientId -CertificateThumbprint $cfg.CertificateThumbprint -ContextScope Process -NoWelcome -ErrorAction Stop | Out-Null
        $Script:GraphConnected=$true
        $Script:txtTenant.Text="Connected Tenant: $($cfg.TenantId) | AppId: $($cfg.ClientId)"
        $Script:btnConnect.Content='Connected'; $Script:btnConnect.Background = '#16A34A'
        Write-Log "Connected to Microsoft Graph app-only. AppId: $($cfg.ClientId)"
    } catch { Write-Log "Graph app-only connection failed: $($_.Exception.Message)" 'ERROR' }
}

function Collect-TenantSummary {
    Write-Log 'Collecting Tenant Summary...'
    try{
        $org=(Get-GraphAll 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName,tenantType,countryLetterCode,verifiedDomains' | Select-Object -First 1)
        $Script:Metrics.Tenant=$org.displayName
        $domains=(($org.verifiedDomains | ForEach-Object {$_.name}) -join '; ')
        Add-Finding 'Tenant Summary' 'Collected' 'Info' $org.displayName "TenantId: $($org.id) | Type: $($org.tenantType) | Domains: $domains" 'Review tenant identity and verified domains periodically.'
    } catch { Add-Finding 'Tenant Summary' 'Error' 'Medium' 'Tenant summary failed' $_.Exception.Message 'Check app permissions.' }
}
function Collect-Licenses {
    Write-Log 'Collecting Licenses...'
    $data=Try-GraphAll 'Licenses' 'Subscribed SKUs' 'https://graph.microsoft.com/v1.0/subscribedSkus'
    $total=0;$used=0;$avail=0;$exhausted=0
    foreach($sku in $data){$t=[int]$sku.prepaidUnits.enabled;$u=[int]$sku.consumedUnits;$a=$t-$u;if($a -lt 0){$a=0};$pct=if($t -gt 0){[math]::Round(($u/$t)*100,2)}else{0};$status=if($a -le 0){'Exhausted'}elseif($pct -ge 90){'Low Availability'}else{'OK'};$Script:Licenses.Add([pscustomobject]@{LicenseName=$sku.skuPartNumber;Total=$t;Assigned=$u;Available=$a;UsagePercent="$pct%";Status=$status})|Out-Null;$total+=$t;$used+=$u;$avail+=$a;if($a -le 0){$exhausted++}}
    if($data.Count -gt 0){$Script:Metrics.TotalLicenses=$total;$Script:Metrics.AssignedLicenses=$used;$Script:Metrics.AvailableLicenses=$avail;Add-Finding 'Licenses' 'Collected' 'Info' 'License summary' "Total: $total | Assigned: $used | Available: $avail | Exhausted SKUs: $exhausted" 'Review license usage and exhausted SKUs.'}
}
function Collect-Users {
    Write-Log 'Collecting Users...'
    $url='https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,mail,department,jobTitle,accountEnabled,createdDateTime,userType,description,assignedLicenses&$top=999'
    $Script:RawUsers=Try-GraphAll 'Users' 'Users' $url
    foreach($u in $Script:RawUsers){$Script:Users.Add([pscustomobject]@{DisplayName=$u.displayName;UserPrincipalName=$u.userPrincipalName;Mail=$u.mail;Department=$u.department;JobTitle=$u.jobTitle;UserType=$u.userType;Description=$u.description;AccountEnabled=$u.accountEnabled;CreatedDateTime=$u.createdDateTime;AssignedLicenseCount=if($u.assignedLicenses){@($u.assignedLicenses).Count}else{0}})|Out-Null}
    if($Script:RawUsers.Count -gt 0){$Script:Metrics.TotalUsers=$Script:RawUsers.Count;$Script:Metrics.ActiveUsers=@($Script:RawUsers|Where-Object{$_.accountEnabled -eq $true}).Count;$Script:Metrics.DisabledUsers=@($Script:RawUsers|Where-Object{$_.accountEnabled -eq $false}).Count;$Script:Metrics.LicensedUsers=@($Script:RawUsers|Where-Object{@($_.assignedLicenses).Count -gt 0}).Count;$Script:Metrics.GuestUsers=@($Script:RawUsers|Where-Object{$_.userType -eq 'Guest'}).Count;Add-Finding 'Users' 'Collected' 'Info' 'User summary' "Total: $($Script:Metrics.TotalUsers) | Active: $($Script:Metrics.ActiveUsers) | Disabled: $($Script:Metrics.DisabledUsers) | Licensed: $($Script:Metrics.LicensedUsers) | Guests: $($Script:Metrics.GuestUsers)" 'Review disabled licensed users and guest users periodically.'}
}
function Collect-Roles {
    Write-Log 'Collecting Directory Roles and Global Admins...'
    $roles=Try-GraphAll 'Directory Roles' 'Roles' 'https://graph.microsoft.com/v1.0/directoryRoles'
    foreach($role in $roles){$members=Try-GraphAll 'Directory Roles' "Members $($role.displayName)" "https://graph.microsoft.com/v1.0/directoryRoles/$($role.id)/members";foreach($m in $members){$row=[pscustomobject]@{RoleName=$role.displayName;MemberObjectId=$m.id;DisplayName=$m.displayName;UserPrincipalName=$m.userPrincipalName;Mail=$m.mail;Description=$m.description;AccountEnabled=$m.accountEnabled};$Script:RawRoleAssignments+=$row;$Script:Roles.Add($row)|Out-Null}}
    $Script:Metrics.DirectoryRoles=$roles.Count;$Script:Metrics.GlobalAdmins=@($Script:RawRoleAssignments|Where-Object{$_.RoleName -eq 'Global Administrator'}).Count
    if($roles.Count -gt 0){Add-Finding 'Directory Roles' 'Collected' 'Info' 'Role summary' "Roles: $($roles.Count) | Assignments: $($Script:RawRoleAssignments.Count) | Global Admins: $($Script:Metrics.GlobalAdmins)" 'Review privileged role assignments regularly.'}
}
function Collect-ConditionalAccess {
    Write-Log 'Collecting Conditional Access Policies and Named Locations...'
    $policies=Try-GraphAll 'Conditional Access' 'Policies' 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies'
    foreach($p in $policies){
        $row=[pscustomobject]@{DisplayName=$p.displayName;State=$p.state;PolicyId=$p.id;CreatedDateTime=$p.createdDateTime;ModifiedDateTime=$p.modifiedDateTime;IncludeUsers=ConvertTo-FlatString $p.conditions.users.includeUsers;ExcludeUsers=ConvertTo-FlatString $p.conditions.users.excludeUsers;IncludeGroups=ConvertTo-FlatString $p.conditions.users.includeGroups;ExcludeGroups=ConvertTo-FlatString $p.conditions.users.excludeGroups;IncludeRoles=ConvertTo-FlatString $p.conditions.users.includeRoles;ExcludeRoles=ConvertTo-FlatString $p.conditions.users.excludeRoles;IncludeApplications=ConvertTo-FlatString $p.conditions.applications.includeApplications;ExcludeApplications=ConvertTo-FlatString $p.conditions.applications.excludeApplications;IncludeLocations=ConvertTo-FlatString $p.conditions.locations.includeLocations;ExcludeLocations=ConvertTo-FlatString $p.conditions.locations.excludeLocations;GrantControls=ConvertTo-FlatString $p.grantControls.builtInControls;SessionControls=try{$p.sessionControls|ConvertTo-Json -Compress -Depth 8}catch{''}}
        $Script:CAPolicies.Add($row)|Out-Null
        foreach($id in @($p.conditions.users.excludeUsers)){if($id -and $id -ne 'GuestsOrExternalUsers'){$ex=[pscustomobject]@{PolicyName=$p.displayName;PolicyState=$p.state;ExcludedUserId=$id;DisplayName='';UserPrincipalName='';Mail='';AccountEnabled='';Description=''};$Script:RawCAExclusions+=$ex;$Script:CAExclusions.Add($ex)|Out-Null}}
    }
    $named=Try-GraphAll 'Named Locations' 'Named Locations' 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations'
    foreach($n in $named){$Script:NamedLocations.Add([pscustomobject]@{DisplayName=$n.displayName;Id=$n.id;CreatedDateTime=$n.createdDateTime;ModifiedDateTime=$n.modifiedDateTime;Type=$n.'@odata.type'; Details=($n|ConvertTo-Json -Compress -Depth 8)})|Out-Null}
    $Script:Metrics.CAPolicies=$Script:CAPolicies.Count;$Script:Metrics.CAExclusions=$Script:CAExclusions.Count;$Script:Metrics.NamedLocations=$Script:NamedLocations.Count
    $Script:Metrics.EnabledPolicies=@($Script:CAPolicies|Where-Object{$_.State -eq 'enabled'}).Count
    $Script:Metrics.ReportOnlyPolicies=@($Script:CAPolicies|Where-Object{$_.State -eq 'enabledForReportingButNotEnforced'}).Count
    $Script:Metrics.DisabledPolicies=@($Script:CAPolicies|Where-Object{$_.State -eq 'disabled'}).Count
    Add-Finding 'Conditional Access' 'Collected' 'Info' 'Conditional Access summary' "Policies: $($Script:CAPolicies.Count) | Enabled: $($Script:Metrics.EnabledPolicies) | Report Only: $($Script:Metrics.ReportOnlyPolicies) | Disabled: $($Script:Metrics.DisabledPolicies) | Exclusions: $($Script:CAExclusions.Count) | Named Locations: $($Script:NamedLocations.Count)" 'Review CA exclusions, policy states, and named locations.'
}
function Collect-Apps {
    Write-Log 'Collecting Enterprise Applications and SSO Applications...'
    $apps=Try-GraphAll 'Enterprise Apps' 'Service Principals' 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=displayName,appId,id,accountEnabled,servicePrincipalType,preferredSingleSignOnMode,publisherName,appOwnerOrganizationId,signInAudience,tags,homepage,loginUrl,logoutUrl&$top=999'
    foreach($a in $apps){$Script:Apps.Add([pscustomobject]@{DisplayName=$a.displayName;AppId=$a.appId;ObjectId=$a.id;AccountEnabled=$a.accountEnabled;ServicePrincipalType=$a.servicePrincipalType;PreferredSingleSignOnMode=$a.preferredSingleSignOnMode;PublisherName=$a.publisherName;AppOwnerOrganizationId=$a.appOwnerOrganizationId;SignInAudience=$a.signInAudience;Tags=($a.tags -join '; ');Homepage=$a.homepage;LoginUrl=$a.loginUrl;LogoutUrl=$a.logoutUrl})|Out-Null}
    $Script:Metrics.EnterpriseApps=$Script:Apps.Count;$Script:Metrics.SSOApps=@($Script:Apps|Where-Object{$_.PreferredSingleSignOnMode}).Count
    if($apps.Count -gt 0){Add-Finding 'Enterprise Apps' 'Collected' 'Info' 'Enterprise app summary' "Enterprise Apps: $($Script:Metrics.EnterpriseApps) | SSO Apps: $($Script:Metrics.SSOApps)" 'Review stale apps, ownership, sign-in settings, and SSO mode.'}
}
function Collect-Groups {
    Write-Log 'Collecting Groups...'
    $groups=Try-GraphAll 'Groups' 'Groups' 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,createdDateTime&$top=999'
    foreach($g in $groups){
        $memberCount=Get-GraphCount "https://graph.microsoft.com/v1.0/groups/$($g.id)/members/`$count"
        if($null -eq $memberCount){$memberCount='N/A'}
        $isUnified = @($g.groupTypes) -contains 'Unified'
        $row=[pscustomobject]@{DisplayName=$g.displayName;Description=$g.description;Mail=$g.mail;MailEnabled=$g.mailEnabled;SecurityEnabled=$g.securityEnabled;GroupTypes=ConvertTo-FlatString $g.groupTypes;CreatedDateTime=$g.createdDateTime;MemberCount=$memberCount}
        if($g.securityEnabled -eq $true -and $isUnified -eq $false){$Script:SecurityGroups.Add($row)|Out-Null}
        if($g.mailEnabled -eq $true -and $isUnified -eq $false){$Script:DistributionGroups.Add($row)|Out-Null}
        if($isUnified){$Script:M365Groups.Add($row)|Out-Null}
        if($memberCount -eq 0){$Script:EmptyGroups.Add($row)|Out-Null}
    }
    $Script:Metrics.SecurityGroups=$Script:SecurityGroups.Count; $Script:Metrics.DistributionGroups=$Script:DistributionGroups.Count; $Script:Metrics.M365Groups=$Script:M365Groups.Count; $Script:Metrics.EmptyGroups=$Script:EmptyGroups.Count
    Add-Finding 'Groups' 'Collected' 'Info' 'Group summary' "Security Groups: $($Script:Metrics.SecurityGroups) | Distribution Groups: $($Script:Metrics.DistributionGroups) | Microsoft 365 Groups: $($Script:Metrics.M365Groups) | Empty Groups: $($Script:Metrics.EmptyGroups)" 'Review empty groups and group sprawl periodically.'
}
function Collect-Privileged {
    Write-Log 'Collecting Privileged Accounts...'
    $keyword=@($Script:RawUsers|Where-Object{($_.description -match 'break|emergency') -or ($_.displayName -match 'break|emergency') -or ($_.userPrincipalName -match 'breakglass|break-glass|emergency|bg')}|ForEach-Object{[pscustomobject]@{Source='Name/UPN/Description Match';DisplayName=$_.displayName;UserPrincipalName=$_.userPrincipalName;Mail=$_.mail;Description=$_.description;AccountEnabled=$_.accountEnabled}})
    $ca=@($Script:RawCAExclusions|ForEach-Object{[pscustomobject]@{Source='Conditional Access Exclusion';DisplayName=$_.DisplayName;UserPrincipalName=$_.UserPrincipalName;Mail=$_.Mail;Description=$_.Description;AccountEnabled=$_.AccountEnabled}})
    $ga=@($Script:RawRoleAssignments|Where-Object{$_.RoleName -eq 'Global Administrator'}|ForEach-Object{[pscustomobject]@{Source='Global Administrator';DisplayName=$_.DisplayName;UserPrincipalName=$_.UserPrincipalName;Mail=$_.Mail;Description=$_.Description;AccountEnabled=$_.AccountEnabled}})
    $all=@($keyword+$ca+$ga)|Where-Object{$_.UserPrincipalName -or $_.DisplayName -or $_.Source -eq 'Conditional Access Exclusion'}|Sort-Object UserPrincipalName,Source -Unique
    foreach($p in $all){$Script:Privileged.Add($p)|Out-Null};$Script:Metrics.PrivilegedAccounts=$Script:Privileged.Count
    Add-Finding 'Privileged Accounts' 'Collected' 'Info' 'Privileged account summary' "Candidates: $($Script:Privileged.Count)" 'Review privileged and CA-excluded accounts.'
}
function Collect-Enhanced {
    Write-Log 'Collecting enhanced operations metrics...'
    try{$domains=Get-GraphAll 'https://graph.microsoft.com/v1.0/domains';$issues=@($domains|Where-Object{-not $_.isVerified}).Count;$Script:Metrics.Domains="$issues issues";Add-Enhanced 'Domains' 'Domain Count' $domains.Count 'Collected' "Unverified: $issues"}catch{Add-Enhanced 'Domains' 'Domain Count' 'N/A' 'Error' $_.Exception.Message}
    $svc=Try-GraphAll 'Service Health' 'Active Issues' 'https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues'; if($svc.Count -gt 0){$active=@($svc|Where-Object{$_.status -notmatch 'serviceRestored|resolved|postIncidentReviewPublished'}).Count;$Script:Metrics.ServiceHealth=$active;Add-Enhanced 'Service Health' 'Active Issues' $active 'Collected' "Total returned: $($svc.Count)"}
    $msg=Try-GraphAll 'Message Center' 'Major Messages' 'https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages'; if($msg.Count -gt 0){$major=@($msg|Where-Object{$_.isMajorChange}).Count;$Script:Metrics.MessageCenter=$major;Add-Enhanced 'Message Center' 'Major Messages' $major 'Collected' "Total returned: $($msg.Count)"}
    try{$r=Invoke-GraphGetPage 'https://graph.microsoft.com/v1.0/security/secureScores?$top=1';$s=@($r.value|Select-Object -First 1)[0];if($s -and [double]$s.maxScore -gt 0){$pct=[math]::Round(([double]$s.currentScore/[double]$s.maxScore)*100,2);$Script:Metrics.SecureScore="$($s.currentScore) / $($s.maxScore) ($pct%)"}else{$Script:Metrics.SecureScore='No data'};Add-Enhanced 'Secure Score' 'Secure Score' $Script:Metrics.SecureScore 'Collected' ''}catch{$Script:Metrics.SecureScore='Permission Required';Add-Enhanced 'Secure Score' 'Secure Score' 'Permission Required' 'Blocked' $_.Exception.Message}
    $risk=Try-GraphAll 'Risky Users' 'Risky User Count' 'https://graph.microsoft.com/v1.0/identityProtection/riskyUsers'; if($risk.Count -gt 0){$Script:Metrics.RiskyUsers=$risk.Count;Add-Enhanced 'Risky Users' 'Risky User Count' $risk.Count 'Collected' ''}
    try{$start=(Get-Date).AddDays(-1).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ');$r=Invoke-GraphGetPage "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$top=100&`$filter=createdDateTime ge $start and status/errorCode ne 0";$cnt=@($r.value).Count;$Script:Metrics.FailedSignIns=$cnt;Add-Enhanced 'Failed Sign-ins' 'Failed Sign-ins 24h' $cnt 'Collected' 'Top 100 returned'}catch{$Script:Metrics.FailedSignIns='Permission Required';Add-Enhanced 'Failed Sign-ins' 'Failed Sign-ins 24h' 'Permission Required' 'Blocked' $_.Exception.Message}
    try{$enabled=@($Script:RawUsers|Where-Object{$_.accountEnabled -eq $true -and $_.userType -ne 'Guest'});$with=0;$without=0;$checked=0;foreach($u in $enabled){$checked++;try{$m=Get-GraphAll "https://graph.microsoft.com/v1.0/users/$($u.id)/authentication/methods";$mfa=@($m|Where-Object{$_.'@odata.type' -notmatch 'passwordAuthenticationMethod'});if($mfa.Count -gt 0){$with++}else{$without++}}catch{}};$pct=if($checked -gt 0){[math]::Round(($with/$checked)*100,2)}else{0};$Script:Metrics.MfaStatus="$pct%";$Script:Metrics.UsersWithoutMfa=$without;Add-Enhanced 'MFA' 'MFA Coverage' "$pct%" 'Collected' "Checked: $checked | With MFA: $with | Without: $without"}catch{$Script:Metrics.MfaStatus='Permission Required';$Script:Metrics.UsersWithoutMfa='Permission Required';Add-Enhanced 'MFA' 'MFA Coverage' 'Permission Required' 'Blocked' $_.Exception.Message}
    $devices=Try-GraphAll 'Intune Devices' 'Managed Devices' 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices'; if($devices.Count -gt 0){$total=$devices.Count;$comp=@($devices|Where-Object{$_.complianceState -eq 'compliant'}).Count;$non=@($devices|Where-Object{$_.complianceState -eq 'noncompliant'}).Count;$pct=if($total -gt 0){[math]::Round(($comp/$total)*100,2)}else{0};foreach($d in $devices|Select-Object -First 500){$Script:Devices.Add([pscustomobject]@{DeviceName=$d.deviceName;UserPrincipalName=$d.userPrincipalName;OperatingSystem=$d.operatingSystem;ComplianceState=$d.complianceState;LastSyncDateTime=$d.lastSyncDateTime})|Out-Null};$Script:Metrics.DeviceCompliance="$pct%";Add-Enhanced 'Intune Devices' 'Device Compliance' "$pct%" 'Collected' "Total: $total | Compliant: $comp | Non-compliant: $non"}
}
function Collect-MailFlow { $Script:Metrics.MailFlowFailures='N/A'; Add-Enhanced 'Mail Flow' 'Failures 24h' 'N/A' 'Skipped' 'Mail flow requires optional Exchange app-only connection in a later build.' }
function Refresh-Dashboard {
    if(-not $Script:GraphConnected){[System.Windows.MessageBox]::Show('Connect first.','Not connected','OK','Warning')|Out-Null;return}
    Clear-Data; Write-Log 'Refreshing dashboard using REST app-only Graph calls...'; $Script:txtLastRefresh.Text='Refreshing...'
    Collect-TenantSummary; Collect-Licenses; Collect-Users; Collect-Roles; Collect-ConditionalAccess; Collect-Apps; Collect-Groups; Collect-Privileged; Collect-Enhanced; Collect-MailFlow
    Update-Cards; $Script:txtLastRefresh.Text="Last Refresh: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"; Write-Log 'Dashboard refresh completed.'
}
function Export-Workbook {
    if(-not (Ensure-ImportExcel)){return}
    $stamp=Get-Date -Format 'yyyyMMdd'
    $path=Join-Path $Script:ReportPath "M365_Enterprise_Command_Center_$stamp.xlsx"
    if(Test-Path $path){Remove-Item $path -Force}
    $summary=@(); foreach($k in $Script:Metrics.Keys){$summary += [pscustomobject]@{Metric=$k;Value=$Script:Metrics[$k]}}
    $summary | Export-Excel -Path $path -WorksheetName 'Executive Summary' -AutoSize -FreezeTopRow -BoldTopRow -TableName 'ExecutiveSummary' -TableStyle Medium2
    @(@($Script:Users,'Users'),@($Script:Licenses,'Licenses'),@($Script:SecurityGroups,'Security Groups'),@($Script:DistributionGroups,'Distribution Groups'),@($Script:M365Groups,'Microsoft 365 Groups'),@($Script:EmptyGroups,'Empty Groups'),@($Script:Roles,'Directory Roles'),@($Script:CAPolicies,'CA Policies'),@($Script:CAExclusions,'CA Exclusions'),@($Script:NamedLocations,'Named Locations'),@($Script:Apps,'Enterprise Apps'),@($Script:Privileged,'Privileged Accounts'),@($Script:Enhanced,'Enhanced Metrics'),@($Script:Devices,'Devices'),@($Script:MailFlow,'Mail Flow'),@($Script:Findings,'Findings')) | ForEach-Object { $_[0] | Export-Excel -Path $path -WorksheetName $_[1] -AutoSize -FreezeTopRow -BoldTopRow -Append }
    Write-Log "Workbook exported: $path"; [System.Windows.MessageBox]::Show("Workbook exported:`n$path",'Export Complete','OK','Information')|Out-Null
}

# ---------------- UI ----------------
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Xaml
$win = New-Object System.Windows.Window
$win.Title='Microsoft 365 Enterprise Command Center v3.2'; $win.Height=1000; $win.Width=1720; $win.WindowStartupLocation='CenterScreen'; $win.Background='#F1F5F9'
$grid = New-Object System.Windows.Controls.Grid; $grid.Margin='14'
@('Auto','Auto','*','210') | ForEach-Object { $rd=New-Object System.Windows.Controls.RowDefinition; $rd.Height=$_; $grid.RowDefinitions.Add($rd) }
$win.Content=$grid

# Header
$header=New-Object System.Windows.Controls.Border; $header.Background='#0F172A'; $header.CornerRadius='16'; $header.Padding='18'; $header.Margin='0,0,0,12'; [System.Windows.Controls.Grid]::SetRow($header,0); $grid.Children.Add($header)|Out-Null
$dock=New-Object System.Windows.Controls.DockPanel; $header.Child=$dock
$titleStack=New-Object System.Windows.Controls.StackPanel; [System.Windows.Controls.DockPanel]::SetDock($titleStack,'Left'); $dock.Children.Add($titleStack)|Out-Null
$t=New-Object System.Windows.Controls.TextBlock; $t.Text='Microsoft 365 Enterprise Command Center'; $t.Foreground='White'; $t.FontSize=28; $t.FontWeight='Bold'; $titleStack.Children.Add($t)|Out-Null
$v=New-Object System.Windows.Controls.TextBlock; $v.Text='Version 3.2'; $v.Foreground='#93C5FD'; $v.FontSize=12; $titleStack.Children.Add($v)|Out-Null
$Script:txtTenant=New-Object System.Windows.Controls.TextBlock; $Script:txtTenant.Text='Connected Tenant: Not connected'; $Script:txtTenant.Foreground='#CBD5E1'; $Script:txtTenant.FontSize=13; $Script:txtTenant.Margin='0,6,0,0'; $titleStack.Children.Add($Script:txtTenant)|Out-Null
$Script:txtLastRefresh=New-Object System.Windows.Controls.TextBlock; $Script:txtLastRefresh.Text='Last Refresh: Never'; $Script:txtLastRefresh.Foreground='#CBD5E1'; $Script:txtLastRefresh.FontSize=13; $titleStack.Children.Add($Script:txtLastRefresh)|Out-Null
$btnStack=New-Object System.Windows.Controls.StackPanel; $btnStack.Orientation='Horizontal'; [System.Windows.Controls.DockPanel]::SetDock($btnStack,'Right'); $dock.Children.Add($btnStack)|Out-Null
function New-Button($text,$color){ $b=New-Object System.Windows.Controls.Button; $b.Content=$text; $b.Width=110; $b.Height=38; $b.Margin='6'; $b.Background=$color; $b.Foreground='White'; $b.FontWeight='SemiBold'; return $b }
$Script:btnConnect=New-Button 'Connect' '#0EA5E9'; $btnStack.Children.Add($Script:btnConnect)|Out-Null
$Script:btnRefresh=New-Button 'Refresh' '#16A34A'; $btnStack.Children.Add($Script:btnRefresh)|Out-Null
$Script:btnExport=New-Button 'Export CSV' '#475569'; $btnStack.Children.Add($Script:btnExport)|Out-Null

# Cards
$scroll=New-Object System.Windows.Controls.ScrollViewer; $scroll.VerticalScrollBarVisibility='Auto'; [System.Windows.Controls.Grid]::SetRow($scroll,1); $grid.Children.Add($scroll)|Out-Null
$cards=New-Object System.Windows.Controls.UniformGrid; $cards.Columns=6; $cards.Margin='0,0,0,12'; $scroll.Content=$cards
foreach($key in $Script:MetricKeys){
    $border=New-Object System.Windows.Controls.Border; $border.Name="card$key"; $border.Background='#64748B'; $border.CornerRadius='14'; $border.Padding='10'; $border.Margin='4'
    $sp=New-Object System.Windows.Controls.StackPanel
    $lab=New-Object System.Windows.Controls.TextBlock; $lab.Text=$Script:CardLabels[$key]; $lab.Foreground='White'; $lab.FontWeight='Bold'
    $val=New-Object System.Windows.Controls.TextBlock; $val.Name="lbl$key"; $val.Text='N/A'; $val.Foreground='White'; $val.FontSize=20; $val.FontWeight='Bold'; $val.TextWrapping='Wrap'
    $sp.Children.Add($lab)|Out-Null; $sp.Children.Add($val)|Out-Null; $border.Child=$sp; $cards.Children.Add($border)|Out-Null
    Set-Variable -Name "card$key" -Value $border -Scope Script
    Set-Variable -Name "lbl$key" -Value $val -Scope Script
}

# Tabs
$tabs=New-Object System.Windows.Controls.TabControl; [System.Windows.Controls.Grid]::SetRow($tabs,2); $grid.Children.Add($tabs)|Out-Null
function Add-Tab($name,$source){ $ti=New-Object System.Windows.Controls.TabItem; $ti.Header=$name; $dg=New-Object System.Windows.Controls.DataGrid; $dg.AutoGenerateColumns=$true; $dg.IsReadOnly=$true; $dg.ItemsSource=$source; $ti.Content=$dg; $tabs.Items.Add($ti)|Out-Null }
Add-Tab 'Findings' $Script:Findings; Add-Tab 'Licenses' $Script:Licenses; Add-Tab 'Users' $Script:Users; Add-Tab 'Roles' $Script:Roles; Add-Tab 'CA Policies' $Script:CAPolicies; Add-Tab 'CA Exclusions' $Script:CAExclusions; Add-Tab 'Named Locations' $Script:NamedLocations; Add-Tab 'Security Groups' $Script:SecurityGroups; Add-Tab 'Distribution Groups' $Script:DistributionGroups; Add-Tab 'Microsoft 365 Groups' $Script:M365Groups; Add-Tab 'Empty Groups' $Script:EmptyGroups; Add-Tab 'Apps' $Script:Apps; Add-Tab 'Privileged' $Script:Privileged; Add-Tab 'Enhanced' $Script:Enhanced; Add-Tab 'Devices' $Script:Devices; Add-Tab 'Mail Flow' $Script:MailFlow

# Log
$logBorder=New-Object System.Windows.Controls.Border; $logBorder.Background='#020617'; $logBorder.CornerRadius='12'; $logBorder.Padding='12'; $logBorder.Margin='0,12,0,0'; [System.Windows.Controls.Grid]::SetRow($logBorder,3); $grid.Children.Add($logBorder)|Out-Null
$logDock=New-Object System.Windows.Controls.DockPanel; $logBorder.Child=$logDock
$logTitle=New-Object System.Windows.Controls.TextBlock; $logTitle.Text='Activity Log'; $logTitle.Foreground='White'; $logTitle.FontSize=14; $logTitle.FontWeight='Bold'; $logTitle.Margin='0,0,0,6'; [System.Windows.Controls.DockPanel]::SetDock($logTitle,'Top'); $logDock.Children.Add($logTitle)|Out-Null
$Script:txtLog=New-Object System.Windows.Controls.TextBox; $Script:txtLog.Background='#020617'; $Script:txtLog.Foreground='#C4B5FD'; $Script:txtLog.BorderBrush='#334155'; $Script:txtLog.FontFamily='Consolas'; $Script:txtLog.FontSize=12; $Script:txtLog.TextWrapping='Wrap'; $Script:txtLog.AcceptsReturn=$true; $Script:txtLog.VerticalScrollBarVisibility='Auto'; $Script:txtLog.IsReadOnly=$true; $logDock.Children.Add($Script:txtLog)|Out-Null

$Script:btnConnect.Add_Click({Connect-GraphAppOnly})
$Script:btnRefresh.Add_Click({Refresh-Dashboard})
$Script:btnExport.Add_Click({Export-Workbook})
Write-Log 'Dashboard v3.2 loaded. Click Connect, then Refresh.'
[void]$win.ShowDialog()
