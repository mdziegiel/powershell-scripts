#Requires -Version 5.1
<#
.SYNOPSIS
    Offboards a user from Hans Kissle - disables AD account, revokes M365 license,
    moves to disabled OU, and clears group memberships.
.DESCRIPTION
    Full offboarding workflow:
    - Disables the AD account
    - Clears the manager field
    - Removes all AD group memberships (except Domain Users)
    - Moves account to the Disabled Users OU
    - Revokes all M365 licenses via Microsoft Graph
    - Logs all actions taken
.PARAMETER SamAccountName
    The SamAccountName of the user to offboard.
.PARAMETER DisabledOU
    DistinguishedName of the OU to move disabled accounts to.
.PARAMETER OutputPath
    Path for the log file. Defaults to script directory.
.EXAMPLE
    .\Invoke-UserOffboarding.ps1 -SamAccountName "jsmith"
    .\Invoke-UserOffboarding.ps1 -SamAccountName "jsmith" -DisabledOU "OU=Disabled,DC=hk,DC=lan"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$SamAccountName,

    [string]$DisabledOU = "OU=Disabled Users,DC=hk,DC=lan",

    [string]$OutputPath = $PSScriptRoot
)

$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$LogPath = Join-Path $OutputPath "Offboarding_${SamAccountName}_$Timestamp.log"
$Log = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param([string]$Message, [string]$Color = "White")
    $Entry = "$(Get-Date -Format 'HH:mm:ss') | $Message"
    $Log.Add($Entry)
    Write-Host $Entry -ForegroundColor $Color
}

Write-Log "=== Offboarding: $SamAccountName ===" "Cyan"

# --- Step 1: Get AD User ---
try {
    $User = Get-ADUser -Identity $SamAccountName -Properties DisplayName, UserPrincipalName, MemberOf, Manager, DistinguishedName -ErrorAction Stop
    Write-Log "Found AD account: $($User.DisplayName) ($($User.UserPrincipalName))" "Green"
}
catch {
    Write-Log "ERROR: Could not find AD user '$SamAccountName'. Exiting." "Red"
    $Log | Out-File $LogPath -Encoding UTF8
    exit 1
}

# --- Step 2: Disable AD Account ---
try {
    Disable-ADAccount -Identity $SamAccountName -ErrorAction Stop
    Write-Log "AD account disabled." "Green"
}
catch {
    Write-Log "ERROR disabling AD account: $($_.Exception.Message)" "Red"
}

# --- Step 3: Clear Manager Field ---
try {
    Set-ADUser -Identity $SamAccountName -Clear Manager -ErrorAction Stop
    Write-Log "Manager field cleared." "Green"
}
catch {
    Write-Log "WARNING: Could not clear manager field: $($_.Exception.Message)" "Yellow"
}

# --- Step 4: Remove Group Memberships ---
$Groups = $User.MemberOf
Write-Log "Removing $($Groups.Count) group membership(s)..." "Cyan"
foreach ($Group in $Groups) {
    try {
        Remove-ADGroupMember -Identity $Group -Members $SamAccountName -Confirm:$false -ErrorAction Stop
        Write-Log "  Removed from: $((($Group -split ",")[0]) -replace 'CN=','')" "Green"
    }
    catch {
        Write-Log "  WARNING: Could not remove from $Group - $($_.Exception.Message)" "Yellow"
    }
}

# --- Step 5: Move to Disabled OU ---
try {
    Move-ADObject -Identity $User.DistinguishedName -TargetPath $DisabledOU -ErrorAction Stop
    Write-Log "Account moved to: $DisabledOU" "Green"
}
catch {
    Write-Log "ERROR moving account to disabled OU: $($_.Exception.Message)" "Red"
}

# --- Step 6: Revoke M365 Licenses ---
Write-Log "Connecting to Microsoft Graph to revoke M365 licenses..." "Cyan"
try {
    Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome -ErrorAction Stop

    $MgUser = Get-MgUser -Filter "userPrincipalName eq '$($User.UserPrincipalName)'" -Property Id, AssignedLicenses -ErrorAction Stop

    if ($MgUser.AssignedLicenses.Count -gt 0) {
        $RemoveLicenses = $MgUser.AssignedLicenses | Select-Object -ExpandProperty SkuId
        Set-MgUserLicense -UserId $MgUser.Id -AddLicenses @() -RemoveLicenses $RemoveLicenses -ErrorAction Stop
        Write-Log "Revoked $($RemoveLicenses.Count) M365 license(s)." "Green"
    }
    else {
        Write-Log "No M365 licenses assigned." "Yellow"
    }

    Disconnect-MgGraph | Out-Null
}
catch {
    Write-Log "ERROR revoking M365 licenses: $($_.Exception.Message)" "Red"
}

# --- Save Log ---
$Log | Out-File $LogPath -Encoding UTF8
Write-Log "`nOffboarding complete. Log saved to: $LogPath" "Yellow"
