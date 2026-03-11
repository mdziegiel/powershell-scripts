#Requires -Version 5.1
<#
.SYNOPSIS
    Resets an Active Directory user password and forces change at next logon.
.DESCRIPTION
    Resets a user's AD password, unlocks the account if locked, and sets the
    must-change-password flag. Optionally generates a random password if none
    is provided. Logs all actions taken.
.PARAMETER SamAccountName
    The SamAccountName of the user to reset.
.PARAMETER NewPassword
    The new password to set. If omitted, a random 12-character password is generated.
.PARAMETER NoForceChange
    Switch to skip forcing a password change at next logon.
.PARAMETER OutputPath
    Path for the log file. Defaults to script directory.
.EXAMPLE
    .\Reset-UserPassword.ps1 -SamAccountName "jsmith"
    .\Reset-UserPassword.ps1 -SamAccountName "jsmith" -NewPassword "TempP@ss123!"
    .\Reset-UserPassword.ps1 -SamAccountName "jsmith" -NoForceChange
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SamAccountName,

    [string]$NewPassword,

    [switch]$NoForceChange,

    [string]$OutputPath = $PSScriptRoot
)

$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$LogPath = Join-Path $OutputPath "PasswordReset_${SamAccountName}_$Timestamp.log"
$Log = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param([string]$Message, [string]$Color = "White")
    $Entry = "$(Get-Date -Format 'HH:mm:ss') | $Message"
    $Log.Add($Entry)
    Write-Host $Entry -ForegroundColor $Color
}

function New-RandomPassword {
    $Upper   = "ABCDEFGHJKLMNPQRSTUVWXYZ".ToCharArray()
    $Lower   = "abcdefghjkmnpqrstuvwxyz".ToCharArray()
    $Numbers = "23456789".ToCharArray()
    $Special = "!@#$%^&*".ToCharArray()

    $Password = @(
        ($Upper   | Get-Random -Count 3)
        ($Lower   | Get-Random -Count 3)
        ($Numbers | Get-Random -Count 3)
        ($Special | Get-Random -Count 3)
    ) | ForEach-Object { $_ } | Sort-Object { Get-Random }

    return -join $Password
}

Write-Log "=== Password Reset: $SamAccountName ===" "Cyan"

# --- Get AD User ---
try {
    $User = Get-ADUser -Identity $SamAccountName -Properties DisplayName, UserPrincipalName, LockedOut, Enabled -ErrorAction Stop
    Write-Log "Found: $($User.DisplayName) ($($User.UserPrincipalName))" "Green"
}
catch {
    Write-Log "ERROR: Could not find AD user '$SamAccountName'. Exiting." "Red"
    $Log | Out-File $LogPath -Encoding UTF8
    exit 1
}

# --- Check if account is disabled ---
if (-not $User.Enabled) {
    Write-Log "WARNING: Account is currently disabled." "Yellow"
}

# --- Generate password if not provided ---
if (-not $NewPassword) {
    $NewPassword = New-RandomPassword
    Write-Log "Generated random password." "Cyan"
}

# --- Reset Password ---
try {
    Set-ADAccountPassword -Identity $SamAccountName -Reset -NewPassword (ConvertTo-SecureString $NewPassword -AsPlainText -Force) -ErrorAction Stop
    Write-Log "Password reset successfully." "Green"
}
catch {
    Write-Log "ERROR resetting password: $($_.Exception.Message)" "Red"
    $Log | Out-File $LogPath -Encoding UTF8
    exit 1
}

# --- Force change at next logon ---
if (-not $NoForceChange) {
    try {
        Set-ADUser -Identity $SamAccountName -ChangePasswordAtLogon $true -ErrorAction Stop
        Write-Log "User will be required to change password at next logon." "Green"
    }
    catch {
        Write-Log "WARNING: Could not set must-change-password flag: $($_.Exception.Message)" "Yellow"
    }
}

# --- Unlock account if locked ---
if ($User.LockedOut) {
    try {
        Unlock-ADAccount -Identity $SamAccountName -ErrorAction Stop
        Write-Log "Account was locked — now unlocked." "Green"
    }
    catch {
        Write-Log "WARNING: Could not unlock account: $($_.Exception.Message)" "Yellow"
    }
}

# --- Display new password ---
Write-Host "`n----------------------------------------" -ForegroundColor Cyan
Write-Host " User:     $($User.DisplayName)" -ForegroundColor White
Write-Host " Password: $NewPassword" -ForegroundColor Yellow
Write-Host "----------------------------------------`n" -ForegroundColor Cyan
Write-Log "Temporary password: $NewPassword"

# --- Save Log ---
$Log | Out-File $LogPath -Encoding UTF8
Write-Log "Complete. Log saved to: $LogPath" "Yellow"
