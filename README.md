# PowerShell Scripts

PowerShell scripts for endpoint management, automation, and system configuration.

---

## Repository Structure

### User-Provisioning
Scripts for user account management and lifecycle.

| Script | Description |
|--------|-------------|
| `NewUserProvisioning.ps1` | Creates new AD users and assigns M365 licenses |
| `Invoke-UserOffboarding.ps1` | Disables account, removes groups, revokes M365 licenses, moves to disabled OU |
| `Reset-UserPassword.ps1` | Resets password, unlocks account, forces change at next logon |
| `HelpDesk.ps1` | HelpDesk account provisioning |
| `New-LocalAdminAccount.ps1` | Local admin account setup |
| `RemediateLocalUser.ps1` | Remediates local user account issues |

---

### Endpoint-Configuration
Scripts for configuring and standardizing Windows endpoints.

| Script | Description |
|--------|-------------|
| `RenamePC.ps1` | Renames the PC and updates the This PC icon |
| `RenamePC_With_Usersfiles.ps1` | Renames PC and enables User Files desktop icon |
| `AlignWindows11taskbarleft.ps1` | Aligns Windows 11 taskbar to the left |
| `Set-DNSSuffixSearchList.ps1` | Configures DNS suffix search list |
| `Set-PowerConfiguration.ps1` | Applies power settings profile |
| `Set-WindowsFeatures.ps1` | Enables or disables Windows optional features |
| `Enable Auto Time Zone updater.ps1` | Enables automatic time zone detection |
| `Remove 260 Character Path Limit.ps1` | Removes the 260 character path length limit |
| `DisableWindowsFirewallProfiles.ps1` | Disables firewall profiles (if required) |
| `DriveMapping.ps1` | Maps network drives |

---

### Software-Management
Scripts for installing, configuring, and removing software.

| Script | Description |
|--------|-------------|
| `FireFoxRemoval.ps1` | Removes Mozilla Firefox |
| `CanonRemoval.ps1` | Removes Canon software |
| `ScreenConnectRemoval.ps1` | Removes ScreenConnect agent |
| `SupportAssistCleanup.ps1` | Removes Dell SupportAssist |
| `DCU Script.ps1` | Dell Command Update configuration |
| `Office Shortcuts.ps1` | Creates Office application shortcuts |
| `Launch OneDrive.ps1` | Launches and configures OneDrive |

---

### Security
Scripts for security auditing, access control, and compliance.

| Script | Description |
|--------|-------------|
| `Get-LocalAdminAudit.ps1` | Audits local Administrators group |
| `Get-DisabledAccountAudit.ps1` | Exports disabled AD accounts |
| `Get-UserLastLogonAudit.ps1` | Exports last logon times |
| `Get-BitLockerAudit.ps1` | Checks BitLocker status |
| `Bitlocker_Sync.ps1` | Syncs BitLocker recovery keys |
| `DisabledAdmin.ps1` | Disables local Administrator |
| `Remove all local accounts.ps1` | Removes non-default local users |
| `RemoveRegistryKeys.ps1` | Removes registry keys |
| `SSTP VPN V2.ps1` | Configures VPN connection |

---

### Reporting
Scripts for reporting across systems.

| Script | Description |
|--------|-------------|
| `Get-InactiveIntuneDevices.ps1` | Reports inactive Intune devices |
| `Get-M365LicensedUsers.ps1` | Exports licensed users |
| `Get-ADGroupMembershipReport.ps1` | Exports group memberships |

---

### Branding
Scripts for branding and desktop configuration.

| Script | Description |
|--------|-------------|
| `Set-WallpaperLockscreen.ps1` | Sets wallpaper and lock screen |
| `CreateShortcut.ps1` | Creates shortcuts |
| `LogoPic/` | Logo assets |
| `Fonts/` | Font files |

---

## Requirements

- PowerShell 5.1 or later
- RSAT ActiveDirectory module (for AD scripts)
- Microsoft Graph PowerShell module (for M365 scripts)
- Administrator privileges for most scripts

---

## ⚠️ Notes

- This repository has been sanitized for public use
- No production credentials or sensitive data are included
- Replace placeholder values before use

---

## 🚀 Usage

Run scripts with appropriate permissions.

Example:
```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\ScriptName.ps1
```

---

## 📌 Disclaimer

These scripts are provided **as-is** without any warranties, express or implied.  
Use at your own risk.

The author assumes no responsibility for:
- System damage
- Data loss
- Security misconfigurations
- Production outages

All scripts should be reviewed and tested in a non-production environment before deployment.

All scripts should be reviewed and tested in a non-production environment before deployment.
