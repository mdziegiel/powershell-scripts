 PowerShell Scripts

PowerShell scripts for endpoint management, automation, and system configuration used across the Hans Kissle environment.

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
| `hkadmin.ps1` | HK Admin account setup |
| `GIDAdmin.ps1` | GID Admin account setup |
| `RemediateLocalUser.ps1` | Remediates local user account issues |

---

### Endpoint-Configuration
Scripts for configuring and standardizing Windows endpoints.

| Script | Description |
|--------|-------------|
| `RenamePC.ps1` | Renames the PC and updates the This PC icon |
| `RenamePC_With_Usersfiles.ps1` | Renames PC and enables User Files desktop icon |
| `AlignWindows11taskbarleft.ps1` | Aligns Windows 11 taskbar to the left |
| `Set-DNSSuffixSearchList.ps1` | Configures DNS suffix search list for hk.lan |
| `Set-PowerConfiguration.ps1` | Applies power settings profile (Workstation/Laptop/Kiosk) |
| `Set-WindowsFeatures.ps1` | Enables or disables Windows optional features in bulk |
| `Enable Auto Time Zone updater.ps1` | Enables automatic time zone detection |
| `Remove 260 Character Path Limit.ps1` | Removes the 260 character path length limit |
| `Autopilot_Turnoff_firewall_9.25.23.ps1` | Disables firewall during Autopilot provisioning |
| `DriveMapping.ps1` | Maps network drives |

---

### Software-Management
Scripts for installing, configuring, and removing software.

| Script | Description |
|--------|-------------|
| `FireFoxRemoval.ps1` | Removes Mozilla Firefox |
| `BostonCanonRemoval.ps1` | Removes Boston Canon software |
| `BostonuniFlowRemoval.ps1` | Removes Boston uniFlow software |
| `ScreenConnectRemoval.ps1` | Removes ConnectWise ScreenConnect agent |
| `SupportAssistCleanup.ps1` | Removes Dell SupportAssist |
| `DCU Script.ps1` | Dell Command Update script |
| `Office Shortcuts.ps1` | Creates Office application shortcuts |
| `Launch Ondrive.ps1` | Launches and configures OneDrive |

---

### Security
Scripts for security auditing, access control, and compliance.

| Script | Description |
|--------|-------------|
| `Get-LocalAdminAudit.ps1` | Audits local Administrators group across all domain computers |
| `Get-DisabledAccountAudit.ps1` | Exports all disabled AD accounts with last logon and group memberships |
| `Get-UserLastLogonAudit.ps1` | Exports last logon times for all AD users with inactive flags |
| `Get-BitLockerAudit.ps1` | Checks BitLocker encryption status across all domain machines |
| `Bitlocker_Sync.ps1` | Syncs BitLocker recovery keys to Active Directory |
| `DisabledAdmin.ps1` | Disables the local Administrator account |
| `Remove all local accounts.ps1` | Removes all non-default local user accounts |
| `RemoveRegistryKeys.ps1` | Removes specified registry keys |
| `SSTP VPN V2.ps1` | Configures SSTP VPN connection |

---

### Reporting
Scripts for generating reports on Intune, Active Directory, and Microsoft 365.

| Script | Description |
|--------|-------------|
| `Get-InactiveIntuneDevices.ps1` | Reports Intune devices not synced in X days |
| `Get-M365LicensedUsers.ps1` | Exports all M365 licensed users with license details and last sign-in |
| `Get-ADGroupMembershipReport.ps1` | Exports AD group memberships by group or by user |

---

### Branding
Scripts and assets for corporate branding and desktop configuration.

| Script | Description |
|--------|-------------|
| `Set-HK-WallpaperLockscreen.ps1` | Sets Hans Kissle wallpaper and lock screen |
| `CreateShortcut.ps1` | Creates desktop or Start Menu shortcuts |
| `LogoPic/` | Logo image assets |
| `Fonts/` | Corporate font files |

---

### Miscellaneous
Scripts that don't fit a specific category.

---

## Requirements

- PowerShell 5.1 or later
- Scripts that query Active Directory require the **RSAT ActiveDirectory** module
- Scripts that interact with Microsoft 365 require the **Microsoft.Graph** PowerShell module
- Most scripts require **local Administrator** or **Domain Admin** privileges

## Notes

- All scripts target the `hk.lan` Active Directory domain
- CSV/log output defaults to the script's directory unless `-OutputPath` is specified
- Always test in a non-production environment before running against all machines
