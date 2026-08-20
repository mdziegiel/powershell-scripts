# PowerShell Scripts

PowerShell scripts for endpoint management, automation, reporting, and system configuration.

Current script count: 154.

---

## Repository Structure

### User-Provisioning
Scripts for user account management and lifecycle.

| Script | Description | Tags |
|--------|-------------|------|
| [Add-ContactToAD.ps1](User-Provisioning/Add-ContactToAD.ps1) | Adds a contact to Active Directory | AD |
| [Create-ContactsOnAD.ps1](User-Provisioning/Create-ContactsOnAD.ps1) | Creates contacts in Active Directory | AD |
| [HelpDesk.ps1](User-Provisioning/HelpDesk.ps1) | Downloads a custom .ico file, creates a local folder to store it, and adds a HelpDesk shortcut to the Public Desktop using the icon. | Endpoint |
| [Invoke-UserOffboarding.ps1](User-Provisioning/Invoke-UserOffboarding.ps1) | Disables the AD account, clears the manager field, removes group memberships except Domain Users, moves the account to the disabled users OU, revokes Microsoft 365 licenses via Microsoft Graph, and logs all actions taken. | AD, Graph, License, M365 |
| [Local_Admin_Script_Name.ps1](User-Provisioning/Local_Admin_Script_Name.ps1) | Creates a local administrator account, sets the password to never expire, and adds the account to the local Administrators group | Endpoint, LocalAdmin, Security |
| [New-LocalAdminAccount.ps1](User-Provisioning/New-LocalAdminAccount.ps1) | Creates new local admin account and adds to Administrators group | Endpoint, LocalAdmin, Security |
| [NewUserProvisioningGUI.ps1](User-Provisioning/NewUserProvisioningGUI.ps1) | WPF form replaces all Read-Host prompts | Endpoint, AD |
| [Reset-UserPassword.ps1](User-Provisioning/Reset-UserPassword.ps1) | Resets a user's AD password, unlocks the account if locked, and forces a password change at next logon | AD, Password, Security |
| [Add-UsersToGroup.ps1](User-Provisioning/Add-UsersToGroup.ps1) | Bulk add users from a CSV file to an Active Directory group. | AD, CSV |
| [CleanupDisabledUsers.ps1](User-Provisioning/CleanupDisabledUsers.ps1) | Stage and delete disabled Active Directory user accounts. | AD |
| [Get-UserFromOU.ps1](User-Provisioning/Get-UserFromOU.ps1) | Export Active Directory users from a specific OU to CSV. | AD, CSV |
| [CleanupInActiveComputers.ps1](User-Provisioning/CleanupInActiveComputers.ps1) | Stage and disable inactive Active Directory computer accounts. | AD, Device |
| [Get-ADComputers.ps1](User-Provisioning/Get-ADComputers.ps1) | Export Active Directory computer objects and properties to CSV. | AD, CSV, Device |
| [Get-ADGroupMembers.ps1](User-Provisioning/Get-ADGroupMembers.ps1) | Export members of one or more Active Directory groups. | AD, CSV |
| [Copy-ADUserGroupMembership.ps1](User-Provisioning/Copy-ADUserGroupMembership.ps1) | Copy group memberships from one Active Directory user to another. | AD |

---

### Endpoint-Configuration
Scripts for configuring and standardizing Windows endpoints.

| Script | Description | Tags |
|--------|-------------|------|
| [AlignWindows11taskbarleft.ps1](Endpoint-Configuration/AlignWindows11taskbarleft.ps1) | Configures the registry value 'TaskbarAl' under HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced to 0, which aligns the taskbar to the left for the current user | Endpoint, Windows |
| [Disable Windows Firewall Profiles.ps1](Endpoint-Configuration/Disable%20Windows%20Firewall%20Profiles.ps1) | Disables Domain, Private, and Public firewall profiles | Endpoint, Firewall, Security |
| [DriveMapping.ps1](Endpoint-Configuration/DriveMapping.ps1) | Performs network drive mappings with PowerShell | Endpoint, Network |
| [Enable Auto Time Zone updater.ps1](Endpoint-Configuration/Enable%20Auto%20Time%20Zone%20updater.ps1) | Sets required Windows services (lfsvc, tzautoupdate, w32time) to Automatic and starts them if needed | Endpoint |
| [poweroff.ps1](Endpoint-Configuration/poweroff.ps1) | This script halts the local computer immediately (needs admin rights). | Device, Endpoint |
| [reboot.ps1](Endpoint-Configuration/reboot.ps1) | This PowerShell script reboots the local computer immediately (needs admin rights). | Device, Endpoint |
| [Remove 260 Character Path Limit.ps1](Endpoint-Configuration/Remove%20260%20Character%20Path%20Limit.ps1) | Remove 260 Character Path Limit | Endpoint, Software |
| [RenamePC.ps1](Endpoint-Configuration/RenamePC.ps1) | Takes ownership of the registry key for "This PC" and updates both the default value and LocalizedName to match the current computer name using $env:COMPUTERNAME | Device, Endpoint |
| [RenamePC_With_Usersfiles.ps1](Endpoint-Configuration/RenamePC_With_Usersfiles.ps1) | Modifies the CLSID registry key for "This PC" by taking ownership and updating the display name and LocalizedName values to reflect the current hostname ($env:COMPUTERNAME) | Device, Endpoint |
| [Set-DNSSuffixSearchList.ps1](Endpoint-Configuration/Set-DNSSuffixSearchList.ps1) | Sets the DNS suffix search list in the registry and on all active network adapters | Endpoint, DNS, Network |
| [Set-PowerConfiguration.ps1](Endpoint-Configuration/Set-PowerConfiguration.ps1) | Applies standardized power settings using powercfg, including sleep, hibernate, display timeout, and fast startup | Endpoint |
| [Set-WindowsFeatures.ps1](Endpoint-Configuration/Set-WindowsFeatures.ps1) | Manages Windows optional features using DISM | Endpoint, Software, Windows |
| [wake-up-host.ps1](Endpoint-Configuration/wake-up-host.ps1) | This PowerShell script sends a magic UDP packet to a computer to wake him up (requires the target computer to have Wake-on-LAN activated). | Device, Endpoint, Network, WOL |
| [new-reboot-task.ps1](Endpoint-Configuration/new-reboot-task.ps1) | Create a timed reboot scheduled task. | Automation, Device, Endpoint, ScheduledTask |
| [Scheduled-Tasks-Inventory.ps1](Endpoint-Configuration/Scheduled-Tasks-Inventory.ps1) | Inventory scheduled tasks on Windows servers. | Automation, CSV, Endpoint, Reporting |
| [CleanUp-Windows.ps1](Endpoint-Configuration/CleanUp-Windows.ps1) | Clean up Windows 10 from bloatware. | Endpoint, Software, Windows |
| [Disable-Services.ps1](Endpoint-Configuration/Disable-Services.ps1) | Disable selected Windows services. | Endpoint, Security, Windows |
| [Remove-Copilot.ps1](Endpoint-Configuration/Remove-Copilot.ps1) | Remove Copilot from Windows 10. | Endpoint, Software, Windows |

---

### Software-Management
Scripts for installing, configuring, and removing software.

| Script | Description | Tags |
|--------|-------------|------|
| [Analyze-WindowsStartup.ps1](Software-Management/Analyze-WindowsStartup.ps1) | Audit startup apps, services, and scheduled tasks | Endpoint, Software, Automation |
| [Bitlocker_Sync.ps1](Software-Management/Bitlocker_Sync.ps1) | Retrieves the BitLocker recovery key protector ID for the system drive and backs it up to Azure AD (Entra ID) | AD, BitLocker, Entra, Security |
| [CanonRemoval.ps1](Software-Management/CanonRemoval.ps1) | Identifies and removes Canon printer objects | Printer, Software |
| [check-drive-space.ps1](Software-Management/check-drive-space.ps1) | This PowerShell script checks the given drive for free space left (10 GB by default). | Endpoint, Disk |
| [check-health.ps1](Software-Management/check-health.ps1) | This PowerShell script queries the system health of the local computer (hardware, software, and network) and prints it. | Endpoint, Health |
| [DCU Script.ps1](Software-Management/DCU%20Script.ps1) | Configures Dell Command Update via registry and CLI, including update scheduling and system behavior | Software |
| [FireFoxRemoval.ps1](Software-Management/FireFoxRemoval.ps1) | Removal of Firefox browser | Browser, Software |
| [Install-7Zip.ps1](Software-Management/Install-7Zip.ps1) | Winget wrapper to install 7-Zip | Software |
| [Install-AdobeReader.ps1](Software-Management/Install-AdobeReader.ps1) | Winget wrapper to install Adobe Acrobat Reader | AD, Software |
| [install-basic-apps.ps1](Software-Management/install-basic-apps.ps1) | This PowerShell script installs basic Windows apps such as browser, e-mail client, etc | Browser, Software |
| [install-chrome.ps1](Software-Management/install-chrome.ps1) | This PowerShell script installs the Google Chrome browser from WinGet. | Browser, Software |
| [Install-Chrome.ps1](Software-Management/Install-Chrome.ps1) | Winget wrapper to install Google Chrome | Browser, Software |
| [install-firefox.ps1](Software-Management/install-firefox.ps1) | This PowerShell script installs the Mozilla Firefox browser from Microsoft Store. | Browser, Software |
| [Install-Firefox.ps1](Software-Management/Install-Firefox.ps1) | Winget wrapper to install Mozilla Firefox | Browser, Software |
| [install-git-for-windows.ps1](Software-Management/install-git-for-windows.ps1) | This PowerShell script installs Git for Windows. | Software |
| [install-github-cli.ps1](Software-Management/install-github-cli.ps1) | This PowerShell script installs the GitHub command-line interface (CLI). | Software |
| [install-ssh-client.ps1](Software-Management/install-ssh-client.ps1) | This PowerShell script installs a SSH client (needs admin rights). | Software |
| [install-ssh-server.ps1](Software-Management/install-ssh-server.ps1) | This PowerShell script installs a SSH server (needs admin rights). | Software |
| [install-wsl.ps1](Software-Management/install-wsl.ps1) | This PowerShell script installs Windows Subsystem for Linux | Software |
| [Invoke-RemoveUnwantedApps.ps1](Software-Management/Invoke-RemoveUnwantedApps.ps1) | Remove a few apps we don't like to have installed any longer, we use WinGet to do so | Software |
| [Launch Onderive.ps1](Software-Management/Launch%20Onderive.ps1) | Downloads the OneDrive installer, performs a silent all-users installation, then creates and runs a temporary scheduled task to launch OneDrive and removes the task after execution | Automation, ScheduledTask, Software |
| [Office Shortcuts.ps1](Software-Management/Office%20Shortcuts.ps1) | Creates public desktop shortcuts for common Microsoft Office applications including Word, Excel, Outlook, and PowerPoint using COM objects | Branding, Exchange, Software |
| [Optimize-WindowsDisk.ps1](Software-Management/Optimize-WindowsDisk.ps1) | Audits Windows storage and optionally runs supported disk optimization and cleanup tasks. | Disk, Reporting, Software |
| [Remove-bloatware.ps1](Software-Management/Remove-bloatware.ps1) | Removes pre-installed Windows 10 apps and common consumer bloatware so fresh endpoints start from a cleaner baseline. | Software |
| [Remove-Widget.ps1](Software-Management/Remove-Widget.ps1) | Removes the Windows 11 Widgets / WebExperience package for existing users and removes the provisioned package for new users. | Software |
| [ScreenConnectRemoval.ps1](Software-Management/ScreenConnectRemoval.ps1) | Removal of ScreenConnect | Software |
| [SupportAssistCleanup.ps1](Software-Management/SupportAssistCleanup.ps1) | Removes Dell SupportAssist for PCs if installed, including associated components and residual files | Software |

---

### Security
Scripts for security auditing, access control, and compliance.

| Script | Description | Tags |
|--------|-------------|------|
| [Bitlocker_Sync.ps1](Security/Bitlocker_Sync.ps1) | Retrieves the BitLocker recovery key protector ID for the system drive and backs it up to Azure AD (Entra ID) | AD, BitLocker, Entra, Security |
| [DisabledAdmin.ps1](Security/DisabledAdmin.ps1) | Checks if the local Administrator account exists and is enabled | LocalAdmin, Security |
| [Get-ADUserLockout.ps1](Security/Get-ADUserLockout.ps1) | Tracking down account lockout sources with PowerShell | AD, Password, Security |
| [Get-BitLockerAudit.ps1](Security/Get-BitLockerAudit.ps1) | Queries Active Directory computers remotely to collect BitLocker encryption status for all drives | AD, BitLocker, Device, Reporting |
| [Get-DisabledAccountAudit.ps1](Security/Get-DisabledAccountAudit.ps1) | Exports disabled user accounts from Active Directory, including last logon time, OU location, and group memberships | AD, Reporting, Security |
| [Get-EntraTenantReview.ps1](Security/Get-EntraTenantReview.ps1) | Exports Entra user, admin, policy, and application reports. | Entra, Reporting, Compliance |
| [Get-LocalAdminAudit.ps1](Security/Get-LocalAdminAudit.ps1) | Queries Active Directory for domain-joined computers and remotely enumerates the local Administrators group on each system | AD, Device, LocalAdmin, Reporting |
| [Get-UserLastLogonAudit.ps1](Security/Get-UserLastLogonAudit.ps1) | Queries Active Directory for user accounts and exports last logon date, account status, and key attributes | AD, Reporting, Security |
| [Invoke-BackupBitlockerRecoveryKey.ps1](Security/Invoke-BackupBitlockerRecoveryKey.ps1) | Backup the BitLocker Recovery Information to the Azure Active Directory If the Boot Drive is not encrypted, the Script will try to enable the quick protection | AD, BitLocker, Security |
| [Invoke-GPO-Win11DiscoveryWorkbook.ps1](Security/Invoke-GPO-Win11DiscoveryWorkbook.ps1) | Invoke-GPO-Win11DiscoveryWorkbook | Compliance, GPO, Endpoint |
| [Invoke-Win11GPOMigrationAssessmentWorkbook.ps1](Security/Invoke-Win11GPOMigrationAssessmentWorkbook.ps1) | Creates the Windows 11 GPO Migration Assessment workbook | Compliance, GPO |
| [Logon-Audit.ps1](Security/Logon-Audit.ps1) | Audits Windows logon and logoff events from Security logs and can send the results to Microsoft Teams. | Reporting, Security, Teams |
| [M365EnterpriseCommandCenter_v3_2.ps1](Security/M365EnterpriseCommandCenter_v3_2.ps1) | Technician console using app-only certificate auth and direct Microsoft Graph REST calls. | Graph, M365, Security |
| [M365EnterpriseOperationsCenter.ps1](Security/M365EnterpriseOperationsCenter.ps1) | Read-only Microsoft 365 operations dashboard for monitoring tenant health, security, licensing, mail flow, and device compliance. | Compliance, Exchange, Health, M365 |
| [Remove all local accounts.ps1](Security/Remove%20all%20local%20accounts.ps1) | Enumerates all local user accounts and deletes any account not included in an approved allow list | Security, Software |
| [Remove-AzureADInactiveDevices.ps1](Security/Remove-AzureADInactiveDevices.ps1) | Remove all inactive AzureAD Devices Inactivity Threshold can be given | Entra, Security |
| [RemoveRegistryKeys.ps1](Security/RemoveRegistryKeys.ps1) | Removal of Registry Keys WindowsUpdate WindowsUpdate\AU | Security, Software |
| [Setup-M365EOC-AppOnly.ps1](Security/Setup-M365EOC-AppOnly.ps1) | One-time setup for Microsoft 365 Enterprise Operations Center v3.0 | M365, Security |
| [SSTP VPN V2.ps1](Security/SSTP%20VPN%20V2.ps1) | Creates or updates a Windows VPN connection using the specified connection settings, applies DNS suffix configuration, enables split tunneling, and adds static routes for internal network access | Network, Security, VPN |
| [windefender.ps1](Security/windefender.ps1) | This script can enable / disable and show Windows defender real time monitoring! | Security, Windows |

---

### Reporting
Scripts for reporting across systems.

| Script | Description | Tags |
|--------|-------------|------|
| [Create-DL-Group-Report.ps1](Reporting/Create-DL-Group-Report.ps1) | Connects to Exchange Online and exports a distribution list membership report to CSV. | CSV, Exchange, Reporting |
| [Export-ADGroupMemberToCSV.ps1](Reporting/Export-ADGroupMemberToCSV.ps1) | Lists users in an AD group and exports name, object class, and SamAccountName to CSV. | AD, CSV, Reporting |
| [Export-AllMailboxesReport.ps1](Reporting/Export-AllMailboxesReport.ps1) | Retrieves mailbox details from Exchange Online, including aliases, mailbox size, full access permissions, send-as permissions, and forwarding configuration, then exports the results to CSV for reporting and audit… | CSV, Exchange, Reporting |
| [Export-GPOSettings.ps1](Reporting/Export-GPOSettings.ps1) | Exports Group Policy settings | Compliance, GPO, Reporting |
| [Get-ADGroupMembershipReport.ps1](Reporting/Get-ADGroupMembershipReport.ps1) | Generates a CSV report of Active Directory group memberships for all users or a specific group | AD, CSV, Reporting |
| [Get-AuditGuestTeams.ps1](Reporting/Get-AuditGuestTeams.ps1) | This Script function will create a report that will help IT PROs to Monitor and Audit Guest users | Reporting, Teams |
| [Get-ComputersToRebootAfterUpdate.ps1](Reporting/Get-ComputersToRebootAfterUpdate.ps1) | Lists AD computers with a pending reboot after Windows Update KB installation and exports them to CSV. | AD, CSV, Device, Reporting |
| [Get-enADForestInformation.ps1](Reporting/Get-enADForestInformation.ps1) | Retrieve information about an Active Directory Forest | AD, Reporting |
| [Get-enMailboxPermissionReport.ps1](Reporting/Get-enMailboxPermissionReport.ps1) | Get a detailed mailbox permission report and exports this report to a given CSV file | CSV, Exchange, Reporting |
| [Get-enMailboxSendAsReport.ps1](Reporting/Get-enMailboxSendAsReport.ps1) | Get a detailed mailbox Send permission report and exports this report to a given CSV file | CSV, Exchange, Reporting |
| [Get-FullTeamsReport.ps1](Reporting/Get-FullTeamsReport.ps1) | Connects to Microsoft Teams and exports a full report of teams, channels, and members. | Reporting, Teams |
| [Get-GraphGuestReportsInEntra.ps1](Reporting/Get-GraphGuestReportsInEntra.ps1) | Report the Sponsors of Entra ID Guest Accounts | Entra, Graph, Reporting |
| [Get-InactiveIntuneDevices.ps1](Reporting/Get-InactiveIntuneDevices.ps1) | Connects to Microsoft Graph and identifies devices that have not checked in within a specified number of days | Device, Graph, Intune, Reporting |
| [Get-M365LicensedUsers.ps1](Reporting/Get-M365LicensedUsers.ps1) | Connects to Microsoft Graph and exports users with assigned Microsoft 365 licenses, including license details, account status, and sign-in information | Graph, License, M365, Reporting |
| [Get-MailboxFolderPermissions.ps1](Reporting/Get-MailboxFolderPermissions.ps1) | Enumerates mailbox folders in a mailbox and reports users with explicit non-default, non-anonymous permissions on each folder. | Exchange, Reporting |
| [Get-MailboxUsage.ps1](Reporting/Get-MailboxUsage.ps1) | Reports mailbox usage | Exchange, Reporting |
| [Get-MFAUserReport.ps1](Reporting/Get-MFAUserReport.ps1) | Get a Azure AD MFA User report, the function can export the report as CSV | AD, CSV, Entra, Reporting |
| [Get-SharedMailboxPermissions.ps1](Reporting/Get-SharedMailboxPermissions.ps1) | Connects to Exchange Online and enumerates shared mailbox permissions. | Exchange, Reporting |
| [Get-SPOSiteStorageUsage.ps1](Reporting/Get-SPOSiteStorageUsage.ps1) | SharePoint Online storage usage reporting | Disk, Reporting, SharePoint |
| [Get-TeamsChannelUsersReport.ps1](Reporting/Get-TeamsChannelUsersReport.ps1) | Connects to Microsoft Graph and Microsoft Teams to report users, roles, and shared-channel membership across teams. | Graph, Reporting, Teams |
| [Invoke-GetAzureADAuditSignInLogs.ps1](Reporting/Invoke-GetAzureADAuditSignInLogs.ps1) | Get the AzureAD Audit Sign-In Logs and create several CSV files | AD, CSV, Entra, Reporting |
| [List-AllMembersOfADistributionList.ps1](Reporting/List-AllMembersOfADistributionList.ps1) | Connects to Exchange Online, expands a distribution list, and prints each member’s display name and primary SMTP address. | Exchange, Reporting |
| [ListADObjectsOnSelectedOU.ps1](Reporting/ListADObjectsOnSelectedOU.ps1) | Lists AD objects from a graphically selected OU and exports common name and description to CSV. | AD, CSV, Graph, Reporting |

---

### Branding
Scripts for branding, desktop layout, and sign-in presentation.

| Script | Description | Tags |
|--------|-------------|------|
| [Copy-Startmenu.ps1](Branding/Copy-Startmenu.ps1) | Copies a Windows 11 Start menu template to all existing user profiles and the default profile | Branding |
| [Create_Common_Desktop_Shortcuts.ps1](Branding/Create_Common_Desktop_Shortcuts.ps1) | Creates a standard set of desktop shortcuts for common applications so the desktop layout stays consistent. | Branding |
| [Create_Common_StartMenu_Shortcuts.ps1](Branding/Create_Common_StartMenu_Shortcuts.ps1) | Creates a standard set of Start Menu shortcuts for common applications so users get a consistent launcher layout. | Branding |
| [CreateShortcut.ps1](Branding/CreateShortcut.ps1) | Creates Measure.lnk on the Public Desktop and generates a marker file used for Intune detection | Branding, Intune |
| [Download_Themes.ps1](Branding/Download_Themes.ps1) | Downloads Microsoft theme packs to a local folder so wallpaper and lock screen branding assets can be distributed consistently. | Branding, Wallpaper |
| [Invoke-InstallCustomBGInfo.ps1](Branding/Invoke-InstallCustomBGInfo.ps1) | Download an install the latest BGInfo on an internal server | Branding, Software |
| [MOTD.ps1](Branding/MOTD.ps1) | Fetches a random quote and returns it as a message of the day for display at sign-in or in a branded console prompt. | Branding |
| [new-shortcut.ps1](Branding/new-shortcut.ps1) | This PowerShell script creates a new shortcut file. | Branding, Endpoint |
| [set-wallpaper.ps1](Branding/set-wallpaper.ps1) | This PowerShell script sets the given image file as desktop wallpaper (.JPG or .PNG supported) | Branding, Wallpaper |
| [Set-WallpaperLockscreen.ps1](Branding/Set-WallpaperLockscreen.ps1) | Copies wallpaper.jpg to a local path, applies it as the device-level lock screen, and schedules the desktop background to be configured at the next user logon | Branding, Device, Wallpaper |
| [Start-Apps.ps1](Branding/Start-Apps.ps1) | Launches a curated set of Start Menu shortcuts and shortcuts-based app entries for quick access to standard tools. | Branding |

---

### Miscellaneous
Scripts that do not fit the other categories cleanly.

| Script | Description | Tags |
|--------|-------------|------|
| [Create-HyperV_VM.ps1](Miscellaneous/Create-HyperV_VM.ps1) | Creates Hyper-V virtual machines from a CSV definition. | CSV, HyperV |
| [Delete-HyperV_VM.ps1](Miscellaneous/Delete-HyperV_VM.ps1) | Deletes Hyper-V virtual machines listed in a CSV file. | CSV, HyperV |
| [Get-SharePermissionAudit.ps1](Miscellaneous/Get-SharePermissionAudit.ps1) | Recursively enumerates folders under a specified UNC path and reports identity, account type, access rights, allow/deny status, inheritance, and whether permissions are explicitly set or inherited | CSV, Network, Reporting |
| [HealthCheck.ps1](Miscellaneous/HealthCheck.ps1) | Checks Windows server health across hardware, software, and network basics. | Health, Network, Software |
| [GitHub-Mass-Query.ps1](Miscellaneous/GitHub-Mass-Query.ps1) | Massively search GitHub for keywords. | Automation, CSV, Reporting |

---

### Cloud-Administration
Scripts for Microsoft 365, Azure, Exchange, Teams, and SharePoint administration.

| Script | Description | Tags |
|--------|-------------|------|
| [Add-DelegationRights.ps1](Cloud-Administration/Add-DelegationRights.ps1) | Adds mailbox delegation rights for Office 365 users, with or without automapping. | AD, Exchange, M365 |
| [Add-UserToSharedMailbox.ps1](Cloud-Administration/Add-UserToSharedMailbox.ps1) | Grants Office 365 users access to a shared mailbox, with or without automapping. | AD, Exchange, M365 |
| [Cleanup-APIPermissions.ps1](Cloud-Administration/Cleanup-APIPermissions.ps1) | To enhance your tenant's security posture, it's crucial to regularly review the API permissions requested by SPFx solutions and compare them with those granted to the ”SharePoint Online Client Extensibility Web… | AD, M365, SharePoint |
| [Export-WTCAPolicy.ps1](Cloud-Administration/Export-WTCAPolicy.ps1) | Connects to Microsoft Graph and exports Conditional Access policies and related settings. | AD, Graph, M365 |
| [Get-Office365Endpoints.ps1](Cloud-Administration/Get-Office365Endpoints.ps1) | Microsoft updates the Office 365 IP address and FQDN entries at the end of each month and occasionally out of the cycle for operational or support requirements | AD, M365 |
| [Get-TeamsServiceNumbers.ps1](Cloud-Administration/Get-TeamsServiceNumbers.ps1) | Get the Phone numbers assigned to Teams/SfB Services Supported are AutoAttendant and/or CallQueue | AD, M365, Teams |
| [Get-WTAzureADAppSP.ps1](Cloud-Administration/Get-WTAzureADAppSP.ps1) | Connects to Microsoft Graph and reports Azure AD or Entra application service principals. | AD, Entra, Graph, M365 |
| [Get-WTAzureADGroup.ps1](Cloud-Administration/Get-WTAzureADGroup.ps1) | Connects to Microsoft Graph and lists Azure AD or Entra groups. | AD, Entra, Graph, M365 |
| [Get-WTCAPolicy.ps1](Cloud-Administration/Get-WTCAPolicy.ps1) | Connects to Microsoft Graph and reads Conditional Access policy details. | AD, Graph, M365 |
| [Move-Team.ps1](Cloud-Administration/Move-Team.ps1) | Copies channels and files from one Team to another, with a REPORT mode for read-only comparison and an EXECUTE mode that can add members, create channels, and copy files between teams. | AD, M365, Reporting, Teams |
| [New-TeamMailbox.ps1](Cloud-Administration/New-TeamMailbox.ps1) | This scripts creates a new shared mailbox (aka team mailbox) and security groups for full access and and send-as delegation | AD, Exchange, M365 |
| [New-WTAzureADGroup.ps1](Cloud-Administration/New-WTAzureADGroup.ps1) | Connects to Microsoft Graph and creates Azure AD or Entra groups. | AD, Entra, Graph, M365 |
| [Set-Office365Signature.ps1](Cloud-Administration/Set-Office365Signature.ps1) | Configures an Office 365 signature for users. | AD, M365 |
| [Teams-Phone-User-Assignment-Phase1.ps1](Cloud-Administration/Teams-Phone-User-Assignment-Phase1.ps1) | Assign voice-related Teams policies to users | AD, M365, Teams |
| [Teams-Phone-User-Assignment-Phase2-Cutover.ps1](Cloud-Administration/Teams-Phone-User-Assignment-Phase2-Cutover.ps1) | Assign phone numbers and complete Teams phone cutover | AD, M365, Teams |
| [Update-UPNandLicence.ps1](Cloud-Administration/Update-UPNandLicence.ps1) | Reads a CSV of users, updates each account’s UPN, usage location, and assigned Microsoft 365 license, and writes transcript and subject logs. | AD, CSV, License, M365 |

---

### Intune
Scripts for Intune, Autopilot, and device-management workflows.

| Script | Description | Tags |
|--------|-------------|------|
| [Application_Get_Assign.ps1](Intune/Application_Get_Assign.ps1) | Authenticates to Microsoft Graph, enumerates Intune applications, and exports assignment and deployment details to a report. | Graph, Intune, Reporting |
| [Change-DeviceCategory.ps1](Intune/Change-DeviceCategory.ps1) | Sets the device category for a single Intune device using the Microsoft Graph beta endpoint via the Intune PowerShell SDK (Connect-MSGraph). | Device, Endpoint, Graph, Intune |
| [Change-DeviceCategoryMulti.ps1](Intune/Change-DeviceCategoryMulti.ps1) | Iterates over all managed devices and assigns a device category depending on whether the device name matches a pattern | Device, Intune |
| [Change-DeviceCategorySingle.ps1](Intune/Change-DeviceCategorySingle.ps1) | Sets the device category for one Intune device using the Microsoft Graph beta endpoint via the Intune PowerShell SDK (Connect-MSGraph). | Device, Endpoint, Graph, Intune |
| [CompliancePolicy_Export.ps1](Intune/CompliancePolicy_Export.ps1) | Authenticates to Microsoft Graph, enumerates Intune compliance policies, and exports their settings to a report for review. | Compliance, Graph, Intune, Reporting |
| [Copy-DeviceConfigurationProfile.ps1](Intune/Copy-DeviceConfigurationProfile.ps1) | Copy a device configuration profile in Intune | Device, Intune |
| [CorpDeviceEnrollment_Add.ps1](Intune/CorpDeviceEnrollment_Add.ps1) | Authenticates to Microsoft Graph and creates corporate device identifiers for Intune enrollment from a Graph- backed input list. | Device, Graph, Intune |
| [Detect-MultipleIntuneMDMCert.ps1](Intune/Detect-MultipleIntuneMDMCert.ps1) | This script checks the Local Machine certificate store for multiple Intune MDM Device CA certificates | Device, Intune |
| [Get-AllAssignmentsError.ps1](Intune/Get-AllAssignmentsError.ps1) | Retrieves all failed configuration profile and app assignments in the tenant and exports them as CSV files | CSV, Intune |
| [Get-AllAssignmentsErrorAppRegistration.ps1](Intune/Get-AllAssignmentsErrorAppRegistration.ps1) | Retrieves all failed configuration profile and app assignments in the tenant, exports them as CSV files, and sends an email report with attachments via Microsoft Graph | CSV, Exchange, Graph, Intune |
| [Get-CleanUpDiskDetection.ps1](Intune/Get-CleanUpDiskDetection.ps1) | Checks free space on the C: drive | Disk, Intune |
| [Get-CleanUpDiskRemediation.ps1](Intune/Get-CleanUpDiskRemediation.ps1) | Configures selected cleanup categories in the registry and invokes CleanMgr.exe to free disk space | Automation, Disk, Intune |
| [Get-NewEnrolledDevicesReport.ps1](Intune/Get-NewEnrolledDevicesReport.ps1) | Queries Microsoft Graph for devices enrolled in the past 7 days, generates a CSV report, and sends it as an email attachment via the Graph Mail API. | CSV, Device, Exchange, Graph |
| [Get-PendingRebootDetection.ps1](Intune/Get-PendingRebootDetection.ps1) | Checks multiple registry keys to determine whether a system reboot is pending | Intune |
| [Get-PendingRebootNotificationRemediation.ps1](Intune/Get-PendingRebootNotificationRemediation.ps1) | Intune Proactive Remediation script that shows a Windows toast notification prompting the user to reboot their system after updates have been installed. | Automation, Intune, Software |
| [Get-UnassignedAppsAndConfigurations.ps1](Intune/Get-UnassignedAppsAndConfigurations.ps1) | Queries Microsoft Graph for all mobile apps with their assignments expanded and returns those that have no assignments configured. | Graph, Intune |
| [Get-Windows11Report.ps1](Intune/Get-Windows11Report.ps1) | Connects to Microsoft Graph, retrieves all managed Windows devices, calculates Windows 11 adoption percentages, builds an HTML report with a pie chart and device table, and sends it as an email attachment. | Device, Exchange, Graph, Intune |
| [ManagedDevices_Get.ps1](Intune/ManagedDevices_Get.ps1) | Authenticates to Microsoft Graph and exports Intune managed device inventory and compliance details. | Compliance, Device, Graph, Intune |
| [Repair-MultipleIntuneMDMCert.ps1](Intune/Repair-MultipleIntuneMDMCert.ps1) | This script removes duplicate Intune MDM Device CA certificates from the Local Machine certificate store, keeping only the most recent one | Device, Intune, Software |

---

## Requirements

- PowerShell 5.1 or later
- RSAT ActiveDirectory and Group Policy modules for AD / GPO scripts
- Microsoft Graph PowerShell module for Microsoft 365 / Intune scripts
- MicrosoftTeams and ExchangeOnlineManagement for cloud admin scripts
- Administrator privileges for most scripts

---

## Notes

- This repository is maintained as a local script library for MRDTech operations.
- Review scripts before production use.
- Replace environment-specific paths, tenant names, and URLs as needed.

---

## Usage

Run scripts with appropriate permissions.

Example:
```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\ScriptName.ps1
```

---

## Disclaimer

These scripts are provided as-is without any warranties, express or implied.
Use at your own risk.

All scripts should be reviewed and tested in a non-production environment before deployment.

---

_Generated from the current repository tree._
