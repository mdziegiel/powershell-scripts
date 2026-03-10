<#
.SYNOPSIS
    --.Replace "This PC" icon name with the actual name of the PC
.DESCRIPTION
    This script takes ownership of the registry value HKEY_CLASSES_ROOT\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}
    It then updates the key name to $env:ComputerName & The LocalizedName as well
#>

#Requires -Version 3.0

#region VariableDeclaration
$ScriptVersion = "21.2.7.1"
#endregion

#region FunctionListings
function enable-privilege {
 param(
  [ValidateSet(
   "SeAssignPrimaryTokenPrivilege","SeAuditPrivilege","SeBackupPrivilege",
   "SeChangeNotifyPrivilege","SeCreateGlobalPrivilege","SeCreatePagefilePrivilege",
   "SeCreatePermanentPrivilege","SeCreateSymbolicLinkPrivilege","SeCreateTokenPrivilege",
   "SeDebugPrivilege","SeEnableDelegationPrivilege","SeImpersonatePrivilege",
   "SeIncreaseBasePriorityPrivilege","SeIncreaseQuotaPrivilege",
   "SeIncreaseWorkingSetPrivilege","SeLoadDriverPrivilege",
   "SeLockMemoryPrivilege","SeMachineAccountPrivilege",
   "SeManageVolumePrivilege","SeProfileSingleProcessPrivilege",
   "SeRelabelPrivilege","SeRemoteShutdownPrivilege",
   "SeRestorePrivilege","SeSecurityPrivilege","SeShutdownPrivilege",
   "SeSyncAgentPrivilege","SeSystemEnvironmentPrivilege",
   "SeSystemProfilePrivilege","SeSystemtimePrivilege",
   "SeTakeOwnershipPrivilege","SeTcbPrivilege","SeTimeZonePrivilege",
   "SeTrustedCredManAccessPrivilege","SeUndockPrivilege",
   "SeUnsolicitedInputPrivilege")]
  $Privilege,
  $ProcessId = $pid,
  [Switch] $Disable
 )

 $definition = @'
 using System;
 using System.Runtime.InteropServices;

 public class AdjPriv {
  [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
  internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall,
   ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);

  [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
  internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr phtok);

  [DllImport("advapi32.dll", SetLastError = true)]
  internal static extern bool LookupPrivilegeValue(string host, string name, ref long pluid);

  [StructLayout(LayoutKind.Sequential, Pack = 1)]
  internal struct TokPriv1Luid {
   public int Count;
   public long Luid;
   public int Attr;
  }

  internal const int SE_PRIVILEGE_ENABLED = 0x2;
  internal const int TOKEN_QUERY = 0x8;
  internal const int TOKEN_ADJUST_PRIVILEGES = 0x20;

  public static bool EnablePrivilege(long processHandle, string privilege, bool disable) {
   TokPriv1Luid tp;
   IntPtr htok = IntPtr.Zero;
   OpenProcessToken(new IntPtr(processHandle), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
   tp.Count = 1;
   tp.Luid = 0;
   tp.Attr = disable ? 0 : SE_PRIVILEGE_ENABLED;
   LookupPrivilegeValue(null, privilege, ref tp.Luid);
   return AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
  }
 }
'@

 $processHandle = (Get-Process -Id $ProcessId).Handle
 $type = Add-Type $definition -PassThru
 $type[0]::EnablePrivilege($processHandle, $Privilege, $Disable)
}
#endregion

#region ScriptBody

# Take ownership of This PC CLSID
enable-privilege SeTakeOwnershipPrivilege
$key = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey(
 "CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}",
 [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree,
 [System.Security.AccessControl.RegistryRights]::TakeOwnership
)

$acl = $key.GetAccessControl([System.Security.AccessControl.AccessControlSections]::None)
$acl.SetOwner([System.Security.Principal.NTAccount]"BUILTIN\Administrators")
$key.SetAccessControl($acl)

$acl = $key.GetAccessControl()
$rule = New-Object System.Security.AccessControl.RegistryAccessRule(
 "BUILTIN\Administrators","FullControl","Allow"
)
$acl.SetAccessRule($rule)
$key.SetAccessControl($acl)
$key.Close()

# Grant rights to Admin & SYSTEM
$RegistryPath = "Registry::HKEY_CLASSES_ROOT\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
$NewAcl = Get-Acl $RegistryPath

$NewAcl.SetAccessRule(
 New-Object System.Security.AccessControl.RegistryAccessRule(
  "BUILTIN\Administrators","FullControl","Allow"
 )
)

$NewAcl.SetAccessRule(
 New-Object System.Security.AccessControl.RegistryAccessRule(
  "NT AUTHORITY\SYSTEM","FullControl","Allow"
 )
)

Set-Acl -Path $RegistryPath -AclObject $NewAcl

# Rename This PC
Set-Item -Path $RegistryPath -Value $env:COMPUTERNAME -Force
Set-ItemProperty -Path $RegistryPath -Name "LocalizedString" -Value $env:COMPUTERNAME -Force

# Show "This PC" icon
Set-ItemProperty `
 -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" `
 -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" `
 -Value 0 -Force

# ================================
# 🔽 ADDED SECTION (ONLY CHANGE)
# Enable "User's Files" desktop icon
# ================================

$UsersFilesGuid = "{59031a47-3f72-44a7-89c5-5595fe6b30ee}"

$Targets = @(
 "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel",
 "HKU:\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
)

$LoadedUserSids = Get-ChildItem Registry::HKEY_USERS |
 Where-Object { $_.PSChildName -match '^S-1-5-21-' } |
 Select-Object -ExpandProperty PSChildName

foreach ($sid in $LoadedUserSids) {
 $Targets += "HKU:\$sid\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
}

foreach ($t in $Targets | Select-Object -Unique) {
 New-Item -Path $t -Force | Out-Null
 Set-ItemProperty -Path $t -Name $UsersFilesGuid -Value 0 -Force
}

# ================================
# 🔼 END ADDED SECTION
# ================================

Get-Item -Path $RegistryPath
exit $exitcode

#endregion
