<#
==============================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : BitLocker Key Sync
SYNOPSIS    : Backs up BitLocker recovery key to Azure AD
DESCRIPTION : Retrieves the BitLocker recovery key protector ID for the
              system drive and backs it up to Azure AD (Entra ID).
              Returns success or failure based on the backup operation
==============================================================================
#>

try{
$BLV = Get-BitLockerVolume -MountPoint $env:SystemDrive
        $KeyProtectorID=""
        foreach($keyProtector in $BLV.KeyProtector){
            if($keyProtector.KeyProtectorType -eq "RecoveryPassword"){
                $KeyProtectorID=$keyProtector.KeyProtectorId
                break;
            }
        }

       $result = BackupToAAD-BitLockerKeyProtector -MountPoint "$($env:SystemDrive)" -KeyProtectorId $KeyProtectorID
return $true
}
catch{
     return $false
}
