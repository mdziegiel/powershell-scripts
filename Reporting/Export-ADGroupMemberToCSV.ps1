<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Export-ADGroupMemberToCSV
SYNOPSIS    : Lists users in an AD group and exports name, object class, and
              SamAccountName to CSV.
DESCRIPTION : Lists users in an AD group and exports name, object class, and
              SamAccountName to CSV.
================================================================================
#>
#Requires -Modules ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
$ErrorActionPreference = "SilentlyContinue"

#----------------------------------------------------------[Declarations]----------------------------------------------------------
#Script Version
$sScriptVersion = "1.0"

$ADGroupName = ""
$CSVExport = ""

#-----------------------------------------------------------[Functions]------------------------------------------------------------
    <#
        Empty
    #>

#------------------------------------------------------------[Actions]-------------------------------------------------------------

Get-ADGroupMember -Identity $ADGroupName | Select-Object name,objectClass,SamAccountName | Export-CSV $CSVExport -Encoding UTF8 -NoType -Force

# Show an information message
[System.Windows.Forms.MessageBox]::Show("All .pst from $PSTfolder were imported to Outlook" , "Information" , 0, [Windows.Forms.MessageBoxIcon]::Information)

# End Script