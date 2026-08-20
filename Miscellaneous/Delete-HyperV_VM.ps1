<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Delete-HyperV_VM
SYNOPSIS    : Deletes Hyper-V virtual machines listed in a CSV file.
DESCRIPTION : Deletes Hyper-V virtual machines listed in a CSV file.
================================================================================
#>
$ErrorActionPreference = "SilentlyContinue"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#Write script directory path on "ScriptDir" variable
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

# Log file creation, similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD].log
$ActualDate = Get-Date -uformat %Y_%m_%d
$ScriptLogFile = "$ScriptDir\$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))" + "_" + $ActualDate + ".log"

# Declare Environment and Hyper-V parameters
$VMLoc = "E:\Hyper-V\"

#-----------------------------------------------------------[Functions]------------------------------------------------------------
function Stop-TranscriptOnLog
{
	Stop-Transcript
	# Add EOL required for Notepad.exe application usage
	[string]::Join("`r`n", (Get-Content $ScriptLogFile)) | Out-File $ScriptLogFile
}

function Select-FileDialog
{
	param ([string]$Title, [string]$Filter = "All files *.*|*.*")
	[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
	$fileDialogBox = New-Object Windows.Forms.OpenFileDialog
	$fileDialogBox.ShowHelp = $false
	$fileDialogBox.initialDirectory = $ScriptDir
	$fileDialogBox.filter = $Filter
	$fileDialogBox.Title = $Title
	$Show = $fileDialogBox.ShowDialog()
	
	If ($Show -eq "OK")
	{
		Return $fileDialogBox.FileName
	}
	Else
	{
		Write-Error "Canceled operation"
		[System.Windows.Forms.MessageBox]::Show("Script is not able to continue. Operation stopped.", "Operation canceled", 0, [Windows.Forms.MessageBoxIcon]::Error)
		Stop-TranscriptOnLog
		Exit
	}
	
}

#------------------------------------------------------------[Actions]-------------------------------------------------------------
# Start of log completion
Start-Transcript $ScriptLogFile | Out-Null

# Import CSV file
[System.Windows.Forms.MessageBox]::Show(
"
Select on this window the CSV file who contains VM names.
Its content must be similar to:
  
Name			(Required line)
VM01
VM02
VM03
", "VM list", 0, [Windows.Forms.MessageBoxIcon]::Question)

$CSVInputFile = Select-FileDialog -Title "Select CSV file" -Filter "CSV File (*.csv) |*.csv"

# Import VM parameters list
$csvValues = Import-Csv $CSVInputFile -Delimiter ';'

# Delete Virtual Machines
foreach ($VMList in $VMList)
{
	$VMName = $VMList.Name
	# Stop, Delete VM & VHDX
	Get-VM $VMName | %{ Stop-VM -VM $_ -Force; Remove-VM -VM $_ -Force; Remove-Item -Path $_.Path -Recurse -Force }
}

# Get VM
Get-VM | Sort-Object Name | Select Name, State, CPUUsage, MemoryAssigned | Export-CliXML $ScriptDir\Get-VM.xml
Import-CliXML $ScriptDir\Get-VM.xml | Out-GridView -Title Get-VM -PassThru

# Stop the log transcript
Stop-TranscriptOnLog