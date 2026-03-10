???# Create needed keys
New-Item -Path HKLM:\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\Settings\Schedule -ErrorAction SilentlyContinue
New-Item -Path HKLM:\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\Settings\General -ErrorAction SilentlyContinue

# Insert desired values
New-ItemProperty -Path HKLM:\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\Settings\Schedule -Name AutomationMode -Value ScanDownloadApplyNotify -PropertyType String -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\Settings\Schedule -Name DayOfWeek -Value Monday -PropertyType String -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\Settings\Schedule -Name DeferRestartCount -Value 3 -PropertyType String -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\Settings\Schedule -Name DeferRestartInterval -Value 8 -PropertyType String -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\Settings\Schedule -Name Time -Value 2022-11-10T11:00:00 -PropertyType String -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\Settings\Schedule -Name ScheduleMode -Value Weekly -PropertyType String -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\Settings\Schedule -Name SystemRestartDeferral -Value 1 -PropertyType DWORD -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\Settings\General -Name MaxRetryAttempts -Value 3 -PropertyType DWORD -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\Settings\General -Name SuspendBitLocker -Value 1 -PropertyType DWORD -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Dell\UpdateService\Clients\CommandUpdate\Preferences\Settings\General -Name UserConsentDefault -Value 0 -PropertyType DWORD -Force

# Set BIOS password (Require Dell Command | Configure in place)
Start-Process -FilePath "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe" -ArgumentList '/configure -biosPassword="NorthernBIOS9"'
