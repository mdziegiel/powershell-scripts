???$ODClient = "https://go.microsoft.com/fwlink/?linkid=844652"
$output = "$ENV:temp"  + '\OneDriveSetup.exe'
$apppath = "C:\Program Files (x86)\Microsoft OneDrive\OneDrive.exe"
$action = New-ScheduledTaskAction -Execute $apppath
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date)
 
Invoke-WebRequest -Uri $ODClient -OutFile $output
Start-Process -FilePath $output -ArgumentList '/allusers', '/silent'
Start-Sleep -Seconds 60

Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "Launch OneDrive" | Out-Null
Start-ScheduledTask -TaskName "Launch OneDrive"
Start-Sleep -Seconds 5
Unregister-ScheduledTask -TaskName "Launch OneDrive" -Confirm:$false
