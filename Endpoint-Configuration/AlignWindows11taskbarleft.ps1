$registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$Al = "TaskbarAl"   # Taskbar alignment
$value = "0"

New-ItemProperty -Path $registryPath -Name $Al -Value $value -PropertyType DWORD -Force -ErrorAction Ignore
