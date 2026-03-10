$TargetFile = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
$ShortcutFile = "$env:Public\Desktop\Word.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = $TargetFile
$Shortcut.Save()

$TargetFile = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
$ShortcutFile = "$env:Public\Desktop\Excel.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = $TargetFile
$Shortcut.Save()

$TargetFile = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
$ShortcutFile = "$env:Public\Desktop\Outlook.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = $TargetFile
$Shortcut.Save()

$TargetFile = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
$ShortcutFile = "$env:Public\Desktop\PowerPoint.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = $TargetFile
$Shortcut.Save()
