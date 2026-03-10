rm -r c:\URLIcon
del "C:\Users\Public\Desktop\HelpDesk.lnk"

#Variables creating local folder and download .ico file
$LocalIconFolderPath = "C:\URLIcon"
$SourceIcon = "https://helpdesk.hanskissle.com/custom/customimages/MEHD.ico"
$DestinationIcon = "C:\URLIcon\HelpDesk.ico"


#Step 1 - Create a folder to place the URL icon
New-Item $LocalIconFolderPath -Type Directory

#Step 2 - Download a ICO file from a website into previous created folder
curl $SourceIcon -o $DestinationIcon

#Step 3 - Add the custom URL shortcut to your Desktop with custom icon
$new_object = New-Object -ComObject WScript.Shell
$destination = $new_object.SpecialFolders.Item('AllUsersDesktop')
$source_path = Join-Path -Path $destination -ChildPath '\\HelpDesk.lnk'
$source = $new_object.CreateShortcut($source_path)
$source.TargetPath = 'https://helpdesk.hanskissle.com/'
$source.IconLocation = ”C:\URLIcon\HelpDesk.ico”
$source.Save()