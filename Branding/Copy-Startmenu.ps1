<#
================================================================================
AUTHOR      : Michael Dziegiel
SCRIPT      : Copy-Startmenu
SYNOPSIS    : Copies a Windows 11 Start menu template to user profiles and the
              default profile.
DESCRIPTION : Copies a Windows 11 Start menu template to all existing user
              profiles and the default profile. This standardizes
              pinned items across new and existing users.
================================================================================
#>
$startmenuTemplate = "\\server\repository\Windows\Startmenu win11\start2.bin"

# Get all user profile folders
$usersStartMenu = get-childitem -path "C:\Users\*\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState"

# Copy Start menu to all users folders
ForEach ($startmenu in $usersStartMenu) {
  Copy-Item -Path $startmenuTemplate -Destination $startmenu -Force
}

# Create default user profile folders
# and copy start menu layout

# Default profile path
$defaultProfile = "C:\Users\default\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState"

# Create folders if they don't exist
if(-not(Test-Path $defaultProfile)) {
   new-item $defaultProfile -ItemType Directory -Force
}

# Copy file
Copy-Item -Path $startmenuTemplate -Destination $defaultProfile -Force