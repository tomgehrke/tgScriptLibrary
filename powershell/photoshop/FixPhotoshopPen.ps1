# PhotoshopPenFixer.ps1
# by Tom Gehrke

# Applies "fixes" for issues caused by the interaction of Wacom drivers and
# Windows Ink functionality within Adobe Photoshop.

# Disable use of Windows pen API

$PhotoshopFolders = Get-ChildItem -Path "$env:APPDATA\Adobe\Adobe Photoshop CC *"

foreach ($PhotoshopVersion in $PhotoshopFolders) {
    $SettingsFolder = "$PhotoshopVersion\$(Split-Path $PhotoshopVersion -Leaf) Settings"
    Add-Content -Path "$SettingsFolder\PsUserConfig.txt" -Value "# Use WinTab`r`nUseSystemStylus 0"
}