# Startup shortcut creation
$WshShell = New-Object -ComObject WScript.Shell
$ShortcutPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\AudioWindowTracker.lnk"
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "c:\Users\cleoz\Projects\Personal\AudioWindowTracker\AudioWindowTracker.exe"
$Shortcut.WorkingDirectory = "c:\Users\cleoz\Projects\Personal\AudioWindowTracker"
$Shortcut.Save()
Write-Host "Startup shortcut created at $ShortcutPath"
