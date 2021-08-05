
# Word
$Shortcut2 = $WScriptShel2.CreateShortcut($ShortcutFileWord)
$Shortcut2.TargetPath = $WordFile
$Shortcut2.Save()

# Excel
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFileExcel)
$Shortcut.TargetPath = $ExcelFile
$Shortcut.Save()

# PowerPoint
$PPShortcut = $WScriptShell.CreateShortcut($ShortcutFilePowerpoint)
$PPShortcut.TargetPath = $PowerpointFile
$PPShortcut.Save()

# "C:\Program Files\Microsoft Office\root"
$WordFile = "$env:SystemRoot\System32\word.exe"
$PowerpointFile = "$env:SystemRoot\System32\powerpoint.exe"
$ExcelFile = "$env:SystemRoot\System32\excel.exe"