$OutlookFile = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
$ExcelFile = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
$WordFile = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
$PowerpointFile = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"


$ShortcutFileOutlook = "$env:Public\Desktop\Outlook.lnk"
$ShortcutFileWord = "$env:Public\Desktop\Word.lnk"
$ShortcutFileExcel = "$env:Public\Desktop\Excel.lnk"
$ShortcutFilePowerpoint = "$env:Public\Desktop\PowerPoint.lnk"

$WScriptShell = New-Object -ComObject WScript.Shell

# Outlook
$OShortcut = $WScriptShell.CreateShortcut($ShortcutFileOutlook)
$OShortcut.TargetPath = $OutlookFile
$OShortcut.Save()

# Word
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFileWord)
$Shortcut.TargetPath = $WordFile
$Shortcut.Save()

# Excel
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFileExcel)
$Shortcut.TargetPath = $ExcelFile
$Shortcut.Save()


# PowerPoint
$PPShortcut = $WScriptShell.CreateShortcut($ShortcutFilePowerpoint)
$PPShortcut.TargetPath = $PowerpointFile
$PPShortcut.Save()



