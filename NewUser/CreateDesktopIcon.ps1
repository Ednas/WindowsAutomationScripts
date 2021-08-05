$TargetFile = "$env:SystemRoot\System32\notepad.exe"
$CalcFile = "$env:SystemRoot\System32\calc.exe"

$ShortcutFile = "$env:Public\Desktop\Notepad.lnk"
$ShortcutFile2 = "$env:Public\Desktop\Calculator.lnk"

$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = $TargetFile
$Shortcut.Save()

$Shortcut2 = $WScriptShell.CreateShortcut($ShortcutFile2)
$Shortcut2.TargetPath = $CalcFile
$Shortcut2.Save()