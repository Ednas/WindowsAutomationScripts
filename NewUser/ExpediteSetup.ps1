# Author: Edna Jonsson
# 8/3/2021
# You might need to change the execution policy in order for this script to run. 
# Get-ExecutionPolicy
# Set-ExecutionPolicy Unrestricted

# Open webpages to download Adobe Reader and ***
Start-Process "https://get.adobe.com/reader/"
Start-Process "https://www.google.com/intl/en_us/chrome/"

# Open Webpage in Chrome 
[system.Diagnostics.Process]::Start("chrome","https://thedri.com")

# Opens Word so you can login to O365 with the license
 $Word = New-Object -ComObject Word.Application
 $Word.Visible = $True




# Set Power Settings to NEVER Sleep while plugged in
Powercfg /Change standby-timeout-ac 0

# Deletes Icons on the desktop
Remove-Item C:\Users\*\Desktop\*lnk –Force

# This will change the OneDrive from being twice in File Explorer
Write-Output  'This will change the Registry for Azure Account computers'
Get-PSDrive -PSProvider 'Registry' | Select-Object -Property Name, Root
Set-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}' -Name 'System.IsPinnedToNameSpaceTree' -value '0'

Get-CimInstance -ClassName Win32_Desktop | Select-Object -ExcludeProperty "CIM*"


# Let's seperate this to a closeout script or a hold to process the end later

# Cleanup from Word
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable Word

# Todo - Set up a timer function to be able to uncomment the last line 
# Set-ExecutionPolicy Restricted
# rundll32.exe user32.dll,LockWorkStation