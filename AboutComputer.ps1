# https://www.educba.com/useful-powershell-scripts/
# Fetch Information Related to System

Write-Host "Welcome to the script of fetching computer Information"
Write-host "The BIOS Details are as follows"
Get-CimInstance -ClassName Win32_BIOS

Write-Host "The systems processor is"
Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -Property SystemType

Write-Host "The computer Manufacture and physical memory details are as follows"
Get-CimInstance -ClassName Win32_ComputerSystem

Write-Host "The installed hotfixes are"
Get-CimInstance -ClassName Win32_QuickFixEngineering

Write-Host "The OS details are below"
Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -Property Build*,OSType,ServicePack*

Write-Host "The following are the users and the owners"
Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -Property *user*

Write-Host "The disk space details are as follows"
Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" |Measure-Object -Property FreeSpace,Size -Sum |Select-Object -Property Property,Sum

Write-Host "Current user logged in to the system"
Get-CimInstance -ClassName Win32_ComputerSystem -Property UserName

Write-Host "Status of the running services are as follows"
Get-CimInstance -ClassName Win32_Service | Format-Table -Property Status,Name,DisplayName -AutoSize -Wrap

