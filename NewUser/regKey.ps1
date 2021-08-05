Write-Output  'This will change the Registry for Azure Account computers'
Get-PSDrive -PSProvider 'Registry' | Select-Object -Property Name, Root
Set-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}' -Name 'System.IsPinnedToNameSpaceTree' -value '0'

# Set-ExecutionPolicy Restricted