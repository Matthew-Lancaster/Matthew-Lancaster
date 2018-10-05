

# powershell -ExecutionPolicy Unrestricted Add-AppxPackage -DisableDevelopmentMode -Register $Env:SystemRoot\WinStore\AppxManifest.xml


Get-AppXPackage | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml"}

Get-AppXPackage -AllUsers -Name Microsoft.WindowsStore | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" -Verbose}

# pause