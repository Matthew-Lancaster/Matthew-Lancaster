<#
----
[SOLVED] Delete folder on Target PC. Using PowerShell. - Spiceworks
https://community.spiceworks.com/topic/1979226-delete-folder-on-target-pc-using-powershell
----
#>

Write-Host "get-host"
get-host

Write-Host "VersionTable.PSVersion.Major"
$PSVersionTable.PSVersion.Major

Write-Host "VersionTable.PSVersion"
$PSVersionTable.PSVersion

Write-Host "VersionTable"
$PSVersionTable

Write-Host "winrm quickconfig -q"
winrm quickconfig -q


# Invoke-Command -ComputerName 1-asus-x5dij -ScriptBlock{Remove-Item "d:\0 00 HTTrack\00 ICONS\Icon Archive Site\*.html" -Recurse}

<#
[1-asus-x5dij] Connecting to remote server 1-asus-x5dij failed with the following error message : The WinRM client cannot process the request. If the authentication scheme is different from Kerberos, or if the
client computer is not joined to a domain, then HTTPS transport must be used or the destination machine must be added to the TrustedHosts configuration setting. Use winrm.cmd to configure TrustedHosts. Note
that computers in the TrustedHosts list might not be authenticated. You can get more information about that by running the following command: winrm help config. For more information, see the
about_Remote_Troubleshooting Help topic.
    + CategoryInfo          : OpenError: (1-asus-x5dij:String) [], PSRemotingTransportException
    + FullyQualifiedErrorId : ServerNotTrusted,PSSessionStateBroken
#>