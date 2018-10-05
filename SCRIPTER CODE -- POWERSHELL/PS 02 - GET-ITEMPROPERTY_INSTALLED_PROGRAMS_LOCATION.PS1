# 32 BIT
Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, InstallLocation |Format-Table -AutoSize

# STILL MISSING WINRAR

Write-Host "---------------------------------------------------------"
Write-Host "---------------------------------------------------------"

# 64 BIT
Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, InstallLocation |Format-Table -AutoSize

# HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
# WINRAR

# BOTH SAME RESULT

# Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, InstallLocation |Format-Table -AutoSize

# HKLM\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall