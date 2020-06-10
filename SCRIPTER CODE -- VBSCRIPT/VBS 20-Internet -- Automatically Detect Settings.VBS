'--------------------------------------------
'This is not working for me 
'And I have to set two setting value
'The binary bit changes as should 
'When the checkbox is checked and unchecked
'But even as the registry value set changed
'It still don't produced the checkbox checked
'Even with a reboot
'the Original Code was to Use HKCU 
'But My Search Found was in a Different Place HKLM
'The 1st Value if Working Maybe True for what you Want
'But My Look at It Changes in Two Places
'--------------------------------------------

'Example use of IEautomaticallydetect 
IEautomaticallydetect "OnlyAutoDetect" 
wscript.quit

SUB IEautomaticallydetect(status)

strComputer = "." ' Use . for current machine

DIM sKey
DIM sValue_1
DIM sValue_2
DIM binaryVal_1 
DIM binaryVal_2
Dim oReg 

'Connect to the Registry

'Set oReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

'Retrieve the current settings.

Const HKCU=&H80000001
Const HKLM=&H80000002
Const HKEY_CLASSES_ROOT   = &H80000000
Const HKEY_CURRENT_USER   = &H80000001
Const HKEY_LOCAL_MACHINE  = &H80000002
Const HKEY_USERS          = &H80000003

Const REG_SZ        = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY    = 3
Const REG_DWORD     = 4
Const REG_MULTI_SZ  = 7

hDefKey = HKEY_LOCAL_MACHINE

'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections

'HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections
'sKey = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"

sValue_1 = "DefaultConnectionSettings"
sKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"
oReg.GetBinaryValue hDefKey, sKey, sValue_1, binaryVal_1
sValue_2 = "SavedLegacySettings"
oReg.GetBinaryValue hDefKey, sKey, sValue_2, binaryVal_2

'09 when only 'Automatically detect settings' is enabled 
'03 when only 'Use a proxy server for your LAN' is enabled 
'0B when both are enabled 
'05 when only 'Use automatic configuration script' is enabled 
'0D when 'Automatically detect settings' and 'Use automatic configuration script' are enabled 
'07 when 'Use a proxy server for your LAN' and 'Use automatic configuration script' are enabled 
'0F when all the three are enabled. 
'01 when none of them are enabled.

select case (status) 
case "proxy" binaryVal_1(8) = 03          'Force Autodetect on 
case "pac" binaryVal_1(8) = 05            'Force Autodetect off 
case "OnlyAutoDetect" binaryVal_1(8) = 09 'Only 'Automatically detect settings' is enabled 
case "OnlyAutoDetect" binaryVal_1(8) = 09 'Only 'Automatically detect settings' is enabled 
end select


binaryVal_2(8)=binaryVal_1(8)

if status<>"" then
	oReg.SetBinaryValue hDefKey, sKey, sValue_1, binaryVal_1
	oReg.SetBinaryValue hDefKey, sKey, sValue_2, binaryVal_2
end if


'----
'LAN Connection settings keep changing back to proxy server after - Microsoft Community
'https://answers.microsoft.com/en-us/ie/forum/ie11-iewindows8_1/lan-connection-settings-keep-changing-back-to/76a0f5d2-167f-41fa-bf40-1461b8c01642?page=2
'----
'----
'VBScript Scripting Techniques: Work with the Registry
'http://www.robvanderwoude.com/vbstech_registry_wshshell.php
'----
' Create a WSH Shell object:
Set wshShell = CreateObject( "WScript.Shell" )
'
' Create a new DWORD value:
wshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\MigrateProxy", 0, "REG_DWORD"
wshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"
wshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyHttp1.1", 0, "REG_DWORD"
wshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", "", "REG_SZ"

'I TRY THESE NOTHING MADE A DIFFERENCE
'-------------------------------------
'Problem Found Video Downloader Professional 
'Seem to Be Buggy with Virus For Days Also
'It is the extension for Chrome
'Video Downloader Ultimate is Okay
'There wasn't Any Icon showing for the Former
'I Thought they were both made by Nero
'The Later One Is
'It is Quite a Good Thing Download Youtube Clips
'And Pay for Only Once
'-------------------------------------
end sub

'---------------------
'CREDIT PAGE
'COMMENT SECTION TAKEN
'----
'Disable Internet Explorer automatically detect settings checkbox
'https://www.htguk.com/disabling-internet-explorer/
'----
'---------------------


'Some settings are managed by system administrator
'Unable to change proxy settings error some settings are managed by system administrator
'----
'Unable to change proxy settings error some settings are managed by - Microsoft Community
'https://answers.microsoft.com/en-us/ie/forum/ie8-windows_7/unable-to-change-proxy-settings-error-some/a7f3e6eb-072c-4c0e-8d49-c2c0673f24f5
'----
'HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings
'Method 1: Check the permissions on Internet Settings in registry
'a)   Click on start button > type in  regedit in the search box and hit enter
'b)   Navigate to the location below
'           HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings
'c)   Right Click on Internet Settings and select Permissions
'd)   Ensure that your user profile has full control