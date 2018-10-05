
@REM -- USE A TWEAK PROGRAM 
@REM -- CAN BE DONE BY REGIRSRY SETTER
@REM -- BUT TWEAK WILL DO MAKE SURE UAC IS NOT DISABLED
@REM -- TO GET MISCROSOFT STORE OPEN WITHOUT ADMIN ERROR

WSRESET 

powershell -noexit -executionpolicy bypass -File "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 03 - MICROSOFT STORE THIS APP CAN'T OPEN _ FIX.PS1"


REM powershell -ExecutionPolicy Unrestricted Add-AppxPackage -DisableDevelopmentMode -Register $Env:SystemRoot\WinStore\AppxManifest.xml

REM PAUSE

REM powershell -noexit -executionpolicy bypass -File "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 03 - MICROSOFT STORE THIS APP CAN'T OPEN _ FIX.PS1"

REM C:\Users\MATT 03\AppData\Local\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe



