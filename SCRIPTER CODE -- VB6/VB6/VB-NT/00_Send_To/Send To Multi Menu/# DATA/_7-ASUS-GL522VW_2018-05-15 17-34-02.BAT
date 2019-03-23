@ECHO OFF
Rem :: Check for ADMIN Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' GTR '2' (
    Rem Get ADMIN Privileges
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    GoTo VERY_ENDER_BLOCK
    EXIT /B
)
ECHO Got ADMIN Privileges
@CD\
@ECHO ON
SET PATH_TO_DO=
Set PATH_TO_DO="C:\Users\MATT 04\AppData\Roaming\*"
@ECHO %PATH_TO_DO%
ICACLS %PATH_TO_DO% /setowner Everyone /T /C /L /Q
ICACLS %PATH_TO_DO% /T /L /C /Q /Grant Everyone:F
@ECHO.
@ECHO OFF
PAUSE
START /D "D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\# DATA\_7-ASUS-GL522VW_2018-05-15 17-34-02.VBS"
:VERY_ENDER_BLOCK
EXIT

