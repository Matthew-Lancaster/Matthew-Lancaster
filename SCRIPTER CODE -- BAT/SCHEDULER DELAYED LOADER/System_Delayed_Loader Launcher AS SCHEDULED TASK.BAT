@ECHO OFF

@REM ---------------------------------------
@ECHO ____ System_Delayed_Loader Launcher AS SCHEDULED TASK.BAT ____ 
TITLE %~f0
@REM ---------------------------------------

@REM :: Check for ADMIN Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
@IF '%errorlevel%' GTR '2' (
	@REM Get ADMIN Privileges
	@ECHO Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
	@ECHO UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
	"%temp%\getadmin.vbs"
	@DEL "%temp%\getadmin.vbs"
	@GOTO VERY_ENDER_BLOCK
	@EXIT /B
)

@ECHO Got ADMIN Privileges

@CD \

REM --------------------------
schtasks /QUERY /TN "\SYSTEM LOGON.BAT" | find "ERROR: The system cannot find the file specified." > nul
IF NOT %errorlevel% EQU 0 GOTO NOT
REM --------------------------
schtasks /CREATE /RU ADMINISTRATORS /TN "\SYSTEM LOGON.BAT" /TR "C:\22_LOGON.BAT" /SC ONLOGON /RL HIGHEST /F
REM --------------------------
:NOT

GOTO NOT

REM --------------------------
schtasks /QUERY /TN "\SYSTEM STARTER.BAT" | find "ERROR: The system cannot find the file specified." > nul
IF NOT %errorlevel% EQU 0 GOTO NOT
REM --------------------------
schtasks /CREATE /RU ADMINISTRATORS /TN "\SYSTEM STARTER.BAT" /TR "C:\24_SYSTEM_START.BAT" /SC ONSTART /RL HIGHEST /F
REM --------------------------
:NOT

GOTO NOT

REM --------------------------
schtasks /QUERY /TN "\SYSTEM_DELAYED_LOADER.EXE" | find "ERROR: The system cannot find the file specified." > nul
IF NOT %errorlevel% EQU 0 GOTO NOT
REM --------------------------
schtasks /CREATE /RU "MATT 01" /RP " " /TN "\SYSTEM_DELAYED_LOADER.EXE" /TR "D:\VB6\VB-NT\00_Best_VB_01\DELAYED_LOADER\SYSTEM_DELAYED_LOADER.exe" /SC ONLOGON /RL HIGHEST /F
REM --------------------------
:NOT

PAUSE
