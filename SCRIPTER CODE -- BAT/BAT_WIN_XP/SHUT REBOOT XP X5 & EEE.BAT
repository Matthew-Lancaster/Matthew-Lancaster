REM :: Check for ADMIN Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' GTR '2' (
	REM Get ADMIN Privileges
	echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
	echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
	"%temp%\getadmin.vbs"
	del "%temp%\getadmin.vbs"
	GOTO VERY_ENDER_BLOCK
	EXIT /B
)

ECHO Got ADMIN Privileges


SHUTDOWN -R -F -M 2-ASUS-EEE -T 10
SHUTDOWN -R -F -M 1-ASUS-X5DIJ -T 10
REM -- WAIT FOR OTHER WITH -T 10
