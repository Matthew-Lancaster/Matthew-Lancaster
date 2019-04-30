@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM RUN FROM HERE NOW NEWER 2018 INFO NOTE 2018-MAY-03
@REM ----------------------------------------------------------------
@REM C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk
@REM IN ANOTHER WAY LIKE THIS
@REM ----------------------------------------------------------------
@REM Run, "%VICE_VERSA_EXE%" "%VICE_VERSA_PATH%\VV C DRIVE ROOT __ %A_ComputerName%.fsf" /hiddenautoexec
@REM ----------------------------------------------------------------
@REM NEW IMPROVED CODE ABOVE AS SYNC WAS NOT POSSIBLE UNLESS EACH 
@REM COMPUTER HAD EACH OWN SYNC DATABASE
@REM ----------------------------------------------------------------
@REM NOW OBSOLETE HERE
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ ALL BATCH FILE COMMAND SETS 
@REM __ GOT AN IMPROVED ADMINISTRATOR LAUNCHER
@REM __ THAT DOESN'T GET STUCK IN A LOOP IF NOT SUCCESSFUL
@REM __ AT THE SACRIFICE OF COMMAND LINE ARGUMENTS BUT EASY TO 
@REM __ ADD CODE IF REQUIRE
@REM ----------------------------------------------------------------
@REM __ FROM ---- Thu 03-May-2018 01:31:51
@REM __ TO   ---- Thu 03-May-2018 03:35:00 __ 2 HOUR
@REM ----------------------------------------------------------------

@CD\

@REM __ SET ADMIN A SAFE DOUBLE CHECKER WAY NOT A LOOPY ONE IF CAN'T GET IT TO GO ADMIN

@ECHO OFF

IF '%1%'=='/ADMIN' (
	@REM :: Check for ADMIN Privileges Again
	>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
	@IF '%errorlevel%' GTR '2' SET NOT_GOT_ADMIN=TRUE
)

@IF "%NOT_GOT_ADMIN%"=="TRUE" (
	ECHO AN ATTEMPT WAS MADE TO SET ADMIN BUT NOT SUCCESSFUL
	ECHO YOUR USER MUST BE A STANDARD USER OR SOMETHING LIKE HER
	ECHO SAVES GET STUCK IN A LOOP ASKING OR ALSO DELAY REQUIRING
	ECHO END
	PAUSE
	GOTO ENDE
	EXIT /B 1
)

@REM :: Check for ADMIN Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
@IF '%errorlevel%' GTR '2' (
	@REM Get ADMIN Privileges
	@ECHO Set UAC = CreateObject^("Shell.Application"^) >> "%temp%\getadmin.vbs"
	@ECHO UAC.ShellExecute "%~s0", "/ADMIN", "", "runas", 1 >> "%temp%\getadmin.vbs"

	"%temp%\getadmin.vbs"
	@DEL "%temp%\getadmin.vbs"
	GOTO ENDE
	@EXIT /B
)

@ECHO Got ADMIN Privileges

@ECHO ON

START "" /WAIT "C:\Program Files\ViceVersa Pro\ViceVersa.exe" "C:\SCRIPTER\SCRIPTER -- ViceVersa PRO\VV C DRIVE ROOT\VV C DRIVE ROOT __ C DRIVE ROOT SOURCE __ C DRIVE.fsf" /hiddenautoexec /autoclose

REM START "" /WAIT "C:\Program Files\ViceVersa Pro\ViceVersa.exe" "C:\SCRIPTER\SCRIPTER -- ViceVersa PRO\VV C DRIVE ROOT\VV C DRIVE ROOT __ C DRIVE ROOT _ SOURCE __ C- _ TARGET.fsf" /hiddenautoexec /autoclose

REM ECHO %TIME% >"%~f0.TIME_EVENT.TXT"

@FOR %%G IN (C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (
	IF EXIST "%%G:\C DRIVE ROOT" ECHO "%%G:\C DRIVE ROOT"
	)
	


	REM PAUSE


REM schtasks /DELETE /TN "z_VICE_VERSA_C_DRIVE_ROOT"

REM schtasks /QUERY /TN "z_VICE_VERSA_C_DRIVE_ROOT" >NUL 2>"%~dp0SCHTasks_Output_2.TXT"
REM @FOR /F "tokens=1 delims= " %%A IN ('TYPE "%~dp0SCHTasks_Output_2.TXT"') DO @IF "%%A"=="ERROR:" SET Var_01=NOT_EXIST_SCHEDULE
REM DEL "%~dp0SCHTasks_Output_2.TXT"

REM @if "%Var_01%"=="" GOTO NOT_HERE

REM schtasks /CREATE /TN ""z_VICE_VERSA_C_DRIVE_ROOT"" /TR "%~f0" /SC MINUTE /MO 60 /RU "%USERNAME%" /RP " " /RL HIGHEST /F

REM schtasks /CREATE /TN ""z_VICE_VERSA_C_DRIVE_ROOT"" /TR "%~f0" /SC MINUTE /MO 60 /U "%USERNAME%" /P " " /RL HIGHEST /F

REM schtasks /CREATE /U "%COMPUTERNAME%\%USERNAME%" /P " " /TN ""z_VICE_VERSA_C_DRIVE_ROOT"" /TR "%~f0" /SC MINUTE /MO 60 /RL HIGHEST /F

REM schtasks /CREATE /RU ADMINISTRATORS /TN ""z_VICE_VERSA_C_DRIVE_ROOT"" /TR "%~f0" /SC MINUTE /MO 60 /RL HIGHEST /F

REM :NOT_HERE

REM /autoexec
REM /dialogautoexec 
REM /hiddenautoexec

REM PAUSE

:VERY_ENDER_BLOCK
:ENDE

exit
