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
@REM __ FROM ---- Thu 03-May-2018 01:31:51 __ INITAL START
@REM __ TO   ---- Thu 03-May-2018 03:35:00 __ 2 HOUR
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- 001 SESSION -- UPDATES HAPPEN ALL THE TIME
@REM -- 
@REM ----------------------------------------------------------------
@REM -- SETTING PERMISSION THIS SCRIPT IS AUTORUN BY AN AUTOHOTKEYS PROGRAM
@REM -- EVERY FEW DAY AND A TIMER MARKER IS WRITTEN TO FILE
@REM -- YOU HAVE OT KEEP IT UP WRITE PERMISSION OR SILLY THING HOPELESS
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- HERE BAT IS NOW SYNC IN 2 LOCATION 
@REM -- CENTRALIZING MY BEST CODE
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 47-OWNER-HARDCODED ANYWHERE.BAT
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\OWNER\#_OWNER-HARDCODED ANYWHERE.BAT
@REM ----------------------------------------------------------------
@REM -- IS IS SYNCED BY HERE
@REM -- C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 32-COPIER_SYNC.VBS
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- 002 SESSION -- 
@REM -- 
@REM ----------------------------------------------------------------
@REM -- CHANGED EnableDelayedExpansion TO USE LATER IN THE PROGRAM WHERE 
@REM -- THE FOR LOOP ARE
@REM -- 
@REM -- PLAYED AROUND LOOKING FOR CODE TO COMPARE TWO VARIABLE IN 
@REM -- EnableDelayedExpansion IN A LOOP AND IN A FUNCTION FIND IMPOSIALE AS FUNCTION WON'T USE GOTO LABEL
@REM -- AND FOR LOOP CAN COUNT TO CERTAIN AMOUNT OR CARRY ON FOREVER
@REM -- AND VARIBALE HAVE TO BE SOMEHOW PUT INTO A FOR LOOP TO USE THERE
@REM -- ANOTHER THING IMPOSIABLE 
@REM -- AT LEAST FOUND OUT ABOUT INFINATE FOR LOOP WITH COUNTER STEPPER SET ON NAUGHT INSTEAD OF 1
@REM -- 
@REM -- ALSO FINDING OUT THAT TAKEOWNERSHIP IS MORE POWERFUL IF NOT GOT ACCESS TO COPIER OF FILE
@REM -- 
@REM -- EVEN IF IT CANT GIVE ACCESS TO ICACLS CHANGE
@REM -- 
@REM -- ALSO "C:\Windows\System32\*.VBS" THIS ONE WHEN TAKEOWNERSHIP HAPPEN
@REM -- AS THERE ARE FEW VBS SCRIPT IN THERE
@REM -- TAKE OWNERSHIP HAPPEN UPTO THIS 4TH FILE C:\Windows\System32\winrm.vbs
@REM -- AND THEN STOP ON EROR NOT GOT ACCESS OF SOME KIND
@REM -- AND WHEN ICACLS RUN IT WITH IT PASS OVER ERROR OPTION SET
@REM -- IT GETS TO A FILE
@REM -- C:\Windows\System32\LogFiles\WMI\RtBackup\*.VBS: Access is denied.
@REM -- THAT INCLUDE A WILDCARD AND THAT WAS WHERE STOPPED AT FOR TAKEOWNERSHIP
@REM -- THE ICACLS SETOWNER HAS PROBLEM WITH FILE NOT BEEN ABLE TO BE SET EVEN WITH TAKEOWNER COMMAND
@REM -- AND THE ICACLS GRANT HAS PROBLEM WITH FILE AFTER THE WILDCARD ONE AS THEY NOT BEEN SET BY TAKEOWNER COMMAND
@REM -- 
@REM -- I FIND SETTER OWNER AND GRANT WITH ICACLS IS BETTER SET TO ADMISTRATIOR FOR SOME THING 
@REM -- RATHER THAN EVERYONE - EVERYONE DOESN'T GET ACCEPTED
@REM -- FOR 1ST FEW THIS SETTING "C:\Windows\System32\*.VBS"
@REM -- AND HERE IS STILL NOT ABLE TO SET DUE TO I THINK THE FOLDER IS DISALLOW
@REM -- C:\Windows\System32\Printing_Admin_Scripts\en-US\
@REM -- 
@REM -- TAKEOWNER COMMAND IS ONLY USED TO SET A PERSON YOUR CURRENT USER OR ADMIN
@REM -- 
@REM -- ONCE BEEN SET TO ADMIN YOU ARE THEN ABLE TO SET EVERYONE
@REM ----------------------------------------------------------------
@REM -- FROM Sun 14-Oct-2018 20:34:24
@REM -- FROM Sun 14-Oct-2018 22:45:00
@REM ----------------------------------------------------------------

@ECHO OFF
@REM ENDLOCAL
@CD\

@REM ----------------------------------------------------------------
@REM __ SET ADMIN A SAFE DOUBLE CHECKER WAY NOT A LOOPY ONE
@REM __ IF CAN'T GET IT TO GO ADMIN DOESN'T RSI INJURY ITSELF AND YOU
@REM ----------------------------------------------------------------

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

@ECHO OFF

SET TITLE_VAR=%~f0
TITLE %TITLE_VAR%

ECHO.

SETLOCAL EnableDelayedExpansion

REM -----------------------------------------------------------------
SET LDA=0
REM -----------------------------------------------------------------
SET /A LDA=LDA+1
SET Sources[!LDA!]="D:\VB6"
SET /A LDA=LDA+1
SET Sources[!LDA!]="D:\VB6-EXE"
SET /A LDA=LDA+1
SET Sources[!LDA!]="E:\01 Desktop"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\Windows\System32.VBS"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\Windows\System32\Printing_Admin_Scripts\en-US.VBS"
SET /A LDA=LDA+1
SET Sources[!LDA!]="D:\VI_ DSC ME"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\SCRIPTOR GITHUB"
SET /A LDA=LDA+1
SET Sources[!LDA!]="D:\DSC"
SET /A LDA=LDA+1
SET Sources[!LDA!]="D:\SNAPSHOT CCSE HIKVISION"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\Users\MATT 01\AppData"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\Users\MATT 02\AppData"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\Users\MATT 03\AppData"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\Users\MATT 04\AppData"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\SCRIPTER"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\DOWNLOADS"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\PStart"
SET /A LDA=LDA+1
SET Sources[!LDA!]="D:\# MY DOCS"
SET /A LDA=LDA+1
SET Sources[!LDA!]="D:\#0 1 INSTALLATIONS"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\BAT"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\Users\Public\Desktop"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\Program Files\WinRAR"
SET /A LDA=LDA+1
SET Sources[!LDA!]="D:\0 CLOUD"
SET /A LDA=LDA+1
SET Sources[!LDA!]="D:\0 00 MUSIC ---"
SET /A LDA=LDA+1
SET Sources[!LDA!]="C:\RF"
SET /A LDA=LDA+1
SET Sources[!LDA!]=""
REM -----------------------------------------------------------------

REM -----------------------------------------------------------------
REM -- SORT THE EMPTY ARRAY FOR VOID VALUE
REM -----------------------------------------------------------------
SET LDB=0
FOR /L %%a in (1,1,!LDA!) DO (
	IF NOT !sources[%%a]!=="" (
		SET /A LDB=LDB+1
		SET Sources_02[!LDB!]=!Sources[%%a]!
	)
)
REM -----------------------------------------------------------------

REM -----------------------------------------------------------------
REM -- PUT BACK IN
REM -----------------------------------------------------------------
FOR /L %%a in (1,1,!LDB!) DO (
	SET Sources[%%a]=!Sources_02[%%a]!
)
SET LDA=%LDB%
REM -----------------------------------------------------------------

REM -----------------------------------------------------------------
REM -- FORMAT FOLDER NAME WITH EXTRA \*
REM -----------------------------------------------------------------
FOR /L %%a in (1,1,!LDA!) DO (
	SET Sources[%%a]=!Sources[%%a]!\*
)
REM -----------------------------------------------------------------

FOR /L %%a in (1,1,!LDA!) DO (
	IF NOT "!sources[%%a]!"=="" (
		ECHO --------------------------------------------------------
		ECHO COMMAND INSTRUCTION POINTER %%a OF !LDA!
		ECHO --------------------------------------------------------
		SET PATH_TO_DO=!Sources[%%a]!
		ECHO PATH_TO_DO = !PATH_TO_DO:"=!
		ECHO --------------------------------------------------------

		SET VAR_FLAG=NOT
		ECHO !PATH_TO_DO!
		ECHO "E:\01 Desktop\*"
		IF !PATH_TO_DO!=="E:\01 Desktop\*" SET VAR_FLAG=YES
		IF !VAR_FLAG!==YES (
			SET COMMAND_VAR=TAKEOWN /F !PATH_TO_DO! /A /R /D Y
			ECHO !COMMAND_VAR!
			!COMMAND_VAR! >NUL
			REM -- TAKEOWN TO ADMIN AND MORE NON STOP AT ERROR WITH /C UNLIKE ABOVE
			REM -- ABLE SET TO EVERYONE AGAIN ONE ADMIN IS DONE BEFORE
		)
		
		SET COMMAND_VAR=ICACLS !PATH_TO_DO! /setowner Everyone /T /C /L /Q 
		ECHO !COMMAND_VAR!
		REM !COMMAND_VAR!
		SET COMMAND_VAR=ICACLS !PATH_TO_DO! /T /L /C /Q /Grant Everyone:F
		ECHO !COMMAND_VAR!
		REM !COMMAND_VAR!
		SET COMMAND_VAR=ICACLS !PATH_TO_DO! /T /L /C /Q  /grant administrators:F
		REM ECHO !COMMAND_VAR!
		REM !COMMAND_VAR!

		REM SET VAR_FLAG=NOT
		REM IF "!PATH_TO_DO!"==C:\SCRIPTOR GITHUB\* SET VAR_FLAG=YES
		REM IF !VAR_FLAG!==YES (
			REM SET COMMAND_VAR=ICACLS !PATH_TO_DO! /setintegritylevel (CI)H
			REM ECHO !COMMAND_VAR!
			REM !COMMAND_VAR!
		REM )
		REM __ 
		REM __ [ Friday 18:08:30 Pm_05 October 2018 ]
		REM __ SET /setintegritylevel HIGH _ IF LOW IT DISPLAY THE IN PROPERTIES AND SECURITY LOW MEAN LESS ACCESS
		REM __ POP UP DIALOG APPEAR FOR ACCESS A .BAT FOR IN FOLDER LIKE IT IS
		REM __ CI MEAN ALL FOLDER ND SUB-FOLDER
		REM ----
		REM iCacls - Modify Access Control List - Windows CMD - SS64.com
		REM https://ss64.com/nt/icacls.html
		REM ----
	)
)


PAUSE

EXIT


REM -----------------------------------------------------------------
REM GOOD CODE BUT NOT WANTER IN END
REM NORTON DOESN'T RESTRICT FROM LOOK FILE NAME IN 
REM C:\PROGRAM FILES FOLDER
REM -----------------------------------------------------------------
REM SETLOCAL EnableDelayedExpansion
REM SET "TARGET_="
REM SET TARGET_=C:\Program Files
REM ECHO !TARGET_!
REM FOR /f "delims=" %%F in ('Dir "!TARGET_!\" /b /ad') do (
		REM SET "VAR_FLAG="
		REM SET VAR_FLAG=NOT
		REM ECHO \%%F | FIND "\Norton"
		REM IF !errorlevel!==0 SET VAR_FLAG=YES
		REM IF "!VAR_FLAG!"=="YES" (
			REM TAKEOWN /F "!TARGET_!\%%F" /A /R /D Y
			REM ICACLS "!TARGET_!\%%F" /T /C /L /Q /setowner "Everyone" >NUL
			REM ICACLS "!TARGET_!\%%F" /T /C /L /Q /grant Everyone:F >NUL
		REM )
	REM )
REM )


SETLOCAL EnableDelayedExpansion
@FOR %%G IN (C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (
	IF EXIST %%G:\. (
		SET "TARGET_="
		SET TARGET_=%%G:
		CALL :IsDeviceReady !TARGET_! isready_
		IF !isready_!==true (
			 FOR /f "delims=" %%F in ('Dir !TARGET_!\ /b /a-d') do (
				SET "VAR_FLAG="
				SET VAR_FLAG=YES
				IF "%%F"=="System Volume Information" SET VAR_FLAG=NOT
				IF "!VAR_FLAG!"=="YES" (
					ICACLS "!TARGET_!\%%F" /C /L /Q /setowner "Everyone" >NUL
					ICACLS "!TARGET_!\%%F" /C /L /Q /grant Everyone:F >NUL
				)
			)
		)
	)
)

@FOR %%G IN (E) DO (
	IF EXIST %%G:\. (
		SET "TARGET_="
		SET TARGET_=%%G:
		CALL :IsDeviceReady !TARGET_! isready_
		IF !isready_!==true (

		FOR /f "delims=" %%F in ('Dir !TARGET_!\ /b /ad') do (
				SET "VAR_FLAG="
				SET VAR_FLAG=YES
				IF "%%F"=="System Volume Information" SET VAR_FLAG=NOT
				IF "%%F"=="$Recycle.Bin" SET VAR_FLAG=NOT
				IF "%%F"=="$RECYCLE.BIN" SET VAR_FLAG=NOT					
				IF "%%F"=="Windows" SET VAR_FLAG=NOT				
				IF "%%F"=="Recovery" SET VAR_FLAG=NOT
				IF "%%F"=="EFI" SET VAR_FLAG=NOT
				IF "%%F"=="ProgramData" SET VAR_FLAG=NOT
				IF "%%F"=="Boot" SET VAR_FLAG=NOT
				IF "%%F"=="Documents and Settings" SET VAR_FLAG=NOT
				IF "!VAR_FLAG!"=="YES" (
					ICACLS "!TARGET_!\%%F" /T /C /L /Q /setowner "Everyone" >NUL
					ICACLS "!TARGET_!\%%F" /T /C /L /Q /grant Everyone:F >NUL
				)
			)
		)
	)
)



@REM START "" /D!PATH_TO_DO! /B /MIN /HIGH /WAIT ICACLS !PATH_TO_DO! /setowner Everyone /T /C /L /Q >NUL
@REM START "" /D!PATH_TO_DO! /B /MIN /HIGH /WAIT ICACLS !PATH_TO_DO! /T /L /C /Q /Grant Everyone:F >NUL


ECHO.
ECHO DONE

IF "%1"=="" PAUSE
GOTO ENDE


GOTO VERY_ENDE


::
:IsDeviceReady
SETLOCAL
set "ready_="
set ready_=true
dir "%~1" > nul 2>&1
if %errorlevel% NEQ 0 set ready_=false
endlocal & set "%2=%ready_%" & goto :EOF

::
:GetLabel
SETLOCAL
set "label_="
for /f "tokens=5*" %%a in (
  'vol "%~1"^|find "Volume in drive "') do (
    set label_=%%b)
endlocal & set "%2=%label_%" & goto :EOF




@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/--------------------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/--------------------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@REM ----------------------------------------------------------------------------
@REM -- THE END AND REFERENCE SECTION -------------------------------------------
@REM -- BYE BYE -----------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/--------------------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/--------------------\/\/\/\/\/\/\/\/\/\/\/\/\/\



:VERY_ENDER_BLOCK

:VERY_ENDE

:ENDE

EXIT