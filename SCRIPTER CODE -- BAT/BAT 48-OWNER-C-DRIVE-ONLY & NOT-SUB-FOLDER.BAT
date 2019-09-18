REM ---------------------------------------------------------------- 
REM I LIKE THIS CODE GIVE ACCESS TO ALL FOLDER QUICKER
REM HAVE A READER READ
REM NOT SET OWNERSHIP ON IMPORTANT FOLDER
REM SO NOT CORRUPT RECLEBIN OR ANY
REM IT DO ALL DRIVE NOT ONLY C -- BUTTY C WAS BASTARD
REM EASY LOOP SET WHICH ONE OR ALL - DO TICKET
REM ---------------------------------------------------------------- 
REM [ Tuesday 22:44:50 Pm_17 September 2019 ]
REM ---------------------------------------------------------------- 


@CD\

REM ---------------------------------------------------------------- 
REM __ SET ADMIN A SAFE DOUBLE CHECKER WAY NOT A LOOPY ONE
REM __ IF CAN'T GET IT TO GO ADMIN DOESN'T RSI INJURY ITSELF AND YOU
REM ---------------------------------------------------------------- 

IF '%1%'=='/ADMIN' (
	REM :: Check for ADMIN Privileges Again
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

REM :: Check for ADMIN Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
@IF '%errorlevel%' GTR '2' (
	REM Get ADMIN Privileges
	@ECHO Set UAC = CreateObject^("Shell.Application"^) >> "%temp%\getadmin.vbs"
	@ECHO UAC.ShellExecute "%~s0", "/ADMIN", "", "runas", 1 >> "%temp%\getadmin.vbs"

	"%temp%\getadmin.vbs"
	@DEL "%temp%\getadmin.vbs"
	GOTO ENDE
	@EXIT /B
)

@ECHO Got ADMIN Privileges




@ECHO.
@ECHO ---------------------------------------
@ECHO ---------------------------------------

@ECHO QUICK ROOT NOT SUB FOLDER OWNER SET NOT GRANT ACCESS PART


@ECHO OFF
setlocal EnableDelayedExpansion
CD\

SET "INPUT_VAR="
REM ---------------------

REM GOTO JUMP_PART0

SET INPUT_VAR="C:\ProgramData\Microsoft\Windows"
ECHO.
ECHO !INPUT_VAR!
ECHO ICACLS __ !INPUT_VAR!
ECHO ---------------------------------------------------------------------------
REM __ TAKEOWN IS THE SAME AS GET GRANT ACCESS EVERYONE
REM __ TAKEOWN /F !INPUT_VAR! /R /A /D Y >NUL
ICACLS !INPUT_VAR! /C /L /Q /setowner "Everyone" >NUL
ICACLS !INPUT_VAR! /C /L /Q /grant Everyone:F >NUL
REM TIMEOUT /T 2

REM REM ___________ MS WINDOW 10
REM SET INPUT_VAR="C:\ProgramData\Microsoft\Windows\Start Menu Places"
REM ECHO.
REM ECHO !INPUT_VAR!
REM ECHO ICACLS __ !INPUT_VAR!
REM ECHO ---------------------------------------------------------------------------
REM REM __ TAKEOWN IS THE SAME AS GET GRANT ACCESS EVERYONE
REM REM __ TAKEOWN /F !INPUT_VAR! /R /A /D Y >NUL
REM ICACLS !INPUT_VAR! /T /C /L /Q /setowner "Everyone" >NUL
REM ICACLS !INPUT_VAR! /T /C /L /Q /grant Everyone:F >NUL
REM ECHO.

TIMEOUT /T 2
ECHO.


REM -------------------------------------------------------------------------------------------------
REM -------------------------------------------------------------------------------------------------

:JUMP_PART0
REM GOTO JUMP_PART2
REM -------------------------------------------------------------------------------------------------

REM @FOR %%G IN (C) DO (
@FOR %%G IN (C,D,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (
	IF EXIST %%G:\. (
		SET "target_="
		SET target_=%%G:
		CALL :IsDeviceReady !target_! isready_
		IF !isready_!==true (
			REM __ /ad MISSING __ TO INCLUDE ALL __ FILE AND DIRECTORY
			FOR /f "delims=" %%F in ('Dir !target_!\ /b') do (
				SET "VAR_FLAG="
				SET VAR_FLAG=YES
				IF "%%F"=="System Volume Information" SET VAR_FLAG=NOT
				IF "%%F"=="$Recycle.Bin" SET VAR_FLAG=NOT
				IF "%%F"=="$RECYCLE.BIN" SET VAR_FLAG=NOT
				IF "%%F"=="Recovery" SET VAR_FLAG=NOT
				IF "%%F"=="Windows" SET VAR_FLAG=NOT
				IF "%%F"=="ProgramData" SET VAR_FLAG=NOT
				IF "%%F"=="Boot" SET VAR_FLAG=NOT
				REM __ Documents and Settings __ A JUNCTION FOR C:\USERS
				REM IF "%%F"=="Documents and Settings" SET VAR_FLAG=NOT
				IF "!VAR_FLAG!"=="YES" (
						ECHO.
						
						REM __ TAKEOWNER IS THE SAME AS GET GRANT ACCESS EVERYONE
						REM TITLE #01 OF #03 __ TAKEOWN ___________ "%%G:\%%F"
						REM ECHO "%%G:\%%F"
						REM ECHO #01 OF 03 __ TAKEOWN /F "%%G:\%%F" /A /D N 
						REM REM /R RECURSIVE MISSING
						REM ECHO ---------------------------------------------------------------------------
						REM TAKEOWN /F "%%G:\%%F" /A >NUL
						REM REM TIMEOUT /T 2
						REM ECHO.
						
						TITLE #02 OF #03 __ ICACLS /SETOWNER __ "%%G:\%%F"
						ECHO "%%G:\%%F"
						ECHO #01 OF 02 __ ICACLS "%%G:\%%F" /setowner "administrators"
						ECHO #01 OF 02 __ ICACLS "%%G:\%%F" /setowner "Everyone"
						ECHO ---------------------------------------------------------------------------
						REM GOT TO DO ADMIN 1ST BEFORE EVERYONE
						
						ICACLS "%%G:\%%F" /C /L /Q /setowner "administrators" >NUL
						ICACLS "%%G:\%%F" /C /L /Q /setowner "Everyone" >NUL
						REM TIMEOUT /T 2
						REM /T RECURSIVE MISSING
						ECHO.
						
						TITLE #03 OF #03 __ ICACLS /GRANT _____ "%%G:\%%F"
						ECHO "%%G:\%%F"
						ECHO #02 OF 02 __ ICACLS "%%G:\%%F" /grant Everyone:F
						ECHO ---------------------------------------------------------------------------
						REM GRANT STILL DOES SUB FOLDER SO HAVE TO LEAVE OUT
						REM ----------------------------------------------------------------------------
						REM ICACLS "%%G:\%%F" /C /L /Q /grant administrators:F >NUL
						REM ICACLS "%%G:\%%F" /C /L /Q /grant Everyone:F >NUL
						REM TIMEOUT /T 2
						ECHO.
				)
			)
		)
	)
)

REM -------------------------------------------------------------------------------------------------
REM -------------------------------------------------------------------------------------------------

:JUMP_PART2


REM PAUSE



GOTO VERY_ENDE


::
:IsDeviceReady
setlocal
set "ready_="
set ready_=true
dir "%~1" > nul 2>&1
if %errorlevel% NEQ 0 set ready_=false
endlocal & set "%2=%ready_%" & goto :EOF




PAUSE
GOTO VERY_ENDE


@ECHO BEGIN -- 01 OF 02 -- SET OWNER EVERYONE
@ECHO NEXT --- 02 OF 02 -- GRANT EVERYONE
@ECHO Q FOR QUITE ABOUT SUCCESS MESSAGING
@ECHO ---------------------------------------
@ECHO ---------------------------------------
@ECHO.

@REM -- OPTIONAL REM OUT PAUSE
@REM -- SEARCH

@REM ICACLS * /t /Q /grant Everyone:F >NUL
@REM ICACLS * /t /Q /grant Everyone:F >NUL
@REM @PAUSE
@REM ICACLS * /setowner "Everyone" /T /C /L /Q >NUL
@REM ICACLS * /t /Q /grant Everyone:F >NUL

@REM ICACLS "C:\Users"  /t /Q /grant Everyone:F >NUL


@REM ----------------------------------------------------------------------------------------------
@REM -- Iterate through...

@REM -- ...files in current dir: for %f in (.\*) do @echo %f
@REM -- ...subdirs in current dir: for /D %s in (.\*) do @echo %s
@REM -- ...files in current and all subdirs: for /R %f in (.\*) do @echo %f
@REM -- ...subdirs in current and all subdirs: for /R /D %s in (.\*) do @echo %s
@REM -- ----
@REM -- windows - Iterate all files in a directory using a for-loop - Stack Overflow
@REM -- http://stackoverflow.com/questions/138497/iterate-all-files-in-a-directory-using-a-for-loop
@REM -- ----
@REM -- Fri 26 August 2016 15:56:00
@REM ----------------------------------------------------------------------------------------------


@REM ----------------------------------------------------------------------------------------------
@REM set "var1="
@REM Fri 26-Aug-2016 16:33:22
@REM ---------------------
@REM Form FindWindow ---
@REM batch file - how to clear a variable which is set from command line by "set /p var1=" - Stack Overflow - Google Chrome
@REM ---------------------
@REM ----
@REM batch file - how to clear a variable which is set from command line by "set /p var1=" - Stack Overflow
@REM http://stackoverflow.com/questions/17106181/how-to-clear-a-variable-which-is-set-from-command-line-by-set-p-var1
@REM ----
@REM ----------------------------------------------------------------------------------------------


@REM ----------------------------------------------------------------------------------------------
@REM -- setlocal EnableDelayedExpansion
@REM -- ----
@REM -- cmd - windows batch SET inside IF not working - Stack Overflow
@REM -- http://stackoverflow.com/questions/9102422/windows-batch-set-inside-if-not-working
@REM -- ----
@REM -- Fri 26 August 2016 16:43:59
@REM ----------------------------------------------------------------------------------------------

@REM ----------------------------------------------------------------------------------------------
@REM -- ----
@REM -- windows - Iterate all files in a directory using a for-loop - Stack Overflow
@REM -- http://stackoverflow.com/questions/138497/iterate-all-files-in-a-directory-using-a-for-loop
@REM -- ----
@REM -- "delims=" is useful to show long filenames with spaces in it....

@REM -- '/b" show only names, not size dates etc..

@REM -- Some things to know about dir's /a argument.

@REM -- Any use of "/a" would list everything, including hidden and system attributes.
@REM -- "/ad" would only show subdirectories, including hidden and system ones.
@REM -- "/a-d" argument eliminates content with 'D'irectory attribute.
@REM -- "/a-d-h-s" will show everything, but entries with 'D'irectory, 'H'idden 'S'ystem attribute.
@REM -- If you use this on the commandline, remove a "%".
@REM -- Fri 26 August 2016 16:50:28
@REM ----------------------------------------------------------------------------------------------


@ECHO OFF
setlocal EnableDelayedExpansion



@REM -- C:
@REM -- CD\

@REM ----------------------------------------------------------------
@REM for /D %%F in (\*) do (
@REM -- EXAMPLE USAGE HERE -- 	IF "%%F"=="\Windows" SET VAR_FLAG=YES
@REM -- KEEP THE \
@REM ----------------------------------------------------------------

ECHO ----------------------------------------------------------------------------
ECHO ---- PART 01 OF 07 -- C:\ ---- ALL ALONG THE ROOT -- NOT DIGGER SUBFOLDERING
ECHO ----------------------------------------------------------------------------

SET "TESTVAR="
@REM SET TESTVAR=YES

for /f "delims=" %%F in ('Dir C:\ /b /ad') do (

	Set "VAR_FLAG="
	SET VAR_FLAG=YES
	@REM -- SET INCLUDE ALL UNLESS NEXT LEARN

	@REM -- EXCLUDE HERE
	IF "%%F"=="System Volume Information" SET VAR_FLAG=NOT
	IF "%%F"=="$Recycle.Bin" SET VAR_FLAG=NOT
	IF "%%F"=="Windows" SET VAR_FLAG=NOT
	IF "%%F"=="Users" SET VAR_FLAG=NOT
	IF "%%F"=="$Windows.~WS" SET VAR_FLAG=NOT
	IF "%%F"=="ProgramData" SET VAR_FLAG=NOT
	IF "%%F"=="PerfLogs" SET VAR_FLAG=NOT
	IF "%%F"=="Boot" SET VAR_FLAG=NOT
	IF "%%F"=="Documents and Settings" SET VAR_FLAG=NOT
	IF "%%F"=="Recovery" SET VAR_FLAG=NOT
	IF "%%F"=="SC" SET VAR_FLAG=NOT

	REM -- WIN 10
	IF "%%F"=="$WINDOWS.~BT" SET VAR_FLAG=NOT
	IF "%%F"=="aws" SET VAR_FLAG=NOT
	IF "%%F"=="Windows.old" SET VAR_FLAG=NOT
	IF "%%F"=="eSupport" SET VAR_FLAG=NOT

	IF "!VAR_FLAG!"=="YES" (
		IF "!TESTVAR!"=="" (
			
			
			SET "PATH_VAR=" 
			SET PATH_VAR=C:\%%F
			call :Set_ICACLS "!PATH_VAR!"
			
			
			)
		ECHO --------------------------------------------------------------------
		ECHO -- Do All in ----
		ECHO -----------------
		ECHO -- C:\%%F
		ECHO --------------------------------------------------------------------
	)
)


GOTO VERY_ENDE

pause

SET "TESTVAR="
SET TESTVAR=YES

ECHO ----------------------------------------------------------------------------
ECHO ---- PART xxxx xxxx -- C:\ ROOT NOT SUB FOLDER
ECHO ----------------------------------------------------------------------------

IF "!TESTVAR!"=="" (
	REM -- SET "PATH_VAR=" 
	REM -- SET PATH_VAR="C:\"
	REM -- call :Set_ICACLS !PATH_VAR!
	
	
	REM ------------------------------------------------------------------------------
	REM -- LOSE THE /T OPTION OR WILL DO SUB-FOLDER
	REM -- WHEN ROOT IS DONE HERE WILL DO IN C: ALL FILE AND FOLDER BUT NOT SUB FOLDER
	REM -- SOME FEW A RESTRICTED
	REM -- DOES NOT APPEAR TO CAUSE PROBLEM LIKE -- START MENU NOT FUNCTION
	REM --
	REM -- BIT ANNOYING THIS METHOD MIGHT DO RECYCLE-BIN AND DAMAGE IT -- LIKE IT DOES
	REM -- MAYBE PROCESS EACH FILE
	REM ------------------------------------------------------------------------------
	REM -- MAYBE USE AS QUICK OPTION SET ALL ROOT LEVEL FOLDER
	REM -- IT IS OKAY ON C DRIVE NOT WHOLE DRIVE
	REM -- ON C DRIVE IF WANT ACCESS TO PROGRAM FILES
	REM -- IT WILL BE TRUSTED INSTALLED BUT THAT DON'T GIVE MUCH ACCESS
	REM -- DO IT MANUALLY -- SET OWNER EVERYONE OR SOMETHING
	REM -- WHEN OWNER IS EVERYONE -- THEN SET EVERYONE HAS ACCESS
	REM -- IDEA WAS SET OWNER EVERYONE AND THEN RUN THIS PROGRAM AGAIN LIKE NOT ON ALL SUBFOLDER
	REM -- AS TIME TAKING
	REM -- IS ABLE TO DO THAT WHEN SET OWNER EVERYONE MANUALLY
	REM -- BUT AND NOT WHEN GIVE EVERYONE PERMISSION TO ACCESS ALL THAT WILL TRY ALL SUB FOLDER
	REM -- GIVE IT A GO NOT MANY TO DO IN PROGRAM FILES
	REM ------------------------------------------------------------------------------
		
	START "C:\ -- ICACLS 01 OF 02 _____OWNERSHIP IN PROGRESS 02/10/2016--13:02:50.32" /HIGH /WAIT ICACLS "C:\*" /C /L /Q /setowner "Everyone"
	START "C:\ -- ICACLS 02 OF 02 _GRANTEVERYONE IN PROGRESS 02/10/2016--13:02:50.32" /HIGH /WAIT ICACLS "C:\*" /C /L /Q /grant Everyone:F
)

ECHO ----------------------------------------------------------------------------
ECHO ---- PART 02 OF 07 -- C:\USERS\All Users ---- Do All In -- All Users
ECHO ----------------------------------------------------------------------------

IF "!TESTVAR!"=="" (
	SET "PATH_VAR=" 
	SET PATH_VAR="C:\USERS\All Users"
	call :Set_ICACLS !PATH_VAR!
)
ECHO --------------------------------------------------------------------
ECHO ---- PART 03 OF 07 -- C:\USERS\All Users ---- Do All In -- Public
ECHO --------------------------------------------------------------------
IF "!TESTVAR!"=="" (
	SET "PATH_VAR="
	SET PATH_VAR="C:\USERS\Public"
	call :Set_ICACLS !PATH_VAR!
)
ECHO --------------------------------------------------------------------
ECHO ---- PART 04 OF 07 -- C:\USERS\Matt 01  ---- Do All In -- My User
ECHO --------------------------------------------------------------------
IF "!TESTVAR!"=="" (
	SET "PATH_VAR="
	SET PATH_VAR="C:\USERS\MATT 01"
	call :Set_ICACLS !PATH_VAR!
)


@REM -- ----
@REM -- windows - Iterate all files in a directory using a for-loop - Stack Overflow
@REM -- http://stackoverflow.com/questions/138497/iterate-all-files-in-a-directory-using-a-for-loop
@REM -- ----	
@REM -- If you would like to echo only the filenames (not the full path) with their extensions in the current directory @REM -- (recursively), you can do it like this: for /r %i in (*) do ( echo %--^^~--nxi ).

@REM for /R /D %%F in (*) do (


@REM -- @setlocal enableextensions enabledelayedexpansion
@REM -- @echo off
@REM -- set str1=%1
@REM ---------- if not x%str1:bcd=%==x%str1% echo It contains bcd
@REM -- endlocal
@REM -- ----
@REM -- windows - Batch file: Find if substring is in string (not in a file) - Stack Overflow
@REM -- http://stackoverflow.com/questions/7005951/batch-file-find-if-substring-is-in-string-not-in-a-file
@REM -- ----
@REM -- Fri 26 August 2016 18:33:05

@REM -- For compatibility and ease of use it's always best to use FIND to do this, as you may be looking to match 
@REM -- output from a command, a variable in a loop, or a pointer variable in your Batch/CMD.
@REM -- ----
@REM -- ECHO.%Variable% | FIND /I "ABC">Nul && ( Echo Found "ABC" ) || ( Echo Did not find "ABC" )
@REM -- Fri 26 August 2016 18:45:07

@REM -- /B -- BARE NOT HEADER INFO
@REM -- /L -- Lower Case
@REM -- /S -- SUB FOLDERING
@REM -- /AD DIR ONLY AND HIDDEN SYS

@REM -- ----
@REM -- 2ND ATTEMPT AT THIS METHOD SYNTAX CORRECT OR SOMETHING
@REM -- ----
@REM -- windows - FORFILES: Only delete file if RELPATH contains substring - Stack Overflow
@REM -- http://stackoverflow.com/questions/33570132/forfiles-only-delete-file-if-relpath-contains-substring
@REM -- ----
@REM -- set VAR=@relpath
@REM -- if not #!VAR:%search%=!==#!VAR! (
@REM -- ----
@REM -- set VAR=@relpath
@REM -- if not #!VAR:%search%=!==#!VAR! (
@REM -- ----
@REM -- if not x%str1:bcd=%==x%str1% echo It contains bcd
@REM -- ----
@REM -- Fri 26 August 2016 20:09:36
@REM -- ----

@REM ---------------------------------------------------
@REM -- Set "str1="
@REM -- Set str1=xbcdxx22
@REM -- if not x%str1:bcdxx22=%==x%str1% echo It contains bcdxx22
@REM -- Set "str1="
@REM -- Set str1="xxstart menu--xx"
@REM -- Set "Search="
@REM -- Set Search="start menu--"
@REM -- if not x%str1:"start"=%==x%str1% echo It contains \start menu--
@REM ---------------------------------------------------
@REM -- CANNOT DO THIS WITH A SPACE IN THE STRING -- BLOCK ABOVE -- BACK TO FIND COMMAND
@REM -- MUST LEARN THIS
@REM -- TAKEN ALL DAY
@REM -- FIND A WEB SITE WITH A LOT OF DETAIL ABOUT DIR COMMAND
@REM -- ----
@REM -- /E --- OOPS OUCH HOUR HURT -- MUST BE SOMETHING ELSE /E DON'T WORK MICROSOFT
@REM -- ONE I FIND HERE IS SET DIR OUTPUT TO UPPERCASE NOT ONLY LOWER
@REM -- DIR - Display information about files & subdirectories
@REM -- https://jpsoft.com/help/dir.htm
@REM -- ----
@REM -- Fri 26 August 2016 21:35:54
@REM -- ALSO I THINK MICROSOFT MESSED UP BECAUSE OF ALL THE THING WHEN READ
@REM -- IT DOES NOT LET YOU SEARCH FOR THE LAST PART OF THE PATH
@REM -- TURN TO VISUAL BASIC 6 ANSWER
@REM -- AND ANOTHER WHEN SEARCH CERTAIN FILE NAME GUESS BY WORD SELECT IT DOES NOT INCLUDE SUB-FOLDER NAME
@REM -- STILL MAKE THIS WORK WITH A BIT OF WILD CARD DIR
@REM ---------------------------------------------------

@Rem -- -----------------------------------------------------
@Rem -- Code With Here TO ENTER PASTE ON COMMAND LINE 
@Rem -- "C:\BAT\OWNER - MY C DRIVE ALL"
@Rem -- Use Quote or Another Code Same Name Smaller Might Run
@Rem -- -----------------------------------------------------

ECHO --------------------------------------------------------------------
ECHO ---- PART 05 OF 07 -- DIR C:\USERS\STARTMENU -- WITH RECURSE SUBFOLDER LOOKING FOR \STARTMENU
ECHO --------------------------------------------------------------------
@REM ---- VERIFY OUTPUT FOR YOURSELF BEFORE RUN IT OVER -- GIVING IT A GO
@REM ---- DIR "C:\USERS\START MENU" /S /B /AD

@REM -- C:
@REM -- CD \USERS

@REM -- TAKE LOWER CASE USAGE SETTING OFF /L BECAUSE HURT MY HAND EDITING
@REM --------------------------------------------------------------------

SET "TESTVAR="
@REM IF "!TESTVAR!"=="" (

for /f "delims=" %%F in ('DIR "C:\USERS\START MENU" /S /B /AD') do (
	
	@REM Set "VAR_FLAG="
	@REM -------------------------------------------------------------------------------------------
	@REM -- INTRO HERE AGAIN -- RESULT USE WILD CARD SEARCH BETTER
	@REM -------------------------------------------------------------------------------------------
	@REM -- FLAG WANT ON
	ECHO.%%F | FIND /I "\Start Menu">Nul && Set "VAR_FLAG=" && SET VAR_FLAG=YES
	@REM -- TEST -- IF "!VAR_FLAG!"=="YES" ECHO %%F

	@REM -------------------------------------------------------------------------------------------
	@REM -- FLAG WANT EXCLUDE OFF -- PRIORITY ORDER
	ECHO.%%F | FIND /I "\USERS\Default">Nul && Set "VAR_FLAG=" && SET VAR_FLAG=NOT
	@REM -- FLAG WANT EXCLUDE OFF -- PRIORITY ORDER -- SOME TO LEAVE ALONE
	@REM -- NOT MY NEW USER CHANGER YET
	ECHO.%%F | FIND /I "\USERS\MATT 02">Nul && Set "VAR_FLAG=" && SET VAR_FLAG=NOT
	@REM -------------------------------------------------------------------------------------------
	IF !VAR_FLAG!==YES (
		IF "!TESTVAR!"=="" (
			ECHO --------------------------------------------------------------------
			ECHO -- Do All in ----
			ECHO -----------------
			ECHO -- %%F
			ECHO --------------------------------------------------------------------
			SET "PATH_VAR="
			SET PATH_VAR=%%F
			call :Set_ICACLS !PATH_VAR!
		)
	)

	@REM -------------------------------------------------------------------------------------------
	@REM -------------------------------------------------------------------------------------------
	@REM -- Set "FILE_NAME_VAR="
	@REM -- Set FILE_NAME_VAR=%%F
	@REM -- Set "VAR_COMPARE="
	@REM -- Set VAR_COMPARE="Start Menu"
	@REM -- if not x%%F:\StartMenu--=%==x%FILE_NAME_VAR% SET VAR_FLAG=YES
	@REM -- NOT WORKING WELL 
	@REM -------------------------------------------------------------------------------------------
	
	@REM -- if not #!VAR:%search%=!==#!VAR! (
	
	@REM -- ---------------------------------------------
	@REM -- ---------------------------------------------
	@REM -- Set "Var_Path="
	@REM -- Set Var_Path=%%F
	@REM -- Set "Search="
	@REM -- Set Search="\start menu--"
	@REM -- ---------------------------------------------
	@rem echo %Var_Path%
    @rem if not x%varpath:sta=%==x%varpath% echo job done
	@rem IMPOSSIBLE TRY NEW METHOD
	@REM -- ---------------------------------------------
    @REM -- if not x!Var_Path:sta=%==x!Var_Path! SET VAR_FLAG=YES
	@REM -- ---------------------------------------------
	@REM -- ---------------------------------------------

	@Rem SET VAR_FLAG=YES

	@Rem -- 	Set "Search="
	@rem -- 	Set Search="\desktop--"
	@rem --     if not #!VAR_PATH:Search=!==#!VAR_PATH! ( SET VAR_FLAG=YES )
	@rem -- 	Set "Search="
	@rem -- 	Set Search="\all users--"
	@rem --     if not #!VAR_PATH:Search=!==#!VAR_PATH! ( SET VAR_FLAG=YES )
	@rem -- 	Set "Search="
	@rem -- 	Set Search=":\users\matt 01--"
	@rem --     if not #!VAR_PATH:Search=!==#!VAR_PATH! ( SET VAR_FLAG=YES )
	@rem -- 	Set "Search="
	@rem -- 	Set Search=":\users\public--"
	@rem --     if not #!VAR_PATH:Search=!==#!VAR_PATH! ( SET VAR_FLAG=YES )
	
	
	@REM -------------------------------------------------------------------------------------------
	@REM -- ----
	@REM -- windows - Batch file: Find if substring is in string (not in a file) - Stack Overflow
	@REM -- http://stackoverflow.com/questions/7005951/batch-file-find-if-substring-is-in-string-not-in-a-file
	@REM -- ----
	@REM -- Fri 26 August 2016 18:33:05
	@REM -- ABORT BLOCK ATTEMPT READ WEB PAGE BLOCK EXPLAIN DEVELOP DOWN TO GET TO GOOD BIT
	@REM -------------------------------------------------------------------------------------------

	@REM -------------------------------------------------------------------------------------------
	@REM -- THE DOUBLE DASH AT END HELP WITH END IF LINE TERMINATOR 
	@REM -- SO END PATH NOT REPEAT RECURSIVELY THIS JOB
	@REM -- BEST ALSO PUT A SLASH IN FRONT
	@REM -------------------------------------------------------------------------------------------
	
	@REM -------------------------------------------------------------------------------------------
	@REM -- TIRED OF THIS ONE TAKE TOO LONG -- CALL OUT TO FIND COMMAND -- CPU USAGE
	@REM -- Fri 26 August 2016 20:06:53
	@REM -------------------------------------------------------------------------------------------
	@REM -- ECHO.%%F-- | FIND /I "\Start Menu--">Nul && ( SET VAR_FLAG=YES)
	@REM -- ECHO.%%F-- | FIND /I "\DeskTop--">Nul && ( SET VAR_FLAG=YES )
	@REM -- ECHO.%%F-- | FIND /I "\All Users--">Nul && ( SET VAR_FLAG=YES )
	@REM -- ECHO.%%F-- | FIND /I ":\Users\MATT 01--">Nul && ( SET VAR_FLAG=YES )
	@REM -- ECHO.%%F-- | FIND /I ":\Users\Public--">Nul && ( SET VAR_FLAG=YES )
	@REM -------------------------------------------------------------------------------------------

	
	@REM -------------------------------------------------------------------------------------------
	@REM -------------------------------------------------------------------------------------------

	
	@REM ----------------------------------------------------------------------------------
	@REM -- \DEFAULT\ -- NOT MESS WITH 
	@REM -- \Templates\ -- MAYBE LATER -- NONE IN-IT EMPTY 
	@REM -- MINE ALREADY BEEN RUN THROUGH EVERYONE PERMISSION
	@REM -- NOT A DESKTOP IN C:\ProgramData -- ONLY A START MENU
	@REM ----------------------------------------------------------------------------------
	@REM -- REMEMBER THE RISK ON WINDOW 10 IS SET PERMISSION EVERYONE ALL C DRIVE THEN START MENU WILL STOP 
	@REM -- SYSTEM RESTORE ONLY OPTION
	@REM ----------------------------------------------------------------------------------
	
	@REM -- IF "!VAR_FLAG!"=="YES" ECHO %%F

	@REM ----------------------------------------------------------------------------------
	@REM ----------------------------------------------------------------------------------
	@REM --- END OF ROUTINE LOOP -- COMMENT BELOW WRAP BACK IN
	@REM ----------------------------------------------------------------------------------
	
	@REM ----------------------------------------------------------------------------------
	@REM ----------------------------------------------------------------------------------
	@REM ----------------------------------------------------------------------------------
	@REM ----------------------------------------------------------------------------------
	@REM Set FILE_NAME_VAR=%%F%^~nxf
	@REM Set FILE_NAME_VAR=%%F%^~n
	
	@REM IF "!VAR_FLAG!"=="YES" ECHO %%F
	@REM ECHO !VAR_FLAG! -- 
	@REM --------------
	@REM -- ECHO %~nx2F
	@REM --------------
	@REM -------------- LAST OF THE CODE EFFORT ABOVE LINE
	@REM --------------
	@REM ECHO %FILE_NAME_VAR% --
	@REM ECHO %%F
	@REM 	-- !VAR_FLAG! -- %%F -- 
	@REM ----------------------------------------------------------------------------------

	@REM ----------------------------------------------------------------------------------
	@REM ----------------------------------------------------------------------------------
	@REM -- GIVE UP ON THIS BLOCK OF CODE AS CAN GET GET FILENAME ONLY BUT NOT LAST OF PATH
	@REM -- HOUR HURT SPEND ON LEARN OUCH
	@REM ----------------------------------------------------------------------------------
	@REM -- AFTER READ HERE
	@REM -- ----
	@REM -- Parameters / Arguments | Windows CMD | SS64.com
	@REM -- http://ss64.com/nt/syntax-args.html
	@REM -- ----
	@REM -- STACK OVERFLOW LOST IT ON CLEAR 
	@REM -- ----
	@REM -- windows - Iterate all files in a directory using a for-loop - Stack Overflow
	@REM -- http://stackoverflow.com/questions/138497/iterate-all-files-in-a-directory-using-a-for-loop
	@REM -- --------
	@REM -- What does %~d0 mean in a Windows batch file? - Stack Overflow
	@REM -- http://stackoverflow.com/questions/112055/what-does-d0-mean-in-a-windows-batch-file
	@REM -- ----
	@REM ----------------------------------------------------------------------------------
	@REM -- Fri 26 August 2016 18:29:14
	
)


ECHO --------------------------------------------------------------------
ECHO ---- PART SIX 06 OF 07 -- DIR C:\USERS\DESKTOP ---- WITH RECURSIVE SUB-FOLDER LOOKING FOR \DESKTOP
ECHO --------------------------------------------------------------------
@REM ---- VERIFY OUTPUT FOR YOURSELF BEFORE RUN IT OVER -- GIVING IT A GO
@REM ---- DIR "C:\USERS\DESKTOP" /S /B /AD

@REM -- TWO COMMAND ONE LINE
@REM -- ----
@REM -- batch file - How to run two commands in one line in Windows CMD? - Stack Overflow
@REM -- http://stackoverflow.com/questions/8055371/how-to-run-two-commands-in-one-line-in-windows-cmd
@REM -- ----


@REM -- TAKE LOWER CASE USAGE SETTING OFF /L BECAUSE HURT MY HAND EDITING
@REM --------------------------------------------------------------------
@REM ---- VERIFY OUTPUT FOR YOURSELF BEFORE RUN IT OVER -- GIVING IT A GO
@REM ---- DIR "C:\USERS\START MENU" /S /B /AD

@REM -- C:
@REM -- CD \USERS

@REM --------------------------------------------------------------------
@REM PART SIX 06 OF 07  -- DIR C:\USERS\DESKTOP
@REM --------------------------------------------------------------------

SET "TESTVAR="
@REM IF "!TESTVAR!"=="" (

for /f "delims=" %%F in ('DIR "C:\USERS\DESKTOP" /S /B /AD') do (
	
	
	@REM -------------------------------------------------------------------------------------------
	@REM -- FLAG WANT ON
	@REM -------------------------------------------------------------------------------------------
	ECHO.%%F | FIND /I "\Desktop">Nul && Set "VAR_FLAG=" && SET VAR_FLAG=YES
	@REM -------------------------------------------------------------------------------------------
	@REM -- FLAG WANT EXCLUDE OFF -- PRIORITY ORDER
	ECHO.%%F | FIND /I "\USERS\Default">Nul && Set "VAR_FLAG=" && SET VAR_FLAG=NOT
	@REM -- FLAG WANT EXCLUDE OFF -- PRIORITY ORDER -- SOME TO LEAVE ALONE
	@REM -- NOT MY NEW USER CHANGER YET
	ECHO.%%F | FIND /I "\USERS\MATT 02">Nul && Set "VAR_FLAG=" && SET VAR_FLAG=NOT
	@REM -------------------------------------------------------------------------------------------

	IF !VAR_FLAG!==YES (
		IF "!TESTVAR!"=="" (
			ECHO --------------------------------------------------------------------
			ECHO -- Do All in ----
			ECHO -----------------
			ECHO -- %%F
			ECHO --------------------------------------------------------------------
			SET "PATH_VAR="
			SET PATH_VAR=%%F
			call :Set_ICACLS !PATH_VAR!
		)
	)
)

ECHO ---- PART 07 OF 07 -- DIR C:\PROGRAMDATA\START MENU* ---- WITH RECURSE SUBFOLDER LOOKING FOR \DESKTOP
@REM --------------------------------------------------------------------
@REM ---- VERIFY OUTPUT FOR YOURSELF BEFORE RUN IT OVER -- GIVING IT A GO
@REM ---- DIR "C:\PROGRAMDATA\START MENU*" /S /B /AD
@REM --------------------------------------------------------------------
@REM ---- HERE WILL GET -- START MENU -- AND -- START MENU PLACES --
@REM ---- TWO IN ONE
@REM ---- NOT A DESKTOP HERE IN PROGRAMDATA
@REM --------------------------------------------------------------------

@REM --------------------------------------------------------------------
@REM ---- PART 07 OF 07 -- DIR C:\PROGRAMDATA\START MENU*
@REM --------------------------------------------------------------------
SET "TESTVAR="
@REM IF "!TESTVAR!"=="" (

@REM --------------------------------------------------------------------
for /f "delims=" %%F in ('DIR "C:\PROGRAMDATA\START MENU*" /S /B /AD') do (
	
	@REM -------------------------------------------------------------------------------------------
	@REM -- FLAG WANT ON
	@REM -------------------------------------------------------------------------------------------
	ECHO.%%F | FIND /I "\Desktop">Nul && Set "VAR_FLAG=" && SET VAR_FLAG=YES
	@REM -------------------------------------------------------------------------------------------
	@REM -- FLAG WANT EXCLUDE OFF -- PRIORITY ORDER
	ECHO.%%F | FIND /I "\USERS\Default">Nul && Set "VAR_FLAG=" && SET VAR_FLAG=NOT
	@REM -- FLAG WANT EXCLUDE OFF -- PRIORITY ORDER -- SOME TO LEAVE ALONE
	@REM -- NOT MY NEW USER CHANGER YET
	ECHO.%%F | FIND /I "\USERS\MATT 02">Nul && Set "VAR_FLAG=" && SET VAR_FLAG=NOT
	@REM -------------------------------------------------------------------------------------------

	IF !VAR_FLAG!==YES (
		IF "!TESTVAR!"=="" (
			ECHO --------------------------------------------------------------------
			ECHO -- Do All in ----
			ECHO.
			ECHO %%F
			ECHO --------------------------------------------------------------------
			ECHO.
			SET "PATH_VAR="
			SET PATH_VAR=%%F
			call :Set_ICACLS !PATH_VAR!
		)
	)
)


@REM ---------------------------------------------------------------------
ECHO DONE ..........
GOTO VERY_ENDE
@REM ---------------------------------------------------------------------

@REM ---------------------------------------------------------------------
@REM PAUSE
@REM ---------------------------------------------------------------------

@REM ---------------------------------------------------------------------
@REM ---------------------------------------------------------------------
@REM ---------------------------------------------------------------------
@REM ---------------------------------------------------------------------
@REM ---- END ----
@REM ---------------------------------------------------------------------
@REM ---------------------------------------------------------------------
@REM -- INTERESTING REF:
@REM -- THIS CHAP -- WRITE ON STACK OVERFLOW AND COPY PASTE TO OWN WEBSITE
@REM -- ---------------------
@REM -- Count = 272 -- Fri 26-Aug-2016 20:32:28
@REM -- ---------------------
@REM -- Form FindWindow ---
@REM -- windows - Batch file: Find if substring is in string (not in a file)_KECHENGPUZI_???? - Google Chrome
@REM -- ---------------------
@REM -- ----
@REM -- windows - Batch file: Find if substring is in string (not in a file)_KECHENGPUZI_????
@REM -- http://kechengpuzi.com/q/s7005951
@REM -- ----
@REM -- ---------------------
@REM -- ---------------------
@REM -- Count = 273 -- Fri 26-Aug-2016 20:35:43
@REM -- ---------------------
@REM -- Form FindWindow ---
@REM -- windows - Batch file: Find if substring is in string (not in a file)_KECHENGPUZI_???? - Google Chrome
@REM -- ---------------------
@REM -- For compatibility and ease of use it's always best to use FIND to do this, as you may be looking to match output from a command, a variable in a loop, or a pointer variable in your Batch/CMD.
@REM -- ---------------------
@REM -- ---------------------
@REM -- Count = 274 -- Fri 26-Aug-2016 20:35:44
@REM -- ---------------------
@REM -- Form FindWindow ---
@REM -- windows - Batch file: Find if substring is in string (not in a file)_KECHENGPUZI_???? - Google Chrome
@REM -- ---------------------
@REM -- ----
@REM -- windows - Batch file: Find if substring is in string (not in a file)_KECHENGPUZI_????
@REM -- http://kechengpuzi.com/q/s7005951
@REM -- ----
@REM -- ---------------------
@REM -- ---------------------
@REM -- Count = 275 -- Fri 26-Aug-2016 20:35:50
@REM -- ---------------------
@REM -- Form FindWindow ---
@REM -- windows - Batch file: Find if substring is in string (not in a file) - Stack Overflow - Google Chrome
@REM -- ---------------------
@REM -- For compatibility and ease of use it's always best to use FIND to do this, as you may be looking to match output from a command, a variable in a loop, or a pointer variable in your Batch/CMD.
@REM -- ---------------------
@REM -- ---------------------
@REM -- Count = 276 -- Fri 26-Aug-2016 20:35:52
@REM -- ---------------------
@REM -- Form FindWindow ---
@REM -- windows - Batch file: Find if substring is in string (not in a file) - Stack Overflow - Google Chrome
@REM -- ---------------------
@REM -- ----
@REM -- windows - Batch file: Find if substring is in string (not in a file) - Stack Overflow
@REM -- http://stackoverflow.com/questions/7005951/batch-file-find-if-substring-is-in-string-not-in-a-file
@REM --




::
:Set_ICACLS
setlocal

@REM -- TIP NOTE DON'T SET A VAR THAT IS OUTSIDE THE SUBROUTINE AND WITHIN 
@REM -- PROBLEM

@REM --------------------------------------------------------------
@rem -- NOT FINISHED YET WAS GOING TO LEARN STRIP THE QUOTES OFF ""
@REM -- GOT IT ----------------------------------------------------
@REM -- ---- 
@REM -- Dequote - remove quotes | Windows CMD | SS64.com
@REM -- http://ss64.com/nt/syntax-dequote.html
@REM -- ----

TITLE %1 -- IN PROGRESS %DATE--%TIME%

SET "INPUT_VAR="
SET INPUT_VAR=%1


SET "VAR_TITLE_1="
SET VAR_TITLE_1="%~1 -- ICACLS 01 OF 02 _____OWNERSHIP IN PROGRESS %DATE%--%TIME%"
SET "VAR_TITLE_2="
SET VAR_TITLE_2="%~1 -- ICACLS 02 OF 02 _GRANTEVERYONE IN PROGRESS %DATE%--%TIME%"

SET "ARGUMENT_OPTION_1="
SET ARGUMENT_OPTION_1=/HIGH /WAIT
SET "ARGUMENT_OPTION_2="
SET ARGUMENT_OPTION_2=/HIGH

SET "COMMAND_LINE_DISPLAY_____OWNERSHIP="
SET COMMAND_LINE_DISPLAY_____OWNERSHIP=ICACLS !INPUT_VAR! /T /C /L /Q /setowner "Everyone"
SET "COMMAND_LINE_DISPLAY_GRANTEVERYONE="
SET COMMAND_LINE_DISPLAY_GRANTEVERYONE=ICACLS !INPUT_VAR! /T /C /L /Q /grant Everyone:F

SET "ICACLS_VAR_____OWNERSHIP="
SET "ICACLS_VAR_GRANTEVERYONE="

SET ICACLS_VAR_____OWNERSHIP=START !VAR_TITLE_1! !ARGUMENT_OPTION_1! !COMMAND_LINE_DISPLAY_____OWNERSHIP!
SET ICACLS_VAR_GRANTEVERYONE=START !VAR_TITLE_2! !ARGUMENT_OPTION_2! !COMMAND_LINE_DISPLAY_GRANTEVERYONE!




ECHO.
ECHO.
ECHO.
ECHO -------------------------------------------------------
ECHO FOLDER TO DO ----
ECHO.
ECHO %~1
ECHO.
ECHO -------------------------------------------------------
ECHO COMMAND LAUNCH WRAPPER 
ECHO !ICACLS_VAR_____OWNERSHIP!
ECHO -------------------------------------------------------
ECHO COMMAND ON IT OWN
ECHO !COMMAND_LINE_DISPLAY_____OWNERSHIP!
ECHO -------------------------------------------------------
REM !ICACLS_VAR_____OWNERSHIP!
ECHO -------------------------------------------------------
ECHO.
ECHO -------------------------------------------------------
ECHO COMMAND LAUNCH WRAPPER
ECHO !ICACLS_VAR_GRANTEVERYONE!
ECHO -------------------------------------------------------
ECHO COMMAND ON IT OWN
ECHO !COMMAND_LINE_DISPLAY_GRANTEVERYONE!
ECHO -------------------------------------------------------
ECHO.
!ICACLS_VAR_GRANTEVERYONE!

REM -- OPTIONAL REM OUT PAUSE
REM PAUSE 

endlocal &  goto :EOF

@rem call :Set_ICACLS !PATH_VAR! RunPath
@rem endlocal & set "%2=%RunPath%" & goto :EOF

:VERY_ENDE
