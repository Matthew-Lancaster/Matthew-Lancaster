@ECHO.
@ECHO ---------------------------------------
@ECHO ---------------------------------------

@ECHO BEGIN -- 01 OF 02 -- SET OWNER EVERYONE
@ECHO NEXT --- 02 OF 02 -- GRANT EVERYONE
@ECHO Q FOR QUITE ABOUT SUCCESS MESSAGING
@ECHO ---------------------------------------
@ECHO ---------------------------------------
@ECHO.

@REM Icacls * /t /Q /grant Everyone:F >NUL
@REM Icacls * /t /Q /grant Everyone:F >NUL
@REM @PAUSE
@REM Icacls * /setowner "Everyone" /T /C /L /Q >NUL
@REM Icacls * /t /Q /grant Everyone:F >NUL

@REM Icacls "C:\Users"  /t /Q /grant Everyone:F >NUL


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

	

	IF "!VAR_FLAG!"=="YES" (
		IF "!TESTVAR!"=="" (
			
			
			SET "PATH_VAR=" 
			SET PATH_VAR=C:\%%F
			call :Set_Icacls "!PATH_VAR!"
			
			
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
ECHO ---- PART 02 OF 07 -- C:\USERS\All Users ---- Do All In -- All Users
ECHO ----------------------------------------------------------------------------

IF "!TESTVAR!"=="" (
	SET "PATH_VAR=" 
	SET PATH_VAR="C:\USERS\All Users"
	call :Set_Icacls !PATH_VAR!
)
ECHO --------------------------------------------------------------------
ECHO ---- PART 03 OF 07 -- C:\USERS\All Users ---- Do All In -- Public
ECHO --------------------------------------------------------------------
IF "!TESTVAR!"=="" (
	SET "PATH_VAR="
	SET PATH_VAR="C:\USERS\Public"
	call :Set_Icacls !PATH_VAR!
)
ECHO --------------------------------------------------------------------
ECHO ---- PART 04 OF 07 -- C:\USERS\Matt 01  ---- Do All In -- My User
ECHO --------------------------------------------------------------------
IF "!TESTVAR!"=="" (
	SET "PATH_VAR="
	SET PATH_VAR="C:\USERS\MATT 01"
	call :Set_Icacls !PATH_VAR!
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
			call :Set_Icacls !PATH_VAR!
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
			call :Set_Icacls !PATH_VAR!
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
			call :Set_Icacls !PATH_VAR!
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
:Set_Icacls
setlocal

@REM -- TIP NOTE DON'T SET A VAR THAT IS OUTSIDE THE SUBROUTINE AND WITHIN 
@REM -- PROBLEM

TITLE %1 -- IN PROGRESS %DATE--%TIME%

SET "INPUT_VAR="
SET INPUT_VAR=%1

SET "VAR_TITLE_1="
SET VAR_TITLE_1="%1 -- ICACLS 01 OF 02 _____OWNERSHIP IN PROGRESS %DATE%--%TIME%"
SET "VAR_TITLE_2="
SET VAR_TITLE_2="%1 -- ICACLS 02 OF 02 _GRANTEVERYONE IN PROGRESS %DATE%--%TIME%"

SET "ARGUMENT_OPTION_1="
SET ARGUMENT_OPTION_1=/HIGH /WAIT
SET "ARGUMENT_OPTION_2="
SET ARGUMENT_OPTION_2=/HIGH

SET "COMMAND_LINE_DISPLAY_____OWNERSHIP="
SET COMMAND_LINE_DISPLAY_____OWNERSHIP=Icacls !INPUT_VAR! /T /C /L /Q /setowner "Everyone"
SET "COMMAND_LINE_DISPLAY_GRANTEVERYONE="
SET COMMAND_LINE_DISPLAY_GRANTEVERYONE=Icacls !INPUT_VAR! /T /C /L /Q /grant Everyone:F

SET "ICACLS_VAR_____OWNERSHIP="
SET "ICACLS_VAR_GRANTEVERYONE="

SET ICACLS_VAR_____OWNERSHIP=START !VAR_TITLE_1! !ARGUMENT_OPTION_1! !COMMAND_LINE_DISPLAY_____OWNERSHIP!
SET ICACLS_VAR_GRANTEVERYONE=START !VAR_TITLE_2! !ARGUMENT_OPTION_2! !COMMAND_LINE_DISPLAY_GRANTEVERYONE!

ECHO -------------------------------------------------------
ECHO COMMAND LAUNCH WRAPPER 
ECHO !ICACLS_VAR_____OWNERSHIP!
ECHO -------------------------------------------------------
ECHO COMMAND ON IT OWN
ECHO !COMMAND_LINE_DISPLAY_____OWNERSHIP!
ECHO -------------------------------------------------------
!ICACLS_VAR_____OWNERSHIP!
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
PAUSE 

endlocal &  goto :EOF

@rem call :Set_Icacls !PATH_VAR! RunPath
@rem endlocal & set "%2=%RunPath%" & goto :EOF

:VERY_ENDE
