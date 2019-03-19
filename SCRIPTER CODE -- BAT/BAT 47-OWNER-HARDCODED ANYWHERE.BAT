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

SETLOCAL EnableDelayedExpansion

REM ---- FILENAME WITHOUT EXTENSION
@SET TITLE_VAR=%~n0
REM ---- FILENAME WITH   EXTENSION
@SET TITLE_VAR=%~n0%~x0
REM ---- FULL SCRIPT PATH
@SET TITLE_VAR=%~f0

REM Usage on variable %1 (but this works for any parameter)
REM %~f1 Expand %1 to a Fully qualified path name - C:\utils\MyFile.txt
REM %~d1 Expand %1 to a Drive letter only - C:
REM %~p1 Expand %1 to a Path only e.g. \utils\ this includes a trailing \ which will be interpreted as an escape character by some commands.
REM %~n1 Expand %1 to a file Name without file extension or path - MyFile 
REM or if only a path is present, with no trailing backslash, the last folder in that path.
REM %~x1 Expand %1 to a file eXtension only - .txt
REM %~s1 Change the meaning of f, n, s and x to reference the Short 8.3 name (if it exists.) 
REM %~1   Expand %1 removing any surrounding quotes (")
REM %~a1 Display the file attributes of %1
REM %~t1 Display the date/time of %1
REM %~z1 Display the file size of %1
REM %~$PATH:1 Search the PATH environment variable and expand %1 to the fully qualified name of the first match found.
REM The modifiers above can be combined:
REM %~dp1 Expand %1 to a drive letter and path only
REM %~sp1 Expand %1 to a path shortened to 8.3 characters
REM %~nx2 Expand %2 to a file name and extension only
REM ----
REM Parameters / Arguments - Windows CMD - SS64.com
REM https://ss64.com/nt/syntax-args.html
REM ----

SET TITLE_VAR=%~f0
TITLE !TITLE_VAR!
ECHO.

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

SET LDA_EXTRA_01=!LDA!
SET LDA_EXTRA_02=0
SET LDA_EXTRA_03=0
SET GET_SIZE_OF_WORK_COUNTER_VALUE=TRUE
CALL :SET_ICACLS_ROOT_OF_ALL_DRIVE_AND_NOT_SOME_FOLDER
CALL :SET_ICACLS_ROOT_OF_ALL_DRIVE_OF_E_DRIVE_WITH_SUBFOLDER_BUT_NOT_SOME_FOLDER

REM ----
REM How can I echo a newline in a batch file? - Stack Overflow
REM https://stackoverflow.com/questions/132799/how-can-i-echo-a-newline-in-a-batch-file
REM ----
REM ---- EXAMPLE ---- ECHO READY!\n!STEADY!\n!GO
REM ----
(SET \n=^
%=Do not remove this line=%
)

SET LDA_ALL=0
SET /A LDA_ALL=LDA_ALL+!LDA_EXTRA_01!
SET /A LDA_ALL=LDA_ALL+!LDA_EXTRA_02!
SET /A LDA_ALL=LDA_ALL+!LDA_EXTRA_03!

CALL :PAD_COUNTER
CALL :SET_COUNTER_FOR_DISPLAY_02

SET LDA_STR=00000!LDA_EXTRA_01!
SET LDA_STR=%LDA_STR:~-4%


SET FILENAME_VBS="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 41-GET DATE AND TIME FORMATTER.VBS"
for /f "tokens=*" %%a in ('cscript //nologo %FILENAME_VBS%') do (set "strVBScriptResults=%%a")
SET DATE_TIME_FORMATER=%strVBScriptResults%
SET FILENAME_FOR_LOGGER=%~dp0%~n0--LOGGER--%COMPUTERNAME%.TXT
SET FILENAME_NOT_PATH_FOR_LOGGER=%~n0--LOGGER--%COMPUTERNAME%.TXT
ECHO ----------------------------------------------------------------
ECHO ---------------------------------------------------------------->>!FILENAME_FOR_LOGGER!
ECHO !DATE_TIME_FORMATER! -- WORK BEGIN !LDA_STR! JOB -- %COMPUTERNAME%
ECHO !DATE_TIME_FORMATER! -- WORK BEGIN !LDA_STR! JOB -- %COMPUTERNAME%>>!FILENAME_FOR_LOGGER!
ECHO ----------------------------------------------------------------
ECHO ---------------------------------------------------------------->>!FILENAME_FOR_LOGGER!
ECHO !I_STRING!
ECHO !I_STRING!>>!FILENAME_FOR_LOGGER!
ECHO ----------------------------------------------------------------
ECHO ---------------------------------------------------------------->>!FILENAME_FOR_LOGGER!
ECHO.

ECHO ----------------------------------------------------------------
ECHO TIME AND DATE -- LOGGER INFO SENT TO
ECHO ----------------------------------------------------------------
ECHO !FILENAME_NOT_PATH_FOR_LOGGER!
REM ECHO !FILENAME_FOR_LOGGER!
ECHO ----------------------------------------------------------------
ECHO.

PAUSE



REM -----------------------------------------------------------------
REM -- FORMAT FOLDER NAME WITH EXTRA \*
REM -- CODING INCLUDE STRIPPER -- Double quotes ; Quotation mark ; speech marks
REM -- AND THEN PUT BACK AGAIN WITH \* ADDER TO END
REM ----
REM ASCII code Double quotes ; Quotation mark ; speech marks, American Standard Code for Information Interchange, The complete ASCII table, characters,letters, vowels with accents, consonants, signs, symbols, numbers double, quotes, quotation, mark, speech, marks,ascii,34, ascii art, ascii table, code ascii, ascii character, ascii text, ascii chart, ascii characters, ascii codes, characters, codes, tables, symbols, list, alt, keys, keyboard, spelling, control, printable, extended, letters, epistles, handwriting, scripts, lettering, majuscules, capitals, minuscules, lower, case, small, acute, accent, sharp, engrave, diaresis, circumflex, tilde, cedilla, anillo, circlet, eñe, enie, arroba, pound, sterling, cent, type, write, spell, spanish, english, notebooks, laptops, ascii, asci, asccii, asqui, askii, aski, aschi, aschii,20190319
REM https://theasciicode.com.ar/ascii-printable-characters/double-quotes-quotation-mark-speech-marks-ascii-code-34.html
REM ----
REM -----------------------------------------------------------------
FOR /L %%a in (1,1,!LDA!) DO (
	SET Sources[%%a]="!Sources[%%a]:"=!\*"
)
REM -----------------------------------------------------------------

WORK_IP=0

FOR /L %%a in (1,1,!LDA!) DO (
	IF NOT "!sources[%%a]!"=="" (
		
		SET /A WORK_IP_ALL=WORK_IP_ALL+1
		SET /A WORK_IP_01=WORK_IP_01+1
		ECHO --------------------------------------------------------
		CALL :SET_COUNTER_FOR_DISPLAY_01
		ECHO --------------------------------------------------------
		SET PATH_TO_DO=!Sources[%%a]!
		ECHO PATH_TO_DO = !PATH_TO_DO:"=!
		ECHO --------------------------------------------------------

		ECHO !PATH_TO_DO!
		SET VAR_FLAG=NOT
		IF !PATH_TO_DO!=="E:\01 Desktop\*" SET VAR_FLAG=YES
		SET VAR_FLAG=NOT
		IF !VAR_FLAG!==YES (
			SET COMMAND_VAR=TAKEOWN /F !PATH_TO_DO! /A /R /D Y
			REM ECHO !COMMAND_VAR!
			REM !COMMAND_VAR! >NUL
			REM -- TAKEOWN TO ADMIN AND MORE NON STOP AT ERROR WITH /C UNLIKE ABOVE
			REM -- ABLE SET TO EVERYONE AGAIN ONE ADMIN IS DONE BEFORE
			SET COMMAND_VAR=ICACLS !PATH_TO_DO! /Setowner Administrators /T /C /L /Q 
			ECHO !COMMAND_VAR!
			!COMMAND_VAR!
			SET COMMAND_VAR=ICACLS !PATH_TO_DO! /T /L /C /Q  /Grant Administrators:F
			ECHO !COMMAND_VAR!
			!COMMAND_VAR!
		)
		
		IF !VAR_FLAG!==NOT (
			SET COMMAND_VAR=ICACLS !PATH_TO_DO! /Setowner Administrators /T /C /L /Q 
			ECHO !COMMAND_VAR!
			!COMMAND_VAR!
			SET COMMAND_VAR=ICACLS !PATH_TO_DO! /T /L /C /Q  /Grant Administrators:F
			ECHO !COMMAND_VAR!
			!COMMAND_VAR!

			SET COMMAND_VAR=ICACLS !PATH_TO_DO! /Setowner Everyone /T /C /L /Q 
			ECHO !COMMAND_VAR!
			!COMMAND_VAR!
			SET COMMAND_VAR=ICACLS !PATH_TO_DO! /T /L /C /Q /Grant Everyone:F
			ECHO !COMMAND_VAR!
			!COMMAND_VAR!
		)

		SET VAR_FLAG=NOT
		IF !PATH_TO_DO!=="C:\SCRIPTOR GITHUB\*" SET VAR_FLAG=YES
		IF !VAR_FLAG!==YES (
			REM -- NOTICE HAD TO USE THE ^(CI^)H THESE Circumflex accent or Caret
			REM -- TO TELL DOS WITH STRING THAT NORMAL CHARACTER -- ( ( ) )
			SET COMMAND_VAR=ICACLS !PATH_TO_DO! /setintegritylevel ^(CI^)H
			ECHO !COMMAND_VAR!
			!COMMAND_VAR!
		)
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

REM ----
REM Batch Script Variables
REM https://www.tutorialspoint.com/batch_script/batch_script_variables.htm
REM ----
REM Example
REM @echo off 
REM set globalvar = 5
REM SETLOCAL
REM set var = 13145
REM set /A var = %var% + 5
REM echo %var%
REM echo %globalvar%
REM ENDLOCAL
REM Few key things to note about the above program.
REM The ‘globalvar’ is defined with a global scope and is available throughout the entire script.
REM The ‘var‘ variable is defined in a local scope because it is enclosed between a ‘SETLOCAL’ and ‘ENDLOCAL’ block. Hence, this variable will be destroyed as soon the ‘ENDLOCAL’ statement is executed.

GOTO JUMP_ENDER


::
:SYSTEM_COUNT_OF_ROOT_FOLDER_ALL_DRIVE
SET SYSTEM_ROOT_FOLDER_COUNT=0
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
					SET /A SYSTEM_ROOT_FOLDER_COUNT=SYSTEM_ROOT_FOLDER_COUNT+1
				)
			)
		)
	)
)
GOTO :EOF


REM [ Tuesday 04:32:00 Pm_19 March 2019 ]
REM ----
REM DOS COMMAND NULL THE ERROR OUTPUT IN FOR LOOP - Google Search
REM https://www.google.com/search?safe=active&ei=PBiRXOCWA8iM1fAPl_u2OA&q=DOS+COMMAND+NULL+THE+ERROR+OUTPUT+IN+FOR+LOOP&oq=DOS+COMMAND+NULL+THE+ERROR+OUTPUT+IN+FOR+LOOP&gs_l=psy-ab.3...99410.107185..107993...0.0..0.175.2516.37j1......0....1..gws-wiz.......0i71.5FCIxxhBjK0
REM --------
REM cmd - Redirect output of command in for loop of batch script - Stack Overflow
REM https://stackoverflow.com/questions/17113343/redirect-output-of-command-in-for-loop-of-batch-script
REM ----
::
:SET_ICACLS_ROOT_OF_ALL_DRIVE_AND_NOT_SOME_FOLDER

@FOR %%G IN (C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (
	IF EXIST %%G:\. (
		SET "TARGET_="
		SET TARGET_=%%G:
		CALL :IsDeviceReady !TARGET_! isready_
		IF !isready_!==true (
			 FOR /f "delims=" %%F in ('Dir !TARGET_!\ /b /a-d 2 2^>NUL') do (
				SET "VAR_FLAG="
				SET VAR_FLAG=YES

				IF "%%F"=="System Volume Information" SET VAR_FLAG=NOT

				IF !GET_SIZE_OF_WORK_COUNTER_VALUE!==TRUE (
					
					SET VAR_FLAG=NOT
					SET /A LDA_EXTRA_02=LDA_EXTRA_02+1
				)

				IF "!VAR_FLAG!"=="YES" (
					SET /A WORK_IP_ALL=WORK_IP_ALL+1
					SET /A WORK_IP_02=WORK_IP_02+1
					ECHO --------------------------------------------------------
					CALL :SET_COUNTER_FOR_DISPLAY_01
					ECHO --------------------------------------------------------
					SET PATH_TO_DO="!TARGET_!\%%F"
					ECHO PATH_TO_DO = !PATH_TO_DO:"=!
					ECHO --------------------------------------------------------
					ECHO !PATH_TO_DO!

					SET COMMAND_VAR="!TARGET_!\%%F" /C /L /Q /setowner "Everyone"
					ECHO !COMMAND_VAR!
					!COMMAND_VAR! >NUL
					SET COMMAND_VAR="!TARGET_!\%%F" /C /L /Q /grant Everyone:F
					ECHO !COMMAND_VAR!
					!COMMAND_VAR! >NUL
				)
			)
		)
	)
)
GOTO :EOF

::
:PAD_COUNTER
	SET LDD_01=00000!WORK_IP_ALL!
	SET LDD_02=00000!WORK_IP_01!
	SET LDD_03=00000!WORK_IP_02!
	SET LDD_04=00000!WORK_IP_03!
	SET LDD_01=!LDD_01:~-4!
	SET LDD_02=!LDD_02:~-4!
	SET LDD_03=!LDD_03:~-4!
	SET LDD_04=!LDD_04:~-4!
	SET LDA_01=00000!LDA_ALL!
	SET LDA_02=00000!LDA_EXTRA_01!
	SET LDA_03=00000!LDA_EXTRA_02!
	SET LDA_04=00000!LDA_EXTRA_03!
	SET LDA_01=!LDA_01:~-4!
	SET LDA_02=!LDA_02:~-4!
	SET LDA_03=!LDA_03:~-4!
	SET LDA_04=!LDA_04:~-4!
GOTO :EOF

::
:SET_COUNTER_FOR_DISPLAY_01
	CALL :PAD_COUNTER

	SET I_STR_1=COMMANDING INSTRUCTION POINTER ALL WORKER                         !LDD_01! OF !LDA_01!
	SET I_STR_2=COMMANDING INSTRUCTION POINTER 01 OF 03 MAIN ARRAY                !LDD_02! OF !LDA_02!
	SET I_STR_3=COMMANDING INSTRUCTION POINTER 02 OF 03 ALL DRIVE ROOT            !LDD_03! OF !LDA_03!
	SET I_STR_4=COMMANDING INSTRUCTION POINTER 03 OF 03 ALL E DRIVE ROOT WITH SUB !LDD_04! OF !LDA_04!
	SET I_STR_5=TOTAL COMPUTER ROOT FOLDER COUNT                                  !LDA_03!

	SET I_STRING=!I_STR_1!!\n!!I_STR_2!!\n!!I_STR_3!!\n!!I_STR_4!!\n!!\n!!I_STR_5!

	ECHO !I_STRING!
	
GOTO :EOF

:SET_COUNTER_FOR_DISPLAY_02
	CALL :PAD_COUNTER
	SET I_STR_1=COMMANDING COUNT ALL WORKER                         !LDA_01!
	SET I_STR_2=COMMANDING COUNT 01 OF 03 MAIN ARRAY                !LDA_02!
	SET I_STR_3=COMMANDING COUNT 02 OF 03 ALL DRIVE ROOT            !LDA_03!
	SET I_STR_4=COMMANDING COUNT 03 OF 03 ALL E DRIVE ROOT WITH SUB !LDA_04!
	SET I_STR_5=TOTAL COMPUTER ROOT FOLDER COUNT                    !LDA_03!

	SET I_STRING=!I_STR_1!!\n!!I_STR_2!!\n!!I_STR_3!!\n!!I_STR_4!!\n!!\n!!I_STR_5!

GOTO :EOF

::
:SET_ICACLS_ROOT_OF_ALL_DRIVE_OF_E_DRIVE_WITH_SUBFOLDER_BUT_NOT_SOME_FOLDER
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

				IF !GET_SIZE_OF_WORK_COUNTER_VALUE!==TRUE (
					SET VAR_FLAG=NOT
					SET /A LDA_EXTRA_03=LDA_EXTRA_03+1
				)

				IF "!VAR_FLAG!"=="YES" (
					SET /A WORK_IP_ALL=WORK_IP_ALL+1
					SET /A WORK_IP_03=WORK_IP_03+1
					ECHO --------------------------------------------------------
					CALL :SET_COUNTER_FOR_DISPLAY_01
					ECHO --------------------------------------------------------
					SET PATH_TO_DO="!TARGET_!\%%F"
					ECHO PATH_TO_DO = !PATH_TO_DO:"=!
					ECHO --------------------------------------------------------
					ECHO !PATH_TO_DO!

					SET COMMAND_VAR=ICACLS "!TARGET_!\%%F" /T /C /L /Q /setowner "Everyone"
					ECHO !COMMAND_VAR!
					!COMMAND_VAR! >NUL
					SET COMMAND_VAR=ICACLS "!TARGET_!\%%F" /T /C /L /Q /grant Everyone:F
					ECHO !COMMAND_VAR!
					!COMMAND_VAR! >NUL
				)
			)
		)
	)
)
GOTO :EOF



@REM START "" /D!PATH_TO_DO! /B /MIN /HIGH /WAIT ICACLS !PATH_TO_DO! /setowner Everyone /T /C /L /Q >NUL
@REM START "" /D!PATH_TO_DO! /B /MIN /HIGH /WAIT ICACLS !PATH_TO_DO! /T /L /C /Q /Grant Everyone:F >NUL



:JUMP_ENDER

ECHO.
ECHO DONE

SET FILENAME_VBS="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 41-GET DATE AND TIME FORMATTER.VBS"
for /f "tokens=*" %%a in ('cscript //nologo %FILENAME_VBS%') do (set "strVBScriptResults=%%a")
SET DATE_TIME_FORMATER=%strVBScriptResults%
ECHO ----------------------------------------------------------------
ECHO !DATE_TIME_FORMATER! -- WORK DONE !LDA_STR! JOB -- %COMPUTERNAME%
ECHO !DATE_TIME_FORMATER! -- WORK DONE !LDA_STR! JOB -- %COMPUTERNAME%>%~dp0%~n0--LOGGER.TXT
ECHO ----------------------------------------------------------------
ECHO ---------------------------------------------------------------->%~dp0%~n0--LOGGER.TXT
ECHO.

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