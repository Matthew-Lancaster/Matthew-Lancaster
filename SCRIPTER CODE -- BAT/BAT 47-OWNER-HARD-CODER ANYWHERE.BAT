@ECHO OFF
REM ---------------------------------------------------------------- 
REM __ C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 47-OWNER-HARD-CODER ANYWHERE.BAT
REM __ 
REM __ BAT 47-OWNER-HARD-CODER ANYWHERE.BAT
REM __ 
REM __ MATTHEW LANCASTER 
REM __ MATT.LAN@BTINTERNET.COM
REM ---------------------------------------------------------------- 

REM -------------------------------------------------------------------
REM -- SESSION 001
REM -- FROM ----------- Thu 03-May-2018 01:31:51 __ INITIAL START
REM -- TO   ----------- Thu 03-May-2018 03:35:00 __ 2 HOUR
REM -------------------------------------------------------------------
REM -------------------------------------------------------------------
REM -- SESSION 002
REM -- FROM ----------- Sun 14-Oct-2018 20:34:24
REM -- TO   ----------- Sun 14-Oct-2018 22:45:00
REM -------------------------------------------------------------------
REM -- SESSION 003
REM -- 0972 * CLIPBOARD'S TAKEN
REM -- Count = 099  -- Tue 19-Mar-2019 12:54:28
REM -- Count = 1071 -- Wed 20-Mar-2019 00:52:37 -- 12 HOUR
REM -------------------------------------------------------------------

REM -------------------------------------------------------------------
REM __ Location Internet
REM -------------------------------------------------------------------
REM __ BAT 47-OWNER-HARD-CODER ANYWHERE.BAT
REM -------------------------------------------------------------------
REM Matthew-Lancaster/BAT 47-OWNER-HARD-CODER ANYWHERE.BAT at master · Matthew-Lancaster/Matthew-Lancaster
REM https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20BAT/BAT%2047-OWNER-HARD-CODER%20ANYWHERE.BAT
REM -------------------------------------------------------------------
REM -------------------------------------------------------------------

REM ---------------------------------------------------------------- 
REM __ ALL BATCH FILE COMMAND SETS 
REM __ GOT AN IMPROVED ADMINISTRATOR LAUNCHER
REM __ THAT DOESN'T GET STUCK IN A LOOP IF NOT SUCCESSFUL
REM __ AT THE SACRIFICE OF COMMAND LINE ARGUMENTS BUT EASY TO 
REM __ ADD CODE IF REQUIRE
REM ---------------------------------------------------------------- 
REM __ FROM ---- Thu 03-May-2018 01:31:51 __ INITIAL START
REM __ TO   ---- Thu 03-May-2018 03:35:00 __ 2 HOUR
REM ---------------------------------------------------------------- 

REM ---------------------------------------------------------------- 
REM -- SESSION 001 OF 003 -- UPDATES HAPPEN ALL THE TIME
REM ---------------------------------------------------------------- 
REM -- SETTING PERMISSION THIS SCRIPT IS AUTORUN BY AN AUTOHOTKEYS PROGRAM
REM -- EVERY FEW DAY AND A TIMER MARKER IS WRITTEN TO FILE
REM -- YOU HAVE OT KEEP IT UP WRITE PERMISSION OR SILLY THING HOPELESS
REM ---------------------------------------------------------------- 
REM ---------------------------------------------------------------- 
REM -- HERE BAT IS NOW SYNC IN 2 LOCATION 
REM -- CENTRALIZING MY BEST CODE
REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 47-OWNER-HARD-CODER ANYWHERE.BAT
REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\OWNER\#_OWNER-HARDCODED ANYWHERE.BAT
REM ---------------------------------------------------------------- 
REM -- IS IS SYNCED BY HERE
REM -- C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 32-COPIER_SYNC.VBS
REM ---------------------------------------------------------------- 

REM ---------------------------------------------------------------- 
REM -- SESSION 002 OF 003 
REM ---------------------------------------------------------------- 
REM -- CHANGED EnableDelayedExpansion TO USE LATER IN THE PROGRAM WHERE 
REM -- THE FOR LOOP ARE
REM -- 
REM -- PLAYED AROUND LOOKING FOR CODE TO COMPARE TWO VARIABLE IN 
REM -- EnableDelayedExpansion IN A LOOP AND IN A FUNCTION FIND IMPOSSIBLE AS FUNCTION WON'T USE GOTO LABEL
REM -- AND FOR LOOP CAN COUNT TO CERTAIN AMOUNT OR CARRY ON FOREVER
REM -- AND VARIABLE HAVE TO BE SOMEHOW PUT INTO A FOR LOOP TO USE THERE
REM -- ANOTHER THING IMPOSABLE 
REM -- AT LEAST FOUND OUT ABOUT INFINITE FOR LOOP WITH COUNTER STEPPER SET ON NAUGHT INSTEAD OF 1
REM -- 
REM -- ALSO FINDING OUT THAT TAKE-OWNERSHIP IS MORE POWERFUL IF NOT GOT ACCESS TO COPIER OF FILE
REM -- 
REM -- EVEN IF IT CANT GIVE ACCESS TO ICACLS CHANGE
REM -- 
REM -- ALSO "C:\Windows\System32\*.VBS" THIS ONE WHEN TAKE-OWNERSHIP HAPPEN
REM -- AS THERE ARE FEW VBS SCRIPT IN THERE
REM -- TAKE OWNERSHIP HAPPEN UP TO THIS 4TH FILE C:\Windows\System32\winrm.vbs
REM -- AND THEN STOP ON EROR NOT GOT ACCESS OF SOME KIND
REM -- AND WHEN ICACLS RUN IT WITH IT PASS OVER ERROR OPTION SET
REM -- IT GETS TO A FILE
REM -- C:\Windows\System32\LogFiles\WMI\RtBackup\*.VBS: Access is denied.
REM -- THAT INCLUDE A WILD-CARD AND THAT WAS WHERE STOPPED AT FOR TAKE-OWNERSHIP
REM -- THE ICACLS SET-OWNER HAS PROBLEM WITH FILE NOT BEEN ABLE TO BE SET EVEN WITH TAKE-OWNER COMMAND
REM -- AND THE ICACLS GRANT HAS PROBLEM WITH FILE AFTER THE WILD-CARD ONE AS THEY NOT BEEN SET BY TAKE-OWNER COMMAND
REM -- 
REM -- I FIND SETTER OWNER AND GRANT WITH ICACLS IS BETTER SET TO ADMINISTRATOR FOR SOME THING 
REM -- RATHER THAN EVERYONE - EVERYONE DOESN'T GET ACCEPTED
REM -- FOR 1ST FEW THIS SETTING "C:\Windows\System32\*.VBS"
REM -- AND HERE IS STILL NOT ABLE TO SET DUE TO I THINK THE FOLDER IS DISALLOW
REM -- C:\Windows\System32\Printing_Admin_Scripts\en-US\
REM -- 
REM -- TAKE-OWNER COMMAND IS ONLY USED TO SET A PERSON YOUR CURRENT USER OR ADMINISTRATOR
REM -- 
REM -- ONCE BEEN SET TO ADMINISTRATOR YOU ARE THEN ABLE TO SET EVERYONE
REM ---------------------------------------------------------------- 
REM -- FROM Sun 14-Oct-2018 20:34:24
REM -- FROM Sun 14-Oct-2018 22:45:00
REM ---------------------------------------------------------------- 

REM -------------------------------------------------------------------
REM -- SESSION 003 OF 003 
REM -------------------------------------------------------------------
REM -- AFTER A WHOLE DAY
REM -- TOTALLY REVAMPED WHOLE THING
REM -- TO USE ARRAY AND A LOT MORE OF FUNCTION AND SUBROUTINE
REM -- NOT COSTING CODING ANYWHERE NEVER REUSE ANY EFFICIENT TO THE MAX
REM -- AND THIS BECOME BECAUSE GOODSYNC
REM -- HAVE BEEN ARGUING FOR A LONGER TIME
REM -- THAT USING THE P2P MODE GOODSYNC DIRECT FOR DRIVE COPIER
REM -- WHICH IS USEFUL EFFICIENCY IS LESS 1/2 
REM -- WHEN NETWORK SETTING STOP GIVER ACCESS CERTAIN FOLDER
REM -- LIKE I'M WRITE NOW AS THIS IS THE SET ADMINISTRATOR MODE OR ACCESS EVERYONE
REM -- BUT AS I WORK I'M NOW DECIDER THAT SET TO EVERYONE ACCESS IS NOT WORKING
REM -- AS IT TAKE TWICE AS LONG
REM -- BECAUSE I FOUND THAT TO SET EVERYONE ACCESS WHICH MEANS LIKE WHOLE DRIVE 
REM -- COMPLETELY EVERYONE HAS ACCESS
REM -- IS NOT ANY POINT AS SOMETIMES TO GET THAT ACCESS GIVER HAVE TO SET ADMINISTRATOR 
REM -- FIRST AND THEN THE OTHER
REM -- AND NOW THAT IS A TIME CONSUMER
REM -- AND WHAT THE HECK
REM -- ADMINISTRATOR IS ACCESS EVERYONE ANYWAY
REM -- UNLESS YOU NOT AN ADMINISTRATOR USER
REM -- BUT ADMINISTRATOR CAN COME ALONG AND LOOK AT IT
REM -- AND THE END OF THE PROBLEM THERE IS 
REM -- IF GIVEN WERE GOING TO CLEAN REINSTALL THE OPERATOR SYSTEM WINDOWS 10 AGAIN
REM -- AND DONE IT IN ADMINISTRATOR 
REM -- THAT USE FILE NOT IN ADMINISTRATOR BEFORE WOULD BECOME ACCESSIBLE AGAIN
REM -- AND JOB TAKE TWICE AS LONG TO GET THERE
REM -- --
REM -- GIVE COULD BE LESS BETTER OFF AS ADMINISTRATOR WORK WELL ON LOCAL COMPUTER
REM -- BUT NETWORK ACCESS MIGHT SUFFER 
REM -- SOON FIND OUT -- NOT CORRECT STILL NOT WORKER
REM -- --
REM -- THE RESULT FOUND BY TEST _ ADMINISTRATOR WON'T EASILY ALLOW ACCESS TO FOLDER
REM -- WHEN EVERYONE SETTING WILL OVER NETWORK
REM -- WHICH IS UNFORTUNATE BECAUSE FOR BEST RESULT GIVE MUST HAVE TO BE SET 
REM -- ADMINISTRATOR AND THEN EVERYONE WHICH IS TWICE A JOB
REM -- MAYBE A LEVEL OF STABILIZE AND SWITCH HALF BETWEEN TWO
REM -- --
REM -- AND GOODSYNC HAVE BEEN VERY ANNOYER
REM -- BECAUSE THEY GOT A WARNING NOT TO USE GOODSYNC DIRECT P2P MODE
REM -- ONLY BECAUSE EFFICIENCY IS LESS 2/2
REM -- AS EXAMPLE SOMETIME WHEN GIVEN COULD ACCESS THE DATA DRIVE DIRECT WHY WANT TO GO 
REM -- THROUGH P2P CLEVER MODE
REM -- ONLY PROBLEM IS NOW THEY DO SENSE THAT COMPLAIN WITH WARNING THAT WAS THERE
REM -- AND NOW THEY TOTALLY BANNED IT 
REM -- SO NEXT THE PROBLEM ARE IS GIVEN WANT TO COPY A LOAD OF JOB FROM GOODSYNC IMPORT EXPORT 
REM -- AND PUT ON ANOTHER MACHINE
REM -- NONE ARE WORKING
REM -- AND IF GIVEN CHANGE THE FIRST SOURCE FOLDER OF SYNC PAIR
REM -- ALL THE FILTER ARE GONE NOT ON THE DESTINATION SYNCER SIDE
REM -- BUT MOST THE REAL ONE GIVEN WANT ARE ON SYNCER SOURCE SIDE
REM -- AND ADDED PAIN GIVEN GOT TO SWAP ALL THE DIRECTORY TO NOT BEING P2P GOODSYNC TO DIRECT 
REM -- ONTO THE HDD DRIVE 
REM -- LETTER LIKE  Local hard drives (C:\;D:\;E:\;G:\;H:\;I:\;K:\;L:\;M:\;N:\;O:\;P:\;Q:\;R:\;S:\;T:\;U:\)
REM -- RATHER THAN EXAMPLE
REM -- gstps://4-asus-gl522vw.matt-lan-2.goodsync/D:/DSC/2015-Now Sony
REM -- 
REM -- A LOT OF STRING MANIPULATION TECHNIQUE CREATED FOR BATCH FILING
REM -- JUST FOR PRETTY DECORATION
REM -- FINALLY FOUND WRITE FUNCTION AND SUBROUTINE EASIER
REM -- BECAUSE PREVIOUSLY HAD A BASH ON THEM
REM -- THE EXAMPLE I GET ARE FOR ALL SETLOCAL ON VARIABLE
REM -- THAT MEAN THEY ARE HARD TO PASS TO EACH OTHER
REM -- LEARN THE GLOBAL VARIABLE SYSTEM MUCH BETTER
REM -- VARIABLE DON'T PASS BETWEEN FUNCTION THE SAME WAY AS BASIC
REM -- BUT NOT FAR FROM IT -- YOU CAN ENTRY A VARIABLE AND USE MODIFY 
REM -- AND EXCEPT IT TO LEAVE THE ROUTINE AGAIN 
REM -- BUT IS A PROCEDURE OF GRAB IT AT ENTRY AND PUT BACK BEFORE LEAVE
REM -- MUCH LIKE THE ENVIRONMENT VARIABLE PARAMETER PASS ON COMMAND LINE
REM -- AS %1 %2 AND ALL AS %*
REM -- FILE WRITING FOR LOGGER IS PRESENT NOW SO ABEL SEE HOW FAR GET
REM -- EACH NEW WORK JOB IS WRITTEN TO SEE HOW FAR THAT OVERWRITE PREVIOUS INFO
REM -- IS TOO MUCH STORE BE UNNECESSARY
REM -- AND APPEND LOG FOR EACH STAT AND FINISH
REM -- THAT INCLUDE DATA HOW MANY DRIVE GOT THAT ARE IS-READY FOR ACCESS 
REM -- IN YOUR SYSTEM
REM -- TOTAL OF HOW MANY ROOT FOLDER ON ALL DRIVE OVER SYSTEM
REM -- THE NEW WORK IS SPLIT IN 3 PART
REM -- FIRST IS MAIN ARRAY THAT IS USER HARD-CODER OR SEMI HARD-CODER EDITABLE NOTEPAD
REM -- FORM NOTEPAD HERE NOT A COMPLIER 
REM -- NEXT ONE 2ND IS ROOT DRIVE ON EVERY DRIVE
REM -- NEXT ONE 3RD IS WHOLE OF ANY DRIVE WHICH I HAVE SMALL E DRIVE ONLY 4GB FROM THEM TPE OF DAY
REM -- WHERE SMALL IS EASIER TO SYNC WHOLE THING QUICKER
REM -- AND I STORE SPECIAL FOLDER LIKE DESKTOP AND START-MENU AND AS MANY OTHER 
REM -- OF SPECIAL FOLDER ABLE TO SET BY REGISTRY ON START OF EVERY SYSTEM INSTALL 
REM -- AND WHENEVER -- RELOCATABLE TO ANY HDD DRIVE
REM -- 
REM -------------------------------------------------------------------
REM -- WORK FROM 
REM -------------------------------------------------------------------
REM -- 1069 * CLIPBOARD'S TAKEN
REM -- Count = 0002 -- Tue 19-Mar-2019 11:23:10
REM -- Count = 1071 -- Wed 20-Mar-2019 00:52:37 -- 12 HOUR
REM -------------------------------------------------------------------

@ECHO OFF
REM ENDLOCAL
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

REM ---------------------------------------------------------------- 
SET LDA=0
REM ---------------------------------------------------------------- 
SET /A LDA=LDA+1
SET Sources[!LDA!]="D:\VB6-ARC"
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
SET /A LDA=LDA+1
SET Sources[!LDA!]=""
SET /A LDA=LDA+1
SET Sources[!LDA!]=""
SET /A LDA=LDA+1
SET Sources[!LDA!]=""
SET /A LDA=LDA+1
SET Sources[!LDA!]=""
SET /A LDA=LDA+1
SET Sources[!LDA!]=""
SET /A LDA=LDA+1
SET Sources[!LDA!]=""
SET /A LDA=LDA+1
SET Sources[!LDA!]=""
REM ---------------------------------------------------------------- 

REM ---------------------------------------------------------------- 
REM -- SORT THE EMPTY ARRAY FOR VOID VALUE
REM ---------------------------------------------------------------- 
SET LDB=0
FOR /L %%a in (1,1,!LDA!) DO (
	IF NOT !sources[%%a]!=="" (
		SET /A LDB=LDB+1
		SET Sources_02[!LDB!]=!Sources[%%a]!
	)
)
REM ---------------------------------------------------------------- 

REM ---------------------------------------------------------------- 
REM -- PUT BACK IN
REM ---------------------------------------------------------------- 
FOR /L %%a in (1,1,!LDB!) DO (
	SET Sources[%%a]=!Sources_02[%%a]!
)

CALL :COMPUTER_COUNT_OF_ALL_DRIVE

SET LDA=%LDB%
REM ---------------------------------------------------------------- 

SET WORK_IP_ALL=0
SET WORK_IP_01=0
SET WORK_IP_02=0
SET WORK_IP_03=0

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

SET DASH_LINE__________=------------------------------------------------------------------------------
SET FILENAME_FOR_LOGGER_01=%~dp0%~n0_LOGGER_APPENDER_%COMPUTERNAME%.TXT
SET FILENAME_FOR_LOGGER_02=%~dp0%~n0_LOGGER_CURRENT__%COMPUTERNAME%.TXT

CALL :SET_COUNTER_FOR_DISPLAY_01
CALL :SET_COUNTER_FOR_DISPLAY_GET_FILENAME_FROM_FULL_PATH
CALL :SET_COUNTER_FOR_DISPLAY_03_WRITE_1ST_LOGGER_INFO

ECHO !DASH_LINE__________!
ECHO TIME AND DATE -- LOGGER INFO SENT TO
ECHO !DASH_LINE__________!
ECHO !FILENAME_FOR_LOGGER_01!
ECHO !FILENAME_FOR_LOGGER_02!
ECHO !FILENAME_WITHOUT_PATH_FOR_LOGGER_01!
ECHO !FILENAME_WITHOUT_PATH_FOR_LOGGER_02!
ECHO !DASH_LINE__________!
ECHO.


REM ---------------------------------------------------------------- 
REM -- FORMAT FOLDER NAME WITH EXTRA \*
REM -- CODING INCLUDE STRIPPER -- Double quotes ; Quotation mark ; speech marks
REM -- AND THEN PUT BACK AGAIN WITH \* ADDER TO END
REM ----
REM ASCII code Double quotes ; Quotation mark ; speech marks, American Standard Code for Information Interchange, The complete ASCII table, characters,letters, vowels with accents, consonants, signs, symbols, numbers double, quotes, quotation, mark, speech, marks,ascii,34, ascii art, ascii table, code ascii, ascii character, ascii text, ascii chart, ascii characters, ascii codes, characters, codes, tables, symbols, list, alt, keys, keyboard, spelling, control, printable, extended, letters, epistles, handwriting, scripts, lettering, majuscules, capitals, minuscules, lower, case, small, acute, accent, sharp, engrave, diaresis, circumflex, tilde, cedilla, anillo, circlet, eñe, enie, arroba, pound, sterling, cent, type, write, spell, spanish, english, notebooks, laptops, ascii, asci, asccii, asqui, askii, aski, aschi, aschii,20190319
REM https://theasciicode.com.ar/ascii-printable-characters/double-quotes-quotation-mark-speech-marks-ascii-code-34.html
REM ----
REM ----------------------------------------------------------------
FOR /L %%a in (1,1,!LDA!) DO (
	SET Sources[%%a]="!Sources[%%a]:"=!\*"
)
REM ---------------------------------------------------------------- 

SET WORK_JOB_01=TRUE
FOR /L %%a in (1,1,!LDA!) DO (
	IF NOT "!sources[%%a]!"=="" (
		
		SET /A WORK_IP_ALL=WORK_IP_ALL+1
		SET /A WORK_IP_01=WORK_IP_01+1
		
		ECHO !DASH_LINE__________!
		CALL :SET_COUNTER_FOR_DISPLAY_02

		ECHO !DASH_LINE__________!
		SET PATH_TO_DO=!Sources[%%a]!
		ECHO PATH_TO_DO = !PATH_TO_DO:"=!
		ECHO !DASH_LINE__________!

		CALL :SET_COUNTER_FOR_DISPLAY_03_WRITE_APPENDER_LOGGER_INFO

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
			ECHO.
			SET COMMAND_VAR=ICACLS !PATH_TO_DO! /T /L /C /Q  /Grant Administrators:F
			ECHO !COMMAND_VAR!
			!COMMAND_VAR!
			ECHO.
		)
		
		IF !VAR_FLAG!==NOT (
			
			REM SET COMMAND_VAR=ICACLS !PATH_TO_DO! /Setowner Administrators /T /C /L /Q 
			REM ECHO !COMMAND_VAR!
			REM !COMMAND_VAR!
			REM SET COMMAND_VAR=ICACLS !PATH_TO_DO! /T /L /C /Q  /Grant Administrators:F
			REM ECHO !COMMAND_VAR!
			REM !COMMAND_VAR!
			REM ECHO.

			SET COMMAND_VAR=ICACLS !PATH_TO_DO! /Setowner Everyone /T /C /L /Q 
			ECHO !COMMAND_VAR!
			!COMMAND_VAR!
			SET COMMAND_VAR=ICACLS !PATH_TO_DO! /T /L /C /Q /Grant Everyone:F
			ECHO !COMMAND_VAR!
			!COMMAND_VAR!
			ECHO.
		)

		SET VAR_FLAG=NOT
		IF !PATH_TO_DO!=="C:\SCRIPTOR GITHUB\*" SET VAR_FLAG=YES
		IF !VAR_FLAG!==YES (
			REM -- NOTICE HAD TO USE THE ^(CI^)H THESE Circumflex accent or Caret
			REM -- TO TELL DOS WITH STRING THAT NORMAL CHARACTER -- ( ( ) )
			SET COMMAND_VAR=ICACLS !PATH_TO_DO! /setintegritylevel ^(CI^)H
			ECHO !COMMAND_VAR!
			!COMMAND_VAR!
			ECHO.
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


SET GET_SIZE_OF_WORK_COUNTER_VALUE=FALSE
SET WORK_JOB_01=
SET WORK_JOB_02=TRUE
CALL :SET_ICACLS_ROOT_OF_ALL_DRIVE_AND_NOT_SOME_FOLDER
SET WORK_JOB_02=
SET WORK_JOB_03=TRUE
CALL :SET_ICACLS_ROOT_OF_ALL_DRIVE_OF_E_DRIVE_WITH_SUBFOLDER_BUT_NOT_SOME_FOLDER

GOTO VERY_ENDE

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
					ECHO !DASH_LINE__________!
					CALL :SET_COUNTER_FOR_DISPLAY_02
					CALL :SET_COUNTER_FOR_DISPLAY_03_WRITE_APPENDER_LOGGER_INFO
					ECHO !DASH_LINE__________!
					SET PATH_TO_DO="!TARGET_!\%%F"
					SET PATH_TO_DO_02="!TARGET_!\%%F"
					ECHO PATH_TO_DO = !PATH_TO_DO_02:"=!
					ECHO !DASH_LINE__________!
					ECHO !PATH_TO_DO!

					SET COMMAND_VAR=ICACLS !PATH_TO_DO! /C /L /Q /setowner "Everyone"
					ECHO !COMMAND_VAR!
					!COMMAND_VAR! >NUL
					ECHO.
					SET COMMAND_VAR=ICACLS !PATH_TO_DO! /C /L /Q /grant Everyone:F
					ECHO !COMMAND_VAR!
					!COMMAND_VAR! >NUL
					ECHO.
					
					REM SET COMMAND_VAR=ICACLS !PATH_TO_DO! /Setowner Administrators /T /C /L /Q 
					REM ECHO !COMMAND_VAR!
					REM !COMMAND_VAR!
					REM ECHO.
					REM SET COMMAND_VAR=ICACLS !PATH_TO_DO! /T /L /C /Q  /Grant Administrators:F
					REM ECHO !COMMAND_VAR!
					REM !COMMAND_VAR!
					REM ECHO.
					
				)
			)
		)
	)
)
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

					ECHO !DASH_LINE__________!
					CALL :SET_COUNTER_FOR_DISPLAY_02
					CALL :SET_COUNTER_FOR_DISPLAY_03_WRITE_APPENDER_LOGGER_INFO
					ECHO !DASH_LINE__________!

					SET PATH_TO_DO="!TARGET_!\%%F"
					SET PATH_TO_DO_02="!TARGET_!\%%F"
					ECHO PATH_TO_DO = !PATH_TO_DO_02:"=!
					ECHO !DASH_LINE__________!
					ECHO !PATH_TO_DO!

					SET COMMAND_VAR=ICACLS !PATH_TO_DO! /T /C /L /Q /setowner "Everyone"
					ECHO !COMMAND_VAR!
					!COMMAND_VAR! >NUL
					ECHO.
					SET COMMAND_VAR=ICACLS !PATH_TO_DO! /T /C /L /Q /grant Everyone:F
					ECHO !COMMAND_VAR!
					!COMMAND_VAR! >NUL
					ECHO.
					
					REM SET COMMAND_VAR=ICACLS !PATH_TO_DO! /Setowner Administrators /T /C /L /Q 
					REM ECHO !COMMAND_VAR!
					REM !COMMAND_VAR!
					REM ECHO.
					REM SET COMMAND_VAR=ICACLS !PATH_TO_DO! /T /L /C /Q  /Grant Administrators:F
					REM ECHO !COMMAND_VAR!
					REM !COMMAND_VAR!
					REM ECHO.
				)
			)
		)
	)
)
GOTO :EOF

::
:COMPUTER_COUNT_OF_ALL_DRIVE
SET COMPUTER_DRIVE_COUNT=0
@FOR %%G IN (C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (
	IF EXIST %%G:\. (
		SET "TARGET_="
		SET TARGET_=%%G:
		CALL :IsDeviceReady !TARGET_! isready_
		IF !isready_!==true (
			SET /A COMPUTER_DRIVE_COUNT=COMPUTER_DRIVE_COUNT+1
		)
	)
)
GOTO :EOF

::
:COMPUTER_COUNT_OF_ROOT_FOLDER_ALL_DRIVE
SET COMPUTER_ROOT_FOLDER_COUNT=0
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
					SET /A COMPUTER_ROOT_FOLDER_COUNT=COMPUTER_ROOT_FOLDER_COUNT+1
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
	SET I_STR_1=COMMANDING COUNT ALL WORKER                                       !LDA_01!
	SET I_STR_2=COMMANDING COUNT 01 OF 03 MAIN ARRAY                              !LDA_02!
	SET I_STR_3=COMMANDING COUNT 02 OF 03 ALL DRIVE ROOT                          !LDA_03!
	SET I_STR_4=COMMANDING COUNT 03 OF 03 ALL E DRIVE ROOT WITH SUB               !LDA_04!
	SET I_STR_5=TOTAL COMPUTER ROOT FOLDER COUNT                                  !LDA_03!
	SET I_STR_8=TOTAL COMPUTER DRIVE COUNT                                        !COMPUTER_DRIVE_COUNT!

	SET I_STRING_01=!I_STR_1!!\n!!I_STR_2!!\n!!I_STR_3!!\n!!I_STR_4!!\n!!\n!!I_STR_5!!\n!!I_STR_8!
	REM -- IPHONE -- I-STRING -- I-SCREAM SHOCK ICE SCREAM ICE CREAM
	SET I_STRING_02=!I_STR_1!!\n!!I_STR_2!!\n!!I_STR_3!!\n!!I_STR_4!!\n!!I_STR_5!!\n!!I_STR_8!

GOTO :EOF

::
:SET_COUNTER_FOR_DISPLAY_02
	CALL :PAD_COUNTER
	
	SET FILENAME_VBS="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 41-GET DATE AND TIME FORMATTER.VBS"
	for /f "tokens=*" %%a in ('cscript //nologo %FILENAME_VBS%') do (set "strVBScriptResults=%%a")
	SET DATE_TIME_FORMATER=%strVBScriptResults%
	SET I_STR_A=!DATE_TIME_FORMATER!
	SET I_STR_1=COMMANDING INSTRUCTION POINTER ALL WORKER                         !LDD_01! OF !LDA_01!
	SET I_STR_2=COMMANDING INSTRUCTION POINTER 01 OF 03 MAIN ARRAY                !LDD_02! OF !LDA_02!
	SET I_STR_3=COMMANDING INSTRUCTION POINTER 02 OF 03 ALL DRIVE ROOT            !LDD_03! OF !LDA_03!
	SET I_STR_4=COMMANDING INSTRUCTION POINTER 03 OF 03 ALL E DRIVE ROOT WITH SUB !LDD_04! OF !LDA_04!
	SET I_STR_5=TOTAL COMPUTER ROOT FOLDER COUNT                                  !LDA_03!
	SET I_STR_8=TOTAL COMPUTER DRIVE COUNT                                        !COMPUTER_DRIVE_COUNT!
	
	IF !WORK_JOB_01!==TRUE (
		SET I_STR_9=COMMAND INSTRUCTION POINTER 01 OF 03 MAIN ARRAY                   !LDD_02! OF !LDA_02!
	)
	IF !WORK_JOB_02!==TRUE (
		SET I_STR_9=COMMAND INSTRUCTION POINTER 02 OF 03 ALL DRIVE ROOT               !LDD_03! OF !LDA_03!
	)
	IF !WORK_JOB_03!==TRUE (
		SET I_STR_9=COMMAND INSTRUCTION POINTER 03 OF 03 ALL E DRIVE ROOT WITH SUB    !LDD_04! OF !LDA_04!
	)
	
	SET I_STR_10= -----------------------------------------------------------------------------

	SET I_STRING_01=!I_STR_A!!\n!!I_STR_1!!\n!!I_STR_2!!\n!!I_STR_3!!\n!!I_STR_4!!\n!!I_STR_5!!\n!!I_STR_8!!\n!
	SET I_STRING_01=!I_STRING_01!!\n!!I_STR_10!!\n!!I_STR_9!
	
	ECHO !I_STRING_01!
	REM -- IPHONE -- I-STRING -- I-SCREAM SHOCK ICE SCREAM ICE CREAM
	
GOTO :EOF

REM ----
REM How to split the filename from a full path in batch? - Stack Overflow
REM https://stackoverflow.com/questions/9252980/how-to-split-the-filename-from-a-full-path-in-batch
REM --------
REM DOS COMMAND CONVERT STRING FULL PATH TO FILENAME - Google Search
REM https://www.google.com/search?q=DOS+COMMAND+CONVERT+STRING+FULL+PATH+TO+FILENAME
REM ----
::
:SET_COUNTER_FOR_DISPLAY_GET_FILENAME_FROM_FULL_PATH
	For %%A in ("!FILENAME_FOR_LOGGER_01!") do (
		Set Folder=%%~dpA
		Set FILENAME_WITHOUT_PATH_FOR_LOGGER_01=%%~nxA
	)
	For %%A in ("!FILENAME_FOR_LOGGER_02!") do (
		Set Folder=%%~dpA
		Set FILENAME_WITHOUT_PATH_FOR_LOGGER_02=%%~nxA
	)
GOTO :EOF
::
:SET_COUNTER_FOR_DISPLAY_03_WRITE_1ST_LOGGER_INFO
	SET VAR_TELL_WHAT_TO_DO=1
	CALL :SET_COUNTER_FOR_DISPLAY_03
GOTO :EOF
::
:SET_COUNTER_FOR_DISPLAY_03_WRITE_APPENDER_LOGGER_INFO
	SET VAR_TELL_WHAT_TO_DO=2
	CALL :SET_COUNTER_FOR_DISPLAY_03
GOTO :EOF
::
:SET_COUNTER_FOR_DISPLAY_03
	IF !VAR_TELL_WHAT_TO_DO!==1 (
		SET FILENAME_FOR_LOGGER=!FILENAME_FOR_LOGGER_01!
	)
	IF !VAR_TELL_WHAT_TO_DO!==2 (
		SET FILENAME_FOR_LOGGER=!FILENAME_FOR_LOGGER_02!
	)
	
	SET FILENAME_VBS="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 41-GET DATE AND TIME FORMATTER.VBS"
	for /f "tokens=*" %%a in ('cscript //nologo %FILENAME_VBS%') do (set "strVBScriptResults=%%a")
	SET DATE_TIME_FORMATER=%strVBScriptResults%
	
	IF !VAR_TELL_WHAT_TO_DO!==1 (
		SET I_STRING_04=WORK BEGIN !LDA_01! TOTAL WORK -- %COMPUTERNAME%
		SET DATE_TIME_BEGIN_FORMATER=!DATE_TIME_FORMATER!
		ECHO !DASH_LINE__________!>>!FILENAME_FOR_LOGGER!
		ECHO !DATE_TIME_FORMATER!>>!FILENAME_FOR_LOGGER!
		ECHO !I_STRING_04!>>!FILENAME_FOR_LOGGER!
		ECHO !DASH_LINE__________!>>!FILENAME_FOR_LOGGER!
		ECHO !I_STRING_02!>>!FILENAME_FOR_LOGGER!
		ECHO !DASH_LINE__________!>>!FILENAME_FOR_LOGGER!
	)
	SET I_STR_12=01 OF 03 MAIN ARRAY
	SET I_STR_13=02 OF 03 ALL DRIVE ROOT
	SET I_STR_14=03 OF 03 ALL E DRIVE ROOT WITH SUB
	REM -- GET THE LONGEST STRING -- I_STR_14
	CALL :STRLEN RESULT I_STR_14
	SET /A RESULT=!RESULT!+4
	SET CURRENT_WORKER_VAR=---- CURRENT WORKER
	
	IF !WORK_JOB_01!==TRUE (
		SET I_STR_10=!LDD_02! OF !LDA_02! -- 01 OF 03 MAIN ARRAY
		SET I_STR_12=!I_STR_12!
		CALL :PADD_EXTRA_ON RESULT I_STR_12
		SET I_STR_12=!I_STR_12!!CURRENT_WORKER_VAR!
		SET I_STR_XX=!I_STR_12!
	)
	IF !WORK_JOB_02!==TRUE (
		SET I_STR_10=!LDD_03! OF !LDA_03! -- 02 OF 03 ALL DRIVE ROOT
		SET I_STR_13=!I_STR_13!
		CALL :PADD_EXTRA_ON RESULT I_STR_13
		SET I_STR_13=!I_STR_13!!CURRENT_WORKER_VAR!
		SET I_STR_XX=!I_STR_13!
	)
	IF !WORK_JOB_03!==TRUE (
		SET I_STR_10=!LDD_04! OF !LDA_04! -- 03 OF 03 ALL E DRIVE ROOT WITH SUB
		SET I_STR_14=!I_STR_14!
		CALL :PADD_EXTRA_ON RESULT I_STR_14
		SET I_STR_14=!I_STR_14!!CURRENT_WORKER_VAR!
		SET I_STR_XX=!I_STR_14!
	)

	SET WORK_SET_VAR_01=WORK SET OF THREE
	CALL :STRLEN RESULT I_STR_XX
	SET DASH_LINE_02=
	FOR /l %%x in (1, 1, !RESULT!) DO SET DASH_LINE_02=!DASH_LINE_02!-
	SET "DASH_LINE_04=!DASH_LINE_02!XxxX"
	CALL :STRLEN RESULT DASH_LINE_04
	SET /A STR_TXT_LEN_MAKER=78-!RESULT!
	FOR /l %%x in (1, 1, !STR_TXT_LEN_MAKER!) DO SET DASH_LINE_04=!DASH_LINE_04!-
	SET DASH_LINE_02=!DASH_LINE_04!
	
	IF !VAR_TELL_WHAT_TO_DO!==2 (
		SET I_STRING_04=WORK CURRENT !LDD_01! OF !LDA_01! -- %COMPUTERNAME%
		SET I_STRING_05=WORK CURRENT !LDD_01! OF !LDA_01! -- WORK TOTAL
		SET I_STRING_08=WORK CURRENT !I_STR_10!
		SET I_STRING_10=!WORK_SET_VAR_01!!\n!!DASH_LINE_02!!\n!
		SET I_STRING_10=!I_STRING_10!!I_STR_12!!\n!!I_STR_13!!\n!!I_STR_14!

		ECHO !DASH_LINE__________!>!FILENAME_FOR_LOGGER!
		ECHO !DATE_TIME_FORMATER!>>"!FILENAME_FOR_LOGGER!"
		ECHO !DATE_TIME_BEGIN_FORMATER! -- BEGIN DATE TIME>>"!FILENAME_FOR_LOGGER!"
		ECHO !DASH_LINE__________!>>!FILENAME_FOR_LOGGER!
		ECHO !I_STRING_04!>>!FILENAME_FOR_LOGGER!
		ECHO !DASH_LINE__________!>>!FILENAME_FOR_LOGGER!
		ECHO !I_STRING_02!>>!FILENAME_FOR_LOGGER!
		ECHO !DASH_LINE__________!>>!FILENAME_FOR_LOGGER!
		ECHO !I_STRING_10!>>!FILENAME_FOR_LOGGER!
		ECHO !DASH_LINE__________!>>!FILENAME_FOR_LOGGER!

		ECHO !I_STRING_05!>>!FILENAME_FOR_LOGGER!
		ECHO !I_STRING_08!>>!FILENAME_FOR_LOGGER!
		ECHO !DASH_LINE__________!>>!FILENAME_FOR_LOGGER!
		REM -- CAN'T MAKE THIS VARIABLE IN ONE LINE REDIRECT TO FILENAME BECAUSE 
		REM -- THEM EQUAL SIGN REQUIRE ESCAPE CHAR BEFORE
		REM -  BUT STILL NOT MAKE IT WORK !PATH_TO_DO:"=! SO USE TWO LINE INSTEAD
		SET PATH_LOGGER=!PATH_TO_DO:"=!
		REM ECHO PATH TO DO =>>!FILENAME_FOR_LOGGER!

		REM ---------------------------------------------------------
		SET DASH_LINE_04__________=_________________________________________________________________________
		SET DASH_LINE_04__________=-------------------------------------------------------------------------
		REM NOT ABLE TO USE VARIABLE IN PLACE WHERE THE NUMERIC ARE %LEN_T% !LEN_T! !DASHLINE:~-40!
		REM BUT AND THEN USE A FOR LOOP AND THEN WE ARE ABLE
		REM MAKE TH LINE LONG AND CHOP TO SIZE USE ELSE WHERE IN CODE HERE TO PADDER NAUGHT NUMBER AT FRONT
		SET DASH_LINE_04__________=-------------------------------------------------------------------------
		SET DASH_LINE_04__________=!DASH_LINE_04__________:~-67!
		REM ---------------------------------------------------------
		
		SET VAR_PATH_TO_DO="PATH TO DO "
		SET VAR_PATH_TO_DO=!VAR_PATH_TO_DO:"=!
		CALL :STRLEN RESULT VAR_PATH_TO_DO
		REM ---------------------------------------------------------
		REM -- PADD EXTRA TO END OF LINE __ YOU WANT LINE TO BE TOTAL 78 CHARACTER 
		REM -- LONG
		REM ---------------------------------------------------------
		SET /A STR_TXT_LEN_MAKER=78-!RESULT!
		SET DASH_LINE_04__________=
		FOR /l %%x in (1, 1, !STR_TXT_LEN_MAKER!) do SET DASH_LINE_04__________=!DASH_LINE_04__________!-
		
		ECHO !VAR_PATH_TO_DO!!DASH_LINE_04__________!>>!FILENAME_FOR_LOGGER!
		ECHO !DASH_LINE__________!>>!FILENAME_FOR_LOGGER!
		ECHO !PATH_LOGGER!>>!FILENAME_FOR_LOGGER!
		ECHO !DASH_LINE__________!>>!FILENAME_FOR_LOGGER!
	)
GOTO :EOF

:PADD_EXTRA_ON <VAR_IN_01> <VAR_IN_02>
REM -----------------------------------------------------------------
REM -- ONLY PASS VARIABLE TO FUNCTION
REM -----------------------------------------------------------------
REM CAN'T USE THE INPUT VARIABLE AND MODIFY 
REM BUT ARE ABLE TO TAKE INTO ANOTHER VARIABLE 
REM AND THEN PUT BACK WHEN DONE ANYTHING ON
REM -----------------------------------------------------------------
(   
    SET "S=!%~2!#"
    SET "STRINGV=!%~2!"
    SET "LEN=0"
    FOR %%P IN (4096 2048 1024 512 256 128 64 32 16 8 4 2 1) DO (
        IF "!S:~%%P,1!" NEQ "" ( 
            SET /A "LEN+=%%P"
            SET "S=!S:~%%P!"
        )
    )
)
REM -- CAN'T DO ADD OR SUBTRACTING TWO VARIABLE WHEN USE WITH FOR LOOP 
SET /A STR_PADD_EXTRA_ON=!%~1!-!LEN!
FOR /l %%x in (1, 1, !STR_PADD_EXTRA_ON!) DO SET "STRINGV=!STRINGV! "
SET "%~2=!STRINGV!"
GOTO :EOF

:STRLEN <RESULTVAR> <STRINGVAR>
(   
    SET "S=!%~2!#"
    SET "LEN=0"
    FOR %%P IN (4096 2048 1024 512 256 128 64 32 16 8 4 2 1) DO (
        IF "!S:~%%P,1!" NEQ "" ( 
            SET /A "LEN+=%%P"
            SET "S=!S:~%%P!"
        )
    )
)
SET "%~1=%LEN%"
GOTO :EOF



:JUMP_ENDER
ECHO.
ECHO !DASH_LINE__________!
ECHO DONE
ECHO !DASH_LINE__________!
ECHO.

SET FILENAME_VBS="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 41-GET DATE AND TIME FORMATTER.VBS"
for /f "tokens=*" %%a in ('cscript //nologo %FILENAME_VBS%') do (set "strVBScriptResults=%%a")
SET DATE_TIME_FORMATER=%strVBScriptResults%
ECHO !DASH_LINE__________!
ECHO !DASH_LINE__________!>>!FILENAME_FOR_LOGGER!
ECHO !DATE_TIME_FORMATER! -- WORK DONE !LDA_STR! JOB -- %COMPUTERNAME%
ECHO !DATE_TIME_FORMATER! -- WORK DONE !LDA_STR! JOB -- %COMPUTERNAME%>>!FILENAME_FOR_LOGGER!
ECHO !DASH_LINE__________!
ECHO !DASH_LINE__________!>>!FILENAME_FOR_LOGGER!
ECHO.
ECHO.>>!FILENAME_FOR_LOGGER!

IF "%1"=="" PAUSE

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


:VERY_ENDE
:ENDE
REM -- VERY_ENDE VERY ENDE VERY WINDY -- HA HA
REM -- IPHONE -- I-STRING -- I-SCREAM SHOCK ICE SCREAM ICE CREAM

@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/----------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/----------\/\/\/\/\/\/\/\/\/\/\/\/\/\
REM ---------------------------------------------------------------- 
REM -- THE END AND REFERENCE SECTION -------------------------------
REM -- BYE BYE -----------------------------------------------------
REM ---------------------------------------------------------------- 
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/----------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/----------\/\/\/\/\/\/\/\/\/\/\/\/\/\

EXIT



@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/----------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/----------\/\/\/\/\/\/\/\/\/\/\/\/\/\
REM ---------------------------------------------------------------- 
REM -- AND REFERENCE SECTION ---------------------------------------
REM -- REFERENCE SET CAME AROUND TO IN USE THE CODE HERE -----------
REM ---------------------------------------------------------------- 
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/----------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/----------\/\/\/\/\/\/\/\/\/\/\/\/\/\

REM START "" /D!PATH_TO_DO! /B /MIN /HIGH /WAIT ICACLS !PATH_TO_DO! /setowner Everyone /T /C /L /Q >NUL
REM START "" /D!PATH_TO_DO! /B /MIN /HIGH /WAIT ICACLS !PATH_TO_DO! /T /L /C /Q /Grant Everyone:F >NUL

REM ---------------------------------------------------------------- 
REM GOOD CODE BUT NOT WANTER IN END
REM NORTON DOESN'T RESTRICT FROM LOOK FILE NAME IN 
REM C:\PROGRAM FILES FOLDER
REM ---------------------------------------------------------------- 
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
