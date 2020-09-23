@ECHO OFF
@CD \
@ECHO ______________________________
@ECHO __ MATTHEW LANCASTER
@ECHO __ MATT.LAN@BTINTERNET.COM
@ECHO __ Tue 21 February 2017 10:50:00 -- ALL-NIGHTER PREP THIS LEVEL ALL DEBUG WORKING
@ECHO __ ONE ROUNTINE FINISH DO I'M KNACKERING
@ECHO ______________________________

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


@REM WORK 
@REM ---------------------------------------
@REM MADE WORK FOR XP
@REM FIND PROCESS EXIST ADD WILDCARD VERSION USING TASKLIST 
@REM ---------------------------------------
@REM FROM    [ Sat 16:07:48 Pm_28 Apr 2018 ]
@REM TO      [ Sat 17:23:40 Pm_28 Apr 2018 ] 1 HOUR 10 MINUTE
@REM ---------------------------------------


@ECHO ON
@ECHO OFF

@ECHO ____ BAT 01 BOOT KILLER.BAT ____ 
@TITLE %~f0

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



@REM -------------------------------------------------------------
GOTO JUMPER_10


@REM C:\TEMP\-1-ASUS-X5DIJ\WER07a0.dir00\sysdata.xml

@REM BSOD

@REM C:\TEMP\-1-ASUS-X5DIJ\WER07a0.dir00\

@REM DISABLE DRIVER IS OKAY 
@REM NOT STOP AUDIO
@REM C:\TEMP\-1-ASUS-X5DIJ\WER07a0.dir00\Mini031817-01.dmp
@REM C:\TEMP\-1-ASUS-X5DIJ\WER07a0.dir00\sysdata.xml

@REM DEVICE MANAGER
@REM SOUND, VIDEO AND GAME CONTROLLERS
@REM VIA HIGH DEFINTION AUDIO


	@REM <DRIVER>
		@REM <FILENAME>viahduaa.sys</FILENAME>
		@REM <FILESIZE>1057280</FILESIZE>
		@REM <CREATIONDATE>11-24-2010 22:05:29</CREATIONDATE>
		@REM <VERSION>6.0.1.7033</VERSION>
		@REM <MANUFACTURER>VIA Technologies, Inc.</MANUFACTURER>
	@REM <PRODUCTNAME>
@REM VIA High Definition Audio Driver</PRODUCTNAME>
	@REM </DRIVER>


	
:JUMPER_10
@REM -------------------------------------------------------------




@ECHO OFF

@REM @SETLOCAL EnableDelayedExpansion


@REM  ______________________
@REM  __ @SETTER HERE _ TO ENABLE USE THE VARIABLE HOLDER WITHER !_! RATHER THAN %_% 
@REM  __ WHICH WOULD ACESS AS ASYCRONOUS AND OUR EVENT WOULD PRODUCE NOTHING 
@REM  __ AS TOO  QUICK OR ELSE DATA VARIABLE WOULD BE 
@REM  __ RANDOM NUMBER OUT OF SYNC CHAOS TO REQUEST WANT RETURN
@REM  ______________________



GOTO JUMPER_OVER_WORK_BEING_MASTER

GOTO JUMPER_REFEREANCES_SOURCER__TEXT_@REMMER_
	
@REM EXAMPLE OF AN ESCAPE CHARATOR NOT WORKING WHEN PASS VARIABLE
@REM THE ONLY ESCAPE CHARATOR I GOT THAT HAS A PROBLEM __'__
@REM AT THE MOMENT IN MY __ EXE __ VB __ IN MY COLLECTION
@REM OF 528

@SET TSK_SCRIPTOR_NAME_=
@SET TSK_SCRIPTOR_NAME_=#0 Send To - Fix I's.exe

@SET WORD_REPLACE=/'
CALL @SET TSK_SCRIPTOR_NAME_=%%TSK_SCRIPTOR_NAME_:%WORD_REPLACE%=%WORD_REPLACE%%%
@ECHO %TSK_SCRIPTOR_NAME_%

@REM ----------------------------------------------------------------------------
@REM EXAMPLE OF STACKOVERFLOW GIVEN COULD OF BEEN MORE SIMPLE 
@REM REUSE THE VARIABLE OF REPLACE TWICE RATHER THAN WRITE TWO
@REM ----------------------------------------------------------------------------

@REM ----------------------------------------------------------------------------
@REM 1..__
@REM REFERENCE TO LEARN &
@REM SEARCH STRING GOOGLE
@REM ----
@REM COMMAND LINE AND THIS CHARACTER IN QUOTE '
@REM ----
@REM windows - Escape double quotes in parameter - Stack Overflow
@REM http://stackoverflow.com/questions/7760545/escape-double-quotes-in-parameter
@REM ----
@REM ----------------------------------------------------------------------------
@REM REFERENCE HERE 1.. OF 02
@REM 2 Actually ^ didn't work for me, but \ did, per this answer: stackoverflow.com/questions/2403647/… – kalenjordan Aug 16 '12 at 22:05
@REM REFERENCE HERE 2.. OF 02 __ 2010 YEAR
@REM ----
@REM How to escape parameter in windows command line? - Stack Overflow
@REM http://stackoverflow.com/questions/2403647/how-to-escape-parameter-in-windows-command-line
@REM ----
@REM ----------------------------------------------------------------------------
@REM 2..__
@REM REFERENCE TO LEARN &
@REM SEARCH STRING GOOGLE
@REM ----
@REM COMMAND LINE VARIABLE SEARCH REPLACE
@REM ----
@REM command line - String replacement in batch file - Stack Overflow
@REM http://stackoverflow.com/questions/2772456/string-replacement-in-batch-file
@REM ----

:JUMPER_REFEREANCES_SOURCER__TEXT_@REMMER_

@REM   __ REQUIRE DIR FILE COUNT TO ESTIMATE LENGTH TIME WITH COUNTER ADDING
@REM   __

@REM   __ FILE COUNTING
@REM   ________________
FOR /F %%A IN ('DIR D:\VB6\VB-NT\*.EXE /S ^| FIND "File(s)"') DO @SET VB_EXE_FILE_COUNT=%%A
@SET ACCUMULATOR_FILE =0

@REM   __ FILE NAME GATHERING
@REM   ________________
@SET TSK_SCRIPTOR_NAME_=
FOR /F "delims=" %%A in ('dir D:\VB6\VB-NT\*.EXE /a-d /b /s') DO (

	@SET /A ACCUMULATOR_FILE=!ACCUMULATOR_FILE!+1
	@SET IX1=!ACCUMULATOR_FILE! of %VB_EXE_FILE_COUNT%

	@ECHO !IX1!__%%~nxA

	@SET TSK_SCRIPTOR_NAME_=%%~nxA

	@SET WORD_REPLACE=/'
	CALL @SET TSK_SCRIPTOR_NAME_=%%TSK_SCRIPTOR_NAME_:%WORD_REPLACE%=%WORD_REPLACE%%%
	
	@REM  @ECHO %TSK_SCRIPTOR_NAME_%

	TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
	@IF %errorlevel% EQU 1 (
		@ECHO !IX1!__%%~nxA
		TASKLIST /FI "IMAGENAME EQ %%~nxA"  
	)
	

	@REM   TASKLIST /FI "IMAGENAME EQ %%~nxG" 
	@REM  @ECHO.
	
	)
	
	@REM   TASKLIST /FI "IMAGENAME EQ %%~nxG" 
	@REM   TASKLIST /FI "IMAGENAME EQ %%~nxG" | @REM  find "INFO: No tasks" > nul
	
	@REM  @IF %errorlevel% equ 0 (
	
		@REM  @ECHO %%
		
		@REM  @ECHO RESULT %%G
		@REM  @ECHO RESULT IN @ECHO ERROR %%G
		@REM  PAUSE
		@REM  TASKLIST /FI "IMAGENAME EQ %%~nxG" 

	@REM  )
	
	@REM  __ LOOP ENDER
	
	PAUSE
	EXIT
	
	@REM  TASKLIST /FI "IMAGENAME EQ %%~nxG"  | find "INFO: No tasks" > nul
	PAUSE
	
	@IF %errorlevel% EQU 0 GOTO NOT_ERROR_TASK_NOT_RUNNING

	@REM   TASKLIST /FI "IMAGENAME EQ %%~nxG"
	
	@ECHO %%G
	@ECHO %%~nxG

PAUSE
PAUSE



:JUMPER_OVER_WORK_BEING_MASTER
@REM -------------------------------------------------------------

@REM   @ECHO HERE

@ECHO OFF

@SETLOCAL DisableDelayedExpansion
@SETLOCAL EnableExtensions

@REM  ----------------------------
@ECHO.
@ECHO ---- HERE DisableDelayedExpansion IS REQUIRE AS ALL 
@ECHO ---- THE %%ERRORLEVEL%% UP AHEAD WOULD CHANGE TO !ERRORLEVEL! IN THE EDITOR TO LEARN
@REM  ----------------------------


@REM  PAUSE 

@ECHO ON
CD \
D:
CD\

@REM ------------------------------------------------
@REM ----
@REM COMMAND+BATCH+FIND+ON+FILENAME+BAT - Google Search
@REM https://www.google.co.uk/search?q=COMMAND%2BBATCH%2BFIND%2BON%2BFILENAME%2BBAT&rlz=1C1CHBD_en-GBGB721GB721&oq=COMMAND%2BBATCH%2BFIND%2BON%2BFILENAME%2BBAT&aqs=chrome..69i57.2537j0j9&sourceid=chrome&ie=UTF-8
@REM ----



@REM   ------------------------------------------------

@ECHO OFF

@REM  ---------------------------------------------------------------
ver | find "10.0." > nul
@IF %errorlevel% EQU 0 @SET WIN10__=YES
ver | find "6.1." > nul
@IF %errorlevel% EQU 0 @SET WIN7___=YES
ver | find "5.1." > nul
@IF %errorlevel% EQU 0 @SET WINXP__=YES
@REM  ---------------------------------------------------------------

@REM  ---------------------------------------------------------------
@IF "%WINXP__%"=="YES" @ECHO WINDOWS XP
@IF "%WIN7___%"=="YES" @ECHO WINDOWS 7
@IF "%WIN10__%"=="YES" @ECHO WINDOWS 10
@REM  ---------------------------------------------------------------


@REM ----------------------------------------------------------------
@IF %COMPUTERNAME%==2-ASUS-EEE       SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==1-ASUS-X5DIJ     SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==3-LINDA-PC       SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==4-ASUS-GL522VW   SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==5-ASUS-P2520LA   SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==7-ASUS-GL522VW   SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==8-MSI-GP62M-7RD  SET MATTHEW_S_COMPUTER__=YES
@REM ----------------------------------------------------------------

@REM [ Sunday 05:50:50 Am_03 March 2019 ]
@REM DON'T KNOW WHAT USE FOR BUT ALWAYS NOT RESPOND ON CHECK
@REM ----------------------------------------------------------------
@IF NOT %COMPUTERNAME%==4-ASUS-GL522VW GOTO JUMP_OVER
TASKKILL /F /IM jusched.exe /T
:JUMP_OVER

@IF NOT %COMPUTERNAME%==7-ASUS-GL522VW GOTO JUMP_OVER
TASKKILL /F /IM jusched.exe /T
:JUMP_OVER

@REM [ Sunday 05:50:50 Am_03 March 2019 ]
@REM DON'T KNOW WHAT USE FOR BUT ALWAYS NOT RESPOND ON CHECK
@REM -- NOT A GOOD KILLER IT COMES BACK AND REMOVE TASKBAR TRAY 
@REM ----------------------------------------------------------------
@IF NOT %COMPUTERNAME%==4-ASUS-GL522VW GOTO JUMP_OVER
REM TASKKILL /F /IM sihost.exe /T
:JUMP_OVER

@IF NOT %COMPUTERNAME%==7-ASUS-GL522VW GOTO JUMP_OVER
REM TASKKILL /F /IM sihost.exe /T
:JUMP_OVER



@ECHO.

@REM   @SETlocal EnableDelayedExpansion
@REM   @SETlocal EnableExpansion

@ECHO OFF

GOTO JUMP_OVER
@IF NOT %COMPUTERNAME%==2-ASUS-EEE GOTO JUMP_NOT_ASUS_EEE
TASKKILL /F /IM CHROME.EXE /T
:JUMP_OVER
:JUMP_NOT_ASUS_EEE

@IF NOT %COMPUTERNAME%==4-ASUS-GL522VW GOTO JUMP_NOT_COMPUTER_NAME
TASKKILL /F /IM DFX.EXE /T
:JUMP_NOT_COMPUTER_NAME

@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
GOTO JUMP_OVER_UPDATE_STOPPER
@IF NOT "%WIN10__%"=="YES" GOTO JUMP_NOT_WINDOWS_10
@IF %COMPUTERNAME%==7-ASUS-GL522VW      GOTO JUMP_NOT_MATTHEW_S_COMPUTER
@IF NOT "%MATTHEW_S_COMPUTER__%"=="YES" GOTO JUMP_NOT_MATTHEW_S_COMPUTER
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -- WINDOWS SERVICE SHUTDOWN AND @REMOVE DATA
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM PART 1 OF 3 
@REM -----------------------------------------------------------------
@REM __ TASKKILL /F /IM wuauclt.exe /T

@REM __ TASKKILL /F /FI "SERVICES eq wuauserv"
@REM __ PROPER WAY __ SERVICING
@REM __ WINDOWS UPDATE SERVICE__

NET STOP WUAUSERV
SC CONFIG WUAUSERV START= DISABLED

@REM __ SC CONFIG WUAUSERV START= DEMAND

@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------

@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM PART 2 OF 3
@REM -----------------------------------------------------------------
@REM LOCATE THE BATCH FILE IN THE SAME DIRECTORY
CALL "%~dp0BAT 11-RESET WINDOWS UPDATE VER 10.BAT" /QUIET

@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM PART 3 OF 3
@REM -----------------------------------------------------------------
@REM TAKEOWN /F C:\$Windows.~BT\* /R /A 
@REM PAUSE
@REM ICACLS C:\$Windows.~BT\*.* /T /grant administrators:F 
TAKEOWN /F C:\$Windows.~BT\* /R /A /D Y
ICACLS "C:\$Windows.~BT\*" /T /C /L /Q /setowner "Everyone"
ICACLS "C:\$Windows.~BT\*" /T /C /L /Q /grant Everyone:F
RMDIR /S /Q C:\$Windows.~BT\

TAKEOWN /F C:\$Windows.~WS\* /R /A /D Y
ICACLS "C:\$Windows.~WS\*.*" /T /C /L /Q /setowner "Everyone"
ICACLS "C:\$Windows.~WS\*" /T /C /L /Q /grant Everyone:F
RMDIR /S /Q C:\$Windows.~WS\

@REM ----
@REM Safe to delete "$WINDOWS.~BT" - Google Search
@REM https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB744GB744&q=Safe+to+delete+%22%24WINDOWS.%7EBT%22&oq=Safe+to+delete+%22%24WINDOWS.%7EBT%22&gs_l=psy-ab.3..0i30k1.6478.6651.0.6915.2.2.0.0.0.0.78.139.2.2.0....0...1.1.64.psy-ab..0.2.137...0i13i30k1.kjHUIirxFxA
@REM --------
@REM Can I delete $Windows.~BT & $Windows.~WS folders?
@REM http://www.thewindowsclub.com/delete-windows-bt-ws-folders-windows-10
@REM ----
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------

:JUMP_NOT_WINDOWS_10
:JUMP_NOT_MATTHEW_S_COMPUTER
:JUMP_OVER_UPDATE_STOPPER
@REM --------------------------------------------------------------------------
@REM --------------------------------------------------------------------------
@ECHO -------------------------------------------------------------------------
@ECHO -- WILD CARD __ 01 TO 05 OF 01 FIRST ONLY __ FOR EVENT NOT RESPONDER NEED CRASHER
@ECHO -------------------------------------------------------------------------
@REM  -------------------------------------------------------------------------
@ECHO 1.. TASKKILL /F /FI "STATUS EQ NOT RESPONDING" _______ WILD CARD ALL NOT RESPONDING
@ECHO -------------------------------------------------------------------------
@REM  __ @ECHO 2.. TASKKILL /FI "IMAGENAME NE Z" /IM * /T     _______ WILD CARD EVERYTHING __ AND __ NOT FORCE 
@REM  __ @ECHO 3.. TASKKILL /F /FI "IMAGENAME NE Z" /IM * /T  _______ WILD CARD EVERYTHING __ YES FORCE BSD BSD XBOX
@REM  __ @ECHO 4.. TASKKILL /FI "WINDOWTITLE NE Z" /IM * /T      ____ WILD CARD EVERYTHING ALL WINDOWTITLE NOT FORCE
@REM  __ @ECHO 5.. TASKKILL /F /FI "WINDOWTITLE NE Z" /IM * /T ______ WILD CARD EVER ALL WINDOWTITLE YES FORCE BSD
@REM  -------------------------------------------------------------------------
@SET VAR_COUNTER=21
@SET ACCUMULATOR =0

@ECHO.
@ECHO -------------------------------------------------------------------------
@ECHO FIND THE NOT RESPONDER 1ST AND TWICE 2ND AT WHERE APPEAR 
@ECHO LATER IN CALL BY SUBROUTINER @SET OF 05
@ECHO -------------------------------------------------------------------------
@ECHO.
@REM CALL :WILD_CARD__01_OF_05
@REM   __ NOT RESPOND PROCESS NAME TO BE CHECKED WHY
@REM   __ 1.. _______
@REM   __ jusched.exe
@REM   __ SUN 26 FEB
@REM --------------------------------------------------------------------------
@REM --------------------------------------------------------------------------
@REM --------------------------------------------------------------------------
@REM --------------------------------------------------------------------------


@SET ACCUMULATOR =0

@ECHO ON
@ECHO OFF

@REM   GOTO HERE22 __ TESTING CODER


@REM  __ MICROSOFT KEYBOARD @REMOTE
@REM  __ REQUIRE UNINSTALL
TASKKILL /F /IM ITYPE.EXE /T
TASKKILL /F /IM IPOINT.EXE /T

@REM  __ FOR XP __ ASUS-EEE
TASKKILL /F /IM RTHDCPL.EXE /T

@IF NOT "%WINXP_%"=="YES" GOTO JUMPER_NOT_XP
	@REM  __ FOR XP
	@ECHO WINDOW XP KILL __ GOOGLEUPDATE.EXE
	TASKKILL /F /IM GOOGLEUPDATE.EXE /T

:JUMPER_NOT_XP



@REM  <> BEGIN BLOCKER SCHEDULER 
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------

@REM  -------------------------------------------------
@REM  ALL COMPUTER THE BOOT UP DELAY AND EMAIL SCHEDULE
@REM  -------------------------------------------------
@ECHO ON
@REM  @ECHO OFF

@REM --------------------------
@schtasks /QUERY /TN "\SYSTEM LOGON.BAT" | find "ERROR: The system cannot find the file spec@IFied." > nul
@IF NOT %errorlevel% EQU 0 GOTO NOT
@REM --------------------------
schtasks /CREATE /RU "MATT 01" /RP " " /TN "\SYSTEM LOGON.BAT" /TR "C:\22_LOGON.BAT" /SC ONLOGON /RL HIGHEST /F
@REM --------------------------
:NOT

@REM --------------------------
@schtasks /QUERY /TN "\SYSTEM STARTER.BAT" | find "ERROR: The system cannot find the file spec@IFied." > nul
@IF NOT %errorlevel% EQU 0 GOTO NOT
@REM --------------------------
schtasks /CREATE /RU "MATT 01" /RP " " /TN "\SYSTEM STARTER.BAT" /TR "C:\24_SYSTEM_START.BAT" /SC ONSTART /RL HIGHEST /F
@REM --------------------------
:NOT

@REM GOTO HERE22 __ TESTING CODER



@REM --------------------------
@schtasks /QUERY /TN "\SYSTEM_DELAYED_LOADER.EXE" | find "ERROR: The system cannot find the file spec@IFied." > nul
@IF NOT %errorlevel% EQU 0 GOTO NOT
@REM --------------------------
schtasks /CREATE /RU "MATT 01" /RP " " /TN "\SYSTEM_DELAYED_LOADER.EXE" /TR "D:\VB6\VB-NT\00_Best_VB_01\DELAYED_LOADER\SYSTEM_DELAYED_LOADER.exe" /SC ONLOGON /RL HIGHEST /F
@REM --------------------------
:NOT

@REM /IT SWITCH
@REM MAYBE RUN WITH USER LOGGED IN @SETTING UNABLE FIND




@REM -----------------------------------------------------------------------
@REM REQUIRE A QUERY CHECK AND CHANGE THAN CREATE AS @SETTING ARE OVER WRITER  
@REM -----------------------------------------------------------------------


@REM --------------------------------------------------------------------
@REM "C:\Windows\system32\mmc.exe" "C:\Windows\system32\taskschd.msc" /s
@REM RUN HERE TO LOAD TASK SCHEDULER AND WILL REQUIRE
@REM @SET PROGRAM TO LOAD ONLY WHEN LOGGED IN
@REM ALSO THEIR OTHER 2 OF 3 @SET
@REM DON'T KNOW THAT BIT
@REM --------------------------------------------------------------------

@REM  @ECHO ON
@ECHO OFF
@REM  ------------------------------------------------------------------------
@REM  COMMAND /F FORCE TO CHANGE THE NEW @SETTING EVEN WHEN EXIST AS A SCHEDULE
@REM  ------------------------------------------------------------------------


@REM  -------------------------------------------------
@REM  ALL COMPUTER WIN 10 OR MAYBE WIN 07
@REM  -------------------------------------------------
@ECHO ON
@ECHO OFF
@IF NOT "%WINXP_%"=="YES" GOTO JUMPER_NOT_XP_02

@REM ------------------------------------------------
@REM BLOCK OF LINE ADD IN FROM WINDOWS STARTER ACER ONE
@REM ------------------------------------------------
@REM WINDOWS 7 STARTER ACER ONE
@REM ------------------------------------------------
@REM wuauserv
@REM WINDOWS UPDATE HOW DISABLE THE SERVER WITH TASKILL SO NOT AUTO COME BACK
@REM C:\Windows\system32\svchost.exe -k netsvcs
@REM FIND LATER IN HERE FILE







TASKKILL /F /IM LiveUpdate.exe /T
@REM --

TASKKILL /F /IM CNMNSST.exe /T
@REM --

TASKKILL /F /IM wcourier.exe /T
@REM -- A WIRELESS NET TOGGGLER TASK AREA ICON
@REM -- WIN 7

TASKKILL /F /IM vcsFPService.exe /T
@REM -- C:\Windows\System32\vcsFPService.exe
@REM -- WIN 7

TASKKILL /F /IM asusfpservice_x64.exe /T
@REM -- 
@REM -- WIN 7

TASKKILL /F /IM usbchargerplus.exe /T
@REM -- 
@REM -- WIN 7



@REM  ======================================== ====================== ===============
@ECHO ON
@REM   @ECHO OFF
schtasks /change /TN "\Microsoft\Windows\Windows Defender\Windows Defender Scheduled Scan" /DISABLE
schtasks /change /TN "\Microsoft\Windows\Windows Defender\Windows Defender Cleanup" /DISABLE
schtasks /change /TN "\Microsoft\Windows\Windows Defender\Windows Defender Scheduled Scan" /DISABLE
schtasks /change /TN "\Microsoft\Windows\Windows Defender\Windows Defender Ver@IFication" /DISABLE
@ECHO ON
@ECHO OFF
@REM  ======================================== ====================== ===============
@REM  Windows Defender Cache Maintenance       N/A                    Ready
@REM  Windows Defender Cleanup                 N/A                    Ready
@REM  Windows Defender Scheduled Scan          N/A                    Ready
@REM  Windows Defender Ver@IFication            N/A                    Ready
@REM  ======================================== ====================== ===============

@ECHO ON
@REM  @ECHO OFF
schtasks /Create /TN "\Adobe Acrobat Update Task" /TR "'C:\Program Files (x86)\Common Files\Adobe\ARM\1.0\AdobeARM.exe'" /SC DAILY /MO 2 /ST 00:00:00 /F

schtasks /Create /TN "\CCleaner Update" /TR "'C:\Program Files\CCleaner\CCUpdate.exe'" /SC DAILY /MO 2 /ST 00:00:00 /F

@REM ___ Empty trigger not ever run
@REM schtasks /Create /TN "\CCleanerSkipUAC" /TR "'C:\Program Files\CCleaner\CCUpdate.exe' $(Arg0)" /SC DAILY /MO 2 /ST 00:00:00 /F

for /f "tokens=2 delims=\" %%x in ('schtasks /query /fo list ^| findstr /C:"OneDrive Standalone Update Task"') do schtasks /delete /tn "%%x" /F

schtasks /Create /TN "\OneDrive Standalone Update Task" /TR "'%localappdata%\Microsoft\OneDrive\OneDriveStandaloneUpdater.exe'" /SC DAILY /MO 2 /ST 00:00:00 /F

for /f "tokens=2 delims=\" %%x in ('schtasks /query /fo list ^| findstr /C:"User_Feed_Synchronization"') do schtasks /delete /tn "%%x" /F

schtasks /Create /TN "\User_Feed_Synchronization" /TR "'C:\Windows\system32\msfeedssync.exe' sync" /SC DAILY /MO 1 /ST 00:00:00 /F


@REM Source CREDIT ----
@REM MAKE A SCHEDULE COMMAND LINE EVERY 2 DAYS - Google Search
@REM https://www.google.co.uk/search?q=MAKE+A+SCHEDULE+COMMAND+LINE+EVERY+2+DAYS&oq=MAKE+A+SCHEDULE+COMMAND+LINE+EVERY+2+DAYS&aqs=chrome..69i57.16700j0j7&sourceid=chrome&ie=UTF-8
@REM --------
@REM How to schedule a windows task to repeat every day within a time range? - Super User
@REM https://superuser.com/questions/345319/how-to-schedule-a-windows-task-to-repeat-every-day-within-a-time-range
@REM --------
@REM find a sceduled task and delete it - Google Search
@REM https://www.google.co.uk/search?q=find+a+sceduled+task+and+delete+it&oq=find+a+sceduled+task+and+delete+it&aqs=chrome..69i57.12765j0j7&sourceid=chrome&ie=UTF-8
@REM --------
@REM Wanted: Delete Scheduled Tasks that contain certain words script - Windows Forum - Spiceworks
@REM https://community.spiceworks.com/topic/1883582-wanted-delete-scheduled-tasks-that-contain-certain-words-script
@REM ----

schtasks /change /TN "\ASUS Live Update1" /ST 00:00:00 /RI 2880 /DU 48:01
schtasks /change /TN "\ASUS Live Update2" /ST 00:00:00 /RI 2880 /DU 48:01
schtasks /change /TN "\ASUS Smart Gesture Launcher" /DISABLE

@REM schtasks /change /TN "\DropboxUpdateTaskMachineCore" /ST 00:00:00 /RI 2880 /DU 48:01
@REM schtasks /change /TN "\DropboxUpdateTaskMachineUA" /ST 00:00:00 /RI 2880 /DU 48:01

schtasks /Create /TN "\DropboxUpdateTaskMachineCore" /TR "'C:\Program Files (x86)\Dropbox\Update\DropboxUpdate.exe' /c" /SC DAILY /MO 2 /ST 00:00:00 /F

schtasks /Create /TN "\DropboxUpdateTaskMachineUA" /TR "'C:\Program Files (x86)\Dropbox\Update\DropboxUpdate.exe' /ua /installsource scheduler" /SC DAILY /MO 2 /ST 00:00:00 /F

schtasks /Create /TN "\GoogleUpdateTaskMachineCore" /TR "'C:\Program Files (x86)\Google\Update\GoogleUpdate.exe' /c" /SC DAILY /MO 2 /ST 00:00:00 /F

schtasks /Create /TN "\GoogleUpdateTaskMachineUA" /TR "'C:\Program Files (x86)\Dropbox\Update\DropboxUpdate.exe' /ua /installsource scheduler" /SC DAILY /MO 2 /ST 00:00:00 /F

@REM schtasks /change /RP " " /TN "\GoogleUpdateTaskUserS-1-5-21-1585157797-3199143874-1481636937-1001UA" /ST 00:00:00 /RI 2880 /DU 48:01
@REM schtasks /change /RP " " /TN "\GoogleUpdateTaskUserS-1-5-21-1585157797-3199143874-1481636937-1001Core" /ST 00:00:00 /RI 2880 /DU 48:01

@REM  -------------------------------------------------------
@REM  DEFAULT 1440
@REM  24 HOUR MOVE UP A BIT HIGHER 
@REM  1440 * 7 = 10080 1 WEEK -- 7 DAY
@REM  -------------------------------------------------------
@REM  YES AS I TALKER LINE BELOW HAVE FOUND THE /DU /DURATION IS THE @SETTING TO CHANGE THAT ONE DAILY
@REM  GUESS ALSO IS MUST BE TO THE ROUND DAY
@REM  -------------------------------------------------------


@REM  -------------------------------------------------------
@REM  UNABLE TO CHANGE THE DAILY SCHEDULE MORE THAN 1 FREQ YET
@REM  ---------------------------------------------------------
@REM  schtasks /change /TN "\User_Feed_Synchronization-{9C3130D6-C560-4557-8BD1-BFF9344CA145}" /ST 00:00:00 /RI 2880
@REM  \User_Feed_Synchronization-{9C3130D6-C560-4557-8BD1-BFF9344CA145}

@REM  BRING THE TIME SCHED TO BEGIN OF DAY NAUGHT HOUR
@REM  /ST 00:00:00
@REM schtasks /change /RP " " /TN "\OneDrive Standalone Update Task v2" /ST 00:00:00 /RI 2880 /DU 48:01



schtasks /change /TN "\ROG Gaming Center" /DISABLE
schtasks /change /TN "\Update Checker" /DISABLE


@REM 2018-07-23
@REM ----------------------------------------------------------------
@REM WHEN NERO TUNE IT UP SERIAL KEY EXPIRED AND THEN DISABLE IT
@REM ----------------------------------------------------------------
schtasks /change /TN "\Nero TuneItUp PRO (Autopilot.exe)" /DISABLE

@REM  __ THE TASK HERE ALLOW ONLY ENABLE AND DISABLE __ FOR ME
@REM schtasks /change /TN "\User_Feed_Synchronization-{9C3130D6-C560-4557-8BD1-BFF9344CA145}" /DISABLE

@REM  ------------------------------------------------------------------------
@REM  COMMAND HELP
@REM  ------------------------------------------------------------------------
@REM  schtasks /?
@REM  @schtasks /QUERY /?
@REM  @schtasks /QUERY /FO LIST /V
@REM  schtasks /CREATE /?
@REM  schtasks /CHANGER /?
@REM  ------------------------------------------------------------------------


@REM  /DELAY   delaytime   Spec@IFies the wait time to delay the running of the
@REM  task after the trigger is fired.  The time format is
@REM  mmmm:ss.  This option is only valid for schedule types
@REM  ONSTART, ONLOGON, ONEVENT.

@REM  SYSTEM LOGON                             N/A                    Ready
@REM  SYSTEM STARTER                           N/A                    Ready

@REM  1440 = 24 HOUR
@REM  40320 = 28 DAY
@REM  40320 = 4 WEEK
@REM  10080 = 1 WEEK

@REM  MAXIMUM = 599940

@REM  @schtasks /QUERY /?
@REM  Folder: \
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  Adobe Acrobat Update Task                10/04/2017 00:00:00    Ready
@REM  ASUS Live Update1                        01/04/2017 00:00:00    Ready
@REM  ASUS Live Update2                        01/04/2017 00:00:00    Ready
@REM  ASUS Smart Gesture Launcher              N/A                    Disabled
@REM  ASUS Splendid ACMON                      N/A                    Running
@REM  ASUS USB Charger Plus                    N/A                    Ready
@REM  ATK Package 36D18D69AFC3                 N/A                    Ready
@REM  ATK Package A22126881260                 N/A                    Ready
@REM  CCleanerSkipUAC                          N/A                    Ready
@REM  CreateExplorerShellUnelevatedTask        N/A                    Ready
@REM  DELAYED_LOADER.exe                       N/A                    Ready
@REM  DropboxUpdateTaskMachineCore             25/04/2017 00:00:00    Ready
@REM  DropboxUpdateTaskMachineUA               25/04/2017 00:00:00    Ready
@REM  GoogleUpdateTaskMachineCore              15/04/2017 00:00:00    Ready
@REM  GoogleUpdateTaskMachineUA                15/04/2017 00:00:00    Ready
@REM  GoogleUpdateTaskUserS-1-5-21-1585157797- 20/03/2017 04:31:55    Ready
@REM  GoogleUpdateTaskUserS-1-5-21-1585157797- 19/03/2017 19:31:55    Ready
@REM  iSCSIAgentAutoStartup                    N/A                    Running
@REM  MyCloudNASConnectAutoStartup             N/A                    Running
@REM  NetBak-7-ASUS-GL522VW-MATT 01-AutoStartu N/A                    Running
@REM  Norton WSC Integration                   N/A                    Ready
@REM  OneDrive Standalone Update Task v2       20/04/2017 00:00:00    Ready
@REM  Open URL by RoboForm                     N/A                    Ready
@REM  Process Explorer-7-ASUS-GL522VW-MATT 01  N/A                    Running
@REM  ROG Gaming Center                        N/A                    Running
@REM  Run RoboForm TaskBar Icon                N/A                    Ready
@REM  SYSTEM LOGON                             N/A                    Ready
@REM  SYSTEM STARTER                           N/A                    Ready
@REM  Update Checker                           N/A                    Disabled
@REM  User_Feed_Synchronization-{9C3130D6-C560 20/03/2017 00:47:44    Ready
@REM  {51E703E0-32FD-488C-ABBD-3EE10DD08298}   N/A                    Ready


@REM  1 WEEK
@REM  ------
@REM  schtasks /change /TN "\Microsoft\Windows\WindowsUpdate\Automatic App Update" /ST 00:00:00 /DU 10080
@REM  UNABLE TO WORK AT MOMENT LIKE BE LONGER INTERVAL

@REM  Folder: \Microsoft\Windows\WindowsUpdate
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  Automatic App Update                     19/03/2017 20:29:58    Ready
@REM  Scheduled Start                          N/A                    Ready
@REM  sih                                      20/03/2017 06:08:24    Ready
@REM  sihboot                                  N/A                    Ready

@REM  ----------------------------------------------------------------
@REM  UNABLE @SET THE START TIME IS NOT WANT WITHER XML FORMAT IN THERE
@REM  schtasks /change /TN "\Microsoft\Windows\Windows Error Reporting\QueueReporting" /ST 00:00:00 /DU
@REM  ----------------------------------------------------------------
@REM  ------------------------------------------------------------------------
@REM  NERO TALK THIS IMPORTANT SERVICE MUST Of CHANGE AND LEAVE GOING
@REM  Windows Error Reporting Service
@REM  MAYBE ABOVE schtasks IS OKAY LESS FREQUENT OR NOT
@REM  ----------------------------------------------------------------
schtasks /change /TN "\Microsoft\Windows\Windows Error Reporting\QueueReporting" /DU 10080
@REM 4 hours
@REM  ----------------------------------------------------------------

@REM  Folder: \Microsoft\Windows\Windows Error Reporting
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  QueueReporting                           19/03/2017 22:29:57    Ready

@REM  Folder: \Microsoft\Windows\User Profile Service
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  HiveUploadTask                           N/A                    Disabled

@REM  UNABLE @SETTER
@REM  schtasks /change /TN "\Microsoft\Windows\UpdateOrchestrator\Schedule Scan" /ST 00:00:00 /DU 10080

@REM  Folder: \Microsoft\Windows\UpdateOrchestrator
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  Maintenance Install                      N/A                    Disabled
@REM  Policy Install                           N/A                    Disabled
@REM  Reboot                                   N/A                    Ready
@REM  Refresh @SETtings                         20/03/2017 12:30:56    Ready
@REM  Resume On Boot                           N/A                    Disabled
@REM  Schedule Scan                            20/03/2017 16:52:02    Ready
@REM  USO_UxBroker_Display                     N/A                    Ready
@REM  USO_UxBroker_ReadyToReboot               N/A                    Ready

@REM  Folder: \Microsoft\Windows\SystemRestore
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  SR                                       N/A                    Ready

@REM  Folder: \Microsoft\Windows\Sysmain
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  HybridDriveCachePrepopulate              N/A                    Disabled
@REM  HybridDriveCacheRebalance                N/A                    Disabled
@REM  ResPriStaticDbSync                       N/A                    Ready
@REM  WsSwapAssessmentTask                     N/A                    Ready

@REM  \Microsoft\Windows\Speech

@REM  Folder: \Microsoft\Windows\Speech
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  SpeechModelDownloadTask                  20/03/2017 00:00:00    Ready

@REM  Folder: \Microsoft\Windows\Shell
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  CreateObjectTask                         N/A                    Ready
@REM  FamilySafetyMonitor                      N/A                    Ready
@REM  FamilySafetyMonitorToastTask             N/A                    Disabled
@REM  FamilySafetyRefreshTask                  N/A                    Ready
@REM  IndexerAutomaticMaintenance              N/A                    Ready

@REM  Folder: \Microsoft\Windows\SharedPC
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  Account Cleanup                          N/A                    Disabled

@REM  Folder: \Microsoft\Windows\Maps
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  MapsToastTask                            N/A                    Ready
@REM  MapsUpdateTask                           N/A                    Disabled

@REM  Folder: \Microsoft\Windows\File Class@IFication Infrastructure
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  Property Definition Sync                 N/A                    Disabled

@REM  Folder: \Microsoft\Windows\DiskDiagnostic
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  Microsoft-Windows-DiskDiagnosticDataColl N/A                    Disabled
@REM  Microsoft-Windows-DiskDiagnosticResolver N/A                    Disabled

@REM  Folder: \Microsoft\Windows\Defrag
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  ScheduledDefrag                          N/A                    Disabled


@REM  EVERY 4 WEEK
@REM Every 6 hour really on checking
schtasks /change /TN "\Microsoft\Windows\Customer Experience Improvement Program\Consolidator" /DU 40320

@REM  Folder: \Microsoft\Windows\Customer Experience Improvement Program
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  Consolidator                             20/03/2017 00:00:00    Ready
@REM  KernelCeipTask                           N/A                    Ready
@REM  UsbCeip                                  N/A                    Ready

@REM  Folder: \Microsoft\Windows\Chkdsk
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  ProactiveScan                            N/A                    Ready

@REM  EVERY 4 WEEK
schtasks /change /TN "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser" /ST 00:00:00 /DU 40320

@REM  Folder: \Microsoft\Windows\Application Experience
@REM  TaskName                                 Next Run Time          Status
@REM  ======================================== ====================== ===============
@REM  Microsoft Compatibility Appraiser        20/03/2017 04:10:23    Ready
@REM  ProgramDataUpdater                       N/A                    Ready
@REM  StartupAppTask                           N/A                    Ready



@REM  WORK TO FIND HOW @SET THE SCHEDULER TO CHANGE THE DAY FREQ NOT JUST ONLY THE MINUTE 
@REM  NOT FOUND OUT IN WHOLE OF DAY
@REM  ----
@REM  windows - Using the DOS "start" command with parameters passed to the started program - Stack Overflow
@REM  http://stackoverflow.com/questions/154075/using-the-dos-start-command-with-parameters-passed-to-the-started-program
@REM  --------
@REM  xml task scheduler - Google Search
@REM  https://www.google.co.uk/search?num=50&q=xml+task+scheduler&oq=XML+TASK+&gs_l=serp.1.5.0l3j0i67k1j0l6.168685.173343.0.179193.42.16.0.0.0.0.480.1888.3-4j1.5.0....0...1c.1.64.serp..37.4.1552.0..0i131k1.5SkNlzJRhrg
@REM  --------
@REM  how change the days interval scheduler - Google Search
@REM  https://www.google.co.uk/search?num=50&q=how+change+the+days+interval+scheduler&oq=HOW+CHANGE+THE++DaysInterval+SCHEDUALLAR&gs_l=serp.1.0.30i10k1.2625.15501.0.19517.47.27.0.0.0.0.743.4949.2-4j2j4j1j1.12.0....0...1c.1.64.serp..35.11.4604.0..0j35i39k1j0i131k1j0i67k1j33i21k1j33i160k1._nE9A1YDr5Y
@REM  --------
@REM  Change Internet Time Update interval in Windows 10 / 8 / 7
@REM  http://www.thewindowsclub.com/change-internet-time-update-interval-windows
@REM  --------
@REM  Create task in to windows task scheduler through XML file | AuthorCode
@REM  http://www.authorcode.com/create-task-in-to-windows-task-scheduler-through-xml-file/
@REM  --------
@REM  how @SET a task schedule disable - Google Search
@REM  https://www.google.co.uk/search?q=how+@SET+a+task+schedule+disable&gws_rd=ssl
@REM  --------
@REM  Windows Automatic Maintenance: Run, Stop, Disable, Schedule
@REM  http://www.thewindowsclub.com/automatic-maintenance-windows-8
@REM  --------
@REM  Rename scheduled task in Windows Task Scheduler
@REM  http://www.thewindowsclub.com/rename-scheduled-tasks-windows-task-scheduler
@REM  --------
@REM  Add Email to Google Tasks from Gmail with RightTasks
@REM  http://www.thewindowsclub.com/add-email-to-google-tasks-righttasks
@REM  --------
@REM  Working with Task Scheduler XML – Hey, Scripting Guy! Blog
@REM  https://blogs.technet.microsoft.com/heyscriptingguy/2013/12/14/working-with-task-scheduler-xml/
@REM  --------
@REM  How to Automatically Run Programs and @SET @REMinders With the Windows Task Scheduler
@REM  https://www.howtogeek.com/123393/how-to-automatically-run-programs-and-@SET-@REMinders-with-the-windows-task-scheduler/
@REM  --------
@REM  Schedule tasks from command line
@REM  https://www.windows-commandline.com/schedule-tasks-command-line/
@REM  --------
@REM  Task Scheduling and VBScript - Scripts to add, @REMove and edit a Task Schedule
@REM  http://www.activexperts.com/admin/vbscript-collection/taskscheduling/
@REM  --------
@REM  Disabling Task Scheduler in win7 - Windows 7 Help Forums
@REM  https://www.sevenforums.com/general-discussion/69801-disabling-task-scheduler-win7.html
@REM  --------
@REM  How to disable a windows schedule task in Windows 10 using the utility schtasks? - Super User
@REM  https://superuser.com/questions/1034673/how-to-disable-a-windows-schedule-task-in-windows-10-using-the-utility-schtasks
@REM  --------
@REM  Batch files - SCHTASKS
@REM  http://www.robvanderwoude.com/schtasks.php
@REM  --------
@REM  Rob van der Woude's Visual Studio Store
@REM  http://www.robvanderwoude.com/storevisualstudio.php
@REM  --------
@REM  ChangeDetection - Monitor a Web Page for Changes
@REM  https://www.changedetection.com/detect.html
@REM  --------
@REM  Microsoft Visual Studio Enterprise 6.0: Amazon.co.uk: Software
@REM  https://www.amazon.co.uk/gp/product/B00004U7PM/ref=ox_sc_mini_detail?ie=UTF8&psc=1&smid=A2D4FJ9J0DN1QW
@REM  --------
@REM  Microsoft Visual BASIC 6.0 - Google Search
@REM  https://www.google.co.uk/search?num=50&biw=1211&bih=502&tbm=shop&q=Microsoft+Visual+BASIC+6.0&oq=Microsoft+Visual+BASIC+6.0&gs_l=serp.3...36554.41114.0.41684.22.15.0.0.0.0.516.2631.0j1j7j1j0j1.10.0....0...1c.1.64.serp..14.3.783...30i10k1._-a4Swv1_mk
@REM  --------
@REM  eBay Global Search
@REM  http://www.geo-ship.com/?c_keywords=Microsoft%20Visual%20BASIC%206.0&c_keywordsTranslated=Microsoft%20Visual%20BASIC%206.0&c_priceRange_min=-1&c_priceRange_max=-1&c_sellerName=&c_category=See-All-Categories&c_priceOrgCurrencyRange_min=-1&c_priceOrgCurrencyRange_max=-1&a_selectedLocales=EBAY-US,EBAY-GB,EBAY-AU,EBAY-DE,EBAY-FR,EBAY-IT,EBAY-NLBE,EBAY-AT,EBAY-ES,EBAY-ENCA,EBAY-IE,EBAY-NL,EBAY-CH,EBAY-PL,EBAY-SG,EBAY-MOTOR&a_searchDesc=false&a_searchPaypalOnly=true&a_showShippingType=true&a_bestOfferOnly=false&a_searchCompleted=false&a_searchSoldOnly=false&a_onlyWithShipping=false&a_@SETPricerangeTotal=false&a_bidsRange_min=-1&a_bidsRange_max=-1&a_feedbackRange_min=-1&a_feedbackRange_max=-1&a_positiveFeedbackRange_min=-1&a_positiveFeedbackRange_max=-1&a_shippingPriceRatio=0&a_buyingFormat=BIN&a_itemCondition=All&a_availableTo=GB&a_locatedIn=ANY&a_resultsPerPage=100&a_ebaySortOrder=PricePlusShippingLowest&a_sellerType=All
@REM  --------
@REM  Microsoft Visual Basic 6.0 w/ Key Professional Visual Studio 6.0 Plus Pack | eBay
@REM  http://www.ebay.com/itm/Microsoft-Visual-Basic-6-0-w-Key-Professional-Visual-Studio-6-0-Plus-Pack-/222436720731?rmvSB=true
@REM  --------
@REM  Microsoft Visual Basic 6.0 Professional Edition (w/ Product Key) pc software 659556149239 | eBay
@REM  http://www.ebay.com/itm/Microsoft-Visual-Basic-6-0-Professional-Edition-w-Product-Key-pc-software/112333291091?_trksid=p2047675.c100005.m1851&_trkparms=aid%3D222007%26algo%3DSIC.MBE%26ao%3D2%26asc%3D20131003132420%26meid%3Dcb621a01beb141a49007916a386e90a7%26pid%3D100005%26rk%3D2%26rkt%3D6%26sd%3D222436720731
@REM  --------
@REM  Microsoft Visual Studio Enterprise 6.0 VE VS 6 C++ Basic FoxPro J++ = 3 CD&#039;s | eBay
@REM  http://www.ebay.com/itm/322454382925?ssPageName=STRK%3AMEBIDX%3AIT&fromMakeTrack=true
@REM  ----

@REM  </> END BLOCKER SCHEDULER 
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------------

:JUMPER_NOT_XP_02

@REM   ------------------------------------------------
@REM   BLOCK OF LINE ADD IN FROM WINDOWS STARTER ACER ONE
@REM   ------------------------------------------------
@REM   WINDOWS 7 STARTER ACER ONE
@REM   ------------------------------------------------
@REM   wuauserv
@REM   WINDOWS UPDATE HOW DISABLE THE SERVER WITH TASKILL SO NOT AUTO COME BACK
@REM   C:\Windows\system32\svchost.exe -k netsvcs
@REM   FIND LATER IN HERE FILE


TASKKILL /F /IM LiveUpdate.exe /T
TASKKILL /F /IM CNMNSST.exe /T
TASKKILL /F /IM wcourier.exe /T
@REM   -- A WIRELESS NET TOGGLER TASK AREA ICON
@REM   -- WIN 7
TASKKILL /F /IM vcsFPService.exe /T
@REM   -- C:\Windows\System32\vcsFPService.exe
@REM   -- WIN 7
TASKKILL /F /IM asusfpservice_x64.exe /T
@REM   -- 
@REM   -- WIN 7
TASKKILL /F /IM usbchargerplus.exe /T
@REM   -- 
@REM   -- WIN 7
TASKKILL /F /IM System56Link.exe /T
@REM -- C:\Program Files (x86)\ASUS\ASUS Manager\Power Saver\System56Link.exe
@REM -- WIN 7




@REM   ------------------------------------------------------------------------
@REM   -- WIN 7
@REM   ------------------------------------------------------------------------
@ECHO OFF
@REM   @IF "%WIN10__%"=="YES" GOTO WINDOW_10
@REM   ------------------------------------------------------------------------
@REM   ------------------------------------------------------------------------


@ECHO ON
@ECHO OFF

@REM   __ __________________________________________________________________
@REM   __ ONLY FOR ____ 5-ASUS-PL2520LA BUTTER AND TASKKILL ANY COMPUTER
@REM   __ ________________________________________________________
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ ________________________________________________________
TASKKILL /F /IM BackgroundWindow.exe /T
@REM   __ ________________________________________________________
@REM   __ C:\Program Files (x86)\ASUS\ASUS Manager\BackgroundWindow.exe
@REM   __ ________________________________________________________
TASKKILL /F /IM SecureDeleteBackground.exe /T
@REM -- C:\Program Files (x86)\ASUS\ASUS Manager\PC Cleanup\SecureDeleteBackground.exe
@REM -- WIN 7

TASKKILL /F /IM SecureDeleteWindow.exe /T
@REM -- C:\Program Files (x86)\ASUS\ASUS Manager\PC Cleanup\SecureDeleteWindow.exe
@REM -- WIN 7

@REM -- TASKKILL /F /IM NclUSBSrv64.exe /T
@REM -- C:\Program Files (x86)\PC Connectivity Solution\Transports\NclUSBSrv64.exe
@REM -- WIN 7 -- USB MEDIA SERVER

TASKKILL /F /IM LWEMon.exe /T
@REM -- C:\Program Files\Logitech\Gaming Software\LWEMon.exe
@REM -- WIN 7

TASKKILL /F /IM AsLdrSrv.exe /T
@REM -- C:\Program Files (x86)\ASUS\ATK Package\ATK Hotkey\AsLdrSrv.exe
@REM -- NOT SURE WIN 7

@REM -- BLUETOOTH MIRCOSOFT -- C:\Program Files (x86)\PC Connectivity Solution\Transports\NclMSBTSrvEx.exe
@REM -- NOT SURE WIN 7

TASKKILL /F /IM AsusTPLoader.exe /T
@REM -- C:\Program Files (x86)\ASUS\ASUS Smart Gesture\AsTPCenter\x64\AsusTPLoader.exe
@REM -- NOT SURE WIN 7

TASKKILL /F /IM QuickGesture.exe /T
@REM -- C:\Program Files (x86)\ASUS\ASUS Smart Gesture\QuickGesture\x86\QuickGesture.exe
@REM -- NOT SURE WIN 7

TASKKILL /F /IM AsusTPHelper.exe /T
@REM -- C:\Program Files (x86)\ASUS\ASUS Smart Gesture\AsTPCenter\x64\AsusTPHelper.exe
@REM -- NOT SURE WIN 7

TASKKILL /F /IM SASrv.exe /T
@REM -- C:\Windows\SysWOW64\SASrv.exe
@REM -- LET BOOT FIRST AND SETTING HAPPEN WIN 7

TASKKILL /F /IM ACMON.exe /T
@REM -- C:\Program Files (x86)\ASUS\Splendid\ACMON.exe
@REM -- WIN 7
@REM -- SEARCH INTERNET 
@REM -- ACMON.exe Windows process - What is it? - File.net
@REM -- www.file.net/process/acmon.exe.html
@REM -- ACMON.exe file information. ... Description: ACMON.exe is not essential for Windows and will often cause problems. ... The process ACMON belongs to the software ASUS Splendid Video by ATK.

@REM -- -------------------------- -------------------------------------------
@REM -- ATKOSD2.exe -- OKAY TO USE
@REM -- ON SCREEN DISPLAY OF volume and display brightness setting information
@REM -- C:\Program Files (x86)\ASUS\ATK Package\ATKOSD2\ATKOSD2.exe
@REM -- USE UNINSTALL
@REM -- ----
@REM -- ATKOSD2 by ASUS - Should I @REMove It?
@REM -- http://www.shouldi@REMoveit.com/ATKOSD2-9057-program.aspx
@REM -- ----
@REM -- ----
@REM -- G74sx ATKOSD2.exe
@REM -- https://rog.asus.com/forum/showthread.php?44806-G74sx-ATKOSD2-exe
@REM -- ----
@REM -- 
@REM -- DEPENDANT AND SHOULD CLOSE BEFORE UNINSTALL HAPPEN
@REM -- 
@REM -- DMedia.EXE
@REM -- C:\Program Files (x86)\ASUS\ATK Package\ATK Media\DMedia.exe
@REM -- DONT KNOW WHAT DOES

@REM -- ATKGFNEX -- SERVICE
@REM -- 

@REM -- MORE INFO BOTTOM THREAD REPLACE ATKOSD3.exe TO ATKOSD2.exe
@REM -- ----
@REM -- G74sx ATKOSD2.exe
@REM -- https://rog.asus.com/forum/showthread.php?44806-G74sx-ATKOSD2-exe
@REM -- -------------------------- -------------------------------------------

TASKKILL /F /IM SmartAudio.exe /T
@REM -- C:\Program Files\CONEXANT\SAII\SmartAudio.exe
@REM -- TEST WIN 7

TASKKILL /F /IM flvga_tray.exe /T
@REM -- C:\Windows\System32\flvga_tray.exe
@REM -- LOOK LIKE USB TO VGA ADAPTOR SOFTWARE
@REM -- ----
@REM -- flvga_tray.exe - Google Search
@REM -- https://www.google.co.uk/search?q=flvga_tray.exe&gws_rd=cr&ei=vG4XWMyBLaSLgAae64fABg
@REM -- ----
@REM -- TEST WIN 7

TASKKILL /F /IM FIRSTRUN.EXE /T
@REM -- C:\Program Files (x86)\Microsoft Office\Office15\FIRSTRUN.EXE
@REM -- OFFICE ADVERT
@REM -- WIN 10

@REM if a service, use NET STOP
@REM --------------------------

@ECHO.
@ECHO THESE EXTRA APP / SERVICE MIGHT NOT WANT TO BE DELETE
@ECHO -----------------------------------------------

@REM IF "%1"=="AT_BOOTER" GOTO ENDING
@REM --------------------------------

@REM PAUSE 


@REM --------------------------------
@REM ABOUT THE --NET STOP-- INFO NOTE 
@REM SOME ARE SORTED IN ORDER
@REM TO MAKE EASY LOOK UP -- LINE UP 
@REM -- TO DO LATER
@REM --------------------------------

@ECHO Enables you to send and receive faxes, using fax resources available on this computer
@REM ----------------------------------------------------------
@REM MANUAL STOP
@REM FAX TELEPHONE
@REM Enables you to send and receive faxes, using fax resources available on this computer or on the 
@IF "%WINXP__%"=="YES" GOTO NOT_WHEN_WIN_XP
@IF "%WIN7___%"=="YES" GOTO NOT_WHEN_WIN_07

NET STOP Fax
@REM C:\windows\system32\fxssvc.exe

:NOT_WHEN_WIN_XP
:NOT_WHEN_WIN_07

@ECHO TO DO WITH FLICKR YAHOO
@REM ----------------------------------------------------------
@REM TASKKILL /F /IM Sync.exe /T

@ECHO -- PAID FOR 
@ECHO C:\Users\MATT 01\AppData\Local\FlickrUploadrWindows\Update.exe" --processStart Flickr.exe
@REM ----------------------------------------------------------
@REM TASKKILL /F /IM Update.exe /T
@REM -- "C:\Users\MATT 01\AppData\Local\FlickrUploadrWindows\Update.exe" --processStart Flickr.exe
@REM -- PAID FOR 
@ECHO TO DO WITH FLICKR YAHOO

@REM   __ ________________________________________________________
@REM   __ C:\Program Files (x86)\ASUS\ASUS Manager\PC Cleanup\SecureDeleteBackground.exe
@REM   __ D@IFFERCULT TO STOP THIS TASK AND END TASK NOT AVAILABLE LIKE OTHER 2
@REM   __ SO CONFIGER STOP ON AC POWER TO TESTER


@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ ________________________________________________________
@REM   __ THIS REQUIRE KILL FROM CERTAIN FOLDER
@REM   __ WORKING ON THAT ONE
@REM   __ ________________________________________________________
@REM   __ TASKKILL /F /IM *USB Lock\svchost.exe /T
@REM   __ ________________________________________________________
@REM   __ C:\Program Files (x86)\ASUS\ASUS Manager\USB Lock\svchost.exe

@REM   __ ________________________________________________________
@REM   __ EXTRA INFO
@REM   __ ________________________________________________________
@REM   __ E:\01 Start Menu\#_5-ASUS-PL2520LA\Programs\Startup-01-Delayed\5-ASUS-P2520LA--MATT 01\# 00 MOD@IFY WORK DEFAULT @SETTING\#0 ____ W@IFI HOT SPOT __ DeviceTool IS ONLY ONE WORTH RUNNING FROM HERE __ I SHUT REST DOWN.exe.txt
@REM   __ ________________________________________________________


@REM   __ ________________________________________________________
@REM   __ THESE BUNCH OF TASK HAVE TO BE STOP IN SCHEDULER
@REM   __ ONLY FOR THE \\5-ASUS-P2520LA\
@REM   __ ________________________________________________________

@REM   __ ________________________________________________________
@REM   __ ALSO TODAY BEEN WORKING ON 
@REM   __ CONTROL PANEL __ ADMINISTRATOR  AND GROUP POLICY
@REM   __ FOUND A WORKING METHOD TO ADMINISTRATOR EVERYTHING
@REM   __ @SET ALL THE SUCCESS GOING ENABLED
@REM   __ ALSO @SET WHAT USER NAME AND ADMINISTRATOR ARE ALLOW THE ACCESS
@REM   __ SHOULD BE AS MUCH FOR THAT ONE
@REM   __ AND REBOOT AFTER
@REM   __ ONLY @IF GOT A PRO VERSION OF WINDOW TO ACCESS THE POLICY
@REM   __ I FOUND ERROR AS UPDATE WAS NOT ALLOW WORKING AS NOT ADMINISTRATOR  APPARENTLY
@REM   __ ________________________________________________________
@REM   __ 
@REM   __ ALSO TODAY CONFIGURE @SETTING UPPER
@REM   __ NOTEPAD PLUG_IN WHAT DON'T WANT BY DEFAULT OR DO WANT
@REM   __ SEARCH AROUND FOR THESE HARDCORE
@REM   __ ________________________________________________________
@REM   __ 1..
@REM   __ LOCATE IN NOTEPAD++ DEFAULT IS USE HIGHLIGHT MARK LINE CHANGER RATHER NOT AND REST IS GOOD JUMP LAST POSITION
@REM   __ 2..
@REM   __ NOTEPAD++ SPELL CHECK USE ROUNDED BOX OR SQUARE _ BOX FOR HIGHLIGHTER BEST
@REM   __ 3..
@REM   __ NOTEPAD++ AUTOSAVE MIGHT BE A BIT MUCH WORKER WHEN SAVE ONE SAVE THEM ALL UN-ENABLE
@REM   __ ________________________________________________________
@REM   __ 4..
@REM   __ NOTEPAD++ PLUGINS RUNME REQUIRE ADDITIONAL OPTION LIKE SHOW TOOL-BAR OTHER MENU OF 3 AND 
@REM   __ LEARN TO DO HOT KEY SH@IFT-F5
@REM   __ PUT AT TOP CONTEXT MENU
@REM   __ RIGHT CLICK ON TAB SHOW CONTEXT MENU
@REM   __ ________________________________________________________
@REM   __ AND ALSO SAVE BEFORE RUN
@REM   __ ________________________________________________________
@REM   __ 5..
@REM   __ NOTEPAD++ PREFERENCES RECENT FILE HISTORY MENU @SET MAX LENGTH FILE-NAME ABOUT 150 @IF LOW RESOLUTION
@REM   _________________________________________________________
@REM   __ +..
@REM   __ NOTEPAD++ PREFERENCES USE THE CLOUD OPTION @SETTER A DIRECTORY
@REM   _________________________________________________________
@REM   __ +..
@REM   __ NOTEPAD++ PREFERENCES DEFAULT DIRECTORY __ USE THE NEW SAVE AS DIALOGUE 
@REM   _________________________________________________________
@REM   __ +..
@REM   __ NOTEPAD++ STYLE CONFIGURATION HAS THEME FOR DARKER COLOURT
@REM   _________________________________________________________
@REM   __ +..
@REM   __ NOTEPAD++ SUMMARY GIVE FIGURE COUNT INFO FOR DOCUMENT LINE CHARACTER AND WORD
@REM   _________________________________________________________
@REM   __ +..
@REM   __ NOTEPAD++ PREFERENCES __ GRNERAL __ TOOLBAR __ STANDARD ICON SEEM BETTER BUT NOT BIGGER
@REM   __ GET THEM ALL ON
@REM   __ AND SHELL EXECUTE IS NEAR END
@REM   _________________________________________________________
@REM   __ +.. AND TOPICA ABOVE LINE
@REM   __ WHEN GET THEM ICON TOOLBAR SORTED __ DO THE @SET RUN 
@REM   __ THIS TYPE OF FILE AS COMMAND SHELLING AND WITH 
@REM   __ @SET A CLOSE BY HOT KEY LIKE SH@IFT AND / 
@REM   GET A EXT@REMELY QUICKER RUNNER GOING
@REM   _________________________________________________________




GOTO JUMPER_TEXT_PART_NOTEPAD_INFO
@REM __ ________________________________________________________
@REM __ ________________________________________________________
@REM __ ________________________________________________________
@REM __ ________________________________________________________
@REM FOR BACKUP OF NOTEPAD INI AND @SETTING
@REM THE PLUGGIN FOLDER HERE
@REM __ ________________________________________________________
@REM C:\Program Files (x86)\Notepad++\plugins
@REM __ ________________________________________________________


@REM Form FindWindow ---
@REM Debug Info
@REM ---------------------
@REM Notepad++ v7.2.2   (32-bit)
@REM Build time : Nov 27 2016 - 16:31:07
@REM Path : C:\Program Files (x86)\Notepad++\notepad++.exe
@REM Admin mode : ON
@REM Local Conf mode : OFF
@REM OS : Windows 7
@REM Plugins : autosave2.dll ComparePlugin.dll DSpellCheck2.dll Explorer.dll FallingBricks.dll HTMLTag_unicode.dll LocationNavigate.dll MenuIcons.dll mimeTools.dll NotepadStarterPlugin.dll NppCalc.dll NppConverter.dll NppExport.dll NppFTP.dll NppSaveAsAdmin.dll PluginManager.dll PythonScript.dll RunMe.dll SpeechPlugin.dll TopMostU.dll VisualStudioLineCopy.Unicode.dll ZenCoding-Python.dll 


@REM FILES IN USE NOTEPADD++ WHILE IN RUNNER __ TRY LOCATE INI FILE MUST BE AT STARTER
@REM 01 TO 07
@REM ------------------------------------------------------------------------
@REM Form FindWindow ---
@REM OpenedFilesView
@REM ---------------------
@REM comctl32.dll.mui	C:\Windows\winsxs\x86_microsoft.windows.c..-controls.resources_6595b64144ccf1df_6.0.7600.16385_en-us_581cd2bf5825dde9\comctl32.dll.mui	
@REM C:\Program Files (x86)\Notepad++\notepad++.exe	
@REM 5-ASUS-P2520LA\MATT 01	

@REM user32.dll.mui	
@REM C:\Windows\SysWOW64\en-US\user32.dll.mui	
@REM C:\Program Files (x86)\Notepad++\notepad++.exe	
@REM 5-ASUS-P2520LA\MATT 01	

@REM shell32.dll.mui	
@REM C:\Windows\SysWOW64\en-US\shell32.dll.mui	
@REM C:\Program Files (x86)\Notepad++\notepad++.exe	
@REM 5-ASUS-P2520LA\MATT 01	

@REM R000000000007.clb	
@REM C:\Windows\Registration\R000000000007.clb	
@REM notepad++.exe	
@REM C:\Program Files (x86)\Notepad++\notepad++.exe	
@REM 5-ASUS-P2520LA\MATT 01	

@REM StaticCache.dat	
@REM C:\Windows\Fonts\StaticCache.dat	
@REM notepad++.exe	
@REM C:\Program Files (x86)\Notepad++\notepad++.exe	
@REM 5-ASUS-P2520LA\MATT 01	

@REM HelpNppCalcRu_1.4.txt	
@REM C:\Program Files (x86)\Notepad++\plugins\doc\NppCalc\HelpNppCalcRu_1.4.txt	
@REM C:\Program Files (x86)\Notepad++\notepad++.exe	
@REM 5-ASUS-P2520LA\MATT 01	

@REM HelpNppCalcEn_1.4.txt	
@REM C:\Program Files (x86)\Notepad++\plugins\doc\NppCalc\HelpNppCalcEn_1.4.txt	
@REM 14/11/2016 09:43:06	
@REM 14/11/2016 09:43:06	
@REM C:\Program Files (x86)\Notepad++\notepad++.exe	
@REM 5-ASUS-P2520LA\MATT 01	
@REM ------------------------------------------------------------------------

@REM MOST THE INI FILE SEEM TO BE HERE
@REM ALSO AS TALKER ABOVE __ DON'T USE THE CLOUD TO A PERSONAL FOLDER REFERENCING 
@REM BECAUSE IT IS RESTRICTED TO STAY WITHIN PROGRAM AREA OR USER 
@REM UNLESS LEARN HOW OPEN IT ON
@REM ------------------------------------------------------------------------
@REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\ProcessActivityView.exe
@REM WITHER THIS CODE PROGRAM
@REM REQUIRED IS FOLLOW @IF MONITOR A 32 BIT CODE MUST USER 32 BIT VERSION
@REM AND NIRSOFT FRONT END HAS ABILITY TO SWAP
@REM ------------------------------------------------------------------------
@REM 1	C:\Program Files (x86)\Notepad++\plugins\APIs\normal.xml	
@REM 2	
@REM C:\Users\MATT 01\AppData\Roaming\Notepad++\plugins\Config\Hunspell\en_GB.usr	
@REM 3	
@REM C:\Users\MATT 01\AppData\Roaming\Notepad++\plugins\Config\LocationNavigate.ini	
@REM 4	
@REM C:\Users\MATT 01\AppData\Roaming\Notepad++\plugins\Config\NppFTP\NppFTP.xml	
@REM 5	
@REM C:\Users\MATT 01\AppData\Roaming\Notepad++\plugins\Config\NppFTP\Cert@IFicates.xml	
@REM 6	
@REM C:\Users\MATT 01\AppData\Roaming\Notepad++\config.xml	
@REM 7	
@REM C:\Users\MATT 01\AppData\Roaming\Notepad++\session.xml	
@REM ------------------------------------------------------------------------


@REM __ ________________________________________________________
@REM __ ________________________________________________________
@REM __ ________________________________________________________
@REM __ ________________________________________________________
:JUMPER_TEXT_PART_NOTEPAD_INFO










@REM ----------------------------------------------------------
@REM AUTO STOP
@REM ---- NOT FOUND NAME YET OF SERVICE - NOT PRESENT AT MOMENT
@REM NET STOP NBService.exe
@REM TASKKILL /F /IM peservice.exe /T


@REM   __ @SET FROM AUTORUNS
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ ________________________________________________________
TASKKILL /F /IM YMUpdater.exe /T
@REM   __ ________________________________________________________
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ OR ALL COMPUTER 
@REM   __ ________________________________________________________
@REM   __ THIS TASK YAHOO MESSENGER NOT ALL THE TIME HAVE BE WANT-ER
@REM   __ C:\Users\MATT 01\AppData\Roaming\Yahoo Messenger\YMUpdater\YMUpdater.exe
@REM   __ @REMOVE OPTION IN PROGRAM TO NOT STARTUP OR ALWAYS BE PUT BACK IN STARTUP AUTORUN REGISTRY
@REM   __ ________________________________________________________


@REM   __ @SET FROM AUTORUNS
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ ________________________________________________________
TASKKILL /F /IM xmarkssync.exe /T
@REM   __ ________________________________________________________
@REM   __ C:\Program Files (x86)\Xmarks\IE Extension\xmarkssync.exe
@REM   __ ________________________________________________________

@REM   __ @SET FROM AUTORUNS
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ ________________________________________________________
@REM TASKKILL /F /IM VideoDownloaderUltimate.exe /T
@REM   __ C:\ProgramData\VideoDownloaderUltimateWinApp\VideoDownloaderUltimate.exe
@REM   __ ________________________________________________________

@REM   __ @SET FROM AUTORUNS
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ ________________________________________________________
TASKKILL /F /IM supercopier.exe /T
@REM   __ ________________________________________________________
@REM   __ C:\Program Files\Supercopier\supercopier.exe
@REM   __ ________________________________________________________

@REM   __ @SET FROM AUTORUNS
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ ________________________________________________________
TASKKILL /F /IM PPOptimizer.exe /T
@REM   __ ________________________________________________________
@REM   __ C:\Program Files\Process Priority Optimizer\PPOptimizer.exe
@REM   __ ________________________________________________________
@REM   __ UNINSTALL


@IF "%WIN10__%"=="YES" GOTO WINDOW_NOT_WIN_10
@REM   __ @SET FROM AUTORUNS
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ DONT WANT NOKIA SUITE __ WITHER ANY COMPUTER BUT BETTER ONE
@REM   __ ________________________________________________________
TASKKILL /F /IM NokiaSuite.exe /T
@REM   __ ________________________________________________________
@REM   __ C:\Program Files (x86)\Nokia\Nokia Suite\NokiaSuite.exe
@REM   __ ________________________________________________________
@REM   __ UNINSTALL
@REM   :WINDOW_NOT_WIN_10

@IF "%WIN10__%"=="YES" GOTO WINDOW_NOT_WIN_10
@REM   __ @SET FROM AUTORUNS
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ DONT WANT FLICKR SUITE __ WITHER ANY COMPUTER BUT BETTER ONE
TASKKILL /F /IM Flickr.exe /T
@REM   __ ________________________________________________________
@REM   __ LAUNCHER 
@REM   __ C:\Users\MATT 01\AppData\Local\FlickrUploadrWindows\Update.exe
@REM   __ RUNNER
@REM   __ C:\Users\MATT 01\AppData\Local\FlickrUploadrWindows\app-1.0.1.292\Flickr.exe
@REM   __ ________________________________________________________
@REM   :WINDOW_NOT_WIN_10

@IF "%WIN10__%"=="YES" GOTO WINDOW_NOT_WIN_10
@REM   __ @SET FROM AUTORUNS
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ DONT WANT FLICKR SYNC PROGRAM __ WITHER ANY COMPUTER BUT BETTER ONE
@REM   __ ________________________________________________________
TASKKILL /F /IM Sync.exe /T
TASKKILL /F /IM peservice.exe /T
@REM -- C:\Program Files (x86)\photoSync\peservice.exe
@REM -- TO DO FLICKR
@REM -- IF A SYNC GOING THE SERVICE WILL CONTINUE UPLOAD 
@REM -- EVEN IF OTHER APP SYNC STOPPED AND QUIT
@REM --------------------------------

@REM ----------------------------------------------------------
@REM WANT THIS ON USE FOR A BIT AND DELAY AND TASKKILL
@REM DETECT WHEN START ON







TASKKILL /F /IM Update.exe /T
@REM -- "C:\Users\MATT 01\AppData\Local\FlickrUploadrWindows\Update.exe" --processStart Flickr.exe
@REM -- PAID FOR 

@REM   __ ________________________________________________________
@REM   __ ________________________________________________________
@REM   __ RUNNER
@REM   __ C:\Program Files (x86)\photoSync\Sync.exe
@REM   __ ________________________________________________________
:WINDOW_NOT_WIN_10

@REM  PAUSE

@REM   __ @SET FROM AUTORUNS
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ ________________________________________________________
TASKKILL /F /IM Qsync.exe /T
@REM   __ ________________________________________________________
@REM   __ ________________________________________________________
@REM   __ RUNNER
@REM   __ C:\Program Files (x86)\QNAP\Qsync\Qsync.exe
@REM   __ ________________________________________________________

@REM   __ @SET FROM AUTORUNS
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ ________________________________________________________
TASKKILL /F /IM  ProcessTamerTray.exe /T
@REM   __ ________________________________________________________
@REM   __ ________________________________________________________
@REM   __ RUNNER
@REM   __ C:\Program Files (x86)\ProcessTamer\ProcessTamerTray.exe
@REM   __ ________________________________________________________

@REM   __ @SET FROM AUTORUNS
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ ________________________________________________________
TASKKILL /F /IM  MSASCui.exe /T
@REM   __ ________________________________________________________
@REM   __ ________________________________________________________
@REM   __ RUNNER
@REM   __ C:\Program Files\Windows Defender\MSASCui.exe
@REM   __ ________________________________________________________

@REM   __ @SET FROM AUTORUNS
@REM   __ FOR ____ 5-ASUS-PL2520LA
@REM   __ ________________________________________________________
@REM   __ Start WingMan Profiler	Logitech WingMan Event Monitor	Logitech Inc. 04/04/2008 18:24	
@REM   __ JOYPAD CONTROLLER __ NOT REQUIRED THAT MUCH MAYBE AT @SETUP LEARN 
@REM   __ ________________________________________________________
TASKKILL /F /IM  LWEMon.exe /T
@REM   __ ________________________________________________________
@REM   __ ________________________________________________________
@REM   __ RUNNER
@REM   __ C:\Program Files\Logitech\Gaming Software\LWEMon.exe
@REM   __ ________________________________________________________

@REM  PAUSE

@IF "%WIN10__%"=="YES" GOTO WINDOW_NOT_WINDOWS_LIVE_MAIL_WHEN_WINDOWS_10
	@REM   __ ________________________________________________________
	@REM   __ @SET FROM AUTORUNS
	@REM   __ FOR ____ 5-ASUS-PL2520LA
	@REM   __ ________________________________________________________
	@REM   __ DONT WANT WINDOWS_LIVE_MAIL __ WITHER ANY COMPUTER BUT BETTER ONE
	@REM   __ ________________________________________________________
	@REM   __ Microsoft Windows	Windows Mail	Microsoft Corporation	13/07/2009 23:58	
	@REM   __ ________________________________________________________
	TASKKILL /F /IM  LWEMon.exe /T
	@REM   __ ________________________________________________________
	@REM   __ ________________________________________________________
	@REM   __ C:\Program Files\Windows Mail\WinMail.exe
	@REM   __ ________________________________________________________
:WINDOW_NOT_WINDOWS_LIVE_MAIL_WHEN_WINDOWS_10

@REM  PAUSE

@IF "%WIN10__%"=="YES" GOTO WINDOW_NOT_ASUS_LIVE_UPDATE_WHEN_WINDOWS_10
	@REM   __ ________________________________________________________
	@REM   __ @SET FROM AUTORUNS __ AND IS IN SCHEDULER
	@REM   __ FOR ____ 5-ASUS-PL2520LA
	@REM   __ ________________________________________________________
	@REM   __ DONT WANT 
	@REM   __ ________________________________________________________
	@REM   __ \ASUS Live Update2		23/06/2016 02:43	
	@REM   __ ________________________________________________________
	TASKKILL /F /IM  UpdateChecker.exe /T
	@REM   __ ________________________________________________________
	@REM   __ ________________________________________________________
	@REM   __ C:\Program Files (x86)\ASUS\ASUS Live Update\UpdateChecker.exe
	@REM   __ ________________________________________________________
:WINDOW_NOT_ASUS_LIVE_UPDATE_WHEN_WINDOWS_10

@REM  PAUSE

@IF "%WIN10__%"=="YES" GOTO WINDOW_NOT_SONY_HOME_NETWORK_LIBRARY_WHEN_WINDOWS_10
	@REM   __ ________________________________________________________
	@REM   __ @SET FROM AUTORUNS __ AND IS IN SCHEDULER
	@REM   __ FOR ____ 5-ASUS-PL2520LA
	@REM   __ ________________________________________________________
	@REM   __ DONT WANT Sony_Home_Network_Library __ WITHER ANY COMPUTER BUT BETTER ONE
	@REM   __ ________________________________________________________
	@REM   __ \Sony Corporation\Sony Home Network Library\SOHLib SOHDms	Sony Digital Media Server	Sony Corporation	15/01/2014 08:22	

	@REM   __ ________________________________________________________
	TASKKILL /F /IM  sohdms.exe /T
	@REM   __ ________________________________________________________
	@REM   __ ________________________________________________________
	@REM   __ c:\program files (x86)\common files\sony shared\sohlib\
	@REM   __ ________________________________________________________
:WINDOW_NOT_SONY_HOME_NETWORK_LIBRARY_WHEN_WINDOWS_10

@REM  PAUSE

@REM   __ FIND OUT THIS ONE 1.. 2.. 3..  IN SCHEDULER 
@REM   __ ________________________________________________________

@REM   __ ________________________________________________________
@REM   __ \ATKOSD2	ATKOSD2	ASUSTek Computer Inc.	c:\program files (x86)\asus\atk package\atkosd2\atkosd2.exe	15/05/2014 07:22	
@REM   __ FIND IS HERE __ ASUS HOT KEY BY ASUS
@REM   __ KEEPER
@REM   __ ATKGFNEXSrv	GFNEXSrv	ASUS	
@REM   __ c:\program files (x86)\asus\atk package\atkgfnex\gfnexsrv.exe	
@REM   __ 21/11/2011 06:19	

@REM   __ \ASUS Splendid ACMON	ACMON 	ASUS	c:\program files (x86)\asus\splendid\acmon.exe	09/06/2015 12:20	

@REM   __ \ASUS Smart Gesture Launcher	ASUS Smart Gesture Launcher	AsusTek	c:\program files (x86)\asus\asus smart gesture\astpcenter\x64\asustplauncher.exe	19/04/2013 08:55	
@REM   __ ________________________________________________________
@REM   __ ________________________________________________________
@REM   __ ________________________________________________________
@REM   __ WORKING BLOCK ABOVE FEW HOUR UNTIL 5:44 PM SUNDAY __ Sun 26 February 2017 17:44:00----------
@REM   __ ________________________________________________________


@REM  PAUSE

GOTO JUMPER_TEXT_AUTORUNS_SERVICE_STOPPER
	@REM --------------------------------------------------------------------------------
	@REM AUTORUNS SERVICE @SET REQUIRE STOP
	@REM --------------------------------------------------------------------------------

	@REM LOW PRIORITY __ AN UPDATER __ NOT ENABLED
	@REM -------------------------------------------------------------------
	@REM AdobeARMservice	Adobe Acrobat Updater keeps your Adobe software up to date.	Adobe Systems Incorporated	c:\program files (x86)\common files\adobe\arm\1.0\armsvc.exe	20/12/2016 06:32	


	@REM MAYBE 
	@REM ----------
	@REM ASNB4LDRSvc	Wireless Console 3 service	ASUS	c:\program files (x86)\asus\wireless console 3\asnb4ldrsvc.exe	11/03/2014 04:04	

	@REM AsusFPService	AsusFPService	ASUSTek Computer Inc.	c:\program files (x86)\asus\fingerprint\511\asusfpservice_x64.exe	25/03/2015 01:12	


	@REM NOT TO USE CHROME @REMOTE DESKTOP ANY MORE
	@REM TEAM VIEWER IS VASTLY BETTER
	@REM ----------------------------------------------------------------------------
	@REM chromoting	This service enables incoming connections from Chrome @REMote Desktop clients.	Google Inc.	c:\program files (x86)\google\chrome @REMote desktop\56.0.2924.51\@REMoting_host.exe	04/01/2017 05:53	


	@REM LOW PRIORITY __  AN UPDATER __ NOT ENABLED
	@REM NOT REQUIRE XP NOW CHECK @IF EXIST THERE WIN 7 CHECKER AT THE MOMENT
	@REM ---------------------------------------------------------------------------------------------
	@REM gupdate	Keeps your Google software up to date. @IF this service is disabled or stopped, your Google software will not be kept up to date, meaning security vulnerabilities that may arise cannot be fixed and features may not work. This service uninstalls itself when there is no Google software using it.	Google Inc.	c:\program files (x86)\google\update\googleupdate.exe	09/01/2016 12:08	

	@REM gupdatem	Keeps your Google software up to date. @IF this service is disabled or stopped, your Google software will not be kept up to date, meaning security vulnerabilities that may arise cannot be fixed and features may not work. This service uninstalls itself when there is no Google software using it.	Google Inc.	c:\program files (x86)\google\update\googleupdate.exe	09/01/2016 12:08	

	
	REM KILL AT BOOT NOT WANT ANY REGISTRY START AR
	REM TASKKILL /F /IM TeamViewer.exe /T
	
	
	
	

	@REM peservice	photoSync syncing service	webEcoz, LLC.	c:\program files (x86)\photosync\peservice.exe	10/11/2008 07:15	

	@REM PMBDeviceInfoProvider	Enables PlayMemories Home to communicate with the device.	Sony Corporation	c:\program files (x86)\sony\playmemories home\pmbdeviceinfoprovider.exe	24/06/2016 02:37	

	@REM LOW PRIORITY __  AN UPDATER __ NOT ENABLED
	@REM -------------------------------------------------------------------
	@REM SkypeUpdate	Enables the detection, download and installation of updates for Skype.	Skype Technologies	c:\program files (x86)\skype\updater\updater.exe	20/09/2016 12:51	

	@REM SOHDms	Sony Digital Media Server	Sony Corporation	c:\program files (x86)\common files\sony shared\sohlib\sohdms.exe	15/01/2014 08:22	
	@REM SOHDs	VAIO Device Searcher	Sony Corporation	c:\program files (x86)\common files\sony shared\sohlib\sohds.exe	29/01/2013 07:17	
	@REM SpfService	VAIO Entertainment Common Service	Sony Corporation	c:\program files\common files\sony shared\vaio entertainment platform\spf\spfservice64.exe	08/09/2011 05:09	

	@REM vcsFPService	Validity Fingerprint Service	Validity Sensors, Inc.	c:\windows\system32\vcsfpservice.exe	11/09/2012 14:40	

	@REM WinDefend	Protection against spyware and potentially unwanted software	Microsoft Corporation	c:\program files\windows defender\mpsvc.dll	27/05/2013 05:51	

TASKKILL /F /IM minibin.exe /T
@REM -------------------------------------------------

@ECHO SONY CYBER SHOT CAMERA
@REM ----------------------------------------------------------
TASKKILL /F /IM PMBGPSSupportTool.exe /T
TASKKILL /F /IM PMBAnnounce.exe /T
TASKKILL /F /IM PMBVolumeWatcher.exe /T
TASKKILL /F /IM PMBDeviceInfoProvider.exe /T
@REM -- CAMERA SONY CYBERSHOT

@REM TASKKILL /F /IM NoikaSuite.exe /T


@ECHO SONY CYBER SHOT CAMERA
@ECHO "C:\Program Files (x86)\Sony\PlayMemories Home\PMBDeviceInfoProvider.exe"
@REM ----------------------------------------------------------
@REM AUTO STOP
NET STOP PMBDeviceInfoProvider
@REM TASKKILL /F /IM PMBDeviceInfoProvider.exe /T
@REM "C:\Program Files (x86)\Sony\PlayMemories Home\PMBDeviceInfoProvider.exe"

@ECHO NoikaSuite.exe
@REM ----------------------------------------------------------
TASKKILL /F /IM NoikaSuite.exe /T




@REM --------------------------------------------------------------------------------
:JUMPER_TEXT_AUTORUNS_SERVICE_STOPPER

@REM  PAUSE

@REM   __ END LEARN BLOCK WIN 7 WORKER

@ECHO ON
TASKKILL /F /IM FIRSTRUN.EXE /T
@REM   -- C:\Program Files (x86)\Microsoft Office\Office15\FIRSTRUN.EXE
@REM   -- OFFICE ADVERT
@REM   -- WIN 10

@REM ---- TASKKILL /F /IM AsusWSPanel.exe /T
@REM  -- C:\Program Files (x86)\ASUS\WebStorage\2.2.7.557\AsusWSPanel.exe


@REM   ------------------------------------------------------------------------

@REM   ------------------------------------------------------------------------
@REM   -- WIN 10
@REM   ------------------------------------------------------------------------
:WINDY_10
@REM   ------------------------------------------------------------------------
@REM   ------------------------------------------------------------------------

@REM   PAUSE
@REM   @IF a service, use NET STOP
@REM   --------------------------
@REM   TASKKILL /F /IM Sync.exe /T
@REM   TASKKILL /F /IM peservice.exe /T
@REM   -- C:\Program Files (x86)\photoSync\peservice.exe
@REM   -- TO DO FLICKR
@REM   -- @IF A SYNC GOING THE SERVICE WILL CONTINUE UPLOAD 
@REM   -- EVEN @IF OTHER APP SYNC STOPPED AND QUIT
@REM   --------------------------------------------------------------------
@REM   TASKKILL /F /IM PMBGPSSupportTool.exe /T
@REM   TASKKILL /F /IM PMBAnnounce.exe /T
@REM   TASKKILL /F /IM PMBDeviceInfoProvider.exe /T
@REM   -- CAMERA SONY CYBERSHOT
@REM   --------------------------------------------------------------------

TASKKILL /F /IM NoikaSuite.exe /T


@ECHO C:\Program Files (x86)\Samsung\Kies\KiesTrayAgent.exe
@REM ----------------------------------------------------------
TASKKILL /F /IM KiesTrayAgent.exe /T
@REM -- C:\Program Files (x86)\Samsung\Kies\KiesTrayAgent.exe
@REM -- 
@REM -- NOT YET WIN 7

@REM   -- C:\Program Files (x86)\Samsung\Kies\KiesTrayAgent.exe
@REM   -- 
@REM   -- NOT YET WIN 7
@REM   __ NOT ON WIN 10 HERE
@REM   __ TASKKILL /F /IM Update.exe /T
@REM   __ "C:\Users\MATT 01\AppData\Local\FlickrUploadrWindows\Update.exe" --processStart Flickr.exe
@REM   __ PAID FOR 
@REM   ------------------------------------------------
@REM   END OF BLOCK LINE FROM ACER-ONE
@REM   ------------------------------------------------
@REM   -- -------------------------- -------------------------------------------
@REM   -- ATKOSD2.exe -- OKAY TO USE
@REM   -- ON SCREEN DISPLAY OF volume and display brightness @SETting information
@REM   -- C:\Program Files (x86)\ASUS\ATK Package\ATKOSD2\ATKOSD2.exe
@REM   -- USE UNINSTALL
@REM   -- ----
@REM   -- ATKOSD2 by ASUS - Should I @REMove It?
@REM   -- http://www.shouldi@REMoveit.com/ATKOSD2-9057-program.aspx
@REM   -- ----
@REM   -- ----
@REM   -- G74sx ATKOSD2.exe
@REM   -- https://rog.asus.com/forum/showthread.php?44806-G74sx-ATKOSD2-exe
@REM   -- ----
@REM   -- 
@REM   -- DEPENDANT AND SHOULD CLOSE BEFORE UN-INSTALL HAPPEN
@REM   -- 
@REM   -- DMedia.EXE
@REM   -- C:\Program Files (x86)\ASUS\ATK Package\ATK Media\DMedia.exe
@REM   -- DONT KNOW WHAT DOES

@REM   -- ATKGFNEX -- SERVICE
@REM   -- 
@REM   -- MORE INFO BOTTOM THREAD REPLACE ATKOSD3.exe TO ATKOSD2.exe
@REM   -- ----
@REM   -- G74sx ATKOSD2.exe
@REM   -- https://rog.asus.com/forum/showthread.php?44806-G74sx-ATKOSD2-exe
@REM   -- -------------------------- -------------------------------------------
@REM   -- -------------------------- -------------------------------------------
@REM   -- TASKKILL /F /IM NclUSBSrv64.exe /T
@REM   -- C:\Program Files (x86)\PC Connectivity Solution\Transports\NclUSBSrv64.exe
@REM   -- WIN 7 -- USB MEDIA SERVER
@REM   -- -------------------------- -------------------------------------------
@REM   -- -------------------------- -------------------------------------------
@REM   -- BLUE-TOOTH MICROSOFT -- C:\Program Files (x86)\PC Connectivity Solution\Transports\NclMSBTSrvEx.exe
@REM   -- NOT SURE WIN 7
@REM   -- -------------------------- -------------------------------------------
@REM   LATER ON IN SCRIPT HERE ---------
@REM   TASKKILL /F /IM LiveUpdate.exe /T
@REM   TASKKILL /F /IM CNMNSST.exe /T
@REM   -- A WIRELESS NET TOGGLER TASK AREA ICON
@REM   -- WIN 7
@REM   TASKKILL /F /IM wcourier.exe /T
@REM   -- C:\Windows\System32\vcsFPService.exe
@REM   -- WIN 7
@REM   TASKKILL /F /IM vcsFPService.exe /T
@REM   -- 
@REM   -- WIN 7
@REM   TASKKILL /F /IM asusfpservice_x64.exe /T
@REM   -- 
@REM   -- WIN 7
@REM   TASKKILL /F /IM usbchargerplus.exe /T
@REM   -- C:\Program Files (x86)\ASUS\ASUS Manager\Power Saver\System56Link.exe
@REM   -- WIN 7
@REM   TASKKILL /F /IM System56Link.exe /T
@REM   -- C:\Program Files (x86)\ASUS\ASUS Manager\PC Cleanup\SecureDeleteBackground.exe
@REM   -- WIN 7
@REM   TASKKILL /F /IM SecureDeleteBackground.exe /T
@REM   -- C:\Program Files (x86)\ASUS\ASUS Manager\PC Cleanup\SecureDeleteWindow.exe
@REM   -- WIN 7
@REM   TASKKILL /F /IM SecureDeleteWindow.exe /T
@REM   -- C:\Program Files\Logitech\Gaming Software\LWEMon.exe
@REM   -- WIN 7
@REM   -- TASKKILL /F /IM LWEMon.exe /T
@REM   -- C:\Program Files (x86)\ASUS\ATK Package\ATK Hotkey\AsLdrSrv.exe
@REM   -- NOT SURE WIN 7
@REM   -- TASKKILL /F /IM AsLdrSrv.exe /T
@REM   -- C:\Program Files (x86)\ASUS\ASUS Smart Gesture\AsTPCenter\x64\AsusTPLoader.exe
@REM   -- NOT SURE WIN 7
@REM   -- TASKKILL /F /IM AsusTPLoader.exe /T
@REM   -- C:\Program Files (x86)\ASUS\ASUS Smart Gesture\QuickGesture\x86\QuickGesture.exe
@REM   -- NOT SURE WIN 7
@REM   -- TASKKILL /F /IM QuickGesture.exe /T
@REM   -- C:\Program Files (x86)\ASUS\ASUS Smart Gesture\AsTPCenter\x64\AsusTPHelper.exe
@REM   -- NOT SURE WIN 7
@REM   -- TASKKILL /F /IM AsusTPHelper.exe /T
@REM   -- C:\Windows\SysWOW64\SASrv.exe
@REM   -- LET BOOT FIRST AND @SETTING HAPPEN WIN 7
@REM   -- TASKKILL /F /IM SASrv.exe /T
@REM   -- C:\Program Files (x86)\ASUS\Splendid\ACMON.exe
@REM   -- WIN 7
@REM   -- SEARCH INTERNET 
@REM   -- ACMON.exe Windows process - What is it? - File.net
@REM   -- www.file.net/process/acmon.exe.html
@REM   -- ACMON.exe file information. ... Description: ACMON.exe is not essential for Windows and will often cause problems. ... The process ACMON belongs to the software ASUS Splendid Video by ATK.
@REM   -- TASKKILL /F /IM ACMON.exe /T
@REM   -- C:\Program Files\CONEXANT\SAII\SmartAudio.exe
@REM   -- TEST WIN 7
@REM   -- TASKKILL /F /IM SmartAudio.exe /T
@REM   -- C:\Windows\System32\flvga_tray.exe
@REM   -- LOOK LIKE USB TO VGA ADAPTOR SOFTWARE
@REM   -- ----
@REM   -- flvga_tray.exe - Google Search
@REM   -- https://www.google.co.uk/search?q=flvga_tray.exe&gws_rd=cr&ei=vG4XWMyBLaSLgAae64fABg
@REM   -- ----


@REM   -- TEST WIN 7
@REM   -- TASKKILL /F /IM flvga_tray.exe /T
@REM   -- C:\Program Files (x86)\Microsoft Office\Office15\FIRSTRUN.EXE
@REM   -- OFFICE ADVERT
@REM   -- WIN 10
TASKKILL /F /IM FIRSTRUN.EXE /T

@ECHO C:\Program Files (x86)\Nero\Nero8\Nero BackItUp\NBKeyScan.exe
@REM ----------------------------------------------------------
@REM NOT SURE THIS ONE MEANT TO BE OTHER ONE BELOW
@REM DON'T USE VERSION 8 ANY MORE AT THE MOMENT -- YET
@REM THINKING ABOUT IT
TASKKILL /F /IM NBKeyScan.exe /T
@REM C:\Program Files (x86)\Nero\Nero8\Nero BackItUp\NBKeyScan.exe

@IF "%WINXP__%"=="YES" GOTO NOT_WHEN_WIN_XP
@IF "%WIN7___%"=="YES" GOTO NOT_WHEN_WIN_07
	@ECHO Nero BackItUp Scheduler 3 is responsible to control all jobs created using Nero BackItUp 3. These jobs can create backups of selected files/folders/partitions or complete hard disk to hard disk, network drive, disc or FTP.
	@ECHO C:\Program Files (x86)\Nero\Nero8\Nero BackItUp\NBService.exe
	@REM ----------------------------------------------------------
	@REM Nero BackItUp Scheduler 3 is responsible to control all jobs created using Nero BackItUp 3. These jobs can create backups of selected files/folders/partitions or complete hard disk to hard disk, network drive, disc or FTP.
	@REM AUTO STOP
	NET STOP "Nero BackItUp Scheduler 3"
	@REM TASKKILL /F /IM NBService.exe /T
	@REM C:\Program Files (x86)\Nero\Nero8\Nero BackItUp\NBService.exe
	@REM ----------------------------------------------------------
:NOT_WHEN_WIN_XP
:NOT_WHEN_WIN_07




@IF "%WINXP__%"=="YES" GOTO NOT_WHEN_WIN_XP
@IF "%WIN7___%"=="YES" GOTO NOT_WHEN_WIN_07
@ECHO Provides access to Nero application updates and manages Nero applications.
@ECHO C:\Program Files (x86)\Nero\Update\NASvc.exe
@REM ----------------------------------------------------------
@REM AUTO-DELAY -- STOP
@REM Nero Update
@REM Provides access to Nero application updates and manages Nero applications.
NET STOP NAUpdate
@REM TASKKILL /F /IM NASvc.exe /T
@REM C:\Program Files (x86)\Nero\Update\NASvc.exe
:NOT_WHEN_WIN_XP
:NOT_WHEN_WIN_07


@IF "%WINXP__%"=="YES" GOTO NOT_WHEN_WIN_XP
@IF "%WIN7___%"=="YES" GOTO NOT_WHEN_WIN_07
@ECHO C:\Program Files (x86)\Common Files\Nero\Lib\NMIndexingService.exe
@REM ----------------------------------------------------------
@REM MANUAL STOPPED
NET STOP NMIndexingService
@REM "C:\Program Files (x86)\Common Files\Nero\Lib\NMIndexingService.exe"
:NOT_WHEN_WIN_XP
:NOT_WHEN_WIN_07

@ECHO ----------------------------------------------------------
@ECHO C:\Program Files (x86)\ASUS\ASUS GIFTBOX Desktop\ASUSGIFTBOXDesktop.exe
@REM ----------------------------------------------------------
@ECHO 01 02 03 -- OF -- 03 ------------
@ECHO ----------------------------------------------------------
@REM 1 2 3 
@REM ------------------------
@REM AUTO STOP
NET STOP ASUSGiftBoxDekstop
@REM TASKKILL /F /IM ASUSGIFTBOXDesktop.exe /T
@REM C:\Program Files (x86)\ASUS\ASUS GIFTBOX Desktop\ASUSGIFTBOXDesktop.exe

@REM @ECHO C:\Program Files (x86)\ASUS\WebStorage\2.2.2.524\AsusWSWinService.exe
@REM ----------------------------------------------------------
@REM AUTO STOP
@REM NET STOP "Asus WebStorage Windows Service"
@REM TASKKILL /F /IM AsusWSWinService.exe /T
@REM C:\Program Files (x86)\ASUS\WebStorage\2.2.2.524\AsusWSWinService.exe

@ECHO C:\Program Files (x86)\ASUS\ROG Game First III\AsusGameFirstService.exe
@REM ----------------------------------------------------------
@REM MAYBE WANT TO KILL THE PROCESS AS WELL AS STOP THE SERVICE
@REM ----------------------------------------------------------
@REM AUTO
NET STOP AsusGameFirstService
@REM TASKKILL /F /IM AsusGameFirstService.exe /T
@REM C:\Program Files (x86)\ASUS\ROG Game First III\AsusGameFirstService.exe
@REM ----------------------------------------------------------

@ECHO "C:\Program Files (x86)\WildTangent Games\App\GamesAppIntegrationService.exe"
@REM ----------------------------------------------------------
NET STOP GamesAppIntegrationService
@REM "C:\Program Files (x86)\WildTangent Games\App\GamesAppIntegrationService.exe"

@ECHO "C:\Program Files (x86)\WildTangent Games\App\GamesAppService.exe"
@REM ----------------------------------------------------------
@REM MANUAL STOP
NET STOP GamesAppService
@REM "C:\Program Files (x86)\WildTangent Games\App\GamesAppService.exe"

@ECHO C:\Program Files (x86)\ASUS\ROG Gaming Center\ROGGamingKey.exe
@REM ----------------------------------------------------------
TASKKILL /F /IM ROGGamingKey.exe /T
@REM C:\Program Files (x86)\ASUS\ROG Gaming Center\ROGGamingKey.exe

@ECHO C:\Program Files (x86)\ASUS\ROG Gaming Center\ROGGamingCenterService.exe
@REM ----------------------------------------------------------
@REM MANUAL NOT RUNNER
NET STOP ROGGamingCenterService
@REM C:\Program Files (x86)\ASUS\ROG Gaming Center\ROGGamingCenterService.exe



@ECHO C:\Program Files (x86)\Skype\Updater\Updater.exe
@ECHO Enables the detection, download and installation of updates for Skype.
@REM ----------------------------------------------------------
@REM DIABLED
@REM Enables the detection, download and installation of updates for Skype.
NET STOP SkypeUpdate
@REM "C:\Program Files (x86)\Skype\Updater\Updater.exe"


@ECHO Manages access to smart cards read by this computer. If this service is stopped, this computer 
@REM ----------------------------------------------------------
@REM 3 NOT RUNNING
@REM -------------
@REM AUTO
@REM Manages access to smart cards read by this computer. If this service is stopped, this computer 
NET STOP "SCardSvr"
@REM C:\windows\system32\svchost.exe -k LocalServiceAndNoImpersonation


@ECHO Creates software device nodes for all smart card readers accessible to a given session. If this service is disabled, WinRT APIs will not be able to enumerate smart card readers.
@REM ----------------------------------------------------------
@REM MANUAL
@REM Creates software device nodes for all smart card readers accessible to a given session. If this service is disabled, WinRT APIs will not be able to enumerate smart card readers.
NET STOP "ScDeviceEnum"
@REM C:\windows\system32\svchost.exe -k LocalSystemNetworkRestricted


@ECHO Allows the system to be configured to lock the user desktop upon smart card @REMoval.
@REM ----------------------------------------------------------
@REM MANUAL
@REM Allows the system to be configured to lock the user desktop upon smart card @REMoval.
NET STOP "SCPolicySvc"
@REM C:\windows\system32\svchost.exe -k netsvcs




@ECHO Enables Touch Keyboard and Handwriting Panel pen and ink functionality
@REM ----------------------------------------------------------
@REM Enables Touch Keyboard and Handwriting Panel pen and ink functionality
NET STOP "TabletInputService"
@REM C:\windows\System32\svchost.exe -k LocalSystemNetworkRestricted


@ECHO Windows Defender Service
@REM ----------------------------------------------------------
@REM MANUAL NOT RUNNING
@REM Windows Defender Service
@REM Helps protect users from malware and other potentially unwanted software
NET STOP WinDefend
@REM "C:\Program Files\Windows Defender\MsMpEng.exe"


@ECHO Windows Error Reporting Service
@REM ----------------------------------------------------------
@REM AUTO -- RUNNING
@REM Windows Error Reporting Service
@REM __ NERO TUNEITUP WANTS THIS ONE LEFT RUNNING
@REM C:\windows\System32\svchost.exe -k WerSvcGroup
@REM SC CONFIG WerSvc START= DISABLED
@REM SC CONFIG WerSvc START= DEMAND
SC CONFIG WerSvc START= AUTO
@REM NET START "WerSvc"

@REM NET STOP WerSvc
@REM SC CONFIG WerSvc START= DISABLED
@REM TASKKILL /F /FI "SERVICES eq WerSvc"

@REM ------------------------------------------------------------------------
@REM ------------------------------------------------------------------------
@REM Sensor Monitoring Service
@REM __ NERO TUNEITUP WANTS THIS ONE LEFT RUNNING ON IF REQUIRED
SC CONFIG SensrSvc START= DEMAND
@REM -- SensrSvc
@REM -- Monitors ambient light sensors to detect changes in ambient light and adjust the display brightness.  @IF this service is stopped or disabled, the display brightness will not adapt to lighting conditions.
@REM -- C:\Windows\system32\svchost.exe -k LocalServiceAndNoImpersonation
@REM ------------------------------------------------------------------------
@REM NET STOP SensrSvc
@REM SC CONFIG SensrSvc START= DISABLED
@REM TASKKILL /F /FI "SERVICES eq SensrSvc"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.


@REM Windows Backup
@REM __ NERO TUNEITUP WANTS THIS ONE LEFT RUNNING ON IF REQUIRED
@REM __ NOT AS THE DEFAULT AUTO _ SERVICES THAT SHOULD START OF NECESSARY
SC CONFIG SDRSVC START= DEMAND
@ECHO Provides Windows Backup and Restore capabilities.
@REM ----------------------------------------------------------
@REM MANUAL
@REM Windows Backup
@REM Provides Windows Backup and Restore capabilities.
@REM C:\windows\system32\svchost.exe -k SDRSVC
@REM   ------------------------------------------------------------------------
@REM   -- SDRSVC
@REM   -- Windows Backup Properties (Local Computer)
@REM   -- Provides Windows Backup and Restore capabilities.
@REM   -- Windows Backup Properties (Local Computer)
@REM   -- C:\Windows\system32\svchost.exe -k SDRSVC
@REM   ------------------------------------------------------------------------
@REM NET STOP SDRSVC
@REM SC CONFIG SDRSVC START= DISABLED
@REM TASKKILL /F /FI "SERVICES eq SDRSVC"


@REM __ NERO TUNEITUP WANTS THIS ONE LEFT RUNNING ON IF REQUIRED
@REM __ NOT AS THE DEFAULT AUTO _ SERVICES THAT SHOULD START OF NECESSARY
SC CONFIG stisvc START= DEMAND
@REM Windows Image Acquisition (WIA)
@REM Provides image acquisition services for scanners and cameras
@REM C:\Windows\system32\svchost.exe -k imgsvc



@REM Windows Media Player Network Sharing Service
@REM __ NERO TUNEITUP WANTS THIS ONE CHANGED FROM AUTO TO ON DEMAND
SC CONFIG WMPNetworkSvc START= DEMAND
@REM NET START WMPNetworkSvc
@REM --------------------------------------------
@REM WMPNetworkSvc
@REM --------------------------------------------
@REM Windows Media Player Network Sharing Service
@REM --------------------------------------------
@REM Shares Windows Media Player libraries with other networked players and media devices using Universal Plug and Play
@REM --------------------------------------------
@REM "C:\Program Files\Windows Media Player\wmpnetwk.exe"
@REM --------------------------------------------
@REM AUTO WITH A DELAY
@REM --------------------------------------------




SC CONFIG W32Time START= DEMAND
@REM __ NERO TUNEITUP WANTS THIS ONE CHANGED FROM AUTO TO ON DEMAND
@REM Windows Time
@REM Maintains date and time synchronization on all clients and servers in the network. If this service is stopped, date and time synchronization will be unavailable. If this service is disabled, any services that explicitly depend on it will fail to start.
@REM C:\Windows\system32\svchost.exe -k LocalService


@REM __ THIS IS DONE ELSE WHERE ON THE SCHEDULED TIMER
@REM __ 
SC CONFIG AdobeARMservice START= DISABLED
@REM __ NERO TUNEITUP WANTS THIS ONE CHANGED DEMAND TO DEACTIVATED
@REM Adobe Acrobat Update Service
@REM Adobe Acrobat Updater keeps your Adobe software up to date.
@REM "C:\Program Files (x86)\Common Files\Adobe\ARM\1.0\armsvc.exe"

@REM This service caches network content from peers on the local subnet.
@REM BranchCache
@REM PeerDistSvc
@REM __ NERO TUNEITUP WANTS THIS ONE CHANGED DEMAND TO DEACTIVATED
SC CONFIG PeerDistSvc START= DISABLED

@REM __ ALL THESE HYPER ONE NERO TUNEITUP WANTS DISABLED
@REM __ THESE HYPER FROM THE ASUS COMPUTER MACHINE
SC CONFIG vmickvpexchange START= DISABLED
SC CONFIG vmicguestinterface START= DISABLED
SC CONFIG vmicshutdown START= DISABLED
SC CONFIG vmicheartbeat START= DISABLED
SC CONFIG vmicrdv START= DISABLED
SC CONFIG vmictimesync START= DISABLED
SC CONFIG vmicvss START= DISABLED




@ECHO WPS Office Update Service.
@REM ----------------------------------------------------------
@REM AUTO -- NOT RUNNING
@REM WPS Office Update Service.
NET STOP Kingsoft_WPS_UpdateService
@REM C:\Program Files (x86)\Kingsoft\WPS Office\9.1.0.4947\wtoolex\wpsupdatesvr.exe

@REM ----------------------------------------------------------
@REM 3 IN MANUAL TOP ONE WAS AUTO
@REM -----------
@REM Xbox Live Auth Manager
@REM Provides authentication and authorisation services for interacting with Xbox Live. If this service is stopped, some applications may not operate correctly.
NET STOP "XblAuthManager"
@REM C:\windows\system32\svchost.exe -k netsvcs

@REM Xbox Live Game Save
@REM This service syncs save data for Xbox Live save enabled games.  If this service is stopped, game 
NET STOP "XblGameSave"
@REM C:\windows\system32\svchost.exe -k netsvcs

@REM Xbox Live Networking Service
@REM This service supports the Windows.Networking.XboxLive application 
NET STOP "XboxNetApiSvc"
@REM "C:\windows\system32\svchost.exe -k netsvcs"
@REM ----------------------------------------------------------






@ECHO Microsoft® Volume Shadow Copy Service
@REM ----------------------------------------------------------
@REM NET STOP VSS
@REM TASKKILL /F /IM VSSVC.exe /T
@REM MicrosoftÂ® Volume Shadow Copy Service
@REM C:\Windows\System32\VSSVC.exe
@REM --------------------

@ECHO ----------------------------------------------------------

@ECHO ENDER
@REM PAUSE



@REM   _______________________________________
@REM   -- WIN 10
@REM   _______________________________________

TASKKILL /F /IM AsusWSPanel.exe /T

@REM   _______________________________________
@REM   __ @IF a service, use NET STOP
@REM   __ OR HERE USER KILL IS ABLE WORKING
@REM   _______________________________________

@REM   ________________________________________________________
@REM   ---- SOME FROM AUTORUNS FINDER
@REM   ________________________________________________________
@REM   C:\program files (x86)\processtamer\processtamertray.exe
@REM   ----------------------------------------- 
TASKKILL /F /IM processtamertray.exe /T
@REM   C:\program files\windows defender\msascui.exe
@REM   ----------------------------------------- 
TASKKILL /F /IM msascui.exe /T
@REM   C:\program files (x86)\qnap\qsync
@REM   ----------------------------------------- 
TASKKILL /F /IM qsync /T
@REM   C:\program files\process priority optimizer\ppoptimizer.exe
@REM   ----------------------------------------- 
TASKKILL /F /IM ppoptimizer.exe /T
@REM   C:\program files (x86)\windows mail
@REM   ----------------------------------------- 
TASKKILL /F /IM winmail.exe /T 
@REM   -------------------------------------------------------
@REM   -------------------------------------------------------
TASKKILL /F /IM minibin.exe /T
@REM   -------------------------------------------------------
@REM   -------------------------------------------------------
@REM   -- C:\Users\MATT 01\AppData\Local\yahoomessenger\app-0.8.231\Yahoo Messenger.exe
TASKKILL /F /IM "Yahoo Messenger.exe" /T
@REM   -------------------------------------------------------
@REM   -------------------------------------------------------
@REM   -------------------------------------------------------
@REM   -------------------------------------------------------
@REM   -- C:\Users\MATT 01\AppData\Local\yahoomessenger\app-0.8.231\Yahoo Messenger.exe
TASKKILL /F /IM "Yahoo Messenger.exe" /T
@REM   -------------------------------------------------------
@REM   -------------------------------------------------------
@REM   --------------------------------------------------------------------
@REM   --------------------------------------------------------------------
@REM   ---- WINDOWS 7 -- 5P
@REM   -------------------------------------------------------
@REM   -------------------------------------------------------
@REM   -- C:\Program Files (x86)\Xmarks\IE Extension\xmarkssync.exe
TASKKILL /F /IM "xmarkssync.exe" /T
@REM   -------------------------------------------------------
@REM   -------------------------------------------------------
@REM   -------------------------------------------------------
@REM   -------------------------------------------------------
@REM   -- C:\Program Files (x86)\ASUS\ASUS Live Update\LiveUpdate.exe
TASKKILL /F /IM "LiveUpdate.exe" /T
@REM -------------------------------------------------------
@REM --------------------------------------------------------------------
@REM ------------------------------------
@REM ------------------------------------ WINDOWS 7 ASUS-P2520LA
@REM ------------------------------------ BELOW BLOCK 
@REM ------------------------------------
@REM ------------------------------------------------------------------------
@REM -- WINDOWS UPDATE SERVICE __ START YOURSELF IF WANT IT
@REM -- REBOOT MAYBE REQUIRE TO START AGAIN
@REM ------------------------------------------------------------------------
@REM -- SOURCE LOCATION
@REM C:\Windows\System32\wuauclt.exe
@REM ------------------------------------------------------------------------

@REM ---------------------------------------
@REM ----
@REM Appendix H: The wuauclt Utility
@REM https://technet.microsoft.com/en-us/library/cc720477(v=ws.10).aspx
@REM ----
@REM Appendix H: The wuauclt Utility

@REM Applies To: Windows Server 2003 with SP2, Windows Server Update Services
@REM The wuauclt utility allows you some control over the functioning of the Windows Update Agent. It is updated as part of Windows Update.

@REM __ FOR WINDOW 7 NOT WIN 10
@REM TASKKILL /F /IM wuauclt.exe /T
@REM ---------------------------------------


@REM -----------------------------------------------
@REM __ NOT THE WAY TO STOP THIS SERVICE __ LATER ON 
@REM __ NOT USED AT MOMENT LET WINDOWS UPDATE ROLL
@REM -----------------------------------------------
@REM TASKKILL /F /FI "SERVICES eq wuauserv"
@REM ALSO BOTH 
@REM ------------------------------------------------------------------------
@REM -- Allows errors to be reported when programs stop working or responding and allows existing solutions to be delivered. Also allows logs to be generated for diagnostic and repair services. If this service is stopped, error reporting might not work correctly and results of diagnostic services and repairs might not be displayed.
@REM ------------------------------------------------------------------------


GOTO JUMPER_TEXT_ABOUT_WINDOW_UPDATE

@REM CALL UP SERVICE COMMAND GUI
@REM %windir%\system32\services.msc

@REM TO GET A SCRIPTOR OF SERVICES 

@REM ________________________________________________________________
@REM ________________________________________________________________
@REM SC CONFIG /?
@REM WILL SHOW HOW TO USER THE START COMMAND WITH SC CONFIG 
@REM start= <boot|system|auto|demand|disabled|delayed-auto>
@REM IS TALKER A STOP WILL RESTART AGAIN
@REM AND THEN DISABLE
@REM @SET UP A BATCH COMMAND TO BEGINNER AGAIN WITH DEMAND OR AUTO

@REM NET START WUAUSERV
@REM ________________________________________________________________
@REM ________________________________________________________________


@REM ________________________________________________________________
@REM CHECKER THE NET COMMAND 
@REM MAKE SURE ALL NETWORK IS HAPPENING 
@REM GET THE XP COMPUTER IN ON WIN 10 AND WIN 7

@REM ALSO FOR HARD TO NETWORK
@REM GOOD SYNC HAS A NEW WORKER PROGRAM THE DRIVES IT PUT A PASSWORD AND USERNAME 
@REM IN PROGRAM
@REM AND WILL DO ALL THE NETWORK EVEN HARD TO REACH XP COMPUTER
@REM UNTIL @SETTER
@REM ________________________________________________________________
@REM ALSO ANOTHER PRORGAM
@REM DOES SOMETHING LIKE IT
@REM ________________________________________________________________
@REM @IF USE ROBOFORM ON MULTI COMPUTER QUITE ESTESION USE SISTER PROGRAM GOODSYNC
@REM TO DO THE SYNC DATABASE BETWEEN ROBOFORM
@REM THEY EVEN RECOMEND THAT ON THEIR BULLITIN BOARD
@REM I ALSO TRYIED TO CHANGE A FILE BY DLEETE AND PUT BACK BUT NOT RECOGNISED AS SYNCER
@REM ONLY WILL SYNC @IF CHNAGE HAPPEN
@REM AND REQUIRE THIS THING REGUALAR
@REM GUESS NOT TAX THE SERVER REQUIRE AT THEIR END
@REM ________________________________________________________________

@REM _________________________________________________________
@REM SO I'M DOING MY DINNER
@REM AND A REQUIRE PROGRAM FOR LAUNCHES AT LOOKER SERVICE IS ANOTHER NIRSOFT ONE
@REM I TRIED CCLEANER
@REM AND ALSO PARTNER PROGRAM SPECCY
@REM AND ALSO A LOOK AT COMMAND LINE IT
@REM BUT LAKC COMMAND FOR DETAIL
@REM SO HERE IS NIRSOFT 
@REM _________________________________________________________
@REM ServiWin	
@REM displays the list of installed drivers and services on your system.	
@REM 1.66	
@REM 29/09/2015 18:23:38	
@REM http://www.nirsoft.net/utils/serviwin.html	
@REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\x64\serviwin.exe	
@REM GUI Application	
@REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\serviwin.chm	____ HELP FILE
@REM _________________________________________________________
@REM _________________________________________________________


@REM C:\Users\MATT 01>NET /?
@REM The syntax of this command is:

@REM NET
@REM [ ACCOUNTS | COMPUTER | CONFIG | CONTINUE | FILE |
  @REM HELPMSG | LOCALGROUP | PAUSE | SESSION | SHARE |
  @REM STATISTICS | STOP | TIME | USE | USER | VIEW ]

	
:JUMPER_TEXT_ABOUT_WINDOW_UPDATE

@REM   ------------------------------------
@REM   ---- WINDOWS 7 ASUS-P2520LA
@REM   ---- BELOW BLOCK 
@REM   ------------------------------------
@REM   ------------------------------------------------------------------------
@REM   -- WINDOWS UPDATE SERVICE __ START YOURSELF @IF WANT IT
@REM   ------------------------------------------------------------------------
@REM   -- SOURCE LOCATION
@REM   -- C:\Windows\System32\wuauclt.exe
@REM   ------------------------------------------------------------------------

@REM   TASKKILL /F /IM wuauclt.exe /T
@REM   TASKKILL /F /FI "SERVICES eq wuauserv"
@REM   __ PROPER WAY AS BEEN PLAY BEFORE SERVICING

@REM __ WINDOWS UPDATE SERVICE__
@REM NET STOP WUAUSERV
@REM SC CONFIG WUAUSERV START= DISABLED
@REM __ SC CONFIG WUAUSERV START= DEMAND
@REM __


@REM   ---- ___ WIN 10 JOB WORKER
@REM   ------------------------------------------------------------------------
@REM  NET STOP WerSvc
@REM  SC CONFIG WerSvc START= DISABLED
@REM  TASKKILL /F /FI "SERVICES eq WerSvc"
@REM   ---- END WIN 10 JOB WORKER

@REM   ALSO BOTH %% WHAT
@REM   ------------------------------------------------------------------------
@REM  NERO TALK THIS IMPORTANT SERVICE MUST OT CHANGE AND LEAVE GOING
@REM  Windows Error Reporting Service
@REM  MAYBE ABOVE schtasks IS OKAY LESS FREQUENT OR NOT
@REM  ----------------------------------------------------------------
@REM   -- Windows Error Reporting Service
@REM   -- Allows errors to be reported when programs stop working or responding and allows existing solutions to be delivered. Also allows logs to be generated for diagnostic and repair services. @IF this service is stopped, error reporting might not work correctly and results of diagnostic services and repairs might not be displayed.
@REM   ------------------------------------------------------------------------
@ECHO 

@REM   ---- ___ WIN 10 JOB WORKER
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM THIS IS SUPPOSED TO BE CORTANA A DESCRIPTION WORD IS NOT LISTED 
@REM WITH __ TAKSLIST /SVR
@REM SETTING DISABLED WILL STILL BE RUNNER
@REM ---------------------------------------------------------------
@REM NET STOP SEARCHUI
@REM SC CONFIG SEARCHUI START= DISABLED
@REM ---------------------------------------------------------------

@REM TASKKILL /F /FI "SERVICES eq WerSvc"
@REM ------------------------------------------------------------------------
@REM -- "C:\Program Files\Common Files\Sony Shared\VAIO Entertainment Platform\SPF\SpfService64.exe"
@REM -- VAIO Entertainment Common Service
@REM ------------------------------------------------------------------------



@REM ----------------------------------------------------------------
@REM C:\Program Files\Common Files\Logishrd\LAClient\laclient.exe
@REM WINDOWS 10
@REM ARRIVE AT BOOTER IN CMD.COM SCREEN
@REM Wednesday 13:34:20 Pm_22 January 2020
@REM ----------------------------------------------------------------
@REM C:\Windows\System32>TASKKILL /F /IM laclient.exe /T
@REM SUCCESS: The process with PID 7116 (child process of PID 7052) has been terminated.
@REM SUCCESS: The process with PID 7052 (child process of PID 5532) has been terminated.
@REM ----------------------------------------------------------------
@REM NOT SURE 2 PROCESS SUPPOSED TO KILL THEM -- ONLY CLOSE DOS CMD.COM APP
@REM ----------------------------------------------------------------
TASKKILL /F /IM laclient.exe /T







@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.

@REM   ---- END WIN 10 JOB WORKER









@ @REM  PAUSE -------------------

@REM   ------------------------------------------------------------------------
@REM   -- "C:\Program Files\Common Files\Sony Shared\VAIO Entertainment Platform\SPF\SpfService64.exe"
@REM   -- VAIO Entertainment Common Service
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
NET STOP WinDefend
SC CONFIG WinDefend START= DISABLED
TASKKILL /F /FI "SERVICES eq WinDefend"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   -- Windows Defender Service
@REM   -- Helps protect users from malware and other potentially unwanted software
@REM   -- "C:\Program Files\Windows Defender\MsMpEng.exe"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
NET STOP Sense
SC CONFIG Sense START= DISABLED
TASKKILL /F /FI "SERVICES eq Sense"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   -- Sense
@REM   -- Windows Defender Advanced Threat Protection Service
@REM   -- "C:\Program Files\Windows Defender Advanced Threat Protection\MsSense.exe"
@REM   -- Windows Defender Advanced Threat Protection service helps protect against advanced threats by monitoring and reporting security events that happen on the computer.
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
NET STOP WdNisSvc
SC CONFIG WdNisSvc START= DISABLED
TASKKILL /F /FI "SERVICES eq WdNisSvc"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   -- Windows Defender Network Inspection Service
@REM   -- "C:\Program Files\Windows Defender\NisSrv.exe"
@REM   -- Helps guard against intrusion attempts targeting known and newly discovered vulnerabilities in network protocols
@REM   ------------------------------------------------------------------------
@REM   -- "C:\Program Files\Common Files\Sony Shared\VAIO Entertainment Platform\SPF\SpfService64.exe"
@REM   -- VAIO Entertainment Common Service
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
NET STOP SpfService
SC CONFIG SpfService START= DISABLED
TASKKILL /F /FI "SERVICES eq SpfService"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- Sony Digital Media Server
@REM   -- "C:\Program Files (x86)\Common Files\Sony Shared\SOHLib\SOHDms.exe"
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
NET STOP SOHDms
SC CONFIG SOHDms START= DISABLED
TASKKILL /F /FI "SERVICES eq SOHDms"
@REM   ------------------------------------------------------------------------
@REM   -- Smart Card Properties (Local Computer)
@REM   -- SCardSvr
@REM   -- Manages access to smart cards read by this computer. @IF this service is stopped, this computer will be unable to read smart cards. @IF this service is disabled, any services that explicitly depend on it will fail to start.
@REM   -- C:\Windows\system32\svchost.exe -k LocalServiceAndNoImpersonation
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
NET STOP SCardSvr
SC CONFIG SCardSvr START= DISABLED
TASKKILL /F /FI "SERVICES eq SCardSvr"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- BIT VIRUS PC CLEANER ONE
@REM   ----------------------------------------------- 
@REM   -- ServiceLayer Properties (Local Computer)
@REM   -- "C:\Program Files (x86)\PC Connectivity Solution\ServiceLayer.exe"
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
NET STOP ServiceLayer
SC CONFIG ServiceLayer START= DISABLED
TASKKILL /F /FI "SERVICES eq ServiceLayer"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- PMBDeviceInfoProvider
@REM   -- PMBDeviceInfoProvider Properties (Local Computer)
@REM   -- Enables PlayMemories Home to communicate with the device.
@REM   -- PMBDeviceInfoProvider Properties (Local Computer)
@REM   -- "C:\Program Files (x86)\Sony\PlayMemories Home\PMBDeviceInfoProvider.exe"
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
NET STOP PMBDeviceInfoProvider
SC CONFIG PMBDeviceInfoProvider START= DISABLED
TASKKILL /F /FI "SERVICES eq PMBDeviceInfoProvider"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- peservice
@REM   -- photoSync service Properties (Local Computer)
@REM   -- photoSync service
@REM   -- photoSync service Properties (Local Computer)
@REM   -- "C:\Program Files (x86)\photoSync\peservice.exe"
@REM   -- 
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
NET STOP peservice
SC CONFIG peservice START= DISABLED
TASKKILL /F /FI "SERVICES eq peservice"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- DiagTrack
@REM   -- Diagnostics Tracking Service Properties (Local Computer)
@REM   -- The Diagnostics Tracking Service enables data collection about functional issues in Windows components.
@REM   -- Diagnostics Tracking Service Properties (Local Computer)
@REM   -- C:\Windows\System32\svchost.exe -k utcsvc
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
NET STOP DiagTrack
SC CONFIG DiagTrack START= DISABLED
TASKKILL /F /FI "SERVICES eq DiagTrack"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- NOT LOOK HEALTHY VALUE
@REM   ------------------------------------------------------------------------
@REM   -- ddmgr
@REM   -- C:\Windows\system32\ddmgr.exe
@REM   -- 
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
NET STOP ddmgr
SC CONFIG ddmgr START= DISABLED
TASKKILL /F /FI "SERVICES eq ddmgr"


@REM wmic service list full

@REM   ------------------------------------------------------------------------
@REM   -- WSearch
@REM   -- Provides content indexing, property caching, and search results for files, e-mail, and other content.
@REM   -- Windows Search Properties (Local Computer)
@REM   -- C:\Windows\system32\SearchIndexer.exe /Embedding
@REM   ------------------------------------------------------------------------

@REM --
@REM GOT TO HAVE THIS RUNNING AS DEPENDENCY  
@REM @REMOTE HAVE DEPENDENCY ON IT OR NOT HAVE ANY HOMEGROUP OUTSIDE WORLD
@REM --------------------------------------------------------------------

@REM The following services are dependent on the Windows Search service.
@REM Stopping the Windows Search service will also stop these services.
@REM Windows Media Player Network Sharing Service

@REM SC CONFIG WMPNetworkSvc START= AUTO
@REM NET STOP WMPNetworkSvc
@REM HAVE TO DISABLE THIS ONE 1ST OR ELSE QUESTION NAGGER GET ASKING AND HALT

@REM START= DISABLED
@REM START= AUTO
SC CONFIG WSearch START= DEMAND
@REM NET START WSearch
@REM NET STOP WSearch
@REM TASKKILL /F /FI "SERVICES eq WSearch"

@REM @ECHO Windows Search -- SearchIndexer
@REM ----------------------------------------------------------
@REM RUNNING AUTO
@REM LOOK FOR TITLE INDEX ALPAHBET ORDER Windows Search
@REM Windows Search
@REM Provides content indexing, property caching, and search results for files, e-mail, and other content.
@REM TASKKILL /F /IM SearchIndexer.exe /T
@REM NET STOP SearchIndexer.exe
@REM NET STOP SearchIndexer
@REM FIND SERVICE NAME IN SERVICES MANAGER
@REM ----------------
@REM NET STOP WSearch
@REM ----------------
@REM C:\Windows\System32\SearchIndexer.exe
@REM KILL HERE AND IT COME BACK
@REM STOP IN SERVICES INSTEAD
@REM --------------------
@REM The following services are dependent on the Windows Search service.
@REM Stopping the Windows Search service will also stop these services.

@REM Windows Media Player Network Sharing Service
@REM Do you want to continue this operation? (Y/N) [N]: Y


@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- ehSched
@REM   -- Windows Media Center Scheduler Service Properties (Local Computer)
@REM   -- Starts and stops recording of TV programs within Windows Media Center
@REM   -- C:\Windows\ehome\ehsched.exe
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
NET STOP ehSched
SC CONFIG ehSched START= DISABLED
TASKKILL /F /FI "SERVICES eq ehSched"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- Windows Defender Properties (Local Computer)
@REM   -- WinDefend
@REM   -- Windows Defender Properties (Local Computer)
@REM   -- Protection against spyware and potentially unwanted software
@REM   -- C:\Windows\System32\svchost.exe -k secsvcs
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
NET STOP WinDefend
SC CONFIG WinDefend START= DISABLED
TASKKILL /F /FI "SERVICES eq WinDefend"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- vcsFPService
@REM   -- Validity VCS Fingerprint Service Properties (Local Computer)
@REM   -- C:\Windows\system32\vcsFPService.exe
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
NET STOP vcsFPService
SC CONFIG vcsFPService START= DISABLED
TASKKILL /F /FI "SERVICES eq vcsFPService"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- SOHDs
@REM   -- Sony Device Searcher Properties (Local Computer)
@REM   -- "C:\Program Files (x86)\Common Files\Sony Shared\SOHLib\SOHDs.exe"
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
NET STOP SOHDs
SC CONFIG SOHDs START= DISABLED
TASKKILL /F /FI "SERVICES eq SOHDs"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- SCPolicySvc
@REM   -- Smart Card @REMoval Policy Properties (Local Computer)
@REM   -- Allows the system to be configured to lock the user desktop upon smart card @REMoval.
@REM   -- Smart Card @REMoval Policy Properties (Local Computer)
@REM   -- C:\Windows\system32\svchost.exe -k netsvcs
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
NET STOP SCPolicySvc
SC CONFIG SCPolicySvc START= DISABLED
TASKKILL /F /FI "SERVICES eq SCPolicySvc"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- SkypeUpdate
@REM   -- Skype Updater Properties (Local Computer)
@REM   -- Enables the detection, download and installation of updates for Skype.
@REM   -- Skype Updater Properties (Local Computer)
@REM   -- "C:\Program Files (x86)\Skype\Updater\Updater.exe"
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
NET STOP SkypeUpdate
SC CONFIG SkypeUpdate START= DISABLED
TASKKILL /F /FI "SERVICES eq SkypeUpdate"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- LBTServ
@REM   -- Logitech Bluetooth Service Properties (Local Computer)
@REM   -- C:\Program Files\Common Files\LogiShrd\Bluetooth\lbtserv.exe
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
NET STOP LBTServ
SC CONFIG LBTServ START= DISABLED
TASKKILL /F /FI "SERVICES eq LBTServ"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM   ------------------------------------------------------------------------
@REM   -- gupdatem
@REM   -- Google Update Service (gupdatem) Properties (Local Computer)
@REM   -- Keeps your Google software up to date. @IF this service is disabled or stopped, your Google software will not be kept up to date, meaning security vulnerabilities that may arise cannot be fixed and features may not work. This service uninstalls itself when there is no Google software using it.
@REM   -- Google Update Service (gupdatem) Properties (Local Computer)
@REM   -- "C:\Program Files (x86)\Google\Update\GoogleUpdate.exe" /medsvc
@REM   ------------------------------------------------------------------------
@REM @ECHO.
@REM @ECHO -------------------------------------------------------------------
@REM @ECHO -------------------------------------------------------------------
@REM @ECHO.
@REM NET STOP gupdatem
@REM SC CONFIG gupdatem START= DISABLED
@REM TASKKILL /F /FI "SERVICES eq gupdatem"
@REM @ECHO.
@REM @ECHO -------------------------------------------------------------------
@REM @ECHO -------------------------------------------------------------------
@REM @ECHO.
@REM   ------------------------------------------------------------------------
@REM   ------------------------------------------------------------------------
@REM   -- dbupdate
@REM   -- Dropbox Update Service (dbupdate) Properties (Local Computer)
@REM   -- Keeps your Dropbox software up to date. @IF this service is disabled or stopped, your Dropbox software will not be kept up to date, meaning security vulnerabilities that may arise cannot be fixed and features may not work. This service uninstalls itself when there is no Dropbox software using it.
@REM   -- "C:\Program Files (x86)\Dropbox\Update\DropboxUpdate.exe" /svc
@REM   ------------------------------------------------------------------------
@REM @ECHO.
@REM @ECHO -------------------------------------------------------------------
@REM @ECHO -------------------------------------------------------------------
@REM @ECHO.
@REM NET STOP dbupdate
@REM SC CONFIG dbupdate START= DISABLED
@REM TASKKILL /F /FI "SERVICES eq dbupdate"
@REM @ECHO.
@REM @ECHO -------------------------------------------------------------------
@REM @ECHO -------------------------------------------------------------------
@REM @ECHO.

@REM   ------------------------------------------------------------------------
@REM   -- 
@REM   -- DbxSvc
@REM   -- DbxSvc Properties (Local Computer)
@REM   -- Dropbox Service
@REM   -- C:\Windows\system32\DbxSvc.exe
@REM   ------------------------------------------------------------------------
@REM @ECHO.
@REM @ECHO -------------------------------------------------------------------
@REM @ECHO -------------------------------------------------------------------
@REM NET STOP DbxSvc
@REM SC CONFIG DbxSvc START= DISABLED
@REM TASKKILL /F /FI "SERVICES eq DbxSvc"
@REM @ECHO.
@REM @ECHO -------------------------------------------------------------------
@REM @ECHO -------------------------------------------------------------------
@REM @ECHO.
@REM @REM ----------------------------------------------------------
@REM TASKKILL /F /IM DropboxOEM.exe /T
@REM C:\Program Files (x86)\Dropbox\DropboxOEM\DropboxOEM.exe

@REM   ------------------------------------------------------------------------
@REM   ------------------------------------------------------------------------
@REM   -- BDESVC
@REM   -- BitLocker Drive Encryption Service Properties (Local Computer)
@REM   -- BDESVC hosts the BitLocker Drive Encryption service. BitLocker Drive Encryption provides secure startup for the operating system, as well as full volume encryption for OS, fixed or @REMovable volumes. This service allows BitLocker to prompt users for various actions related to their volumes when mounted, and unlocks volumes automatically without user interaction. Additionally, it stores recovery information to Active Directory, @IF available, and, @IF necessary, ensures the most recent recovery cert@IFicates are used.  Stopping or disabling the service would prevent users from leveraging this functionality.
@REM   -- C:\Windows\System32\svchost.exe -k netsvcs
@REM   ------------------------------------------------------------------------
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
NET STOP BDESVC
SC CONFIG BDESVC START= DISABLED
TASKKILL /F /FI "SERVICES eq BDESVC"
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.

@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.

@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.



@REM   ------------------------------------
@REM   ------------------------------------WINDOWS 7 ASUS-P2520LA
@REM   ------------------------------------ ABOVE BLOCK 
@REM   ------------------------------------
@REM   ------------------------------------------------------------------------
@REM   -- 
@REM   ------------------------------------------------------------------------
@REM   -- TASKKILL /F /FI "SERVICES eq "
@REM   -------------------------------------------------------------------------------------------
@REM   -- @REMMER -- RIMMER -- RED DWARF
@REM   -------------------------------------------------------------------------------------------
@REM   -------------------------------------------------------------------------------------------
@REM   -------------------------------------------------------------------------------------------
@REM   -- HERE RUN EVERY HOUR NOT MUCH TIME 30 SEC
@REM   -- WRONG SCHEDULE
@REM   -- TERMINATE IN PROCESS LASSO
@REM   -- DOES AN UPDATE CHECKER EVERY HOUR FAR TOO MANY @REMOVE 
@REM   -- TERMINATE IN PROCESS LASSO
@REM   -- @IF SYSTEM UNDER LOAD SOMETIME RUN HALF A MINUTE
@REM   -- C:\Program Files (x86)\Dropbox\Update\DropboxUpdate.exe
@REM TASKKILL /F /IM DropboxUpdate.exe /T
@REM   -------------------------------------------------------------------------------------------
@REM   -- ALSO HERE RUN EVERY HOUR NOT MUCH TIME 10 SEC
@REM   -- WRONG SCHEDULE
@REM   -- TERMINATE IN PROCESS LASSO
@REM   C:\Program Files (x86)\ASUS\ASUS Live Update\UpdateChecker.exe
TASKKILL /F /IM UpdateChecker.exe /T
@REM   -------------------------------------------------------------------------------------------
@REM   -- ALSO HERE RUN EVERY HOUR NOT MUCH TIME 30 SEC
@REM   -- WRONG SCHEDULE
@REM   -- TERMINATE IN PROCESS LASSO
@REM   C:\Program Files (x86)\Google\Update\GoogleUpdate.exe
@REM TASKKILL /F /IM GoogleUpdate.exe /T
@REM   -------------------------------------------------------------------------------------------
@REM   -------------------------------------------------------------------------------------------

@REM   -----------------------------------------------------------------------------------
@REM   ----
@REM   SOURCE INFO LOCATION FINDER
@REM   ----
@REM   BATCHFILE ASK QUESTION YES - Google Search
@REM   https://www.google.co.uk/search?q=BATCHFILE+ASK+QUESTION+YES&oq=BATCHFILE+ASK+QUESTION+YES&aqs=chrome..69i57.16125j0j7&sourceid=chrome&ie=UTF-8
@REM   ----
@REM   ----
@REM   batch file - How do I request and receive user input in a .bat and use it to run a certain program? - Stack Overflow
@REM   http://stackoverflow.com/questions/12021033/how-do-i-request-and-receive-user-input-in-a-bat-and-use-it-to-run-a-certain-pr
@REM   ----
@REM   -----------------------------------------------------------------------------------



@REM   ----------------------------------------------------------------------------------------
@REM   ----------------------------------------------------------------------------------------
@REM   ----------------------------------------------------------------------------------------
@REM   ----------------------------------------------------------------------------------------
@REM   ----------------------------------------------------------------------------------------
@REM   ----------------------------------------------------------------------------------------
@REM   ----------------------------------------------------------------------------------------
@REM   ----------------------------------------------------------------------------------------
@REM @ECHO OFF

@REM   ----------------------------------------------------------------------------------------
@REM   ---- BOTH PAIR OF TWO @SET -- YAHOO AND ANOTHER
@REM   ---- BOTH IN THE @SET AND YAHOO PAID FOR RENTAL
@REM   ----  NOT SURE @IF WANT 2TB OF SPACE ANYHOW AND WHAT HAPPEN STOP PAYING
@REM   TASKKILL /F /IM Flickr.exe /T
@REM   -- "C:\Users\MATT 01\AppData\Local\FlickrUploadrWindows\Update.exe" --processStart Flickr.exe
@REM   TASKKILL /F /IM Update.exe /T
@REM   ----------------------------------------------------------------------------------------

@REM   -----------------------------------------------------------------------------------
@REM   -----------------------------------------------------------------------------------
@REM   -----------------------------------------------------------------------------------
@REM   -----------------------------------------------------------------------------------

@REM   ---- TASKLIST WON'T SHOW WITH /M -- ITSELF CMD CONSOLE
@REM   -----------------------------------------------------------------------------------
@REM   
@REM   @ECHO -------------------------------------------------------------------
@REM   @ECHO THE TASK_LIST SHOW MIGHT NOT BE EXACTLY WHAT WILD-CARD IS GOING TO BE DELETER
@REM   @ECHO.
@REM   @ECHO.
@REM   
@REM   TASKLIST /M Flickr*
@REM   TASKLIST /M Update*
@REM   TASKLIST /M Sync*
@REM   TASKLIST /M peservice*

@REM   @ECHO.
@REM   @SET /A ACCUMULATOR=ACCUMULATOR+1
@REM   @ECHO %ACCUMULATOR% OF %VAR_COUNTER%  -- DO YOU WANT KILL ---- FLICKR SYNC PROGRAM @SET
@REM   @ECHO ____________________ ?(Y/N)

@REM   @SET INPUT=
@REM   @SET INPUT=Y

@REM   @IF "%TEST_MODE%" == "0" @SET /P INPUT=Type input: %=%
@REM   @IF "%INPUT%" == "y" GOTO YES
@REM   
@REM   @IF "%INPUT%" == "Y" GOTO YES
@REM   

@REM   GOTO NOT
@REM   :YES
@REM   ----------------------------------------------------------------------------------------
@REM   ---- BOTH PAIR OF TWO @SET -- YAHOO AND ANOTHER
@REM   ---- BOTH IN THE @SET AND YAHOO PAID FOR RENTAL
@REM   ----  NOT SURE @IF WANT 2TB OF SPACE ANYHOW AND WHAT HAPPEN STOP PAYING
@REM   TASKKILL /F /IM Flickr.exe /T
@REM   -- "C:\Users\MATT 01\AppData\Local\FlickrUploadrWindows\Update.exe" --processStart Flickr.exe
@REM   TASKKILL /F /IM Update.exe /T
@REM   ----------------------------------------------------------------------------------------

@REM   -- PAID FOR 

@REM   ----------------------------------------------------------------------------------------
@REM   -- TO DO WITH FLICKR
@REM   -- C:\Program Files (x86)\photoSync\peservice.exe
@REM   -- @IF A SYNC GOING THE SERVICE WILL CONTINUE UPLOAD 
@REM   -- EVEN @IF OTHER APP SYNC STOPPED AND QUIT
@REM   TASKKILL /F /IM Sync.exe /T
@REM   TASKKILL /F /IM peservice.exe /T
@REM   ----------------------------------------------------------------------------------------
@REM   :NOT

@ECHO OFF
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM  -- BY NERO ROM GERMANY SOURCE AUTHOR EXPENSIVE PER YEAR REVENUE LEARNING ----
@REM ----------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@SET TSK_SCRIPTOR_NAME_=VideoDownloaderUltimate*
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_% 
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
@REM -- ABOVE WIN_XP FORMAT "INFO: No tasks"
@REM -- AND WIN_XP FORMAT  "INFO: No tasks running"
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
rem FOR /F %%x IN ('TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"') DO IF %%x==Image SET SET_GO_VAR=YES
rem @IF %SET_GO_VAR%==NO GOTO NOT
@IF %errorlevel% EQU 1 GOTO NOT
	@ECHO.
	@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
	@ECHO %IX1% ------------------------------------------------------------------------
	@ECHO.
	@ECHO %IX1% -- Do You Want Kill ---- %TSK_SCRIPTOR_NAME_%
	@ECHO ______________________________ ?(Y/N)
	@SET INPUT=
	@REM @SET /P INPUT=Type input: %=%
	@IF "%INPUT%" == "y" GOTO YES
	@IF "%INPUT%" == "Y" GOTO YES
	GOTO NOT
	:YES
		@REM   -- PROBLEM HERE WITH ADVERT POP UP
		@REM   -- C:\ProgramData\VideoDownloaderUltimateWinApp\VideoDownloaderUltimate.exe
		@REM   -- PAID FOR  NERO -- WHEN ASK EXIT HAS STILL RUNNING
		@ECHO ON 
		TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T
		@ECHO OFF
		@ECHO.
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@SET TSK_SCRIPTOR_NAME_=Flickr.EXE
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_% 
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
	@ECHO.
	@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
	@ECHO %IX1% -- Do You Want Kill ---- %TSK_SCRIPTOR_NAME_%
	@ECHO ______________________________ ?(Y/N)
	@SET INPUT=
	@REM @SET /P INPUT=Type input: %=%
	@IF "%INPUT%" == "y" GOTO YES
	@IF "%INPUT%" == "Y" GOTO YES
	GOTO NOT
	:YES
		@ECHO ON 
		TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T 
		@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@REM  -- FLICKR AND THE OTHER ONE WITH FLICKR PAIR FROM APP GARDEN OF FLICKR OTHER ASK DONATION FOR LOOK BETTER LEARNING Sync.exe
@SET TSK_SCRIPTOR_NAME_=peservice.exe
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_% 
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
	@ECHO.
	@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
	@ECHO %IX1% -- Do You Want Kill ---- FLICKR SYNC SERVICE THING %TSK_SCRIPTOR_NAME_% 
	@ECHO ____________________ ?(Y/N)
	@SET INPUT=
	@SET /P INPUT=Type input: %=%
	@IF "%INPUT%" == "y" GOTO YES
	@IF "%INPUT%" == "Y" GOTO YES
	GOTO NOT
	:YES
		@ECHO ON
		TASKKILL /F /IM Flickr.exwuauclt.exee /T
		TASKKILL /F /IM peservice.exe /T
		TASKKILL /F /IM Sync.exe /T
		TASKKILL /F /IM Update.exe /T
		@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@SET TSK_SCRIPTOR_NAME_=PMB*
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_% 
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
	@ECHO.
	@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
	@ECHO %IX1% -- Do You Want Kill ---- CAMERA SONY CYBERSHOT @SET
	@ECHO ____________________ ?(Y/N)
	@SET INPUT=
	@SET /P INPUT=Type input: %=%
	@IF "%INPUT%" == "y" GOTO YES
	@IF "%INPUT%" == "Y" GOTO YES
	GOTO NOT
	:YES
		@ECHO ON
		TASKKILL /F /IM PMBGPSSupportTool.exe /T
		TASKKILL /F /IM PMBAnnounce.exe /T
		TASKKILL /F /IM PMBDeviceInfoProvider.exe /T
		@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@SET TSK_SCRIPTOR_NAME_=supercopier.EXE
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_% 
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
	@ECHO.
	@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
	@ECHO %IX1% -- Do You Want Kill ---- %TSK_SCRIPTOR_NAME_%
	@ECHO ____________________ ?(Y/N)
	@SET INPUT=
	@SET /P INPUT=Type input: %=%
	@IF "%INPUT%" == "y" GOTO YES
	@IF "%INPUT%" == "Y" GOTO YES
	GOTO NOT
	:YES
		@ECHO ON 
		TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T 
		@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@REM  ---- SAMSUNG MOBILE ANDROID SYSTEM
@SET TSK_SCRIPTOR_NAME_=KiesTrayAgent*
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_% 
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
	@ECHO.
	@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
	@ECHO %IX1% -- Do You Want Kill ---- SAMSUNG\KIES\KIESTRAYAGENT.EXE
	@ECHO ____________________ ?(Y/N)
	@SET INPUT=
	@SET /P INPUT=Type input: %=%
	@IF "%INPUT%" == "y" GOTO YES
	@IF "%INPUT%" == "Y" GOTO YES
	GOTO NOT
	:YES
		@ECHO ON 
		TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T 
		@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.

@SET TSK_SCRIPTOR_NAME_=MediaHome.EXE
TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T

@REM
@REM
@REM NeroMediaHome
@SET TSK_SCRIPTOR_NAME_=MediaHome.EXE
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_% 
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
	@ECHO.
	@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
	@ECHO %IX1% -- Do You Want Kill ---- NERO %TSK_SCRIPTOR_NAME_%
	@ECHO ____________________ ?(Y/N)
	@SET INPUT=
	@SET /P INPUT=Type input: %=%
	@IF "%INPUT%" == "y" GOTO YES
	@IF "%INPUT%" == "Y" GOTO YES
	GOTO NOT
	:YES
		@REM  -- C:\Program Files (x86)\Nero\KM\MediaHome.exe
		@ECHO ON 
		TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T
		@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@REM  -- C:\Program Files (x86)\Nokia\Nokia Suite\NokiaSuite.exe
@SET TSK_SCRIPTOR_NAME_=NokiaSuite.exe
@SET /A ACCUMULATOR=ACCUMULATOR+1
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
	@ECHO.
	@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
	@ECHO %IX1% -- Do You Want Kill ---- %TSK_SCRIPTOR_NAME_%
	@ECHO ____________________ ?(Y/N)
	@SET INPUT=
@REM @SET /P INPUT=Type input: %=%
	@IF "%INPUT%" == "y" GOTO YES
	@IF "%INPUT%" == "Y" GOTO YES
	GOTO NOT
	:YES
		@ECHO ON 
		TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T 
		@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@REM  -- C:\Program Files\IrfanView\i_view64.exe
@SET TSK_SCRIPTOR_NAME_=I_VIEW*
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_% 
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
	@ECHO.
	@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
	@ECHO %IX1% -- Do You Want Kill ---- IRFANVIEW %TSK_SCRIPTOR_NAME_%
	@ECHO ____________________ ?(Y/N)
	@SET INPUT=
	@SET /P INPUT=Type input: %=%
	@IF "%INPUT%" == "y" GOTO YES
	@IF "%INPUT%" == "Y" GOTO YES
	GOTO NOT
	:YES
		@ECHO ON 
		TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T 
		@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@REM  -- VISUAL BASIC 6 PROGRAM FOLDER
@REM  -- TRY A WILDCARD


@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@SET TSK_SCRIPTOR_NAME_=VB6.EXE
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_% 
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
	@ECHO.
	@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
	@ECHO %IX1% -- Do You Want Kill ---- %TSK_SCRIPTOR_NAME_%
	@ECHO ____________________ ?(Y/N)
	@SET INPUT=
	@SET /P INPUT=Type input: %=%
	@IF "%INPUT%" == "y" GOTO YES
	@IF "%INPUT%" == "Y" GOTO YES
	GOTO NOT
	:YES
		@ECHO ON 
		TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T 
		@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@REM  -- 
@ECHO -- VB __ 01 OF 02 NOT WILDING
@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@REM  D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe
@REM  
@SET TSK_SCRIPTOR_NAME_=ClipBoard Logger.exe
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_%
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
	@ECHO.
	@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
	@ECHO %IX1% -- Do You Want Kill ---- %TSK_SCRIPTOR_NAME_%
	@ECHO ____________________ ?(Y/N)
	@SET INPUT=
	@SET /P INPUT=Type input: %=%
	@IF "%INPUT%" == "y" GOTO YES
	@IF "%INPUT%" == "Y" GOTO YES
	GOTO NOT
	:YES
		@ECHO ON 
		TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T
		@ECHO OFF
:NOT
@ECHO OFF
@ECHO.

@REM  -- Tue 21 February 2017 02:53:40----------
@REM  -- ----
@REM  -- TASKLIST ALL FROM A FOLDER - Google Search
@REM  -- https://www.google.co.uk/search?q=TASKLIST+ALL+FROM+A+FOLDER&gws_rd=ssl
@REM  -- ----
@REM  -- ----
@REM  -- Windows - Taskkill for all running processes that exist in a spec@IFic directory - Super User
@REM  -- http://superuser.com/questions/728595/taskkill-for-all-running-processes-that-exist-in-a-spec@IFic-directory
@REM  -- ----
@REM  -- CARE THE EXAMPLE SOURCE USE 
@REM  -- @IF "%ERRORLEVEL%"=="0" TASKKILL /F /IM %%A
@REM  -- 
@REM  -- OUR PERSON USE THIS ONE INSTEAD OR
@REM  -- GETTER CRASH THE BATCH .BAT SCRIPTOR AS COMPLING THROUGH 
@REM  -- @IF %errorlevel% EQU 0 TASKKILL /F /IM %%A
@REM  -- 
@REM  --  @IF WONDERING
@REM  -- TASKLIST /FI "IMAGENAME EQ %%A" 2>NUL | FIND /I /N "%%A">NUL
@REM  TASKLIST /FI "IMAGENAME EQ %%A" 2>NUL | FIND /I /N "%%A">NUL
@REM  -- 
@REM  --  @IF WONDERING -- NOT THE ERROR AS MENTION MOMENT AGO BUT MOVED ON UN@IFORM REST OF CODE
@REM  --  BRACKET SYNTAX AROUND THE @IF CONDICTION
@REM  -- 

@REM   :HERE22 ____ @REMMER OUT ABOVE __ LET CODE FLOW ON



@REM   __________________________________________________________________
@REM   __ WANT THIS CODE TO DO WILD CARD EVERYTHING FROM FOLDER KILL
@REM   __ I'M WORKING ON HER
@REM   __ USEFUL ELSE WHERE FOR ONE WITH PATH FULL
@REM   __ AND MIGHT HAVE ODD CATCHY SMALL FILENAME WILDCARD INTO OTHER
@REM   __________________________________________________________________
GOTO JUMP_HERE23
@REM   _______________________________________________________________
@REM   ____ NOT @SET PROPER YET CODE ROUTINE BELOW
@REM   ____ MARKER BLOCK AGAIN BELOW
GOTO JUMP_HERE23
@REM  -- 
@REM  -- VISUAL BASIC 6 PROGRAM FOLDER 
@REM  -- 
@ECHO ON
@ECHO -------------------------------------------------------------------
@ECHO -- VB __ 02 OF 02 __ VB6 __ WILDCARD-ING ____ ALL IN FOLDER ____ D:\VB6\VB-NT
@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM  
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- 
@ECHO OFF
@SET TASK_NOT_RUNNING=0
@REM  for /r c:\apps\ %%G in (*.exe) do taskkill /F /IM %%~nxG

@REM  FOR /R D:\VB6\VB-NT %%G IN (*.EXE) DO TASKLIST /F /IM %%~NXG

@REM  FOR /R D:\VB6\VB-NT %%G IN (*.EXE) DO TASKLIST /IM %%~NXG >NUL

@REM  -- TASKILLL DON'T WORK LONG PATHING 
@REM  -- POWERSHELL MIGHT DO

@ECHO OFF
	
FOR /F "delims=" %%G in ('dir D:\VB6\VB-NT\*.EXE /a-d /b /s') DO (
	TASKLIST /FI "IMAGENAME EQ %%~nxG" 
	TASKLIST /FI "IMAGENAME EQ %%~nxG" | find "INFO: No tasks" > nul

	@IF %errorlevel% EQU 0 @ECHO %%
	
		@REM  @ECHO RESULT %%G
		@REM  @ECHO %%G
		@REM  PAUSE
		@REM  TASKLIST /FI "IMAGENAME EQ %%~nxG" 

	)
PAUSE

@REM  TASKLIST /FI "IMAGENAME EQ %%~nxG"  | find "INFO: No tasks" > nul

PAUSE

@IF %errorlevel% EQU 0 GOTO NOT_ERROR_TASK_NOT_RUNNING

TASKLIST /FI "IMAGENAME EQ %%~nxG"
@ECHO %%G
@ECHO %%~nxG


PAUSE


TASKLIST /FI "IMAGENAME EQ D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe"

PAUSE
FOR /F "delims=" %%G in ('dir D:\VB6\VB-NT\*.EXE /a-d /b /s') do @ECHO %%~nxG
PAUSE

PAUSE
FOR /F "delims=" %%G in ('dir D:\VB6\VB-NT\*.EXE /a-d /b /s') do @ECHO %%~nxG


PAUSE
FOR /F "delims=" %%G in ('dir D:\VB6\VB-NT\*.EXE /a-d /b /s') do @ECHO  /IM %%G



PAUSE

FOR /F "delims=" %%G in ('dir D:\VB6\VB-NT\*.EXE /a-d /b /s') do TASKLIST /IM %%G | FIND /I /N find "INFO: No tasks" > nul

PAUSE


FOR /F "delims=" %%G in ('dir D:\VB6\VB-NT\*.EXE /a-d /b /s') do TASKLIST /IM %%G~nxG | FIND /I /N find "INFO: No tasks" > nul


@REM  FOR /R D:\VB6\VB-NT %%G IN (*.EXE) DO (
	@REM  @ECHO %%~NXG
	@REM  @ECHO %%G~nxa
	@REM  @TASKLIST /IM %%G~NXG > NUL
	@REM  @IF %errorlevel% EQU 0 (
	@REM  @SET TASK_KILL_WORKING=1
@REM  	)
@REM  )
	
	
@REM  @ECHO %%~NXG 

@REM  @FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (

PAUSE

PUSHD D:\VB6\VB-NT    
FOR %%A IN (*.EXE) DO (
	TASKLIST /FI "IMAGENAME EQ %%A" 2>NUL | FIND /I /N find "INFO: No tasks" > nul
	@IF NOT %errorlevel% EQU 0 @SET TASK_KILL_WORKING=1
rem TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
rem @IF %errorlevel% EQU 1 GOTO NOT

)

PAUSE

	PUSHD D:\VB6\VB-NT
	FOR %%A IN (*.EXE) DO  @ECHO %%A 

PAUSE

@IF %TASK_KILL_WORKING%==1 (
	PUSHD D:\VB6\VB-NT
	FOR %%A IN (*.EXE) DO  @ECHO %%A 
)

PAUSE

@IF %TASK_KILL_WORKING%==1 (
	PUSHD D:\VB6\VB-NT
	FOR %%A IN (*.EXE) DO (
		@ECHO %%A
		@REM  TASKLIST /FI "IMAGENAME EQ %%A" 
	)
)

PAUSE

@IF %TASK_KILL_WORKING%==1 GOTO ASKER
GOTO NOT
:ASKER
@ECHO ON
@ECHO.
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% -- VB6 __ WILDCARD-ING ____ ALL IN FOLDER ____ D:\VB6\VB-NT
@ECHO ____________________ ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES
@IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES
PUSHD D:\VB6\VB-NT
FOR %%A IN (*.EXE) DO (
TASKLIST /FI "IMAGENAME EQ %%A" 2>NUL | FIND /I /N "%%A">NUL
rem TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
rem @IF %errorlevel% EQU 1 GOTO NOT

@IF %errorlevel% EQU 0 (
	@ECHO ON 
	TASKKILL /F /IM %%A
	@ECHO OFF
	)
)

:NOT
@ECHO OFF
@ECHO.
@REM pause


@REM  TASKLIST /FI "IMAGENAME EQ %%A" 2>nul | FIND /i /n "%%A">nul


@REM   ____ NOT @SET PROPER YET CODE ROUTINE ABOVE
:JUMP_HERE23


@REM  ---- BIGGER PROGRAM
@ECHO -------------------------------------------------------------------
@ECHO 1.. TSK_SCRIPTOR_NAME_=CMD.EXE
@ECHO 2.. TSK_SCRIPTOR_NAME_=EXPLORER.EXE
@ECHO 3.. TSK_SCRIPTOR_NAME_=CHROME.EXE
@ECHO -------------------------------------------------------------------


@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@SET TSK_SCRIPTOR_NAME_=CMD.EXE
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_% 
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
@ECHO.
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% -- Do You Want Kill ---- CMD __ WILL KILLER SELF WHEN /T SWITCH @SET LIKE IT IS
@ECHO ____________________ ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES
@IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES
@ECHO ON 
TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T
@ECHO OFF
@ECHO -- CMD.EXE -- ITSELF -- WON'T ALLOW ITSELF BUT AND WILL OTHER UNLESS /T 
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@SET TSK_SCRIPTOR_NAME_=EXPLORER.EXE
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_% -- WILL KILLER THIS CMD.BAT SCRIPT @IF CALLED BY THE EXPLORER.EXE MOUSE CLICKER 
@ECHO AND HAS /T OPTION WHICH IT DO ____ ANYTHING SUB DEPENDANT
@ECHO Mean -- And Any Child Processes Which Were Started By It.
@ECHO ____ EXPLORER WILL ATTEMPT RE@SET GO AGAIN AFTER A KILLER
@ECHO.
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% -- Do You Want Kill ---- %TSK_SCRIPTOR_NAME_%
@ECHO.
@ECHO ____________________ ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES
@IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES
@ECHO ON 
TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T

for /l %%x in (1, 1, 100) do (
	
	----
	BATCH COMMAND FOR LOOP - Google Search
	https://www.google.co.uk/webhp?sourceid=chrome-instant&rlz=1C1CHBD_en-GBGB721GB721&ion=1&espv=2&ie=UTF-8#q=BATCH+COMMAND+FOR+LOOP&*
	----
	@REM   ---- FOR LOOP 
    @REM   @ECHO %%x
    @REM   copy %%x.txt z:\whatever\etc
)
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO RE@SET_YES

@REM   ---- FOR LOOP 
GOTO RE@SET_NOT
:RE@SET_YES
EXPLORER
@ECHO ---- EXPLORER.EXE RESET AND RELOAD HAPPEN SUCCESFULLY
:RE@SET_NOT
:NOT
@ECHO OFF
@ECHO.


@REM  -----------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM
@REM
@SET TSK_SCRIPTOR_NAME_=CHROME.EXE
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% ---- %TSK_SCRIPTOR_NAME_% 
@ECHO ON
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%"
@ECHO OFF
TASKLIST /FI "IMAGENAME EQ %TSK_SCRIPTOR_NAME_%" 2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
@ECHO.
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO %IX1% -- Do You Want Kill ---- %TSK_SCRIPTOR_NAME_%
@ECHO ____________________ ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES
@IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES
@ECHO ON
TASKKILL /F /IM "%TSK_SCRIPTOR_NAME_%" /T
@ECHO OFF
:NOT
@ECHO OFF
@ECHO.
@REM  -----------------------------------------------------------------------------------
@REM  --
@REM  -----------------------------------------------------------------------------------
@REM  -- IN CHROME PRESS SH@IFT ESCAPE TO LOOK AT LIST OF PROCESS DETAIL
@REM  -- YOU MIGHT FIND OTHER USER LOGGED IN AND KILL PROCESS NOT LOG THEM OUT AT RELOAD
@REM  -- AND TAKE UP MEMORY TO USER TWO AND MORE
@REM  -- SO YOU GOT TO LOGG THEM OUT
@REM  -----------------------------------------------------------------------------------
@REM  -- LOG OFF OUT WON'T STOP THE PROCESS 
@REM  -- @SET THE OPTION TO STAY LOGIN ON WHEN CLOSE HAS TO BE OFF -- BUT THAT IS GLOBAL
@REM  -----------------------------------------------------------------------------------
@REM  -- AND @IF WANT FRESH START EVERY TIME THE USER THAT NOT LOG IN ON THE PROCESS WILL CLOSE UPON EXIT AND NOT RESTART WITHOUT LOGIN ON
@REM  -----------------------------------------------------------------------------------
@REM  -- @IF FRESH START USE IS ON --  OTHER USER ARE ABLE STILL BE LOGIN ON BUT THE EXTENSION PROCESS NOT BE USE  UNTIL SWITCH USER
@REM  -----------------------------------------------------------------------------------
@REM  -- OTHER OPTION -- CLOSING PAGE OF EXTRA USER WON'T CLOSE PROCESS FOR THEM
@REM  -- ONLY WHEN CLOSE IT ALL CHROME
@REM  -----------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------


@REM  -----------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------
@REM  WILD CARD BATCH @SET AND THEN END
@REM  -----------------------------------------------------------------------------------
@REM  -----------------------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@REM  --------------------------------------------------------------------

@REM  --------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------
@ECHO.
@ECHO -- WILD CARD __ 01 TO 05
@REM  --------------------------------------------------------------------
@ECHO 1.. TASKKILL /F /FI "STATUS EQ NOT RESPONDING" _______ WILD CARD ALL NOT RESPONDING
@ECHO 2.. TASKKILL /FI "IMAGENAME NE Z" /IM * /T     _______ WILD CARD EVERYTHING __ AND __ NOT FORCE 
@ECHO 3.. TASKKILL /F /FI "IMAGENAME NE Z" /IM * /T  _______ WILD CARD EVERYTHING __ YES FORCE BSD BSD XBOX
@ECHO 4.. TASKKILL /FI "WINDOWTITLE NE Z" /IM * /T      ____ WILD CARD EVERYTHING ALL WINDOWTITLE NOT FORCE
@ECHO 5.. TASKKILL /F /FI "WINDOWTITLE NE Z" /IM * /T ______ WILD CARD EVER ALL WINDOWTITLE YES FORCE BSD
@ECHO.
@ECHO WHEN DOUBLE VER@IFY ASKING QUESTION HERE ____________________
@ECHO TASKLIST PREVIEW OF 02 AND 03 IS THE SAME NOT WILDCARD VIEWER
@ECHO TASKLIST PREVIEW OF 04 AND 05 IS THE SAME NOT WILDCARD VIEWER
@ECHO.
@ECHO -------------------------------------------------------------------
@ECHO -------------------------------------------------------------------

@REM  --------------------------------------------------------------------

CALL :WILD_CARD__01_OF_05
GOTO JUMPER_ROUTINE_WILD_CARD__01_OF_05
::
:WILD_CARD__01_OF_05
@REM   @SETlocal
@REM  -------------------------------------------------------------------
@ECHO --------------------------------------------------------------------
@ECHO.
@ECHO -- WILD CARD __ 01 OF 05
@REM  --------------------------------------------------------------------
@REM  
@REM  
@REM  ------------------------------------------------------------------------
@REM  -- NOT RESPONDER -------------------------------------------------------------------
@REM  -- NOT REQUIRE A DOUBLE QUESTION EXAMPLE SHOWN GIVE
@REM  ------------------------------------------------------------------------

@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO -------------------------------------------------------------------
@ECHO.
@ECHO.
@ECHO %IX1%
@ECHO ON
@SET TSK_SCRIPTOR_NAME_="EQUAL NOT RESPONDING"
TASKLIST /FI "STATUS EQ NOT RESPONDING"
@ECHO OFF
@ECHO.
TASKLIST /FI "STATUS EQ NOT RESPONDING"  2>nul | find /I /N "Image Name">nul
@IF %errorlevel% EQU 1 GOTO NOT
@ECHO.
@ECHO __ Do You Want Kill ---- ALL NOT RESPONDING SERVICE @SET AS SHOWN AND OR PROCESS @SET 
@ECHO __ WILL MAYBE KILLER CMD.COM THIS BAT SCRIPT LEARN 
@ECHO __ ITSELF ONLY __ @IF SLACKING INTO NOT RESPONDING
@ECHO.
@ECHO %IX1% -- Do You Want Kill ---- %TSK_SCRIPTOR_NAME_%
@ECHO.
@ECHO ____________________ ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES
@IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES
@ECHO.
@ECHO ON
TASKKILL /F /FI "STATUS EQ NOT RESPONDING"
@ECHO OFF
:NOT
@ECHO OFF
@ECHO.
endlocal & goto :EOF
@REM   @REMMER --------------------------------------------------------------------------------
:JUMPER_ROUTINE_WILD_CARD__01_OF_05

@ECHO -------------------------------------------------------------------
@ECHO.
@REM  --------------------------------------------------------------------
@ECHO -- WILD CARD __ 02 OF 05
@REM  --------------------------------------------------------------------
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO -------------------------------------------------------------------
@ECHO.
@ECHO %IX1%
@ECHO.
@ECHO __ Do You Want KILLER ---- WILD CARD EVERYTHING __ AND __ NOT FORCE 
@ECHO __ BIG TIME WILDCARD @IF _YES_ _DOUBLE QUESTION_ ARE YOU SURE TASKLIST PREVIEW
@ECHO.
@ECHO ____________________ Q 01 OF 02 ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES
@IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES
@ECHO.
@ECHO ON
TASKLIST /FI "IMAGENAME NE Z" 
@ECHO OFF
@ECHO ____________________ Q 02 OF 02 YES SURE ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES_2
@IF "%INPUT%" == "Y" GOTO YES_2
GOTO NOT
:YES_2
@ECHO ON
TASKKILL /FI "IMAGENAME NE Z" /IM * /T
@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.
@REM  --------------------------------------------------------------------
@ECHO -- WILD CARD __ 03 OF 05
@REM  --------------------------------------------------------------------
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO -------------------------------------------------------------------
@ECHO.
@ECHO %IX1%
@ECHO.
@ECHO __ Do You Want Kill ---- WILD CARD EVERYTHING 
@ECHO __ YES FORCE BSD BLUE SCREEN OF DEATH
@ECHO __ FIND CAME FROM SEARCHER WEB _XBOX_ USER NONSE __ 
@ECHO __ IS MEANT TO BE USEFUL COMMAND CLEAR THE SYSTEM QUICKER
@ECHO __ BIG TIME WILDCARD @IF _YES_OPTION REPSOND WITHER 
@ECHO __ AND THEN  _DOUBLE QUESTION_ ARE YOU SURE TASKLIST PREVIEW
@ECHO.
@ECHO ____________________ Q 01 OF 02 ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES
@IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES
@ECHO.
@ECHO ON
TASKLIST /FI "IMAGENAME NE Z" 
@ECHO OFF
@ECHO ____________________ Q 02 OF 02 YES SURE ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES_2
@IF "%INPUT%" == "Y" GOTO YES_2
GOTO NOT
:YES_2
@ECHO ON
TASKKILL /F /FI "IMAGENAME NE Z" /IM * /T
@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.
@REM  --------------------------------------------------------------------
@ECHO --- WILD CARD __ 04 OF 05
@REM  --------------------------------------------------------------------
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO -------------------------------------------------------------------
@ECHO.
@ECHO %IX1%
@ECHO.
@ECHO __ Do You Want Kill ---- WILD CARD EVERYTHING ALL WINDOWTITLE NOT FORCE
@ECHO __ BIG TIME WILDCARD @IF _YES_OPTION REPSOND WITHER 
@ECHO __ AND THEN  _DOUBLE QUESTION_ ARE YOU SURE TASKLIST PREVIEW
@ECHO.
@ECHO ____________________ Q 01 OF 02 ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES
@IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES
@ECHO.
@ECHO ON
TASKLIST /FI "WINDOWTITLE NE Z" 
@ECHO OFF
@ECHO ____________________ Q 02 OF 02 YES SURE ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES_2
@IF "%INPUT%" == "Y" GOTO YES_2
GOTO NOT
:YES_2
@ECHO ON
TASKKILL /FI "WINDOWTITLE NE Z" /IM * /T
@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.
@REM  --------------------------------------------------------------------
@ECHO -- WILD CARD __ 05 OF 05
@REM  --------------------------------------------------------------------
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO -------------------------------------------------------------------
@ECHO.
@ECHO %IX1%
@ECHO.
@ECHO __ Do You Want Kill ---- WILD CARD EVER ALL WINDOWTITLE YES FORCE ____BSD____
@ECHO __ ALMOST SAME OPERATION AS ABOVE BUTTER BIGGER AND /F (IS __/FORCE__) OPTION USE 
@ECHO __ ALSO AND TASKLIST DON'T BOTHER THEM /F OPTION AND THEN SAME RESULT AS ABOVE TASKLIST __ @IF SELECTED
@ECHO __ BIG TIME WILDCARD @IF _YES_OPTION REPSOND WITHER 
@ECHO __ AND THEN  _DOUBLE QUESTION_ ARE YOU SURE TASKLIST PREVIEW
@ECHO.
@ECHO ____________________ Q 01 OF 02 ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES
@IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES
@ECHO.
@ECHO ON
TASKLIST /FI "WINDOWTITLE NE Z"
@ECHO OFF
@ECHO ____________________ Q 02 OF 02 YES SURE ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES_2
@IF "%INPUT%" == "Y" GOTO YES_2
GOTO NOT
:YES_2
@ECHO ON
TASKKILL /F /FI "WINDOWTITLE NE Z" /IM * /T
@ECHO OFF
:NOT
@ECHO OFF
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.
@REM  --------------------------------------------------------------------
@ECHO -- WILD CARD __ 05 OF 05
@REM  --------------------------------------------------------------------
@SET /A ACCUMULATOR=ACCUMULATOR+1
@SET IX1=%ACCUMULATOR% of %VAR_COUNTER%
@ECHO -------------------------------------------------------------------
@ECHO.
@ECHO %IX1%
@ECHO.
@ECHO __ Do You Want RE_BOOT THE WIN XP COMPUTER TWO __1-ASUS-X5DIJ AND __2-ASUS-EEE 
@ECHO.
@ECHO ____________________ Q 01 OF 02 ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES
@IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES
@ECHO ON
SHUTDOWN -R -F -M 2-ASUS-EEE -T 10
SHUTDOWN -R -F -M 1-ASUS-X5DIJ -T 10
@REM  -- WAIT FOR OTHER WITH -T 10
@ECHO OFF
:NOT
@ECHO OFF
@ECHO.



@ECHO -------------------------------------------------------------------
@ECHO.
@REM  --------------------------------------------------------------------
@REM  -- FINALLY EDITOR NOTEPAD LEARN 
@REM  --------------------------------------------------------------------
@SET /A ACCUMULATOR=ACCUMULATOR+1
@ECHO -------------------------------------------------------------------
@ECHO.
@ECHO %ACCUMULATOR% OF %VAR_COUNTER%  -- DO YOU WANT EDIT HERE NOTEPAD++ C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT 
@ECHO ____________________ ?(Y/N)
@SET INPUT=
@SET /P INPUT=Type input: %=%
@IF "%INPUT%" == "y" GOTO YES
@IF "%INPUT%" == "Y" GOTO YES
GOTO NOT
:YES

@REM START "" /MAX /HIGH "C:\Program Files (x86)\Notepad++\notepad++.exe" "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT"

START "" /MAX /HIGH "C:\Program Files (x86)\Notepad++\notepad++.exe" %~f0


GOTO ENDING
:NOT
@ECHO.


@ECHO -------------------------------------------------------------------
@ECHO.

@ECHO -------------------------------------------------------------------
@ECHO.
@ECHO THE END ----------
@ECHO.
@ECHO -------------------------------------------------------------------

@REM  PAUSE 

:ENDING

GOTO ENDING2
@ECHO -------------------------------------------------------------------
:ENDING2

@ECHO.

@REM 40 SECONDS
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@ECHO OFF

@IF "%WINXP__%"=="YES" (
@ECHO WINDOWS XP
@REM 40 SECONDS
FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
FOR /L %%# IN (40,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)
)

@IF NOT "%WINXP__%"=="YES" (
	TIMEOUT /T 40
)

:VERY_ENDER_BLOCK

:JUMPER_END_E
:ENDE

EXIT