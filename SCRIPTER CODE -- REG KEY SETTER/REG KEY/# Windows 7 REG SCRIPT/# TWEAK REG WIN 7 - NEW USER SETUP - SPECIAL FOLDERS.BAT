@ECHO OFF
@CD\

ECHO ____ # TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT ____ 
TITLE %~f0

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


REM GOTO BEGIN_JUMP_COMMENT_HEADER_BLOCK
REM -----------------------------------------------------------------------
REM -----------------------------------------------------------------------
REM -----------------------------------------------------------------------

REM COMMENT
REM ---------------------------------------------------------
REM CODE HERE TO RUN REG WIN 7 AND WIN XP AND UP WIN 10 SPECIAL FOLDER TWEAKER COMPLEX
REM ---------------------------------------------------------
REM AND FIRST A NOTE FOUND ABOUT SETTING ADMIN FOR ALL PROGRAM
REM ---------------------------------------------------------
REM ENDCOMMENT


REM COMMENT
REM ---------------------------------------------------------

REM NEW LINK SHOW HOW TO RUN ALL PROGRAM ALWAYS IN ADMIN
REM ----
REM How to grant or get Elevated Privileges in Windows 10/8/7
REM http://www.thewindowsclub.com/elevated-privileges-windows
REM ----

REM --------------------------------------------------------------------
REM RUN THROUGH 1 STEP 2 3 
REM --------------------------------------------------------------------
REM OPEN COMMAND PROMPT WINDOW WITH ADMINISTRATOR 
REM --------------------------------------------------------------------
REM --------------------------------------------------------------------
REM WITH SHIFT OR CONTROL KEY WHILE RIGHT CLICK IN EXPLORER TO ENABLE THE CONTEXT OPTION TO SHOW
REM --------------------------------------------------------------------

REM --------------------------------------------------------------------
REM TYPE THIS LINE IN ---- PASTE IT IN ---- IF POSSIBLE
REM MIGHT TAKE FEW SECOND TO LOAD FIRST TIME
REM --------------------------------------------------------------------
REM SECPOL.MSC
REM --------------------------------------------------------------------

REM LOOK AT THE TREE AND NODE PLUS CLICKER
REM -------------------------------------------------------
REM 1.. 2ND DOWN -- LOCAL POLICE'S
REM NEXT BRANCH 
REM -------------------------------------------------------
REM 2.. SECURITY OPTIONS
REM NEXT BRANCH FEW PAGE DOWN AND THIRD FROM BOTTOM
REM -------------------------------------------------------
REM 3.. User Account Control: RUN ALL ADMINISTRATORS IN ADMIN APPROVAL MODE

REM -------------------------------------------------------
REM HERE SET FROM ENABLED TO DISABLED
REM REQUIRE A REBOOT
REM -------------------------------------------------------

REM INTRUCTION FROM GUI FORM
REM -------------------------------------------------------
REM User Account Control: Turn on Admin Approval Mode

REM This policy setting controls the behavior of all User Account Control (UAC) policy settings for the computer. If you change this policy setting, you must restart your computer.

REM The options are:

REM • Enabled: (Default) Admin Approval Mode is enabled. This policy must be enabled and related UAC policy settings must also be set appropriately to allow the built-in Administrator account and all other users who are members of the Administrators group to run in Admin Approval Mode. 

REM • Disabled: Admin Approval Mode and all related UAC policy settings are disabled. Note: If this policy setting is disabled, the Security Center notifies you that the overall security of the operating system has been reduced.
REM -------------------------------------------------------

REM -------------------------------------------------------
REM WENT THOUGHT A LOAD OF WEB SITE 
REM ALL TALK ABOUT LOADING ONE UP
REM NOT GOOD IN MY EYE FOR A PROGRAMMER THAT CHANGE RECOMPILE THE EXE
REM AND THINGS CHANGE

REM AND FINALLY AT MS WINDOWS CLUB
REM LOOK LIKE THERE ARE ADVERTISEMENT ON SITE
REM AND IDIOT WORDIN
REM -------------------------------------------------------

REM 1 STEP 2 STEP 3 DONE ALL I WANT WITH THE DRIBBLE DROOL

REM -------------------------------------------------------
REM -------------------------------------------------------
REM ENDCOMMENT 

REM PAUSE



REM COMMENT 
REM ---------------------------------
REM ---------------------------------

REM ---------------------------------
REM HERE BEGIN THE REGISTRY SPECIAL FOLDER 
REM ---------------------------------

:JUMP_OVER_COMMENT_BLOCK

REM -- FOUND 2015
REM -- NOTEPAD++ CHANGE HERE
REM -- https://peter.hahndorf.eu/blog/post/2011/12/23/ChangingTheLocationOfWindowsSpecialFoldersOnTheCommandLine

REM -- IF VALUENAME IS DEFAULT THEN YOU SET /v "" Default Sets All Values as Said
REM -- /v       The value name, under the selected Key, to add.
REM -- /ve      adds an empty value name (Default) for the key.

REM -- WELL MOST SOURCE INCLUDING MICROSOFT SAID ADD A CAROT ^ IN FRONT OF % BUT DOUBLE %% IS ANSWER
REM -- WELL THAT SOURCE WAS FOR THE WINDOWS 7 UPWARDS

REM -- AND YOU CAN HAVE SOMETHING LIKE THIS 
REM -- REG add HKCU\SOFTWARE\Classes\sourcecode\shell\open\command /v "" /t REG_SZ /d "\"C:\Program Files\Vim\vim72\gvim.exe\" --remote-tab-silent \"%%1\"" /f
REM -- http://vim.wikia.com/wiki/Windows_file_associations

REM -- ---------------------------------------------------------------------------------- 
REM -- WIN XP NOTE TAKING --------------------------------------------------------------- 
REM -- ONE GOOD HELPFUL NOTE IS IF YOU EVER CHANGE YOU COMPUTER NAME
REM -- YOU GOT TO SET THE OWNERSHIP TO YOUR USERS FROM ADJUSTED TO DEFAULT ADMINISTRATORS
REM -- BE GOING IN SAFE-MODE AND THAT WAY
REM -- ---------------------------------------------------------------------------------- 
REM -- A NOTE FOR THE DAY - WHEN I APPLY THESE SETTING IN SAFE MODE
REM -- AND RELOGON TO DO FOR ALL USER
REM -- WON'T ACCEPT PASSWORD TO GET IN UNLESS REBOOT FIRST 
REM -- ---------------------------------------------------------------------------------- 

REM -- ---------------------------------------------------------------------------------- 
REM -- PROJECT TIME CONSUMPTION 
REM -- ---------------------------------------------------------------------------------- 
REM -- PROJECT ORIGINAL LARGER WORK REGISTRY SET-UP WITH WINDOWS XP PRO 
REM -- PROJECT TIME DAY 1
REM -- PROJECT WORK FRI 22 JULY START    -- Fri 22-Jul-2016 12:09:37
REM -- PROJECT WORK FRI 22 JULY MID WAY  -- Fri 22-Jul-2016 21:06:55
REM -- PROJECT WORK FRI 22 JULY ENDE     -- Fri 22-Jul-2016 23:21:00
REM -- PROJECT WORK FRI 22 JULY ENDE     -- Sat 23-Jul-2016 00:38:31
REM -- PROJECT TIMER DAY 1
REM -- Calculate duration between two dates – results
REM -- http://www.timeanddate.com/date/durationresult.html?d1=22&m1=7&y1=2016&d2=23&m2=7&y2=2016&h1=12&i1=09&s1=27&h2=00&i2=38&s2=31
REM -- Result: 12 hours, 29 minutes and 4 seconds
REM -- ------------------------------------ 12 HOUR WORK PROJECT -- DAY 1 ---------------
REM -- PROJECT TIME DAY 2 
REM -- PROJECT WORK SAT 23 JULY START    -- Sat 23-Jul-2016 08:20:19
REM -- PROJECT WORK SAT 23 JULY ENDE     -- Sat 23-Jul-2016 11:05:40
REM -- PROJECT WORK ----------- START    -- Thu 04-Aug-2016 19:34:00
REM -- PROJECT WORK ----------- ENDE     -- Thu 04-Aug-2016 

REM -- PROJECT TIMER DAY 2 
REM -- Calculate duration between two dates – results
REM -- http://www.timeanddate.com/date/durationresult.html?d1=23&m1=7&y1=2016&h1=08&i1=20&s1=19
REM -- Result: 1 hour, 50 minutes and 55 seconds
REM -- ------------------------------------ 02 HOUR WORK PROJECT -- DAY 2 ---------------
REM -- PROJECT TOTAL WORK TIME 2 DAY
REM -- 

REM ----------------------------------------------------------------------------------
REM ----------------------------------------------------------------------------------
REM ENDCOMMENT



ECHO #### ---------------------------------------------------- ####
ECHO #### -- MATT.LAN@BTINTERNET.COM JANUARY 201SIXTEEN ------ ####
ECHO #### -- 1ST ROLL OUT VERSION WINDOWS XP SP3 PRO --------- ####
ECHO #### ---------------------------------------------------- ####
ECHO #### -- MATT.LAN@BTINTERNET.COM -- JULY 201SIXTEEN ------ ####
ECHO #### -- 2ND ROLL OUT VERSION WINDOWS 7 ------------------ ####
ECHO #### -- UPTO ONLY SPECIAL FOLDRER BUT AND INCLUDE ------- ####
ECHO #### -- WHOLE SPECIAL FOLDER TO SET ATTENTION ----------- ####
ECHO #### ---------------------------------------------------- ####
ECHO #### ---------------------------------------------------- ####
ECHO #### -- SAFE GUARD ONLY RUN ON KNOWN COMPUTER ----------- ####
ECHO #### -- NOT ACCIDENTAL RUN WITHJOUT QUESTION ------------ ####
ECHO #### -- AND THIS VERSION ONLY WINDOWS 7 64 BIT ---------- ####
ECHO #### ---------------------------------------------------- ####
ECHO #### -- THESE ARE ONLY THE SPECIAL FOLDER SETTING ------- ####
ECHO #### -- INCLUDE AS MANY AS POASIABLE COMMON ALSO  ------- ####
ECHO #### -- NOTICE THESE WIN 7 64 BIT AS SAME AS ------------ #### 
ECHO #### -- REGESTRY AS 32 BIT WIN XP ORGINAL WORK  --------- ####
ECHO #### -- BUT MAYBE 64 BIT DIFFERENT IN BINARY VALUE SET -- ####
ECHO #### ---------------------------------------------------- ####
ECHO #### -- READY TO GO ------------------------------------- ####
ECHO #### ---------------------------------------------------- ####

@REM -------------------------------------------------------------------------------------
@REM -- FIND OUR OWN PROGRAM FOR REFERENCE FOR THE MAIN DRIVE PATH WE USE LIKE GRUB4DOS 
@REM -- METHOD OF DRIVE STEPPING WITH PATH  FILE
@REM -- MAKE SURE FIRST FIND IS ONE -- IS THE ONE WANT AND NOT ANOTHER LATER ON THE DRIVE
@REM -- TALK HERE BUT NOT A CHECK DONE YET
@REM -------------------------------------------------------------------------------------

@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM ----------- FIND OUR DRIVE FAT32 OF LINK AND SHORT-CUT
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------

REM --------------------------------------------------------------
REM --------------------------------------------------------------
REM --------------------------------------------------------------
REM --------------------------------------------------------------
REM --------------------------------------------------------------
REM --------------------------------------------------------------
REM --------------------------------------------------------------
:BEGIN_JUMP_COMMENT_HEADER_BLOCK



@ECHO OFF
setlocal EnableDelayedExpansion
@REM -------------------------------------------------------------------------------------
SET DRIVEPATH=NOT_BEEN_SET
@REM -- SET HARDCODED WILL FIND OUR PATH LATER OR IF NOT VALUE HERE WILL SHOW AND QUESTION
@REM -------------------------------------------------------------------------------------

FOR %%G IN (A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@IF EXIST "%%G:\REG KEY SETTINGS\REG KEY\# Windows 7 REG SCRIPT\# TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT" IF "!DRIVEPATH!"=="NOT_BEEN_SET" SET DRIVEPATH=%%G:)

FOR %%G IN (A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@IF EXIST "%%G:\REG KEY SETTINGS\REG KEY\# Windows 7 REG SCRIPT\# TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT" ECHO #### -- -------------------------------------------------------------- #### & ECHO #### -- SCAN FOUND DRIVEPATH UNIQUE FILE IS ON ------------ DRIVE %%G -- ####)

setlocal DISableDelayedExpansion

ECHO #### -- USING FIRST PATH FOUND DRIVE ---------------------- DRIVE %DRIVEPATH% - ####


@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- --------------------------------------------------------
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- NOT A MATCHING FILE OF THIS FOLLOWING NAME 
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- ON ALL DRIVE CHECKED
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- --------------------------------------------------------
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- USUALLY THE NAME IS OF THIS BATCH FILE
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- --------------------------------------------------------
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- *:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\# TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- --------------------------------------------------------
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- FILE SEARCHED TO FIND THE DRIVE TO USE AS AN RELOCATABLE
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- THE RELOCATE IS A IDEA TO HAVE SPECIAL FOLDDERS ON A ---
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- SMALL PARTICTION OF DRIVE AND KEEP IT RELOCATEABLE -----
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- --------------------------------------------------------
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- WILL ABANDON JOB WITHOUT THIS FILE FOUND IN PLACE
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" ECHO -- --------------------------------------------------------
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" PAUSE
@IF "%DRIVEPATH%"=="NOT_BEEN_SET" GOTO VERY_END
@ECHO.
PAUSE

@ECHO #### -- --------------------------------------------------------    -- ####
@ECHO #### -- DOUBLE CHECK THIS RESULT IS NOT MORE THAN ONE
@ECHO #### -- --------------------------------------------------------    -- ####

@ECHO #### -- SET THE RELOCATABLE DRIVE OF OUR *DRIVEPATH* VARIABLE -- %DRIVEPATH% -- ####
@ECHO #### -- --------------------------------------------------------    -- ####
@ECHO.
PAUSE

CD \
@ECHO OFF
@SET VARLATCH=0

@REM OS=Windows_NT

@IF NOT "%OS%"=="Windows_NT" ECHO NOT A MATCHING WINDOWS 7 OR XP OS DETECTED -- %OS% -- GOTO END_SUB2
@IF NOT "%OS%"=="Windows_NT" GOTO END_SUB2
ECHO #### --------------------------------------------------------------------------------
ECHO #### -- CHECK WINDOWS 7 OR XP OPERATING SYSTEM DETECTED OKAY HERE -- %OS% -----


@IF NOT "%PROCESSOR_ARCHITECTURE%"=="AMD64" ECHO NOT A MATCHING 64 BIT COMPUTER -- %PROCESSOR_ARCHITECTURE% -- GOTO END
@REM @IF NOT "%PROCESSOR_ARCHITECTURE%"=="AMD64" GOTO END_SUB
@REM -- THIS WILL BE THE SPECIAL FOLDER SETUP BATCH UNIVERSAL ANY OPERATING SYSTEM WIN 10 AND 32 BIT OR 64 BIT
@REM -- AS TEXT REQUIRED ONLY OR IF BINARY DATA  VALUE USE 64 BIT

ECHO #### -- CHECK 64 BIT COMPUTER DETECTED OKAY HERE -------------- %PROCESSOR_ARCHITECTURE% ---------------
@REM -- AMD TRENDY


IF "%COMPUTERNAME%" == "--------------" SET VARLATCH=1
IF "%COMPUTERNAME%" == "5-ASUS-P2520LA" SET VARLATCH=1
IF "%COMPUTERNAME%" == "4-ASUS-GL522VW" SET VARLATCH=1
IF "%COMPUTERNAME%" == "4-ASUS-GL522VW" SET VARLATCH2=WIN10
IF "%COMPUTERNAME%" == "7-ASUS-GL522VW" SET VARLATCH=1
IF "%COMPUTERNAME%" == "7-ASUS-GL522VW" SET VARLATCH2=WIN10
IF "%COMPUTERNAME%" == "3-LINDA-PC" SET VARLATCH=1
IF "%COMPUTERNAME%" == "1-ASUS-X5DIJ" SET VARLATCH=1
IF "%COMPUTERNAME%" == "1-ASUS-X5DIJ" SET VARLATCH2=WINXP
IF "%COMPUTERNAME%" == "2-ASUS-EEE" SET VARLATCH=1
IF "%COMPUTERNAME%" == "2-ASUS-EEE" SET VARLATCH2=WINXP
IF "%COMPUTERNAME%" == "8-MSI-GP62M-7RD" SET VARLATCH=1
IF "%COMPUTERNAME%" == "8-MSI-GP62M-7RD" SET VARLATCH2=WIN10

@REM ---------------------------------------------------------------
ver | find "5.1." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WINXP
ver | find "6.1." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WIN07
ver | find "10.0." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WIN10
@REM ---------------------------------------------------------------

IF %VARLATCH% == 0 GOTO NOT_A_MATCH_RESULT

IF "%USERNAME%"=="MATT 01" SET VARLATCH=2
IF "%USERNAME%"=="MATT 02" SET VARLATCH=2
IF "%USERNAME%"=="MATT 03" SET VARLATCH=2
IF "%USERNAME%"=="MATT 04" SET VARLATCH=2
IF "%USERNAME%"=="MATT 05" SET VARLATCH=2

:NOT_A_MATCH_RESULT

@REM TEST TOGGLE
@REM @SET VARLATCH=0

REM ECHO. = METHOD TO GET A BLANK LINE
REM @IF %VARLATCH% == 1 ECHO.
REM @IF %VARLATCH% == 2 ECHO. 
@ECHO #### ----------------------------------------------------- ####
ECHO #### -- COMPUTER NAME = %COMPUTERNAME% -- USER NAME = %USERNAME%
@ECHO #### ----------------------------------------------------- ####

@IF NOT %VARLATCH% == 0 @ECHO #### ----------------------------------------------------- ####
@IF %VARLATCH% == 1 ECHO #### -- MATCHING COMPUTER DETECTED - BUT NOT USER NAME
@IF %VARLATCH% == 2 ECHO #### -- MATCHING COMPUTER AND USER NAME DETECTED SUCCESS RESULT OKAY ----------- ####
@IF %VARLATCH% == 2 ECHO #### -- -------------------------------------------------RESULT OKAY ----------- ####
@IF %VARLATCH% == 0 ECHO #### -- NOT A MATCHING COMPUTER NAME DETECTED -- GOTO END 
@IF %VARLATCH% == 0 GOTO END_SUB2
@IF %VARLATCH% == 1 GOTO END_SUB2

@REM -- GOT A MATCHING COMPUTER AND USER ADJUST TO 1 TO GO
@SET VARLATCH=1

@ECHO #### ----------------------------------------------------- ####
@ECHO #### -- MAKE SURE YOU WANT TO GO ON WITH THIS CODE *CTRL C* IF YOU DON'T ------- ####
@ECHO #### ----------------------------------------------------- ####
@ECHO.
@PAUSE

@SET VARSTRING=----------------------------------------------------------------------
@ECHO OFF

@REM -- MIGHT AS WELL - SET THE DRIVE AND PATH FIRST AS IT IS DONE LATER
@%DRIVEPATH%
@CD \



@ECHO #### ----------------------------------------------------- ####
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------

reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseHoverTime" /t REG_SZ /d "20" 
@REM MouseHoverTime SPEEDY TIME HERE BUT AND LIKE A DECAY TIME AWAY AGAIN SEEM TO GET STUCK OPEN
reg.exe ADD "HKCU\Control Panel\Keyboard" /f /v "KeyboardDelay" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\Keyboard" /f /v "KeyboardSpeed" /t REG_SZ /d "31"
REM reg.exe ADD "HKCU\Control Panel\Keyboard" /f /v "KeyboardSpeed" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\Keyboard" /f /v "InitialKeyboardIndicators" /t REG_SZ /d "1"

REM TIMEOUT /T 20


@REM -- How to Fix High CPU Usage by RuntimeBroker.exe on Windows 10 - Google Chrome
@REM ----------------------
reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Services\TimeBroker]" /f /v "Start” /t REG_DWORD /d 3

reg.exe ADD "HKCU\Control Panel\Desktop\WindowMetrics" /f /v "MinAnimate" /t REG_SZ /d "0"
reg.exe ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OptimalLayout" /f /v "EnableAutoLayout" /t REG_DWORD /d 0
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "PaintDesktopVersion" /t REG_DWORD /d 0

reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CursorBlinkRate" /t REG_SZ /d "10"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "DragFullWindows" /t REG_SZ /d "1"

reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseTrails" /t REG_SZ /d "4"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseSpeed" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseSensitivity" /t REG_SZ /d "4"

reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "WheelScrollLines" /t REG_SZ /d "5"
reg.exe ADD "HKCU\SOFTWARE\Microsoft\IntelliPoint\Scrolling" /f /v "ScrollLines" /t REG_DWORD /d 5
reg.exe ADD "HKCU\SOFTWARE\Microsoft\IntelliPoint\Scrolling" /f /v "AccelerationCurve" /t REG_DWORD /d 5

@rem __ reg.exe ADD "HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\Enable" /f /v "SUCCESS /t REG_SZ /d	"Y"	

@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------


@ECHO ON
@ECHO.
@ECHO %VARSTRING%
@ECHO -- SET THE SPECIAL SYSTEM FOLDERS ------------------------------------------------------------

@REM ---------------------
@REM -- SET DESKTOP FOLDER
@REM ---------------------
@ECHO -- AFTER DESKTOP SET TO A NEW FOLDER -- YOU DON'T NEED REQUIRE LOG OFF - ON -- ONLY F5 REFRESH

%DRIVEPATH%
CD \

@SET FOLDER_NAME1=01 Desktop
@SET FOLDER_NAME2=01 Desktop\#_%COMPUTERNAME%
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME1%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME1%\"
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%\"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Desktop" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Desktop" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
@SET Desktop_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%

@SET FOLDER_NAME2=01 Favs
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%\"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Favorites" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Favorites" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
@SET Favs_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%

@SET FOLDER_NAME1=01 SendTo
@SET FOLDER_NAME2=01 SendTo\#_%COMPUTERNAME%
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME1%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME1%\"
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%\"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "SendTo" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "SendTo" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
@SET SendTo_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%

@REM ---------------------------------------------------------------------
@REM ---------------------------------------------------------------------
@REM ---- Start Menu 1 2 3
@REM ---------------------------------------------------------------------
@REM ---------------------------------------------------------------------
@SET FOLDER_NAME1=01 Start Menu
@SET FOLDER_NAME2=01 Start Menu\#_%COMPUTERNAME%
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME1%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME1%\"
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%\"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Start Menu" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Start Menu" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
@SET Start_Menu_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%

@SET FOLDER_NAME2=01 Start Menu\#_%COMPUTERNAME%\Programs
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Programs" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Programs" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"


@SET FOLDER_NAME2=01 Start Menu\#_%COMPUTERNAME%\Programs\Startup
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%"

reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Startup" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Startup" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
set Start_Menu_Startup_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%
REM -----------------------------------------------------------------------------
REM MAKE SURE YOU HAVE STARTUP ITEM IN HERE TO USE 
REM TRY MAKE SURE NOT TO RENAME THE FOLDER AND LOSE COUNT OF HERE HAS THE SETTING
REM -----------------------------------------------------------------------------

REM PROBLEM WITH STARTUP JUMPTO LOOK HERE 
REM HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders

@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ---- MY FOLDERS 1 2 3 4
@REM ---- MY FOLDERS -- MY DOCS -- MY PICTURES -- MY MUSIC -- MY VIDEO 
@REM ----------------------------------------------------------------
@SET FOLDER_NAME1=D:\# MY DOCS
@SET FOLDER_NAME2=D:\# MY DOCS\# 01 My Documents
@IF NOT EXIST "%FOLDER_NAME1%\" MKDIR "%FOLDER_NAME1%\"
@IF NOT EXIST "%FOLDER_NAME2%\" MKDIR "%FOLDER_NAME2%\"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Personal /t REG_EXPAND_SZ /d "%FOLDER_NAME2%"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Personal /t REG_SZ /d "%FOLDER_NAME2%"

@SET My_Documents_FOLDER=%FOLDER_NAME2%

@SET FOLDER_NAME2=D:\# MY DOCS\# 01 My Documents\My Pictures
@IF NOT EXIST "%FOLDER_NAME2%\" MKDIR "%FOLDER_NAME2%\"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "My Pictures" /t REG_EXPAND_SZ /d "%FOLDER_NAME2%"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "My Pictures" /t REG_SZ /d "%FOLDER_NAME2%" 
@SET IMAGE_FOLDER=%FOLDER_NAME2%

@SET FOLDER_NAME2=D:\# MY DOCS\# 01 My Documents\My Music
@IF NOT EXIST "%FOLDER_NAME2%\" MKDIR "%FOLDER_NAME2%\"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "My Music" /t REG_EXPAND_SZ /d "%FOLDER_NAME2%"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "My Music" /t REG_SZ /d "%FOLDER_NAME2%" 
@SET MUSIC_FOLDER=%FOLDER_NAME2%

@SET FOLDER_NAME2=D:\# MY DOCS\# 01 My Documents\My Video
@IF NOT EXIST "%FOLDER_NAME2%\" MKDIR "%FOLDER_NAME2%\"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "My Video" /t REG_EXPAND_SZ /d "%FOLDER_NAME2%"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "My Video" /t REG_SZ /d "%FOLDER_NAME2%"
@SET VIDEO_FOLDER=%FOLDER_NAME2%
 
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- SET THE COMMON START MENU PROGRAMS - FOLDER ONLY FOUND TO BE SETTABLE IN TWEAK XP PRO
@REM -- STARTMENU COMMONS
@REM ----------------------------------------------------------------
@REM -- THE COMMONS ARE SET IN HK LOCAL MACHINE HKLM
@REM -- AND OTHER CURRENT USER
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- ORIGINAL -- %PUBLIC%\Desktop
@REM -- ORIGINAL -- C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools
@REM ----------------------------------------------------------------

@REM -- DON'T REQUIRE 64 BIT SETTING HERE ON FAVS -- BUT WHY DO I MERGE HEM WHEN TALK FORBIDDEN
@REM -- OR WHY A FAULT CONFLICT MERGE HERE
@REM -- NOT SURE
@REM ----------------------------------------------------------------
@SET FOLDER_NAME2=01 Favs Common
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%\"
reg.exe ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Favorites" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Common Favorites" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
@SET Favourite_Common_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- FAV'S NONE TO DO WITH 64 BIT 
@REM -- AND STAY SAME FOLDER ALL COMPUTER OS VERSION
@REM -- MAYBE SOON TEST ON NTFS DRIVE
@REM -- AS ERROR IN OUR HISTORY WITH URL SHORT CUT ON NTFS CORRUPT THEM IN WIN XP
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM -- DETECT 64 BIT WOW SYS
@ECHO -- 64 BIT OR 32 BIT VARIABLE HOLDER

SET FOLDER_NAME64BIT_HKLM=HKLM\SOFTWARE
SET FOLDER_NAME64BIT_HKCU=HKCU\SOFTWARE

@IF "%PROCESSOR_ARCHITECTURE%"=="AMD64" SET FOLDER_NAME64BIT_HKLM=HKLM\SOFTWARE\WOW6432Node
@IF "%PROCESSOR_ARCHITECTURE%"=="AMD64" SET FOLDER_NAME64BIT_HKCU=HKCU\SOFTWARE\WOW6432Node
@ECHO -- PROCESSOR ENVIROMENT VAR IS -- %PROCESSOR_ARCHITECTURE%
@ECHO -- VARIABLE 32 OR 64 VAR SET AS -- FOLDER_NAME64BIT_HKCU -- SET %FOLDER_NAME64BIT_HKCU% -- 
@IF "%PROCESSOR_ARCHITECTURE%"=="AMD64" @ECHO PROCESSOR SET AS USE 64 BIT IS = TRUE -- "AMD64"="%PROCESSOR_ARCHITECTURE% IS EQUAL = TRUE 


@REM @IF "%PROCESSOR_ARCHITECTURE%"=="AMD64" SET FOLDER_NAME64BIT_HKU=HKU\SOFTWARE\WOW6432Node
@REM -- HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders
@REM -- COMMON FAVS IS ONLY ONE NOT A 64 BIT SETTING
@REM -- TEMPORARY FOLDER DOESN'T MOVE AS 32 BIT SETTING
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------


@SET FOLDER_NAME1=01 Desktop Common
@SET FOLDER_NAME2=01 Desktop Common\#_%COMPUTERNAME%
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME1%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME1%\"
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%\"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Desktop" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Common Desktop" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
@SET Desktop_Common_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%

@REM ---------------------------------------------------------------------
@REM ---- Start Menu Common 1 2 3
@REM ---------------------------------------------------------------------
@REM ---------------------------------------------------------------------
@REM -- ORGINAL -- %ProgramData%\Microsoft\Windows\Start Menu\Programs
@REM -- ORGINAL -- C:\ProgramData\Microsoft\Windows\Start Menu\Programs>
@REM ---------------------------------------------------------------------
@SET FOLDER_NAME1=01 Start Menu Common
@SET FOLDER_NAME2=01 Start Menu Common\#_%COMPUTERNAME%

@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME1%" MKDIR "%DRIVEPATH%\%FOLDER_NAME1%"
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%"

reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Start Menu" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Common Start Menu" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 

@SET FOLDER_NAME2=01 Start Menu Common\#_%COMPUTERNAME%\Programs
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Programs" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Common Programs" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 

@SET FOLDER_NAME2=01 Start Menu Common\#_%COMPUTERNAME%\Programs\Startup
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Startup" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Common Startup" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
@SET Start_Menu_Common_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM --- MY FOLDERS COMMON -- MY DOCS -- MY PICTURES -- MY MUSIC -- MY VIDEO 
@REM ----------------------------------------------------------------
@REM -- ORGINAL -- %PUBLIC%\Documents
@REM -- ORGINAL -- %PUBLIC%\
@REM -- ORGINAL -- %PUBLIC%\
@REM -- ORGINAL -- %PUBLIC%\

@SET FOLDER_NAME2=D:\# MY DOCS COMMON
@IF NOT EXIST "%FOLDER_NAME2%\" MKDIR "%FOLDER_NAME2%\"
@SET My_Documents_Common_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%

reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Documents" /t REG_EXPAND_SZ /d "%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Common Documents" /t REG_SZ /d "%FOLDER_NAME2%"

@SET FOLDER_NAME2=D:\# MY DOCS COMMON\My Pictures
@IF NOT EXIST "%FOLDER_NAME2%\" MKDIR "%FOLDER_NAME2%\"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "CommonPictures" /t REG_EXPAND_SZ /d "%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "CommonPictures" /t REG_SZ /d "%FOLDER_NAME2%" 
@SET IMAGE_COMMON_FOLDER=%FOLDER_NAME2%

@SET FOLDER_NAME2=D:\# MY DOCS COMMON\My Music
@IF NOT EXIST "%FOLDER_NAME2%\" MKDIR "%FOLDER_NAME2%\"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "CommonMusic" /t REG_EXPAND_SZ /d "%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "CommonMusic" /t REG_SZ /d "%FOLDER_NAME2%" 
@SET MUSIC_COMMON_FOLDER=%FOLDER_NAME2%


@SET FOLDER_NAME2=D:\# MY DOCS COMMON\My Video
@IF NOT EXIST "%FOLDER_NAME2%\" MKDIR "%FOLDER_NAME2%\"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "CommonVideo" /t REG_EXPAND_SZ /d "%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "CommonVideo" /t REG_SZ /d "%FOLDER_NAME2%" 
@SET VIDEO_COMMON_FOLDER=%FOLDER_NAME2%

@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------

@REM ----------------------------------------
@REM ---- ADMIN TOOLS SET 
@REM ----------------------------------------
@REM ---- ORIGINAL -- C:\Users\MATT 01\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Administrative Tools
@REM ----------------------------------------
@SET FOLDER_NAME1=01 Start Menu\#_%COMPUTERNAME%\Programs
@SET FOLDER_NAME2=01 Start Menu\#_%COMPUTERNAME%\Programs\Administrative Tools
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME1%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME1%\"
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%\"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Administrative Tools" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Administrative Tools" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
@SET ADMIN_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%

GOTO JUMP_NOT_ANY_MORE
@REM ----------------------------------------
@REM ---- COMMON ADMIN
@REM ----------------------------------------
@REM ---- ORGINAL -- C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools
@REM ----------------------------------------
@SET FOLDER_NAME2=01 Start Menu Common\#_%COMPUTERNAME%\Programs\Administrative Tools Common
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME1%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME1%\"
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%\"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Common Administrative Tools" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Administrative Tools" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
@SET ADMIN_COMMON_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%
:JUMP_NOT_ANY_MORE

@REM PROBLEM NOT ABLE TO SEE Administrative Tools IN WIN XP UNLESS DO THIS
IF NOT "%VARLATCH2%" == "WINXP" GOTO JUMP_NOT_WIN_XP
reg.exe DELETE "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Common Administrative Tools"
reg.exe DELETE "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Administrative Tools"
:JUMP_NOT_WIN_XP



@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------
@REM ---- SET TEMPLATES
@REM ----------------------------------------
@SET FOLDER_NAME1=01 Templates
@SET FOLDER_NAME2=01 Templates\--%COMPUTERNAME%
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME1%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME1%\"
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%\"
@SET Templates_FOLDER=%FOLDER_NAME2%
reg.exe ADD "%FOLDER_NAME64BIT_HKCU%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Templates" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKCU%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Templates" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
@SET Templates_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%


@REM ---- ORGINAL -- %ProgramData%\Microsoft\Windows\Templates -- NOT ANY FILES INSIDE
@SET FOLDER_NAME1=01 Templates Common
@SET FOLDER_NAME2=01 Templates Common\--%COMPUTERNAME%
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME1%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME1%\"
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%\"
@SET Templates_Common_FOLDER=%FOLDER_NAME2%
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Templates" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Common Templates" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
@SET Templates_Common_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%

@REM ----------------------------------------
@REM ----------------------------------------
@REM -- SET HISTORY FOLDER
@REM ----------------------------------------

@%DRIVEPATH%
@CD \
@SET FOLDER_NAME1=01 History
@SET FOLDER_NAME2=01 History\--%COMPUTERNAME%--%USERNAME%
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME1%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME1%\"
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%\"
@SET HISTORY_FOLDER=%FOLDER_NAME2%
reg.exe ADD "%FOLDER_NAME64BIT_HKCU%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "History" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKCU%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "History" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
@SET HISTORY_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%
@ECHO.
@ECHO %VARSTRING%


@REM ----------------------------------------
@REM ----------------------------------------
@REM -- SET RECENT FOLDER
@REM -- THIS IS A READ ONLY REGISTRY SETTING ON WINDOWS 10 NIRSOFT SAID
@REM ----------------------------------------
%DRIVEPATH%
CD \
@SET FOLDER_NAME1=01 Recent
@SET FOLDER_NAME2=01 Recent\--%COMPUTERNAME%--%USERNAME%
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME1%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME1%\" 
@IF NOT EXIST "%DRIVEPATH%\%FOLDER_NAME2%\" MKDIR "%DRIVEPATH%\%FOLDER_NAME2%\" 
@SET RECENT_FOLDER=%FOLDER_NAME2%
reg.exe ADD "%FOLDER_NAME64BIT_HKCU%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Recent" /t REG_EXPAND_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKCU%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Recent" /t REG_SZ /d "%DRIVEPATH%\%FOLDER_NAME2%" 
@SET RECENT_FOLDER=%DRIVEPATH%\%FOLDER_NAME2%

@REM ----------------------------------------
@REM ----------------------------------------
@REM -- SET DOWNLOAD FOLDER WINDOWS 10
@REM ----------------------------------------
IF NOT %VARLATCH2% == WIN10 GOTO JUMP_OVER_WIN10
%DRIVEPATH%
CD \
@SET FOLDER_NAME1=C:\DOWNLOADS
@SET DOWNLOAD_FOLDER=%FOLDER_NAME1%
@IF NOT EXIST "%FOLDER_NAME1%\" MKDIR "%FOLDER_NAME1%\" 
reg.exe ADD "%FOLDER_NAME64BIT_HKCU%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "{374DE290-123F-4565-9164-39C4925E467B}" /t REG_EXPAND_SZ /d "%FOLDER_NAME1%"
reg.exe ADD "%FOLDER_NAME64BIT_HKCU%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "{374DE290-123F-4565-9164-39C4925E467B}" /t REG_SZ /d "%FOLDER_NAME1%" 
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "{3D644C9B-1FB8-4f30-9B45-F670235F79C0}" /t REG_EXPAND_SZ /d "%FOLDER_NAME1%"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "{3D644C9B-1FB8-4f30-9B45-F670235F79C0}" /t REG_SZ /d "%FOLDER_NAME1%" 
:JUMP_OVER_WIN10

@REM - ------------------------------------------------------------------------------
@REM - AFTER THAT YOU SEE SETTING FAV SPECIAL FOLDERS REQUIRES A DUPLICATE USE OF 
@REM - REG_EXPAND_SZ AS WELL AS REG_SZ OR REG_MULTI_SZ
@REM - THE REG_MULTI_SZ STORES THE STRING AS A TYPE OF HEXADECIMAL
@REM - THERE ARE 8 TYPES FROM -- CMD> REG /?-- AND REG_EXPAND_SZ AND REG_SZ ARE THERE
@REM - 3 USAGE TYPE'S ARE GIVEN
@REM - ------------------------
@REM -   /t       RegKey data types
@REM -            [ REG_SZ    | REG_MULTI_SZ  | REG_DWORD_BIG_ENDIAN    |
@REM -              REG_DWORD | REG_BINARY    | REG_DWORD_LITTLE_ENDIAN |
@REM -              REG_NONE  | REG_EXPAND_SZ ]
@REM -            If omitted, REG_SZ is assumed
@REM - ------------------------------------------------------------------------------

@REM - ------------------------------------------------------------------------------
@REM - DO ANOTHER VERSION IN VBS SUPPOSE
@REM - COMPATIBLE IN WINDOWS WITHOUT VISUAL BASIC LANGUAGE IDE STUDIO INTEGRATED DEVELOPMENT ENVIRONMENT
@REM - STUDIO AUDIO VIDEO IMAGE
@REM - BETTER AGAIN COMPLETE IT IN THIS BATCH
@REM - ------------------------------------------------------------------------------

@REM - ------------------------------------------------------------------------------
@REM - WELL I DIDN'T WORK OUT HISTORY AND RECENT WERE DIFFERENT - RECENT DOCUMENT AND INTERNET EXPLORER
@REM - SET INTERNET EXPLORER HISTORY
@REM - CAREFUL TO USE A COMPUTER NAME AND USER NAME - WHEN SYNC-ER
@REM - ------------------------------------------------------------------------------

@REM --------------------------------------------------------------------------------
@REM - SPECIAL FOLDERS UPDATE ITS REG ENTRY JUST BY FOLDER MOVE IT - LIKE MANY SPECIAL SYSTEM FOLDERS
@REM --------------------------------------------------------------------------------

@REM -------------------------------------------------------------------------------------
@REM - LAST NOTE OF INFO -- THERE IS A TALK THAT YOUR NOT SUPPOSED TO MAKE COMMON FOLDER SAME AS ANOTHER USED
@REM - AND CHECK NIRSOFT SPECIAL_FOLDER_VIEW SOME ARE READ ONLY BUT OKAY TO CHANGE
@REM ----------------------------------------------------------------

@REM ----------------------------------------
@REM ----------------------------------------
@REM - SET TEMP FOLDER IN THE SPECIAL FOLDERS
@REM ----------------------------------------
C:
CD \
@SET FOLDER_NAME1=C:\TEMP
@SET FOLDER_NAME2=C:\TEMP\-%COMPUTERNAME%
@IF NOT EXIST "%FOLDER_NAME1%\" MKDIR "%FOLDER_NAME1%\"
@IF NOT EXIST "%FOLDER_NAME2%\" MKDIR "%FOLDER_NAME2%\"
@SET TEMP_FOLDER=%FOLDER_NAME2%

@REM -- THIS IS THE TEMP GLOBAL FOLDER FOR ALL WINDOWS USERS SO CAN'T HAVE TEMP USERNAME
@REM -- REBOOT IS REQUIRED TO SET CHANGE

@REM -- NOTE TEST 1 -- THE TEMP FOLDER DON'T LIKE UNDERSCORES IN PATH NAME WITH INSTALL SOME PROGRAMS
@REM ----------------- SEEMS TO STRIP IT OUT LIKE DON'T WANT IT AND FAIL CAN'T FIND INSTALL PATH WANTED
@REM ----------------- REMOVE THE HASH ALSO I THINK
@REM ----------------- CAN CONFIRM DON'T LIKE THE HASH
@REM ----------------- NOW SAME TO DO WITH HISTORY AND RECENT - THINK

@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM -- DETECT 64 BIT WOW SYS -- SYSWOW -- WOW6432Node
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@SET FOLDER_NAME64BIT_HKLM=HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment
@SET FOLDER_NAME64BIT_HKCU=HKCU\Environment
@SET FOLDER_NAME64BIT_HKU=HKU\.DEFAULT\Environment
@REM 
@ECHO TEST WORK TO DO TO CHECK WORK ON 64 BIT AS HAD ERROR
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------

@REM @IF "%PROCESSOR_ARCHITECTURE%"=="AMD64" SET FOLDER_NAME64BIT_HKLM=HKLM\SOFTWARE\WOW6432Node
@REM IF "%PROCESSOR_ARCHITECTURE%"=="AMD64" SET FOLDER_NAME64BIT_HKCU=HKCU\Environment\WOW6432Node
@REM @IF "%PROCESSOR_ARCHITECTURE%"=="AMD64" SET FOLDER_NAME64BIT_HKU=HKCU\Environment\WOW6432Node
@REM -- HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders
@REM -- COMMON FAVS IS ONLY ONE NOT A 64 BIT SETTING
@REM -- TEMPORARY FOLDER DOESN'T MOVE AS 32 BIT SETTING
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------


IF NOT "%VARLATCH2%" == "WINXP" GOTO JUMP_WIN_XP_NOT_DO
reg.exe ADD "%FOLDER_NAME64BIT_HKCU%" /f /v "TEMP" /t REG_MULTI_SZ /d "%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKCU%" /f /v "TMP" /t REG_MULTI_SZ /d "%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKU%" /f /v "TEMP" /t REG_MULTI_SZ /d "%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKU%" /f /v "TMP" /t REG_MULTI_SZ /d "%FOLDER_NAME2%"
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%" /f /v "TEMP" /t REG_MULTI_SZ /d "%FOLDER_NAME2%" 
reg.exe ADD "%FOLDER_NAME64BIT_HKLM%" /f /v "TMP" /t REG_MULTI_SZ /d "%FOLDER_NAME2%"
:JUMP_WIN_XP_NOT_DO

@REM -- FOR WIN 7 UP OR VISTA
SETX TMP %FOLDER_NAME2%
SETX TEMP %FOLDER_NAME2%
@REM ------------------------
@REM THE RESULT IN VAR OF TMP TEMP WILL BE IN SHORT NAME FORMAT
@REM ------------------------

@REM 03 OF 03 
@REM ---------------------
@REM c:/Windows/System32/SystemPropertiesAdvanced.exe
@REM ----
@REM Windows Environment Variables Tutorial
@REM http://xahlee.info/mswin/env_var.html
@REM ----
@REM 02 OF 03 
@REM ----
@REM shell - how to set ENV variable in windows CMD line - Super User
@REM http://superuser.com/questions/212150/how-to-set-env-variable-in-windows-cmd-line
@REM ----
@REM 01 OF 03 
@REM ----
@REM MY TEMP VARIABLE NOT IN COMMAND SET - Google Search
@REM https://www.google.co.uk/search?q=MY+TEMP+VARIABLE+NOT+IN+COMMAND+SET&oq=MY+TEMP@REM +VARIABLE+NOT+IN+COMMAND+SET&aqs=chrome..69i57.8332j0j7&sourceid=chrome&ie=UTF-8
@REM ----



@REM -- ORIGINAL COMMON
@REM -- ---------------
@REM -- Common Start Menu	No	Yes	No	C:\ProgramData\Microsoft\Windows\Start Menu	0x16	CSIDL_COMMON_STARTMENU	%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu	
@REM -- Common Start Menu Programs	No	Yes	No	C:\ProgramData\Microsoft\Windows\Start Menu\Programs	0x17	CSIDL_COMMON_PROGRAMS	%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs	
@REM -- Common Administrative Tools	No	Yes	No	C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools	0x2f	CSIDL_COMMON_ADMINTOOLS	%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Administrative Tools	
@REM -- Common Startup	No	Yes	No	C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup	0x18	CSIDL_COMMON_STARTUP	%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Startup	
@REM -- Common Templates	No	No	No	C:\ProgramData\Microsoft\Windows\Templates	0x2d	CSIDL_COMMON_TEMPLATES	%ALLUSERSPROFILE%\Microsoft\Windows\Templates	
@REM -- Common Desktop	Yes	Yes	No	C:\Users\Public\Desktop	0x19	CSIDL_COMMON_DESKTOPDIRECTORY	%SystemDrive%\Users\Public\Desktop	
@REM -- Common Documents	No	Yes	No	C:\Users\Public\Documents	0x2e	CSIDL_COMMON_DOCUMENTS	%SystemDrive%\Users\Public\Documents	
@REM -- Common Music	No	Yes	No	C:\Users\Public\Music	0x35	CSIDL_COMMON_MUSIC	%SystemDrive%\Users\Public\Music	
@REM -- Common Pictures	No	Yes	No	C:\Users\Public\Pictures	0x36	CSIDL_COMMON_PICTURES	%SystemDrive%\Users\Public\Pictures	
@REM -- Common Video	No	Yes	No	C:\Users\Public\Videos	0x37	CSIDL_COMMON_VIDEO	%SystemDrive%\Users\Public\Videos	

REM ____ SEE END NIRSOFT


@REM -- MIGHT AS WELL SET THIS ANY COMPUTER AT MOMENT MEANT TO BE WIN 10
@REM IF "%COMPUTERNAME%" == "4-ASUS-GL522VW" SET VARLATCH=1
@REM ---- ----
@REM ---- Fixing Windows 10 Automatic Updates Install Problem - AskVG
@REM ---- http://www.askvg.com/fixing-windows-10-automatic-updates-install-problem/
@REM ---- ----
@REM HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU




@ECHO --------------------------------------------------------------------------------
@ECHO --- CODE DONE ------------------------------------------------------------------
@ECHO --------------------------------------------------------------------------------

:VERY_END
@ECHO OFF

@ECHO ####----------------------------------------------------------------------- ####
@ECHO --------------------------------------------------------------------------------
@ECHO REGISTRY CHANGES GOT MADE 
@ECHO AND THESE 15 MAIN FOLDERS MAYBE BEEN CREATED IF NOT EXIST
@ECHO NOT INCLUDING ANY SUB-BRANCH FROM THEM -- LIKE STARTMENU AND 
@ECHO MY DOC -- MUSIC PICTURE VIDEO
@ECHO --------------------------------------------------------------------------------
@ECHO ####----------------------------------------------------------------------- ####


@ECHO ------------ SPECIAL SYSTEM FOLDER SET
@ECHO --------------------------------------
@ECHO 01 OF 24 --- Send To______________=%SendTo_FOLDER%
@ECHO 02 OF 24 --- Desktop______________=%Desktop_FOLDER%
@ECHO 03 OF 24 --- Desktop Common_______=%Desktop_Common_FOLDER%
@ECHO 04 OF 24 --- Favs_________________=%Favs_FOLDER%
@ECHO 05 OF 24 --- Favs Common__________=%Favourite_Common_FOLDER%
@ECHO 06 OF 24 --- Start Menu___________=%Start_Menu_FOLDER%
@ECHO 07 OF 24 --- Start Menu Startup___=%Start_Menu_Startup_FOLDER%
@ECHO 08 OF 24 --- Start Menu Common____=%Start_Menu_Common_FOLDER%
@ECHO 09 OF 24 --- My Documents_________=%My_Documents_FOLDER%
@ECHO 10 OF 24 --- My Documents Common__=%My_Documents_Common_FOLDER%
@ECHO 11 OF 24 --- Templates____________=%Templates_FOLDER%
@ECHO 12 OF 24 --- Templates Common_____=%Templates_Common_FOLDER%
@ECHO 13 OF 24 --- ADMIN________________=%ADMIN_FOLDER%
@ECHO 14 OF 24 --- ADMIN COMMON_________=%ADMIN_COMMON_FOLDER%
@ECHO 15 OF 24 --- HISTORY______________=%HISTORY_FOLDER%
@ECHO 16 OF 24 --- RECENT READ ONLY W10_=%RECENT_FOLDER%
@ECHO 17 OF 24 --- TEMP_________________=%TEMP_FOLDER%
@ECHO 18 OF 24 --- MUSIC________________=%MUSIC_FOLDER%
@ECHO 19 OF 24 --- VIDEO________________=%VIDEO_FOLDER%
@ECHO 20 OF 24 --- IMAGE PHOTO PICTURE__=%IMAGE_FOLDER%
@ECHO 21 OF 24 --- MUSIC COMMON_________=%MUSIC_COMMON_FOLDER%
@ECHO 22 OF 24 --- VIDEO COMMON_________=%VIDEO_COMMON_FOLDER%
@ECHO 23 OF 24 --- IMAGE COMMON_________=%IMAGE_COMMON_FOLDER%
@ECHO 24 OF 24 --- DOWNLOAD_FOLDER _____=%DOWNLOAD_FOLDER%


@REM ECHO 4 OF 4 --- "%DRIVEPATH%\01 WINDOWS XP Activation Backup\%VARDATE%_%COMPUTERNAME%_%USERNAME%"

:END_SUB1

@ECHO OFF
@IF %VARLATCH% == 0 ECHO #### ------------------------------------------------------------------- ####
@IF %VARLATCH% == 0 ECHO #### -- NOTHING WAS DONE                            -------------------- ####
@IF %VARLATCH% == 0 ECHO #### -- NOT A RECOGNISED COMPUTER NAME OR USER NAME -------------------- ####
@IF %VARLATCH% == 0 ECHO #### ------------------------------------------------------------------- ####

@ECHO #### ---------------------------------------------------------------------- ####
@ECHO #### -- READ ME - IN THE BATCH FILE SCRIPT FOR - INFO ON DIY -------------- ####
@ECHO #### -- FINISHED  --------------------------------------------------------- ####
@ECHO #### ---------------------------------------------------------------------- ####

@IF NOT %VARLATCH% == 0 @ECHO #### -- CAREFULLY CHECK SUCCESS RESULTS BEFORE HITT KEY END THIS PROGRAM -- ####
@ECHO #### -- PAUSE PRESS A KEY TO END ------------------------------------------ ####
@ECHO #### ---------------------------------------------------------------------- ####

@PAUSE
@EXIT



@GOTO VERY_END
@GOTO VERY_END
@REM -----------------------------------------------------------------------------
@REM --- CODE DONE ---
@REM -----------------------------------------------------------------------------
@REM --- USE HERE CHECK RESULT 1... 2...
@REM -----------------------------------------------------------------------------
@REM ------- Nirsoft_Package -----------------------------------------------------
@REM --- --- CHECK ALL THE SPECIAL FOLDERING
@REM --- 1... C:\PStart\Progs\#_Nirsoft_Package\NirSoft\specialfoldersview.exe
@REM -----------------------------------------------------------------------------
@REM ------- WWW.System Internals.com --------------------------------------------
@REM --- --- USE THIS REG JUMP TO FROM COPY OF REG KEY ON CLIPBOARD
@REM --- --- VERIFY ALL OKAY PREF BEFORE RUN OR IF POSIABLE ERROR AFTER
@REM --- 2... C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\regjump.exe -c
@REM -----------------------------------------------------------------------------
@REM -----------------------------------------------------------------------------
GOTO VERY_END
GOTO VERY_END
@REM -----------------------------------------------------------------------------
@REM -----------------------------------------------------------------------------
GOTO VERY_END
GOTO VERY_END
@REM -------- ################################# ----------------------------------
@REM -------- NOT SET NOTEPAD++ AT MOMENT 64 BIT RELOCATE SORT TO DO
@REM -------- SET NOTEPAD++ ------------------------------------------------------
@REM -----------------------------------------------------------------------------
@REM @ECHO ---------- SET NOTEPAD++ FOR EDIT FILES
@REM -----------------------------------------------------------------------------
@REM reg.exe ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "C:\Program Files\Notepad++\notepad++.exe %%1"
@REM -----------------------------------------------------------------------------
@REM @ECHO ---------- RESTORE ORGINAL NOTEPAD ------------------------------------
@REM reg.exe ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE %%1"
@REM -----------------------------------------------------------------------------
@REM -------- ################################# ----------------------------------
@REM -----------------------------------------------------------------------------
@REM -----------------------------------------------------------------------------
GOTO VERY_END
GOTO VERY_END



@REM - JUST DOWNLOAD 12GHOST.COM REG JUMP - COMPLEX PROGRAM OF FUNCTIONS YET EASY TO USE
@REM - OR A VBS SCRIPT
@REM - 12GHOST CAN JUMP EVEN WHEN REGEDIT IS OPEN
@REM - BUT NOTICED REGMON CAN'T JUMP TO WHEN OPEN
@REM - PITY NOT TOO EASY BUTTON TO JUMP FROM CLIPBOARD 



@REM BE CAREFUL NOTE HERE - BECAUSE User Shell Folders GOT TO BE - REG_EXPAND_SZ - TO THE HEXIDECIMAL CODE
@REM IT DOESN'T TELL YOU THAT IN REGMON - SO CHECK WITH REGMONS - REGEDIT JUMP
@REM - BTW REGEDIT REGISTRY CALLS REG_SZ FOR REG_MULTI_SZ

@REM - REGMON EXCLUDE SOME ACTIVITY AND 2 PATHS FROM TWEAK XP PRO NARROWS IT DOWN

@REM - CANT SETT QUICK LAUNCH FOLDER IT IS DOWN TO INTERNET EXPLORER
@REM - MIGHT DO

@REM - REGMON IS GOOD LOOK FOR CREATE KEY AND SET VALUE *SET* - BUT DONT SEEM TO CAPTURE UNLESS SUCCESS RESULT ARE ENABLED
@REM - THE SELECT THE LINES AS MOST ARE MULTI AND COPY AND PASTE INTO TEXT DOCUMENT

@REM 84.59259796	Tweak-xp.exe:5920	SetValue	HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Common Programs	SUCCESS	"%DRIVEPATH%\01 Start Menu Common\Programs"	
@REM 84.59375763	Tweak-xp.exe:5920	SetValue	HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Common Programs	SUCCESS	"%DRIVEPATH%\01 Start Menu Common\Programs"	

@REM 204.59127808	Tweak-xp.exe:8772	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\History	SUCCESS	"%DRIVEPATH%\01 History\ASUS-EEE\MATT 02"	
@REM 204.59255981	Tweak-xp.exe:8772	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\History	SUCCESS	"%DRIVEPATH%\01 History\ASUS-EEE\MATT 02"	



@REM - TRIPLE CHECK HERE
@REM - USED REGMON EVEN WITH CRASH - THE FOLDER DON'T HAVE TO EXIST TO SET IN REGISTRY
@REM - WHILE IN REGEDIT TO SEE THE UPDATE MOVE FROM ONE KEY BACK TO THE OTHER
@REM - USE REGMON TO LOCATE THE EXACT KEY JUST IN CASE TYPO
@REM - TO SET THESE SETTING I GUESS YOU HAVE TO BE LOGGED ON AS USER YOU WANT TO MODIFIY
@REM - OR MAYBE YOU CAN USE THE REMOTE NAME 
@REM - REG ADD KeyName [/v ValueName | /ve] [/t Type] [/s Separator] [/d Data] [/f]
@REM - 
@REM -   KeyName  [\\Machine\]FullKey
@REM -            Machine  Name of remote machine - omitting defaults to the current
@REM -                     machine Only HKLM and HKU are available on remote machines
@REM -            FullKey  ROOTKEY\SubKey
@REM -            ROOTKEY  [ HKLM | HKCU | HKCR | HKU | HKCC ]
@REM -            SubKey   The full name of a registry key under the selected ROOTKEY


@REM - 'REG_MULTI_SZ
@REM - I GOT THIS FROM WEB AND PEOPLE USE A SHORT CUT FOR -- REG_MULTI_SZ -- VS -- REG_SZ
@REM - BTW I FOUND - REGEDIT REGISTRY CALLS REG_SZ FOR REG_MULTI_SZ
@REM - UNLESS THEY ARE DIFFERENT - BOTH WORK IN MY EXAMPLE FOR THE TEMP FOLDER


@REM @%DRIVEPATH%\
@REM @CD \

@REM http://stackoverflow.com/questions/3068929/how-to-read-file-contents-into-a-variable-in-a-batch-file
@REM DATE /T >%DRIVEPATH%\TEMP\DATE.TXT"
@REM SET /p VARDATE=<"%DRIVEPATH%\TEMP\DATE.TXT"


@REM http://www.dostips.com/DtTipsStringManipulation.php
@REM --- UK TIME - CHECK
@SET VARDATE=%DATE:~6,4%-%DATE:~3,2%-%DATE:~0,2%


@REM PATH 1
@IF NOT EXIST "%DRIVEPATH%\01 WINDOWS XP Activation Backup\" MKDIR "%DRIVEPATH%\%DRIVEPATH%\01 WINDOWS XP Activation Backup\" 

@REM PATH 1 AND 2
@IF NOT EXIST "%DRIVEPATH%\01 WINDOWS XP Activation Backup\%VARDATE%_%COMPUTERNAME%_%USERNAME%" MKDIR "%DRIVEPATH%\01 WINDOWS XP Activation Backup\%VARDATE%_%COMPUTERNAME%_%USERNAME%"

@COPY %SystemRoot%\system32\wpa.dbl "%DRIVEPATH%\01 WINDOWS XP Activation Backup\%VARDATE%_%COMPUTERNAME%_%USERNAME%\wpa.dbl"
@COPY %SystemRoot%\system32\wpa.bak "%DRIVEPATH%\01 WINDOWS XP Activation Backup\%VARDATE%_%COMPUTERNAME%_%USERNAME%\wpa.bak"

@ECHO DO A DAILY BACKUP OF WINDOWS ACTIVATION FILE
@ECHO DESTINATION -- "%DRIVEPATH%\01 WINDOWS XP Activation Backup\%VARDATE%_%COMPUTERNAME%_%USERNAME%\wpa.dbl"
@ECHO DON'T BOTHER WITH OFFICE XP
@ECHO USE A PROGRAM TO GET SERIALS
@ECHO.
@ECHO %VARSTRING%


@REM TWEAK XP PRO LETS GO
@REM STARTING FROM THE TOP
@REM ---- CACHE OPTIMIZATION ----
@REM SYTEMS WITH RAM OVER 512 MB TWEAK ENABLE TO FORCE THE CORE TO BE KEPT IN MEMORY AND NOT PAGED TO DISK 
@REM #01
@REM - ENABLE CORE SYSTEM PERFORMANCE 
@REM
reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /f /v "DisablePagingExecutive" /t REG_DWORD /d 1 
@ECHO.
@ECHO %VARSTRING%

@REM - DISABLE CORE SYSTEM PERFORMANCE 
@REM reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /f /v "DisablePagingExecutive" /t REG_DWORD /d 0 
@REM 

@REM #02
@REM INCREASE NTFS PEFORMANCE BY DISABLING THE LAST ACCESS TIME STAMP
@REM 
reg.exe ADD "HKLM\System\CurrentControlSet\Control\FileSystem" /f /v "NtfsDisableLastAccessUpdate" /t REG_DWORD /d 1 
@ECHO.
@ECHO %VARSTRING%
@REM - SET REG_DWORD /d 0 -- TO UNDO
@REM -- YOUR NEED THIS AS PEN DRIVES GET A ACCESS OVERLOAD WHEN SET TO USE 
@REM -- LAST ACCESS TIME IN NTFS


@REM #03 --------------
@REM ENABLE BOOT DEFRAG FOR FASTER/QUICKER BOOTING
@REM BOOT TIME DEFRAGGER = YES
reg.exe ADD "HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction" /f /v "Enable" /t REG_SZ /d "Y" 
reg.exe ADD "HKLM\SOFTWARE\Microsoft\Dfrg" /f /v "Enable" /t REG_SZ /d "Y" 
@ECHO.
@ECHO %VARSTRING%
@REM -- HAVE A CONFLICT BECAUSE PIREA(SP) BOOT DEFRAGGLER DOES MY SYSTEM BOOT



@ECHO http://www.optimizingpc.com/optimize/registrytweaks.html

@REM -- Disabling wrapping icon/shortcut labels
@REM -- Probably your desktop contains icons (shortcuts) with long names, which aren't shown completely (the are wrapped......). With a registry tweak you can disable the title wrapping, create a String value IconTitleWrap with value 0 in the following registry key:
@REM -- HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics

@REM 
@REM reg.exe ADD "HKCU\Control Panel\Desktop\WindowMetrics" /f /v "NtfsMftZoneReservation" /t REG_DWORD /d 2 
@ECHO.
@ECHO %VARSTRING%
@REM - SET REG_DWORD /d 0 -- TO UNDO



@REM -- A faster Windows shutdown
@REM -- A QUICKER Windows shutdown
@REM -- Shutting down Windows can take a long time. You can speed this up (at your own risk) with the following tweaks in the registry key:
@REM -- HKEY_CURRENT_USER\Control Panel\Desktop
@REM -- Change the following String values:
@REM -- AutoEndTasks=1
@REM -- HungAppTimeout=4000
@REM -- WaitToKillAppTimeout=4000
@REM -- WaitToKillServiceTimeout=4000
@REM -- (the last registry value is also present in HKLM\SYSTEM\CurrentControlSet\Control)
@REM
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "AutoEndTasks" /t REG_SZ /d "1" 
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "HungAppTimeout" /t REG_SZ /d "180000" 
@REM SOME OF MY PROGRAMS TAKE A WHILE TO CLOSE DOWN LIKE GOODSYNC 
@REM BUT WITHOUT THIS YOU WONT GET PROGRAMS TO SHUT -- SET IT LONGER
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "WaitToKillAppTimeout" /t REG_SZ /d "180000" 
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "WaitToKillServiceTimeout" /t REG_SZ /d "80000" 
reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Control" /f /v "WaitToKillServiceTimeout" /t REG_SZ /d "80000" 
@ECHO.
@ECHO %VARSTRING%
@REM -- WELL THIS DON'T MAKE MUCH DIFFERENCE WITH A DODGY BOOT I GOT CLEAN BOOT WORK WITH AT THE MOMENT
@REM -- AND ALSO SHUTDOWN /F -- FOR FORCE DON'T DO A DODGY STUBBORN BOOT DOWN EITHER
@REM -- WELL UPDATE I FOUND OUT SHUTDOWN -F WONT WORK UNLESS THESE REGISTRY SETTINGS ALSO PRESENT
@REM -- PUT IT IN YOUR TASK-MANAGER TO TRY AT DODGY BOOT TIME

@REM -- NOT SURE IF THIS IS TO SHUTDOWN HUNG APP WHILE IN SYSTEM NOT ONLY BOOT SHUTDOWN
@REM -- SEEMS DODGY IF SO - WHEN PROGRAMS NEED MAKE CRASH REPORT

@REM -- I WOULD LIKE TO MAKE ONE OF MY RUNNING PROGRAM DETECT REQUEST TO BOOT DOWN SO 
@REM -- IT CLOSES ITSELF QUICKER -- PROJECT FOR NEAR FUTURE
@REM -- ON THAT NOTE I DON'T KNOW IF FORCE SHUTDOWN GIVE OPTION FOR PROGRAM TO QUIT GRACEFULLY
@REM -- THINK MAYBE SO THAT WHAT THE TIMEOUT IS FOR
@REM -- SHUTDOWN -F IS LOGGOFF REQUIRE SHUTDOWN -R -F 
@REM -- ALSO WITH THIS REGISTRY DON'T NEED SHUTDOWN -F
@REM -- AND SHUTDOWN -R -F DISPLAY DIALOG TIMECOUNNT DOWN TO SHUTDOWN -- 
@REM -- AND SEEMS TO CLOSE ALL PROGRAMS QUICK WITH ANY GRACE

@REM -- LAST ODD THING IT STILL REPORTS PROGRAM NOT RESPONDING DO YOU WANT TO TERMINATE IT
@REM -- YES UNTIL THE REGISTRY IS SET ALSO
@REM -- YES ALSO THERE IDS A TWEAK FOR SET NON-RESPONDING PROGRAMS TO CLOSE IN TWEAK XP PRO

@REM ----------------------------
@REM -- A faster NTFS file system
@REM -- You probably use the NTFS file system (see the properties of your C: partition). You can improve performance by making some changes in the following registry key:
@REM -- HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Filesystem
@REM -- Change the following registry values, as you wish:
@REM -- DisableNTFSLastAccessUpdate = 1 (last file access is no longer registered).
@REM -- DONE THIS ALREADY
@REM --
@REM -- NtfsDisable8dot3NameCreation = 1 (no more double filenames: long names and the old 8.3 variant (the 8.3 file notation is used for the older Windows versions).
@REM -- WANT THIS ON OUR SYSTEM
@REM --
@REM -- NtfsMftZoneReservation = 2 (if there are many files on a partition, the master file table (MFT) can get fragmented, with this tweak there will be no more fragmenting of the file master file table (default value = 1, range 1-4).
@REM 
reg.exe ADD "HKLM\System\CurrentControlSet\Control\FileSystem" /f /v "NtfsMftZoneReservation" /t REG_DWORD /d 2 
@ECHO.
@ECHO %VARSTRING%
@REM - SET REG_DWORD /d 0 -- TO UNDO

@REM -- BIT OF A LOSER WROTE THIS TO REMOVE 8.3 FILE NAME - SO VERY USEFUL TO HAVE IN COMMAND LINE 
@REM -- 8.3 DOS FILENAME'S = ALTERNATE FILENAME'S



@REM -- How to Remove Shared Documents Icon from My Computer in Windows XP
@REM --  http://www.howtogeek.com/howto/windows/how-to-remove-shared-documents-icon-from-my-computer-in-windows-xp/


@REM -
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /f /v "NoSharedDocuments" /t REG_DWORD /d 1 
@ECHO.
@ECHO %VARSTRING%
@REM -


@REM -- Count = 018 -- Fri 11-Dec-2015 11:06:40
@REM -- ---------------------
@REM -- Registry Monitor - Sysinternals: www.sysinternals.com
@REM -- ---------------------
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchRows	SUCCESS	"3"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchColumns	SUCCESS	"7"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Mouse\MouseHoverTime	SUCCESS	"200"	
@REM --
@REM -- SETTING HERE SPEED OF HOOVER OVER BUT NOT TOO QUICK OF THE SEND TO DIALOGUE APPEARS AND TAKES TIME

@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchRows" /t REG_SZ /d "3" 
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchColumns" /t REG_SZ /d "7" 
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseHoverTime" /t REG_SZ /d "200" 
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- THESE ARE IN A PAIR ON TWEAKUI 
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchColumns	SUCCESS	"7"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Mouse\MouseHoverWidth	SUCCESS	"2"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Mouse\MouseHoverHeight	SUCCESS	"2"	
@REM -- TO SET 1 PIXEL VALUE = 2 OR DOUBLE WHATEVER VALUE
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchColumns" /t REG_SZ /d "7" 
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseHoverWidth" /t REG_SZ /d "2" 
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseHoverHeight" /t REG_SZ /d "200" 
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchRows	SUCCESS	"3"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchColumns	SUCCESS	"7"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\MenuShowDelay	SUCCESS	"117"	
@REM -- I LIKE THIS MOUSE SPEED NIPPY NEAT
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchRows" /t REG_SZ /d "3" 
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchColumns" /t REG_SZ /d "7" 
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "MenuShowDelay" /t REG_SZ /d "100" 
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM 
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchRows	SUCCESS	"3"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchColumns	SUCCESS	"7"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\DragWidth	SUCCESS	"4"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\DragHeight	SUCCESS	"4"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Mouse\DoubleClickWidth	SUCCESS	"6"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Mouse\DoubleClickHeight	SUCCESS	"6"	
@REM -- MOUSE SENSITIVITY -- DRAG IS EQUAL -- DOUBLE CLICK IS DOUBLE
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchRows" /t REG_SZ /d "3" 
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchColumns" /t REG_SZ /d "7" 
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "DragWidth" /t REG_SZ /d "4" 
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "DragHeight" /t REG_SZ /d "4" 
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "DoubleClickWidth" /t REG_SZ /d "6" 
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "DoubleClickHeight" /t REG_SZ /d "6" 
@ECHO.
@ECHO %VARSTRING%
@REM -


@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Sound\Beep	SUCCESS	"No"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\UserPreferencesMask	SUCCESS	9A 7E 03 80 	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\UserPreferencesMask	SUCCESS	9A 5E 03 80 	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\UserPreferencesMask	SUCCESS	92 5E 03 80 	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\UserPreferencesMask	SUCCESS	90 5E 03 80 	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\UserPreferencesMask	SUCCESS	90 5C 03 80 	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\UserPreferencesMask	SUCCESS	90 58 03 80 	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\UserPreferencesMask	SUCCESS	10 58 03 80 	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\UserPreferencesMask	SUCCESS	10 50 03 80 	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\UserPreferencesMask	SUCCESS	10 40 03 80 	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\WindowMetrics\MinAnimate	SUCCESS	"0"	
@REM -- TweakUI.exe:17492	SetValue	HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OptimalLayout\EnableAutoLayout	SUCCESS	0x0	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\PaintDesktopVersion	SUCCESS	0x0	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchRows	SUCCESS	"3"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchColumns	SUCCESS	"7"	

@REM -- RESET EVERYTHING UNCHECKED FIRST AND THEN WHAT WANT CHECKED AFTER
@REM -- I SET EVERYTHING CHECKED BEFORE THIS SO ALL SETTINGS CHANGE

reg.exe ADD "HKCU\Control Panel\Desktop\Sound" /f /v "Beep" /t REG_SZ /d "4" 
@ECHO.
@ECHO %VARSTRING%
@REM -


@REM -- THIS BATCH SEEMS SET SAME SETTING WITH DIFFERENT RESET VALUES AND A QUERY BETWEEN BEFORE
@REM -- MIGHT ONLY NEED THE LAST ONE TO GET IT RIGHT AS I NOTICED LATER -- TEST NOW
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d 9A7E0380
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d 9A5E0380 
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d 925E0380
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d 905E0380
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d 905C0380
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d 90580380
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d 10580380
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d 10500380
@REM -- SET THIS ONE AGAIN LATER 
@REM -- reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d 10400380
reg.exe ADD "HKCU\Control Panel\Desktop\WindowMetrics" /f /v "MinAnimate" /t REG_SZ /d "0"
reg.exe ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OptimalLayout" /f /v "EnableAutoLayout" /t REG_DWORD /d 0
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "PaintDesktopVersion" /t REG_DWORD /d 0
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchRows" /t REG_SZ /d "3" 
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchColumns" /t REG_SZ /d "7" 
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Sound\Beep	SUCCESS	"Yes"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\UserPreferencesMask	SUCCESS	90 40 03 80 	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\PaintDesktopVersion	SUCCESS	0x1	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchRows	SUCCESS	"3"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchColumns	SUCCESS	"7"	

reg.exe ADD "HKCU\Control Panel\Desktop\Sound" /f /v "Beep" /t REG_SZ /d "4" 
@REM -- THIS BATCH SEEMS SET SAME SETTING WITH DIFFERENT RESET VALUES AND A QUERY BEFORE  
@REM -- THIS MUST BE THE ENABLE MOUSE HOT TRACKING EFFECTS FOR TOOL-TIPS AND OTHER THINGS
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d "90400380" 
@REM --
@REM -- I DON'T REQUIRE CURSOR SHADOW ONE
@ECHO.
@ECHO %VARSTRING%
@REM -

reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "PaintDesktopVersion" /t REG_DWORD /d 0
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchRows" /t REG_SZ /d "3" 
@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchColumns" /t REG_SZ /d "7" 
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- FOUND THESE IN --
@REM Windows Registry Editor Version 5.00
@REM [HKEY_CURRENT_USER\Control Panel\Desktop]
@REM -- THE COMPLETE SET SCRIPT ON MY WINDOWS XP PRO
@REM -- DON'T KNOW
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "ActiveWndTrkTimeout" /t REG_DWORD /d 0
@REM -- reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "AutoEndTasks" /t REG_SZ /d "0"
@REM -- GOT THIS ABOVE LINE BEFORE
REM ' SETTINGS - PREFERENCES - EDITING -----
REM ' CARET - CURSOR WIDTH 3
REM  __ REG_DWORD NOT AS MISTAKER REG_SZ 
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CaretWidth" /t REG_DWORD /d 3
@REM -- DON'T KNOW
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitch" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchColumns" /t REG_SZ /d "7"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CoolSwitchRows" /t REG_SZ /d "3"
@REM -- DON'T KNOW
@REM -- reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CTTuneMakeSettingsDefault" /t REG_DWORD /d 0
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CursorBlinkRate" /t REG_SZ /d "10"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "DragFullWindows" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "FontSmoothing" /t REG_SZ /d "2"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "FontSmoothingGamma" /t REG_DWORD /d 1000
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "FontSmoothingOrientation" /t REG_DWORD /d 1
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "FontSmoothingType" /t REG_DWORD /d 2
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "ForegroundFlashCount" /t REG_DWORD /d 3
@REM -- DON'T KNOW
@REM -- HEX VALUE HERE
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "ForegroundLockTimeout" /t REG_DWORD /d 200000
@REM -- DON'T KNOW
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "LowPowerActive" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "LowPowerTimeOut" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "MenuShowDelay" /t REG_SZ /d "100"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "OriginalWallpaper" /t REG_SZ /d ""
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "PaintDesktopVersion" /t REG_DWORD /d 1
@REM -- DON'T KNOW
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "Pattern Upgrade" /t REG_SZ /d "TRUE"
@REM -- DON'T KNOW
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "PowerOffActive" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "PowerOffTimeOut" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "ScreenSaveActive" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "ScreenSaverIsSecure" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "ScreenSaveTimeOut" /t REG_SZ /d "800"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "TileWallpaper" /t REG_SZ /d "1"
@REM -- SET THIS IN BGINFO
@REM -- reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "Wallpaper" /t REG_SZ /d "C:\TEMP\-ASUS-EEE\BGInfo.bmp"

reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "WallpaperStyle" /t REG_SZ /d "0"

reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "WheelScrollLines" /t REG_SZ /d "3"

@ECHO.
@ECHO %VARSTRING%
@REM -

@REM Windows Registry Editor Version 5.00
@REM [HKEY_CURRENT_USER\Control Panel\Mouse]
@REM -- NOT THE COMPLETE SET SCRIPT ON MY WINDOWS XP PRO
@REM -- ONLY ONES I THINK NOT SET BEFORE OR DEFAULT AS NEW USER OR OPERATING SYSTEM

@REM THESE 2 SET THE HIDE POINTER WHILE TYPE AND
@REM SHOW LOCATION OF POINTER WHEN CTRL KEY
@REM ---
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d "90400280" 
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d "90400380" 
@REM ---

reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseTrails" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseThreshold1" /t REG_SZ /d "6"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseThreshold2" /t REG_SZ /d "10"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseSpeed" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseSensitivity" /t REG_SZ /d "4"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "SnapToDefaultButton" /t REG_SZ /d "1"

@REM -- MOUSE WHEEL - HAMPSTER
@REM --
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "WheelScrollLines" /t REG_SZ /d "3"

reg.exe ADD "HKCU\SOFTWARE\Microsoft\IntelliPoint\Scrolling" /f /v "ScrollLines" /t REG_DWORD /d 3
reg.exe ADD "HKCU\SOFTWARE\Microsoft\IntelliPoint\Scrolling" /f /v "AccelerationCurve" /t REG_DWORD /d 6

reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "UserPreferencesMask" /t REG_BINARY /d "90000380" 


reg.exe ADD "HKCU\Control Panel\Cursors\(Default)" /f /v "(Default)" /t REG_SZ /d "Windows Default"
reg.exe ADD "HKCU\Control Panel\Cursors\(Default)" /f /v "Scheme Source" /t REG_DWORD /d 2

reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "Arrow" /t REG_SZ /d "%SYSTEMROOT%\cursors\arrow_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "Help" /t REG_SZ /d "%SYSTEMROOT%\cursors\help_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "AppStarting" /t REG_SZ /d "%SYSTEMROOT%\cursors\wait_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "Wait" /t REG_SZ /d "%SYSTEMROOT%\cursors\busy_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "Crosshair" /t REG_SZ /d "%SYSTEMROOT%\cursors\cross_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "IBeam" /t REG_SZ /d "%SYSTEMROOT%\cursors\beam_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "NWPen" /t REG_SZ /d "%SYSTEMROOT%\cursors\pen_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "No" /t REG_SZ /d "%SYSTEMROOT%\cursors\no_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "SizeNS" /t REG_SZ /d "%SYSTEMROOT%\cursors\size4_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "SizeWE" /t REG_SZ /d "%SYSTEMROOT%\cursors\size3_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "SizeNWSE" /t REG_SZ /d "%SYSTEMROOT%\cursors\size2_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "SizeNESW" /t REG_SZ /d "%SYSTEMROOT%\cursors\size1_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "SizeAll" /t REG_SZ /d "%SYSTEMROOT%\cursors\move_il.cur"
reg.exe ADD "HKCU\Control Panel\Cursors" /f /v "UpArrow" /t REG_SZ /d "%SYSTEMROOT%\cursors\up_il.cur"
	





@REM -- DOUBLE CLICK SPEED
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "DoubleClickSpeed" /t REG_SZ /d "1050"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "DoubleClickHeight" /t REG_SZ /d "5"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "DoubleClickWidth" /t REG_SZ /d "5"



@REM ----

@ECHO.
@ECHO %VARSTRING%
@REM -

@REM Windows Registry Editor Version 5.00
@REM [HKEY_CURRENT_USER\Control Panel\Keyboard]
@REM -- COMPLETE SET SCRIPT ON MY WINDOWS XP PRO
@REM -- DON'T KNOW
@REM reg.exe ADD "HKCU\Control Panel\Keyboard" /f /v "InitialKeyboardIndicators" /t REG_SZ /d "0"
@REM -- LESS-EST DELAY
reg.exe ADD "HKCU\Control Panel\Keyboard" /f /v "KeyboardDelay" /t REG_SZ /d "0"

@REM -- THINK THIS SET QUICKER THAN TWEAKUI ALLOW
@REM -- YES SET IN CONTROL PANEL QUICKEST 31
reg.exe ADD "HKCU\Control Panel\Keyboard" /f /v "KeyboardSpeed" /t REG_SZ /d "31"

@REM -- QUICKEST
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CursorBlinkRate" /t REG_SZ /d "10"

TIMEOUT /T 20


@REM -- CAP LOCK WANTING
@REM -- ----
@REM -- Turn NUM LOCK on at start-up - Microsoft Community
@REM -- https://answers.microsoft.com/en-us/windows/forum/windows_7-hardware/turn-num-lock-on-at-start-up/b6f5c70d-a053-4311-946d-ba9e8879f644?page=4
@REM -- ----
@REM -- CAP LOCK = 1
reg.exe ADD "HKCU\Control Panel\Keyboard" /f /v "InitialKeyboardIndicators" /t REG_SZ /d "1"




@ECHO.
@ECHO %VARSTRING%
@REM -

@REM Windows Registry Editor Version 5.00
@REM [HKEY_CURRENT_USER\Control Panel\PowerCfg]
@REM -- COMPLETE SET SCRIPT ON MY WINDOWS XP PRO -- SUB KEYS EXIST
reg.exe ADD "HKCU\Control Panel\PowerCfg" /f /v "CurrentPowerPolicy" /t REG_SZ /d "3"
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM Windows Registry Editor Version 5.00
@REM [HKEY_CURRENT_USER\Control Panel\Sound]
@REM -- COMPLETE SET SCRIPT ON MY WINDOWS XP PRO -- SUB KEYS EXIST
reg.exe ADD "HKCU\Control Panel\Sound" /f /v "Beep" /t REG_SZ /d "Yes"
reg.exe ADD "HKCU\Control Panel\Sound" /f /v "ExtendedSounds" /t REG_SZ /d "Yes"
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- MAY WANT TO SET THIS ONE
@REM -- HKEY_CURRENT_USER\Control Panel\Accessibility
reg.exe ADD "HKCU\Control Panel\Accessibility\Blind Access" /f /v "On" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\Accessibility\Keyboard Preference" /f /v "On" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\Accessibility\ShowSounds" /f /v "On" /t REG_SZ /d "0"
@ECHO.
@ECHO %VARSTRING%
@REM -


@REM -- HKEY_CURRENT_USER\Console
@REM Windows Registry Editor Version 5.00
@REM [HKEY_CURRENT_USER\Console]
@REM -- NOT THE COMPLETE SET SCRIPT ON MY WINDOWS XP PRO
@REM -- ONLY ONES I THINK NOT SET BEFORE OR DEFAULT AS NEW USER OR OPERATING SYSTEM

@REM -- HOW TO PUT %SYSTEMROOT% IN COMMAND BATCH FILE - Google Search
@REM -- https://www.google.co.uk/search?num=100&q=HOW+TO+PUT+%25SYSTEMROOT%25+IN+COMMAND+BATCH+FILE&spell=1&sa=X&ved=0ahUKEwj-joaoltTJAhXIxxQKHeN4CjYQvwUIGigA&biw=991&bih=475
@REM -- "%systemroot%" chages to "C:\windows" - DosTips.com
@REM -- http://www.dostips.com/forum/viewtopic.php?f=3&t=2151
@REM -- Post subject: Re: "%systemroot%" chages to "C:\windows"PostPosted: 21 Aug 2011 19:19 
@REM -- Joined: 12 Feb 2011 21:02	
@REM -- Posts: 1568	
@REM -- Location: United States (east coast)	
@REM -- from the command line:
@REM -- echo ^%systemroot^%
@REM -- from within a batch file:
@REM -- echo %%systemroot%%

@REM -- PROBLEM ALL DWORD ARE DEFAULT DECIMAL NOT HEX

@REM -- MEDIUM CURSOR SIZE

REM CARET - CURSOR WIDTH 3
REM FOR THE COMMAND CONSOLE NOT WINDOW_S _ MAYBE
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "CursorSize" /t REG_DWORD /d 50
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "HistoryBufferSize" /t REG_DWORD /d 90
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "InsertMode" /t REG_DWORD /d 1
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "NumberOfHistoryBuffers" /t REG_DWORD /d 4


REM :: 655294544 equals 9999 lines in the GUI
reg.exe add "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /v ScreenBufferSize /t REG_DWORD /d 655294544 /f

REM :: We also need to change the Window Height, 3276880 = 50 lines
@REM reg.exe add "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /v WindowSize /t REG_DWORD /d 3276880 /f
@REM -- IF YOU HAVE DEFAULT 25 LINES WONT BE A REGISTRY KEY GET DELETED

@REM -- HKEY_CURRENT_USER\Console\%SystemRoot%_system32_cmd.exe

reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "FaceName" /t REG_SZ /d "Lucida Console"
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "FontFamily" /t REG_DWORD /d 54
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "FontSize" /t REG_DWORD /d 1048576
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "FontWeight" /t REG_DWORD /d 700
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "HistoryNoDup" /t REG_DWORD /d 0
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "QuickEdit" /t REG_DWORD /d 2048
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- AD086944503 ARGOS ORDER NUMBER CHANGE BUT HAS CHARS INIT NOT ONLY NUMBERS AND OLD ORDER NUMBER STILL WORKS

@REM -- Windows Registry Editor Version 5.00
@REM -- [HKEY_CURRENT_USER\Control Panel\International]
@REM -- WHOLE SET AND ONE KEY IN BRANCH FOR MY UNITED KINGDOM 
reg.exe ADD "HKCU\Control Panel\International" /f /v "iCountry" /t REG_SZ /d "44"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iCurrDigits" /t REG_SZ /d "2"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iCurrency" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iDate" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iDigits" /t REG_SZ /d "2"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iLZero" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iMeasure" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iNegCurr" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iTime" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iTLZero" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\International" /f /v "Locale" /t REG_SZ /d "00000809"
reg.exe ADD "HKCU\Control Panel\International" /f /v "s1159" /t REG_SZ /d "AM"
reg.exe ADD "HKCU\Control Panel\International" /f /v "s2359" /t REG_SZ /d "PM"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sCountry" /t REG_SZ /d "United Kingdom"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sCurrency" /t REG_SZ /d "£"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sDate" /t REG_SZ /d "/"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sDecimal" /t REG_SZ /d "."
reg.exe ADD "HKCU\Control Panel\International" /f /v "sLanguage" /t REG_SZ /d "ENG"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sList" /t REG_SZ /d ","
reg.exe ADD "HKCU\Control Panel\International" /f /v "sLongDate" /t REG_SZ /d "dd MMMM yyyy"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sShortDate" /t REG_SZ /d "dd/MM/yyyy"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sThousand" /t REG_SZ /d ","
reg.exe ADD "HKCU\Control Panel\International" /f /v "sTime" /t REG_SZ /d ":"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sTimeFormat" /t REG_SZ /d "HH:mm:ss"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iTimePrefix" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sMonDecimalSep" /t REG_SZ /d "."
reg.exe ADD "HKCU\Control Panel\International" /f /v "sMonThousandSep" /t REG_SZ /d ","
reg.exe ADD "HKCU\Control Panel\International" /f /v "iNegNumber" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sNativeDigits" /t REG_SZ /d "0123456789"
reg.exe ADD "HKCU\Control Panel\International" /f /v "NumShape" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iCalendarType" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iFirstDayOfWeek" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\International" /f /v "iFirstWeekOfYear" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sGrouping" /t REG_SZ /d "3;0"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sMonGrouping" /t REG_SZ /d "3;0"
reg.exe ADD "HKCU\Control Panel\International" /f /v "sPositiveSign" /t REG_SZ /d ""
reg.exe ADD "HKCU\Control Panel\International" /f /v "sNegativeSign" /t REG_SZ /d "-"
reg.exe ADD "HKEY_CURRENT_USER\Control Panel\International\Geo" /f /v "Nation" /t REG_SZ /d "242"
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- THINK THAT IS ALL I WANT FOR CONTROL PANEL 

@REM -- BACK TO TWEAKUI AND 
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchRows	SUCCESS	"3"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Control Panel\Desktop\CoolSwitchColumns	SUCCESS	"7"	
@REM -- TweakUI.exe:17492	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ThumbnailQuality	SUCCESS	0x5F	
@REM -- TweakUI.exe:17492	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ThumbnailSize	SUCCESS	0x64	

@REM -- SIZE 64 = 100 PIXELS OF HEX YES
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /f /v "ThumbnailQuality" /t REG_DWORD /d 95
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /f /v "ThumbnailSize" /t REG_DWORD /d 100
@ECHO.
@ECHO %VARSTRING%
@REM -


@REM -- WELL BEEN SETTING ALL THESE TODAY 11-12-2015 FOUND NEW ICON WITH REGMON TO JUMP TO REGISTRY
@REM -- MAKES EASIER -- AND REGMON NOT CRASH ON MY LOW END ASUS EEE 
@REM -- EXPLORER CUSTOMIZE
@REM -- SET FOLDERS TO REMEMBER 
@REM -- TweakUI.exe:17492	SetValue	HKCU\SOFTWARE\Microsoft\Windows\ShellNoRoam\Bag MRU Size	SUCCESS	0x1F4	
@REM -- TweakUI.exe:17492	SetValue	HKCU\SOFTWARE\Microsoft\Windows\Shell\Bag MRU Size	SUCCESS	0x1F4	
@REM -- SEEMS TO BE A FAULT WITH ANOTHER ALMOST SAME SETTING HERE -- BagMRU Size
@REM -- HEX 1F4 = 500
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\ShellNoRoam" /f /v "Bag MRU Size" /t REG_DWORD /d 500
@REM -- SEEMS TO BE A FAULT WITH ANOTHER ALMOST SAME SETTING HERE -- BagMRU Size
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\Shell" /f /v "Bag MRU Size" /t REG_DWORD /d 500
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- EXPLORER CUSTOMIZE
@REM -- SLIDE SHOW DELAY
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellImageView" /f /v "Timeout" /t REG_DWORD /d 1
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- EXPLORER -- NEED SOME WORK - FEW CHECK-BOXES
@REM -- BIT HARD WORK THIS ONE -- DO LATER LAST ONE TO DO AS WELL
@rem ------------------------------------------------------------

@REM -- COMMON DIALOGUES
@REM -- TweakUI.exe:14396	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoComplete\AutoComplete In File Dialog	SUCCESS	0x1	
@REM -- TweakUI.exe:14396	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\comdlg32\NoFileMru	SUCCESS	0x0	
@REM -- TweakUI.exe:14396	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\comdlg32\NoBackButton	SUCCESS	0x0	
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\AutoComplete" /f /v "AutoComplete" /t REG_DWORD /d 1
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\comdlg32" /f /v "NoFileMru" /t REG_DWORD /d 0
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\comdlg32" /f /v "NoBackButton" /t REG_DWORD /d 0
@ECHO.
@ECHO %VARSTRING%
@REM -


@REM -- TASKBAR AND STARTMENU
@REM -- TweakUI.exe:14396	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoUserNameInStartMenu	SUCCESS	01 00 00 00 	
@REM -- TweakUI.exe:14396	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoTrayItemsDisplay	SUCCESS	01 00 00 00 	
@REM -- TweakUI.exe:14396	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoLowDiskSpaceChecks	SUCCESS	0x1	
@REM -- TweakUI.exe:14396	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\EnableBalloonTips	SUCCESS	0x1	
@REM -- TweakUI.exe:14396	SetValue	

@REM SHOW CHECKED ENABLED = 1
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /f /v "NoUserNameInStartMenu" /t REG_BINARY /d "00000000"
@REM SHOW CHECKED ENABLED = 1
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /f /v "NoTrayItemsDisplay" /t REG_BINARY /d "00000000"

@REM WARN WHEN LOW = CHECKED ENABLED = 0
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /f /v "NoLowDiskSpaceChecks" /t REG_DWORD /d 0

reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /f /v "EnableBalloonTips" /t REG_DWORD /d 1
@ECHO.
@ECHO %VARSTRING%
@REM -


@REM -- TASKBAR AND STARTMENU - GROUPING
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /f /v "TaskbarGroupSize" /t REG_DWORD /d 0
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- DESKTOP ICONS -- NEED SOME WORK - FEW CHECK-BOXES
@REM -- CAN'T BE DONE ALL IN SERIAL NUMBERS
@REM -- USE TWEAK UI


@REM -- DESKTOP FIRST ICON -- NEED SOME WORK - FEW CHECK-BOXES
@REM -- WONT SET THIS ONE PROPER -- USES SERIAL NUMBER
@REM -- HKEY_CLASSES_ROOT\CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}\SortOrderIndex = DWORD 54 FOR MY COMPUTER FIRST -- MY COMPUTER OR MY DOCUMENTS 54 OR 48 HEX

reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /f /v "NoDriveTypeAutoRun" /t REG_DWORD /d 95
@REM -- SET TWICE BUT THINK END RESULT IS 95 FOR CD AUTOPLAY BUT NOT REMOVABLE DRIVES

reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\UnreadMail" /f /v "MessageExpiryDays" /t REG_DWORD /d 0
@ECHO.
@ECHO %VARSTRING%
@REM -
@REM -- 
@REM -------------------------
@REM -- TWEAKUI DONE FOR ME -- 
@REM -------------------------

@REM - QUICK FIX FOR OFFICE TOOLBAR MOVING RELOCATING
@REM -------------------------------------

reg.exe ADD "HKCU\SOFTWARE\Microsoft\Office\9.0\Shortcut Bar\Toolbars\1 OTB" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\1 OTB"
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Office\9.0\Shortcut Bar\Toolbars\2 OTB" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\2 OTB"
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Office\9.0\Shortcut Bar\Toolbars\3 OTB" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\3 OTB"
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Office\9.0\Shortcut Bar\Toolbars\4 OTB" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\4 OTB"
reg.exe ADD "HKCU\SOFTWARE\Microsoft\Office\9.0\Shortcut Bar\Toolbars\5 OTB" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\5 OTB"

@REM -- 7.88941622	Tweak-xp.exe:19696	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\TransparentIcons	SUCCESS	""C:\Program Files\Tweak-XP Pro\tranicon.exe" -ex"	

@REM -- THIS IS GOOD BUT IF YOU GOT DESKTOP COLOR ALSO YOU NEED TO SET IN DESKTOP AND APPEARANCE DESKTOP
@REM TWEAK XP USES ITS OWN PROGRAM TO DO IT



@REM -----------------------------------------------------------
@REM -----------------------------------------------------------
@ECHO OFF
@ECHO.

@REM DISABLE - DELETE - NOT TESTED YET - PUT /f BACK BEFORE /v FOR FORCE DELETE WHEN DONE
@REM reg.exe DELETE "HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction" /v "Enable" 
@REM reg.exe DELETE "HKLM\SOFTWARE\Microsoft\Dfrg" /v "Enable" 

@REM -------------

REM - #04 YOU MIGHT WANT 
REM - OPTIMIZE WINDOWS XP PREFETCHER SETTINGS - MINE IS ENABLED

@REM -------------

REM - THAT'S THE FIRST PAGE DONE FROM TWEAK XP PRO
REM - GOT REGMON ON QUICK LAUNCH NOW IT DONT STAY UP LONG - AT LEAST IT KEEPS SETTINGS

REM -- -------------------------------------------------------------------------------
REM -- -------------------------------------------------------------------------------

REM RISKY ONE DELETE - CAN DEL ALL KEY IN REGISTRY BY FORCE 
REM ----- GOING TO HAVE A REST NOW - REST-ART - REBOOT 
REM ----- ALL DAYER
REM ----- 2:14AM
REM - REBOOT TIME
REM IMPLIMENT SETTINGS 




:VERY_END
GOTO END_SUB2

@ECHO OFF
 
@ECHO ####--------------------------------------------------------------------####
@ECHO REGISTRY CHANGES GOT MADE AND 4 DIRECTORYS MADE IF THEY DIDN'T EXIST ALREADY
@ECHO ####--------------------------------------------------------------------####

@REM - WHAT EVER THE SYS OPP PATH IS FOR TEMP
@ECHO 1 OF 3 --- "C:\TEMP\-%COMPUTERNAME%"
@ECHO 2 OF 3 --- "%DRIVEPATH%\01 Recent\#_%COMPUTERNAME%_%USERNAME%"
@ECHO 3 OF 3 --- "%DRIVEPATH%\01 History\#_%COMPUTERNAME%_%USERNAME%"

@REM ECHO 4 OF 4 --- "%DRIVEPATH%\01 WINDOWS XP Activation Backup\%VARDATE%_%COMPUTERNAME%_%USERNAME%"

:END_SUB2

@ECHO OFF
@IF %VARLATCH% == 0 ECHO #### ------------------------------------------------------------------- ####
@IF %VARLATCH% == 0 ECHO #### -- NOTHING WAS DONE                            -------------------- ####
@IF %VARLATCH% == 0 ECHO #### -- NOT A RECOGNISED COMPUTER NAME OR USER NAME -------------------- ####
@ECHO #### ------------------------------------------------------------------- ####
@ECHO #### -- READ ME - IN THE BATCH FILE SCRIPT FOR - INFO ON DIY ----------- ####
@ECHO #### -- FINISHED                                             ----------- ####

@IF NOT %VARLATCH% == 0 @ECHO #### -- CAREFULLY CHECK SUCCESS RESULTS BEFORE HITT KEY END THIS PROGRAM ####
@ECHO #### -- PAUSE PRESS A KEY TO END --------------------------------------- ####
@ECHO #### ------------------------------------------------------------------- ####

@PAUSE
@EXIT



GOTO VERY_ENDER_BLOCK


REM '----------------------------
REM '-Summary
REM '----------------------------
REM '-Full file path: %DRIVEPATH%\REG KEY SETTINGS\ASUS EEE NEW USERS SETTING\SPECIAL FOLDERS.BAT
REM '-Created: 2015-08-23 01:23:40
REM '-Modified: 2015-11-06 00:59:00
REM '-File length (in byte): 27672
REM '
REM '-Characters (without blanks): 28493
REM '-Words: 4187
REM '-Lines: 701
REM '-Current document length: 29893
REM '-0 selected characters (0 bytes) in 0 ranges 

---------------------------
Summary
---------------------------
Full file path: %DRIVEPATH%\REG KEY SETTINGS\ASUS EEE NEW USERS SETTING\SPECIAL FOLDERS.BAT
Created: 2015-08-23 01:23:40
Modified: 2015-11-06 03:25:00
File length (in byte): 30857

Characters (without blanks): 31013
Words: 4499
Lines: 782
Current document length: 32575
0 selected characters (0 bytes) in 0 ranges

---------------------------
OK   
---------------------------






4.07501602	Tweak-xp.exe:5920	SetValue	HKCU\Environment\TMP	SUCCESS	"C:\TEMP\#_TEMP_MATT_02"	
4.07636213	Tweak-xp.exe:5920	SetValue	HKU\.DEFAULT\Environment\TMP	SUCCESS	"C:\TEMP\#_TEMP_MATT_02"	
4.07708693	Tweak-xp.exe:5920	SetValue	HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\TEMP	SUCCESS	"C:\TEMP\#_TEMP_MATT_02"	
4.07781458	Tweak-xp.exe:5920	SetValue	HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\TMP	SUCCESS	"C:\TEMP\#_TEMP_MATT_02"	




-----------------------------------
README ----------------------------
-----------------------------------

' DONT MATTER IF YOU GOT COPY AND PASTE DOUBLE SLASHES FROM REG EXPORTS - CONVERTS TO SINGLE

' DONT FORGET THE THEME SETTING ONLY TAKE A SECOND TO BACKUP AND RESTORE THE VITAL ONES

' OR OTHER WAY
' USE DEFAULT SETTINGS IF YOU WANT CHANGE FROM YOUR OLD ONE

' ##########
' DESKTOP BACKGROUND NONE
' DESKTOP COLOUR DARK BLUE
' POWER ALWAYS ON
' NOT A PASSWORD AFTER WELCOME SCREEN
' CLOSE THE LID NOTHING
' ###
' DESKTOP PROPERTIES - RIGHT CLICK DESKTOP PROPERTIES
' CUSTOMIZE DESKTOP
' DESKTOP ICONS - ALL SHOW
' UNCHECK DESKTOP CLEANUP

'########## -----
' NOTEPAD++ SETTING VAN BE OR ARE GLOBAL IF YOU SETUP NEW USER - IF RECENTLY
' WITH REGARD TO EDIT OF BATCH FILES BEING MADE BY NOTEPAD++ 
' SO DON'T REALY REQUIRE REG SETTING - BUT NOT HARM TO DO AGAIN
'
' MIGHT HAVE TO SET PLUGINS SPELLING TO GLOBAL FOR NEXT TIME
' SPELL SIMPLE --------
' DICTIONARY PATH APPLY TO ALL USERS - BEST IF NOT TO
'
' SPELL ADVANCED IGNOR WORDS ###
' UNCHECK HAVING NOR FIRST CAPITAL LETTER
' UNCHECK ALL CAPITAL LETTERS
' UNCHECK HAVING _
' UNCHECK START OR END WITH '
'###
' SETTINGS - PREFERANCES - GENERAL -----
' MULTI LINE
' DARKEN INACTIVE TABS - UNCHECK
' DOUBLE CLICK TO CLOSE
'###
' SETTINGS - PREFERANCES - EDITING -----
' CARET - CURSOR WIDTH 3
' BLINK RATE SPEED UP
' BORDER WIDTH MINIMUM NAUGHT
' LINE WRAP - MAYBE FROM ALIGNED
' & LOOK AT MENU - VIEW - WORD WRAP VBCHECKED
' REMEMBER NOTEPAD++ MENU - VIEW HAS GOOD SUMMERY OF TEXT INCLUDE WORD COUNT
'### -----
' FILE ASSOCIATION - MIGHT NEED SOMETHING
'### -----
' BACKUP ON SAVE - VERBOSE SET CUSTOM FOLDER
' MY FOLDER IS D:\NOTEPAD++ BACKUP
'### -----
' AUTO COMPLETION
' BOTH SET OFF AND AUTO INSTERT
' #############################

'### -----
' TASKBAR AND START MENU PROPERTIES - CUSTOMMIZE START MENU - RIGHT CLICK START BUTTON - PROPERTIES
' GENERAL TAB ----------------
' NUMBER OF PROGRAMS ON STARTMENU - LOW END ASUS EEE AND ACER ONE MAX IS 7 
' ADVANCED TAB -- MINE = -----
' TOP CHECK BOX OF TWO HIGHLIGHT NEWLY INSTALLED PROGRAMS
' SET NETWORK PLACES VBCHECKED
' SET ADMINS TOOLS SHOW ALL PLACES
' SET PICTURES SHOW
' CLOSE CUSTOMIZE START MENU -----
' ### TASK BAR TAB ----  
' SHOW QUICK LAUNCH
' UNCHECK HIDE INACTIVE ITEMS
' TASK BAR WHERE PROGRAMS ARE - RIGHT CLICK
' LOCK TASKBAR VBUNCHECKED
' HAVE A LOOK AT QUICK LINKS

'### -----
' MY COMPUTER RIGHT CLICK - PROPERTIES
' ADVANCED - VISUAL EFFECTS - SETTINGS
' AJUST FOR BEST PERFORMACE = CUSTOM
' SHOW WINDOW CONTENTS WHILE DRAGGING
' SMOOTH EDGES OF SCREEN FONTS
' SMOOTH-SCROLL LIST BOXES
' USE VISUAL STYLES ON WINDOWS AND BUTTONS
' THAT SETS XP THEME BLUE AND RED
' REMBER YOU CAN FURTHER DISABLE XP THEME IN THEME SERVICE
'
' VIRTUAL MEMORY ----- 
' CHANGE - GLOBAL SETTING I THINK
' SYSTEM MANAGED PAGING FILE SIZE AND ONLY ON C DRIVE
'
' CHECK ENVIROMENT VARS - ARE VERIFYED TO CHANGE AFTER REG SETTING OF TEMP FOLDER
' SOME PROGRAMS STILL USE TEMP FOLDER SEEMS SET BY ENVIOMENT VAR
' -----
' ADVANCED - STARTUP AND RECOVERY
' UNCHECK AUTO RESTART - I LIKE THE LOOK OF BLUE SCREEN OF DEATH
' GET A WARNING - IF PAGE FILE SIZE IS NAUGHT - CANNOT CREATE A SYSTEM ERROR LOGG AT STOP
' THINK THIS GLOBAL
'### -----

'### -----
' MY COMPUTER RIGHT CLICK - MANAGE OPTION
' SERVICES ARE GLOBAL
' TURN OF INDEXING
' TURN OFF *HELP AND SUPPORT* IN CASE F1 KEY PRESSED
' -----
' MY COMPUTER RIGHT CLICK - PROPERTIES OPTION -----
' SYSTEM PROPERTIES -----
' SYSTEM RESTORE TAB - TURN OFF
'
' ADVANCED TAB -----
' PERFORMANCE TOP BOX - SETTINGS - ADVANCED TAB - CHANGED - VIRTUAL MEMORY -----

' REMOTE TAB - REMOTE ASSISTANCE 
' VBUNCHECK ALLOW REMOTE ASSISTANCE - MIGHT SAVE SERVICE BEING OPEN RUNNING
' REMOTE USER DESKTOP TOO DIFFERCULT
' DEFAULT ENABLED

' ADVANCED TAB ----- 
' BOTTOM RIGHT -----
' ERROR REPORTING - DEFAULT DISABLED - BUT ENABLE IF CRITICAL ERROR

'### -----
' RECYCLE BIN RIGHT CLICK PROPERTIES - ICON ON THE DESKTOP
' CHECK IT OUT

'### -----
' NORTON IS GLOBAL - HARAY FOR THAT
' DISABLE IT IN CHROME
' ABLE IDLE

' EXPLORER - VIEW - DETAILS
' VIEW - STATUS BAR ENABLED
' AJUST THE COLUMNS - ADD COLUMN DATE CREATED 
' TOOLS - FOLDER OPTIONS - VIEW
' APPLY TO ALL FOLDERS
'
' FOLDER OPTIONS - GENERAL -----
' USE CLASSIC WINDOWS FOLDERS
' FOLDER OPTIONS - VIEW --------
' DISPLAY CONTENTS OF SYSTEM FOLDERS
' DISPLAY FULL PATH IN ADDRESS BAR
' DISPLAY FULL PATH IN TITLE BAR
' SHOW HIDDEN FILES AND FOLDERS
' UNCHECK HIDE EXTENTIONS
' UNCHECK HIDE PROTECTED OPERATING SYSTEM FILES
' SHOW CONTROL PANEL
' KEEP USE SIMPLE SHARING RECOMMENED
' MENU BAR ICONS - RIGHT CLICK BELOW FILE MENU - CUSTOMIZE
' PUT CUT COPY PASTE - DEL - UNDO:ON - REMOVE - MOVE TO - COPY TO

' MOUSE - CONTROL PANEL - MOUSE
' SPEED IT UP A BIT - NIB BIP BUT TAD NAD NAB NOB
' AUTOMATIC MOVE POINTER TO THE DEFAULT BUTTON IN A DIALOG BOX
' SHOW LOCATION WHEN PRESS THE CTRL KEY
' DEVICE SETTINGS - ENABLE TOUCH PAD - MAYBE AND AGAIN AFTER LOGG-OFF

'### -----
' FIND - C:\WINDOWS\system32\TweakUI.exe
' SYSINTERNALS POWERTOY

' GENERAL
' DISABLE OPTIMIZE HARD DRIVE WHEN IDLE
' OPTIMIZE - SODIMIZE
' ENABLE SHOW WINDOWS VERSION ON DESKTOP

' MOUSE BRING HOOVER TIME UP SO ACCIDENTLY LAND ON SEND TO WAIT PROBLEM
' HOOVER SENSITIVITY ON 1 
' YOU WANT THE WHOLE AREA HOOVERING OVER TO COUNNT

' EXPLORER - 

' UNCHECK ALLOW HELP ON START MENU
' BUT SERVICE WE DISABLED ALSO

' UNCHECK ALLOW WEB CONTENT TO BE ADDED TO THE DESKTOP
' ENABLE DETECT ACCIDENTLY DOUBLE CLICK
' ENABLED SMOOTH SCROLLING
' UNCHECK PREFIX SHORTCUT
' UNCHECK SHOW ENCRYPT ON CONTECT MENU
' UNCHECK SHOW PICTURES START MENU
' ENABLE - EXPLORER - CLASSIC SEARCH IN EXPLORER
' ENABLE - TAB TO AUTO COMPLETE

' EXPLORER - THUMBNAILS - IMAGE QUALITY HIGHER - PIXEL SIZE 120 DEFAULT 96
' SLIDESHOW SPEED UP MAX

' DESKTOP - UNCHECK DISPLAY INTERNET EXPLORER

' MY COMPUTER CHANGE TO FIRST ICON
' MY COMPUTER - SECIAL FOLDERS - CHANGE INSTALL PATH
' MY COMPUTER - SECIAL FOLDERS - SHARED FOLDERS - CHANGE - HERE - IF YOU WANT
' MY COMPUTER AUTOPLAY NOT FOR REMOVEABLE DRIVES

' GO TO LOGON
' REMOVE IF ANY MESS - WHAT YOU DONT WANT IN WELCOME LOGON MENU
' TRY LOGON SETTING MAKE A MESS OF THNIGS

'### -----
' TWEAK XP PRO -------------- SOME ARE REPEAT FROM TWEAK UI
' SYSTEM PERFORMANCE ----
' CACHE OPTIMIZATION - ENABLED
' ENABLED BOOT DEFRAG
'
' TRICKS AND TWEAKS ----- WINDOWS TWEAKS -----
' TWEAK - TWEAT BIRD - TEAK FINISH - ANTIUQES
' -- START MENU TWEAKS ------
' -- CLASSIC START MENU ----
' HIDE FAVORITES FOLDER
' -- NEW STARTMENU ---------
' HIDE PICTURES
' HIDE MY MUSIC
' CASCADE CONTROL PANEL
' HIDE FAVORITES
' NOTIFY NEW APPLICATIONS - MINE = ON
' DISPLAY ADMIN TOOLS - ENABLED
' RESORT STARTMENU
' -- DESKTOP TWEAKS --------
' ENABLE FONT SMOOTHING - WITH CLEAR TYPE FONT SMOOTHING
' THAT DISABLES MAKE DESKTOP ICONS TRANSPARENT
' -- MISCELLANEOUS TWEAKS -----
' ENABLE REMOVE SHORTCUT ON NEW LINKS
' STARTMENU SPEED QUICKER DEFAULT 400 TRY 202
' -- WINDOWS TWEAKS II -----
' DISABLE DESKTOP CLEANUP WIZARD
' DISABLE SHOW ENCRYPTION COMMAND ON THE FILE CONTECT-MENU
' MOUSE TWEAK 
' VBCHECK ACCIDENTLY DOUBLE CLICKS - NOT SURE THIS IS ENABLE OR DISABLE - DEFAULT UNCHECKED
' ENABLE WINDOWS XP ERROR REPORTING DEFAULT VBCHECKED
'
' CONFIGER CLEAR TYPE
' CHOOSE FIRST ONE - REBOOT REQUIED

' COPY QUICKLAUNCH LINKS FOLDER OVER MANUALLY - LINK TO NEW USER FACTORY SET FOLDER IN E DRIVE QUICK FOLDER
' THIS IS THE FOLDER
' "C:\Documents and Settings\MATT 02\Application Data\Microsoft\Internet Explorer\Quick Launch"

' ### -----
' IMPORTANT WINRAR CONFIG SETTINGS

' ### -----
' GOTO LANGUAGE BAR SETTINGS AND LANGUAGE BAR BUTTON AND TURN OFF

' ### -----
' SET THE DATE FORMAT FROM *LOUSY* FORMAT USA
' CONTROL PANEL ---------------------
' REGIONAL AND LANGUAGE OPTIONS -----
' CUSTOMIZE -------------------------
' CUSTOMISE REGOINAL OPTIONS
' SET SHORT DATE FORMAT TO DD/MM/YYYY

' ### -----
' GO TO DESKTOP - MY COMPUTER DRAG AND DROP TO LEFT MOST AND GET A TOOLBAR 
' NOT FOR WINDOWS 7
' SET ICONS SMALL
' SOMETHING ELSE YOU DO - LIKE SET ONLY DISPLAY MY COMPUTER TOOLBAR

'### -----
' BIT OF OFFICE TOOL BAR SETUP FOLDERS AGAIN - DONE

'### -----
' AFTER BOOT SEE NEXT STEP BEOFRE FIRST
' AND SYNC YOUR STARTUP FOLDERS WITH ANOTHER COMPUTER ON NETWORK
'### -----
' USE SYSINTERNALS SUITE BGINFO IF YOU DO USE THAT
' YOUR NEED TO UPDATE BECAUSE OF THE NEW LOGON
' I RUN IT FROM HERE C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\Bginfo.exe
' BUT I HAVE'NT GOT A LINK NEAR ONE IN START FOLDER
 
' ###############
' ###############
' RUN THIS SCRIPT LOGG OFF ON AGAIN
' SOME SPECIAL FOLDERS WIL BE SET BEFORE RELOGON 
' BUT NOT STARTMENU UNTIL REBOOT OR RELOGON
' BUT IF YOU EDIT CHANGE TO START MENU - RIGHT CLICK START MENU CUSTOMIZE 
' IT WILL BRING THE START MENU FROM THE NEW REG SETTING WITHOUT LOGG OFF

' ### -----
' SET THE DOS COMMAND PROMPT TO LARGE HISTORY BUFFER - BUT MAY FIND SCROLL TO LAST PLACE LEFT 
' A BIT IN THE MIDDLE - UNLESS REACHED BOTTOM
' THINK ITS GLOBAL BUT ALSO SET QUICK EDIT IF WANT TO
' SET LUCIDA CONSOLE FONT

' AND KATMOUSE HAS TO SET NOT TO SCROLL ON BROWSER
' NORTON WANTS TO INSTALL TO BROWSER - DONT HAVE IT
' FEEDLY WANTS LOGON - GMAIL ACCOUNT IS THE ONE TO HAVE MULTIPLE FACEBOOK ACCOUNTS
' SOCIAL FIXER TAKES A FEW SETTING
' ROBOFORM REQUIRES A SYNC - SOMETIMES BEEN LEFT FOR LONG - SO SYNC ON A OLD USER FIRST



'----------------------------------------
' ALSO ONE FOR NOTE'S
' GOTO RIGHT CLICK MY COMPUTER PROPERTIES
' OR
' CONTROL PANEL SYSTEM 
' LOOK AT SYSTEM PROPERTIES
' LOOK AT ADVANCED
' LOOK AT USER PROFILES - DESKTOP SETTINGS RELATED TO YOUR LOGON
' AND YOUR FIND A RESULT IF YOU MANUALLY USED THIS SCRIPT
' TO SET UP YOUR NEW USER
' BECAUSE I DONE A TRANSFER SETTINGS TO A NEW USER #02
' AND THESE RESULTS
' PROFILE STORE ON COMPUTER
' #1 OLD MAIN USER LEAVING BEHIND 
' 110MB USAGE
' #02 USER STARTED FROM TRANSFER SETTINGS
' 40.8MB TO SET WITH
' #03 CLEAN NEW USER 
' 2.43MB
' RESULT CLEAN USER WITHOUT TAKING UP SPACE

' AND A LOOK AT DOCUMENTS AND SETTINGS
' SHOW TWO NEWEST USER #02 #03
' #01 1309 MB OF SPACE
' #02 73.3 MB OF SPACE
' #03 36 MB OF SPACE
'
'THINK I'LL USE #03 STRAIGHT AWAY AND SKIP THE #02
'MIGHT BE #02 HAS BEEN LONGER ON THE SYSTEM
'
'AND THEN I ALWAYS NEED A EXTRA ONE FORWARD AT TOP OR BOTTOM 
'INCASE OF DISARSTER
'SO HERE GO AGAIN

'2015 NOV - AND THEN YOUR ON YOUR OWN
 

REG ADD "HKCR\exefile\shell\Run with Above Normal Priority" /ve /f
REG ADD "HKCR\exefile\shell\Run with Above Normal Priority\Command" /ve /t REG_SZ /d "cmd.exe /c start \"Run with Above Normal Priority\" /AboveNormal \"%%1\"" /f




REM -------------------------
REM reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "My Video" /t REG_EXPAND_SZ /d "M:\Media\Videos"
REM reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "My Video" /t REG_SZ /d "M:\Media\Videos"
REM -------------------------



EXIT





---------------------
Count = 083 -- Sun 23-Aug-2015 01:23:07
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
67.71509552	Tweak-xp.exe:86024	SetValue	HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\DisablePagingExecutive	SUCCESS	0x1	

---------------------
---------------------
Count = 084 -- Sun 23-Aug-2015 01:25:28
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
18.45926476	Tweak-xp.exe:86024	SetValue	HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\Enable	SUCCESS	"Y"	

---------------------
---------------------
Count = 085 -- Sun 23-Aug-2015 01:26:07
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
81.05335236	Tweak-xp.exe:86024	SetValue	HKLM\System\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnablePrefetcher	SUCCESS	0x3	

---------------------
---------------------
Count = 086 -- Sun 23-Aug-2015 01:27:36
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
171.67987061	Tweak-xp.exe:86024	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\StartMenuFavorites	SUCCESS	0x0	

---------------------
---------------------
Count = 087 -- Sun 23-Aug-2015 01:28:09
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
205.61871338	Tweak-xp.exe:86024	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_ShowMyPics	SUCCESS	0x0	

---------------------
---------------------
Count = 088 -- Sun 23-Aug-2015 01:28:23
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
223.71253967	Tweak-xp.exe:86024	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_ShowMyMusic	SUCCESS	0x0	

---------------------
---------------------
Count = 089 -- Sun 23-Aug-2015 01:30:04
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
43.50534821	Tweak-xp.exe:86024	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\StartMenuFavorites	SUCCESS	0x0	

---------------------
---------------------
Count = 090 -- Sun 23-Aug-2015 01:30:35
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
77.38096619	Tweak-xp.exe:86024	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_NotifyNewApps	SUCCESS	0x0	

---------------------
---------------------
Count = 091 -- Sun 23-Aug-2015 01:31:23
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
105.33364105	Tweak-xp.exe:86024	SetValue	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\StartMenuAdminTools	SUCCESS	"YES"	

---------------------
---------------------
Count = 092 -- Sun 23-Aug-2015 01:32:49
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
148.41154480	Tweak-xp.exe:86024	OpenKey	HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MenuOrder\Start Menu2\	SUCCESS	Access: 0x1 	

---------------------
---------------------
Count = 093 -- Sun 23-Aug-2015 01:34:59
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
14.71375275	Tweak-xp.exe:86024	SetValue	HKCU\Control Panel\Desktop\DragFullWindows	SUCCESS	"1"	

---------------------
---------------------
Count = 094 -- Sun 23-Aug-2015 01:36:13
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
63.95655823	Tweak-xp.exe:86024	CreateKey	HKLM\System\CurrentControlSet\Control\DeviceClasses	SUCCESS	Access: 0xF003F 	

---------------------
---------------------
Count = 095 -- Sun 23-Aug-2015 01:36:47
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
63.95739746	Tweak-xp.exe:86024	CreateKey	HKLM\System\CurrentControlSet\Control\Class\{4D36E96C-E325-11CE-BFC1-08002BE10318}\0015	SUCCESS	Access: 0xF003F 	

---------------------
---------------------
Count = 096 -- Sun 23-Aug-2015 01:36:50
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
63.95743179	Tweak-xp.exe:86024	CreateKey	HKLM\System\CurrentControlSet\Control\Class\{4D36E96C-E325-11CE-BFC1-08002BE10318}\0015\GlobalSettings	SUCCESS	Access: 0xF003F 	

---------------------
---------------------
Count = 097 -- Sun 23-Aug-2015 01:36:53
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
63.95762634	Tweak-xp.exe:86024	CreateKey	HKLM\System\CurrentControlSet\Control\Class\{4D36E96C-E325-11CE-BFC1-08002BE10318}\0015	SUCCESS	Access: 0xF003F 	

---------------------
---------------------
Count = 098 -- Sun 23-Aug-2015 01:36:56
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
63.95765305	Tweak-xp.exe:86024	CreateKey	HKLM\System\CurrentControlSet\Control\Class\{4D36E96C-E325-11CE-BFC1-08002BE10318}\0015\GlobalSettings	SUCCESS	Access: 0xF003F 	

---------------------
---------------------
Count = 099 -- Sun 23-Aug-2015 01:38:37
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
215.74893188	Tweak-xp.exe:86024	DeleteValueKey	HKCR\lnkfile\IsShortcut	SUCCESS		

---------------------
---------------------
Count = 100 -- Sun 23-Aug-2015 01:38:40
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
215.74942017	Tweak-xp.exe:86024	DeleteValueKey	HKCR\piffile\IsShortcut	SUCCESS		

---------------------





8-MSI-GP62M-7RD __ ORIGINAL
System Directory x86	No	No	No	C:\Windows\SysWOW64	0x29	CSIDL_SYSTEMX86	%SystemRoot%\SysWOW64	
System Directory	No	No	No	C:\Windows\system32	0x25	CSIDL_SYSTEM	%SystemRoot%\system32	
Fonts	No	Yes	Yes	C:\Windows\Fonts	0x14	CSIDL_FONTS	%SystemRoot%\Fonts	
Windows Directory	No	No	No	C:\Windows	0x24	CSIDL_WINDOWS	%SystemRoot%	
Common Video	No	Yes	No	C:\Users\Public\Videos	0x37	CSIDL_COMMON_VIDEO	%SystemDrive%\Users\Public\Videos	
Common Pictures	No	Yes	No	C:\Users\Public\Pictures	0x36	CSIDL_COMMON_PICTURES	%SystemDrive%\Users\Public\Pictures	
Common Music	No	Yes	No	C:\Users\Public\Music	0x35	CSIDL_COMMON_MUSIC	%SystemDrive%\Users\Public\Music	
Common Documents	No	Yes	No	C:\Users\Public\Documents	0x2e	CSIDL_COMMON_DOCUMENTS	%SystemDrive%\Users\Public\Documents	
Common Desktop	Yes	Yes	No	C:\Users\Public\Desktop	0x19	CSIDL_COMMON_DESKTOPDIRECTORY	%SystemDrive%\Users\Public\Desktop	
My Video	No	Yes	No	C:\Users\MATT 01\Videos	0x0e	CSIDL_MYVIDEO	%USERPROFILE%\Videos	
My Pictures	No	Yes	No	C:\Users\MATT 01\Pictures	0x27	CSIDL_MYPICTURES	%USERPROFILE%\Pictures	
My Music	No	Yes	No	C:\Users\MATT 01\Music	0x0d	CSIDL_MYMUSIC	%USERPROFILE%\Music	
Common Favorites	No	Yes	No	C:\Users\MATT 01\Favorites	0x1f	CSIDL_COMMON_FAVORITES	%USERPROFILE%\Favorites	
Favorites	No	Yes	No	C:\Users\MATT 01\Favorites	0x06	CSIDL_FAVORITES	%USERPROFILE%\Favorites	
My Documents	No	Yes	No	C:\Users\MATT 01\Documents	0x05	CSIDL_PERSONAL	%USERPROFILE%\Documents	
Desktop	No	Yes	No	C:\Users\MATT 01\Desktop	0x10	CSIDL_DESKTOPDIRECTORY	%USERPROFILE%\Desktop	
Templates	No	No	No	C:\Users\MATT 01\AppData\Roaming\Microsoft\Windows\Templates	0x15	CSIDL_TEMPLATES	%APPDATA%\Microsoft\Windows\Templates	
Startup	No	Yes	No	C:\Users\MATT 01\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup	0x07	CSIDL_STARTUP	%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup	
Administrative Tools	No	Yes	No	C:\Users\MATT 01\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Administrative Tools	0x30	CSIDL_ADMINTOOLS	%APPDATA%\Microsoft\Windows\Start Menu\Programs\Administrative Tools	
Start Menu Programs	No	Yes	No	C:\Users\MATT 01\AppData\Roaming\Microsoft\Windows\Start Menu\Programs	0x02	CSIDL_PROGRAMS	%APPDATA%\Microsoft\Windows\Start Menu\Programs	
Start Menu	No	Yes	No	C:\Users\MATT 01\AppData\Roaming\Microsoft\Windows\Start Menu	0x0b	CSIDL_STARTMENU	%APPDATA%\Microsoft\Windows\Start Menu	
Send To	No	Yes	No	C:\Users\MATT 01\AppData\Roaming\Microsoft\Windows\SendTo	0x09	CSIDL_SENDTO	%APPDATA%\Microsoft\Windows\SendTo	
Recent	No	Yes	No	C:\Users\MATT 01\AppData\Roaming\Microsoft\Windows\Recent	0x08	CSIDL_RECENT	%APPDATA%\Microsoft\Windows\Recent	
PrintHood	No	No	No	C:\Users\MATT 01\AppData\Roaming\Microsoft\Windows\Printer Shortcuts	0x1b	CSIDL_PRINTHOOD	%APPDATA%\Microsoft\Windows\Printer Shortcuts	
NetHood	No	No	No	C:\Users\MATT 01\AppData\Roaming\Microsoft\Windows\Network Shortcuts	0x13	CSIDL_NETHOOD	%APPDATA%\Microsoft\Windows\Network Shortcuts	
Quick Launch	No	Yes	No	C:\Users\MATT 01\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch			%APPDATA%\Microsoft\Internet Explorer\Quick Launch	
Application Data	No	No	No	C:\Users\MATT 01\AppData\Roaming	0x1a	CSIDL_APPDATA	%APPDATA%	
Temporary Folder	No	No	No	C:\Users\MATT 01\AppData\Local\Temp\			%USERPROFILE%\AppData\Local\Temp\	
Cookies	Yes	No	Yes	C:\Users\MATT 01\AppData\Local\Microsoft\Windows\INetCookies	0x21	CSIDL_COOKIES	%USERPROFILE%\AppData\Local\Microsoft\Windows\INetCookies	
Temporary Internet Files	Yes	No	Yes	C:\Users\MATT 01\AppData\Local\Microsoft\Windows\INetCache	0x20	CSIDL_INTERNET_CACHE	%USERPROFILE%\AppData\Local\Microsoft\Windows\INetCache	
History	No	No	Yes	C:\Users\MATT 01\AppData\Local\Microsoft\Windows\History	0x22	CSIDL_HISTORY	%USERPROFILE%\AppData\Local\Microsoft\Windows\History	
CD Burning	No	Yes	No	C:\Users\MATT 01\AppData\Local\Microsoft\Windows\Burn\Burn	0x3b	CSIDL_CDBURN_AREA	%USERPROFILE%\AppData\Local\Microsoft\Windows\Burn\Burn	
Local Application Data	No	No	No	C:\Users\MATT 01\AppData\Local	0x1c	CSIDL_LOCAL_APPDATA	%USERPROFILE%\AppData\Local	
Profile Folder	No	No	No	C:\Users\MATT 01	0x28	CSIDL_PROFILE	%USERPROFILE%	
Common Templates	No	No	No	C:\ProgramData\Microsoft\Windows\Templates	0x2d	CSIDL_COMMON_TEMPLATES	%ALLUSERSPROFILE%\Microsoft\Windows\Templates	
Common Startup	No	Yes	No	C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup	0x18	CSIDL_COMMON_STARTUP	%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Startup	
Common Administrative Tools	No	Yes	No	C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools	0x2f	CSIDL_COMMON_ADMINTOOLS	%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Administrative Tools	
Common Start Menu Programs	No	Yes	No	C:\ProgramData\Microsoft\Windows\Start Menu\Programs	0x17	CSIDL_COMMON_PROGRAMS	%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs	
Common Start Menu	No	Yes	No	C:\ProgramData\Microsoft\Windows\Start Menu	0x16	CSIDL_COMMON_STARTMENU	%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu	
Common Application Data	Yes	No	No	C:\ProgramData	0x23	CSIDL_COMMON_APPDATA	%ALLUSERSPROFILE%	
Program Files - Common	No	No	No	C:\Program Files\Common Files	0x2b	CSIDL_PROGRAM_FILES_COMMON	%ProgramFiles%\Common Files	
Program Files - Common x86	No	No	No	C:\Program Files (x86)\Common Files	0x2c	CSIDL_PROGRAM_FILES_COMMONX86	%ProgramFiles% (x86)\Common Files	
Program Files x86	No	Yes	No	C:\Program Files (x86)	0x2a	CSIDL_PROGRAM_FILESX86	%ProgramFiles% (x86)	
Program Files	No	Yes	No	C:\Program Files	0x26	CSIDL_PROGRAM_FILES	%ProgramFiles%	
Downloads	No	Yes	No	C:\DOWNLOAD			%SystemDrive%\DOWNLOAD	


:VERY_ENDER_BLOCK
EXIT /B
)
