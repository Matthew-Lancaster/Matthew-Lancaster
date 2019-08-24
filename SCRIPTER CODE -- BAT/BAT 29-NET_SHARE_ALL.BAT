@ECHO OFF

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

REM ----------------------------------------------------------------------

REM ------------------------------------------------------------------------
REM NET SHARE "Users"="C:\Users" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM Users was shared successfully.
REM ------------------------------------------------------------------------

REM ------------------------------------------------------------------------
REM [/CACHE:Manual | Documents | Programs | BranchCache | None]
REM /CACHE -- NONE
REM ------------------------------------------------------------------------
REM -- NOT WANT OFFLINE USER -----------------------------------------------
REM ------------------------------------------------------------------------

ECHO ____ NET_SHARE_ALL.BAT ____ 
TITLE %~f0

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


ECHO -----------------------------------------------------------------------
ECHO -----------------------------------------------------------------------
ECHO -----------------------------------------------------------------------
ECHO -----------------------------------------------------------------------

REM ------------------------------------------------------------------------
REM GOTO JUMP_OVER
REM ------------------------------------------------------------------------
REM -- WIN XP DON'T APPEAR TO USE COMPUTER NAME AT ALL -- AND STILL WORK WITH IT IN COMMAND LINE
REM -- BUT AND THEN NOT HARM TO KEEP IT IN THE COMMAND LINE
REM ------------------------------------------------------------------------
@D:
@CD\


@REM ---------------------------------------------------------------
@ECHO OFF
@REM ---------------------------------------------------------------
SET WINXP__=NOT
SET WIN07__=NOT
SET WIN10__=NOT
ver | find "5.1." > nul
IF %errorlevel% EQU 0 SET WINXP__=YES
ver | find "6.1." > nul
IF %errorlevel% EQU 0 (
	SET WIN07__=YES
	SET WIN10__=YES
)
ver | find "10.0." > nul
IF %errorlevel% EQU 0 (
	SET WIN10__=YES
	SET WIN07__=YES
)
@REM ---------------------------------------------------------------

@REM ---------------------------------------------------------------
IF "%WINXP__%"=="YES" ECHO WINDOWS XP
IF "%WIN07__%"=="YES" ECHO WINDOWS 7
IF "%WIN10__%"=="YES" (
	ECHO WINDOWS 10
)


REM Setlocal enabledelayedexpansion

REM Set "Pattern=-"
REM Set "Replace=_"

REM Set "COMPUTERNAME_DASH_CONVERT_VAR=%COMPUTERNAME%"
REM REM %%~a"
REM Ren "%%a" "!File:%Pattern%=%Replace%!"
	
REM -- ----
REM -- cmd windows replace spaces with underscores - Stack Overflow
REM -- http://stackoverflow.com/questions/20791264/cmd-windows-replace-spaces-with-underscores/37853017#37853017
REM -- ----	
Setlocal enabledelayedexpansion
set COMPUTERNAME_DASH_CONVERT_VAR=%COMPUTERNAME%
set COMPUTERNAME_DASH_CONVERT_VAR=%COMPUTERNAME:-=_%
SET NAME_VAR=%COMPUTERNAME:-=_%

REM ------------------------------------------------------------------------
call :COMPUTER_NAME_COMMON_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
IF %COMPUTERNAME%==1-ASUS-X5DIJ     call :1-ASUS-X5DIJ_NET_SHARE_DELETE
IF %COMPUTERNAME%==2-ASUS-EEE       call :2-ASUS-EEE_NET_SHARE_DELETE
IF %COMPUTERNAME%==3-LINDA-PC       call :3-LINDA-PC_NET_SHARE_DELETE
IF %COMPUTERNAME%==4-ASUS-GL522VW   call :4-ASUS-GL522VW_NET_SHARE_DELETE
IF %COMPUTERNAME%==5-ASUS-P2520LA   call :5-ASUS-P2520LA_NET_SHARE_DELETE
IF %COMPUTERNAME%==7-ASUS-GL522VW   call :7-ASUS-GL522VW_NET_SHARE_DELETE
IF %COMPUTERNAME%==8-MSI-GP62M-7RD  call :8-MSI-GP62M-7RD_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
:JUMP_OVER
REM ------------------------------------------------------------------------


REM ------------------------------------------------------------------------
REM -- SET THE NET SHARE SET
REM ------------------------------------------------------------------------
CALL :COMPUTER_NAME_COMMON_NET_SHARE_CREATE
REM ------------------------------------------------------------------------
IF %COMPUTERNAME%==2-ASUS-EEE       call :2-ASUS-EEE
IF %COMPUTERNAME%==1-ASUS-X5DIJ     call :1-ASUS-X5DIJ
IF %COMPUTERNAME%==3-LINDA-PC       call :3-LINDA-PC
IF %COMPUTERNAME%==4-ASUS-GL522VW   call :4-ASUS-GL522VW
IF %COMPUTERNAME%==5-ASUS-P2520LA   call :5-ASUS-P2520LA
IF %COMPUTERNAME%==7-ASUS-GL522VW   call :7-ASUS-GL522VW
IF %COMPUTERNAME%==8-MSI-GP62M-7RD  call :8-MSI-GP62M-7RD
REM ------------------------------------------------------------------------

ECHO -----------------------------------------------------------------------
ECHO -----------------------------------------------------------------------
REM  -----------------------------------------------------------------------
REM  --- SHOW THE AFTER RESULT
REM  -----------------------------------------------------------------------
ECHO ON

NET SHARE

@ECHO OFF
ECHO -----------------------------------------------------------------------
ECHO -----------------------------------------------------------------------

REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
REM --- END PRESS A KEY TO ENDE
PAUSE
GOTO ENDE
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------


REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
REM -- THE DELETE NET SHARE SET COMMON ---- 1ST OF 2
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------

::
:COMPUTER_NAME_COMMON_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
ECHO ON
NET SHARE "My Music" \\%COMPUTERNAME% /DELETE
NET SHARE "Music" \\%COMPUTERNAME% /DELETE
NET SHARE "Pictures" \\%COMPUTERNAME% /DELETE
NET SHARE "My Pictures" \\%COMPUTERNAME% /DELETE
NET SHARE "My Video" \\%COMPUTERNAME% /DELETE
NET SHARE "Users" \\%COMPUTERNAME% /DELETE
NET SHARE "SharedDocs" \\%COMPUTERNAME% /DELETE
REM NET SHARE "Print$" \\%COMPUTERNAME% /DELETE
NET SHARE "C$" \\%COMPUTERNAME% /DELETE
NET SHARE "D$" \\%COMPUTERNAME% /DELETE
NET SHARE "E$" \\%COMPUTERNAME% /DELETE
NET SHARE "F$" \\%COMPUTERNAME% /DELETE
NET SHARE "G$" \\%COMPUTERNAME% /DELETE
REM NET SHARE "Admin$" \\%COMPUTERNAME% /DELETE

NET SHARE "# 01 My Documents" /DELETE

REM NET SHARE "IPC$" /DELETE
REM ------------------------------------------------------------------------
REM NET SHARE Admin$ /DELETE
REM IF SET HERE TO DELETE AND PUT BACK HAVE TO REBOOT MACHINE TO GAIN ACCESS
REM GAIN ACCESS THE REMOTE CALLER COMPUTER WONT HAVE ANY UNTIL REBOOT
REM ------------------------------------------------------------------------


REM --------------------------------------------------------------------
REM DIFFICULT TO DO DELETE IF USER ARE CONNECTED THEN ASK A QUESTION Y/N
REM CAN'T SEND INPUT TO ANSWER
REM --------------------------------------------------------------------
REM DON'T MUST HAVE TO INCLUDE COMPUTER NAME WHEN DELETER
REM --------------------------------------------------------------------

NET SHARE "%NAME_VAR%_01_C-DRIVE" /DELETE
NET SHARE "%NAME_VAR%_02_D-DRIVE" /DELETE
NET SHARE "%COMPUTERNAME%_03_FAT32_4GB" /DELETE

REM NET SHARE "%NAME_VAR%_01_C-DRIVE" \\%COMPUTERNAME% /DELETE
REM NET SHARE "%NAME_VAR%_02_D-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "%NAME_VAR%_01_C_DRIVE" /DELETE
NET SHARE "%NAME_VAR%_02_D_DRIVE" /DELETE
NET SHARE "%NAME_VAR%_03_FAT32_4GB" /DELETE
NET SHARE "%COMPUTERNAME%_01_C_DRIVE" /DELETE
NET SHARE "%COMPUTERNAME%_02_D_DRIVE" /DELETE
NET SHARE "%COMPUTERNAME%_03_FAT32_4GB" /DELETE
REM NET SHARE "%NAME_VAR%_04_USB32_GB" \\%COMPUTERNAME% /DELETE
REM 2_ASUS_X5DIJ_01_C_DRIVE
NET SHARE "2_ASUS_X5DIJ_01_C_DRIVE" /DELETE
NET SHARE "2_ASUS_X5DIJ_02_D_DRIVE" /DELETE
NET SHARE "2_ASUS_X5DIJ_03_FAT32_4GB" /DELETE
REM PAUSE

NET SHARE "%NAME_VAR%_05_1SAMSUNG_4TB" /DELETE
NET SHARE "%NAME_VAR%_07_2SAMSUNG_4TB" /DELETE
NET SHARE "%NAME_VAR%_08_1ELEMENTS4TB" /DELETE
NET SHARE "%NAME_VAR%_09_2ELEMENTS4TB" /DELETE

@REM ---- LINE 1.. WAS UNABLE TO DELETER
@REM ---- LINE 2.. WAS OKAY ABOUT DELETE
NET SHARE "4_ASUS_GL522VW_08_1ELEMENTS4TB" /DELETE
NET SHARE 4_ASUS_GL522VW_08_1ELEMENTS4TB \\%COMPUTERNAME% /DELETE

NET SHARE "%NAME_VAR%_08_1ELEMENTS_4TB" /DELETE
NET SHARE "%NAME_VAR%_09_2ELEMENTS_4TB" /DELETE
NET SHARE "%NAME_VAR%_05_1_SAMSUNG_4TB" /DELETE
NET SHARE "%NAME_VAR%_07_2_SAMSUNG_4TB" /DELETE
NET SHARE "%NAME_VAR%_08_1_ELEMENTS_4TB" /DELETE
REM NET SHARE "%NAME_VAR%_09_2_ELEMENTS_4TB" /DELETE
REM NET SHARE "4_ASUS_GL522VW_08_1ELEMENTS4TB" \\%COMPUTERNAME% /DELETE
REM NET SHARE I:\ /DELETE

REM NET SHARE "%NAME_VAR%_04_2_SAMSUNG_4TB_C" /DELETE

REM NET SHARE "%NAME_VAR%_10_2_SAMSUNG_4TB_D" /DELETE
REM NET SHARE "7_ASUS_GL522VW_10_2_SAMSUNG_4TB_D" /DELETE
REM NET SHARE "7_ASUS_GL522VW_10_2_SAMSUNG_4TB_D" \\%COMPUTERNAME% /DELETE

REM NET SHARE "7_ASUS_GL522VW_11_2_SAMSUNG_4TB_D" \\%COMPUTERNAME% /DELETE
REM NET SHARE "7-ASUS-GL522VW_11_2_SAMSUNG_4TB_D" \\%COMPUTERNAME% /DELETE

REM COMES FROM ONE DRIVE
NET SHARE "DOCUMENTS" /DELETE

goto :EOF
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------

REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
REM -- THE DELETE NET SHARE SET 1 TO 7 ROUTINE 1 MISSING ---- 2ND OF 2
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------

::
:2-ASUS-EEE_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
NET SHARE "asus-eee e-win 7" \\%COMPUTERNAME% /DELETE
NET SHARE "asus-eee e-win 7" \\%COMPUTERNAME% /DELETE
NET SHARE "asus-eee wd-55gb" \\%COMPUTERNAME% /DELETE
NET SHARE "SharedDocs" \\%COMPUTERNAME% /DELETE
NET SHARE "Printer2" \\%COMPUTERNAME% /DELETE
NET SHARE "Printer4" \\%COMPUTERNAME% /DELETE
goto :EOF
REM ------------------------------------------------------------------------

::
:1-ASUS-X5DIJ_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
NET SHARE "SharedDocs" \\%COMPUTERNAME% /DELETE
NET SHARE "Printer" \\%COMPUTERNAME% /DELETE
NET SHARE "Printer4" \\%COMPUTERNAME% /DELETE
goto :EOF
REM ------------------------------------------------------------------------

::
:3-LINDA-PC_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
REM ECHO ON
NET SHARE "My Music" \\%COMPUTERNAME% /DELETE
NET SHARE "My Music2" \\%COMPUTERNAME% /DELETE
NET SHARE "My Pictures" \\%COMPUTERNAME% /DELETE
NET SHARE "Pictures2" \\%COMPUTERNAME% /DELETE
NET SHARE "My Video" \\%COMPUTERNAME% /DELETE
NET SHARE "My Video2" \\%COMPUTERNAME% /DELETE
NET SHARE "SharedDocs" \\%COMPUTERNAME% /DELETE
NET SHARE "Users" \\%COMPUTERNAME% /DELETE

REM My Music     D:\# MY DOCS\# 01 My Documents\My Music
REM My Music2    D:\# MY DOCS COMMON\My Music
REM My Pictures  D:\# MY DOCS\# 01 My Documents\My Pictures
REM My Pictures2 D:\# MY DOCS COMMON\My Pictu...
REM My Video     D:\# MY DOCS\# 01 My Documents\My Video
REM My Video2    D:\# MY DOCS COMMON\My Video
REM SharedDocs   C:\USERS\PUBLIC\DOCUMENTS
REM Users        C:\Users

goto :EOF
REM ------------------------------------------------------------------------

::
:4-ASUS-GL522VW_NET_SHARE_DELETE
REM ------------------------------------------------------------------------

NET SHARE "SharedDocs" \\%COMPUTERNAME% /DELETE
NET SHARE "Users" \\%COMPUTERNAME% /DELETE
NET SHARE "%NAME_VAR%_05_NVR_1000_DRIVE" \\%COMPUTERNAME% /DELETE

NET SHARE "%NAME_VAR%_04_1_SAMSUNG_4TB_C" \\%COMPUTERNAME% /DELETE
NET SHARE "%NAME_VAR%_05_1_SAMSUNG_4TB_D" \\%COMPUTERNAME% /DELETE
NET SHARE "%NAME_VAR%_08_1_SAMSUNG_4TB_E" \\%COMPUTERNAME% /DELETE

NET SHARE "4_ASUS_GL522VW_05_1_SAMSUNG_4TB_D" \\%COMPUTERNAME% /DELETE
NET SHARE "4_ASUS_GL522VW_08_1_SAMSUNG_4TB_E" \\%COMPUTERNAME% /DELETE

REM IF SHARE PATH NOT DELETE THEN MOST LIKELY THE DRIVE NAME TO THEM IS NOT THERE
REM USE THEN
REM SUBST G: E:\


rem 2015 2016                          D:\DSC\2015 2016
rem 4-ASUS-GL522VW_01_C-DRIVE          C:\
rem 4-ASUS-GL522VW_02_D-DRIVE          D:\
rem 4-ASUS-GL522VW_03_FAT32            E:\
rem 4-ASUS-GL522VW_04_NVR-1000-DRIVE   G:\
rem SharedDocs   C:\USERS\PUBLIC\DOCUMENTS
rem Users        C:\Users

goto :EOF
REM ------------------------------------------------------------------------

::
:5-ASUS-P2520LA_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
REM -- ECHO ON
NET SHARE "5-ASUS-P2520L-2SAMASUNG4TB" \\%COMPUTERNAME% /DELETE
NET SHARE "#C" \\%COMPUTERNAME% /DELETE
REM -- BOTH ON K DRIVE AND DRIVE DON'T EXIST -- AND BUT WON'T DELETE
REM -- USE WORK-AROUND AND SUBST WON'T DO IT
REM -- CREATE A PARTITION WITH PART-WIZARD 
REM -- AND AND NOT AT ROOT THE PATH HAS TO BE MADE FOR IT ABLE TO DELETE
REM -- ECHO OFF

NET SHARE "Music" \\%COMPUTERNAME% /DELETE
NET SHARE "Pictures" \\%COMPUTERNAME% /DELETE
NET SHARE "Users" \\%COMPUTERNAME% /DELETE
NET SHARE "SharedDocs" \\%COMPUTERNAME% /DELETE
NET SHARE "01 Banging Tunes" \\%COMPUTERNAME% /DELETE
NET SHARE "2015 2016" \\%COMPUTERNAME% /DELETE
NET SHARE "5-ASUS-P2520LA_REMOTE_2-SAMSUNG_4TB" \\%COMPUTERNAME% /DELETE
NET SHARE "5-ASUS-P2520LA_REMOTE_2-SAMSUNG_4TB" \\%COMPUTERNAME% /DELETE

REM #C                                   K:\#\#C
REM 01 Banging Tunes                     D:\0 00 MUSIC ---\04 My Music\01 Banging Tunes
REM 2015 2016                            D:\DSC\2015 2016
REM 5-ASUS-P2520L FAT32-4TB              F:\
REM 5-ASUS-P2520L-2SAMASUNG4TB           K:\
REM 5-ASUS-P2520LA_01_C-DRIVE            C:\
REM 5-ASUS-P2520LA_02_D-DRIVE            D:\
REM 5-ASUS-P2520LA_03_FAT32              E:\
REM 5-ASUS-P2520LA_REMOTE_2-SAMSUNG_4TB  H:\
REM SharedDocs                           C:\USERS\PUBLIC\DOCUMENTS

goto :EOF
REM ------------------------------------------------------------------------

::
:7-ASUS-GL522VW_NET_SHARE_DELETE
REM ------------------------------------------------------------------------

NET SHARE "Music" \\%COMPUTERNAME% /DELETE
NET SHARE "Pictures" \\%COMPUTERNAME% /DELETE
NET SHARE "Users" \\%COMPUTERNAME% /DELETE

REM 7-ASUS-GL522VW_01_C-DRIVE        C:\
REM 7-ASUS-GL522VW_02_D-DRIVE        D:\
REM 7-ASUS-GL522VW_03_FAT32          E:\
REM Music        D:\ON-LINE-SYNC\OneDrive\Music
REM Pictures     D:\ON-LINE-SYNC\OneDrive\Pictures
REM Users        C:\Users

goto :EOF
REM ------------------------------------------------------------------------

::
:8-MSI-GP62M-7RD_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
NET SHARE "D DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "START MENU" \\%COMPUTERNAME% /DELETE
NET SHARE "STARTUP" \\%COMPUTERNAME% /DELETE
NET SHARE "# 01 My Documents" \\%COMPUTERNAME% /DELETE
NET SHARE "Music" \\%COMPUTERNAME% /DELETE
NET SHARE "Pictures" \\%COMPUTERNAME% /DELETE
NET SHARE "Users" \\%COMPUTERNAME% /DELETE

REM Music        D:\ON-LINE-SYNC\OneDrive\Music
REM Pictures     D:\ON-LINE-SYNC\OneDrive\Pictures
REM Users        C:\Users

goto :EOF
REM ------------------------------------------------------------------------



REM ------------------------------------------------------------------------
REM ---- CREATE __ 2ND
REM ------------------------------------------------------------------------


REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
::
:COMPUTER_NAME_COMMON_NET_SHARE_CREATE
REM ------------------------------------------------------------------------
ECHO ON

REM ---------------------------------------------------------------------------
REM STRING WITH $ AFTER ARE HIDDEN
REM IPC$ USER ANONYMOUS AND ALLOW PROCESS TO ENUMERATE THE DOMAIN REMOTELY __
REM __ USEFUL
REM ADMIN$ HAS TO ENABLE HERE IS ENUMERATE DOMAIN WITH PSEXEC.EXE 
REM BY SYSINTERNALS
REM TO USE RUNAS COMMAND SET REMOTELY EACH DOMAIN
REM -- NET SHARE ADMIN$ __ TYPE COMMAND WILL TELL MORE
REM -- FORMAT IS DIFFERENT IPC$ AND NOTHING
REM -- ADMIN$ WITHOUT C:\WINDOWS
REM ---------------------------------------------------------------------------


IF "%WINXP__%"=="YES" (
	NET SHARE IPC$
	NET SHARE ADMIN$ /USERS:99 /UNLIMITED /CACHE:NONE
	NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /USERS:99 /UNLIMITED /CACHE:NONE
	NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /USERS:99 /UNLIMITED /CACHE:NONE
	NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /USERS:99  /UNLIMITED /CACHE:NONE
)
IF "%WIN10__%"=="YES" (
	NET SHARE IPC$
	NET SHARE ADMIN$ /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	
	REM NET SHARE "Clipboard_Logger_#_DATA$"="D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "ART_LOGGERS_-_WEBCAM$"="D:\0 00 ART LOGGERS - WEBCAM" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "ART_LOGGERS_#_APP_AND_SCREEN$"="D:\0 00 Art Loggers\# APP AND SCREEN" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "MY_DOCS_#_01_My_Documents$"="D:\# MY DOCS\# 01 My Documents" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "VIDEO_NOT_ME$"="D:\VIDEO\NOT\ME" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "BT_Cloud_Sync_OLD$"="D:\BT Cloud Sync OLD" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "VB_NT_00_Best_VB_01_Clip_Type_TEXTBOX$"="D:\VB6\VB-NT\00_Best_VB_01\Clip_Type\TEXTBOX" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "VB_NT_00_Best_VB_01_Clip_Type_TEXTBOX_BIGGER$"="D:\VB6\VB-NT\00_Best_VB_01\Clip_Type\TEXTBOX BIGGER" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "My_Documents_00_BlueTooth_SMS's_From_CLipType$"="D:\# MY DOCS\# 01 My Documents\00 BlueTooth SMS's From CLipType" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	
)
goto :EOF


::
:2-ASUS-EEE
REM ---------------------------------------------------------------------------
REM -- WIN XP DON'T USE THIS ____ /GRANT:EVERYONE,FULL
REM -- BUT AND DOES USE THIS ____ /USERS:99 
REM -- IN THE GUI SETTING IT DON'T APPEAR TO USE UNLIMITED ACCESS IN ALL VERSION WIN AND SET MAX AT 20 USER
REM -- WITH THIS EXTRA SETTING AT 99 USER
REM -- IN THE DOS COMMAND NET SHARE AND THE COMPUTER NAME IN THE FORMAT SHOWN IT REPORT THE SHARE HAS MAXIMUM USERS - NOT ANY LIMIT
REM ---------------------------------------------------------------------------
ECHO ON
REM NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /USERS:99 /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /USERS:99 /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /USERS:99  /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /USERS:99 /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_MC_SD_16GB"="H:" /USERS:99  /UNLIMITED /CACHE:NONE

REM --------------------------
REM NET SHARE "%NAME_VAR%_WD-55GB"  /USERS:99  /UNLIMITED /CACHE:NONE
@ECHO OFF
goto :EOF

::
:1-ASUS-X5DIJ
REM ---------------------------------------------------------------------------
REM -- WIN XP DON'T USE THIS ____ /GRANT:EVERYONE,FULL
REM -- BUT AND DOES USE THIS ____ /USERS:99 
REM -- IN THE GUI SETTING IT DON'T APPEAR TO USE UNLIMITED ACCESS IN ALL VERSION WIN AND SET MAX AT 20 USER
REM -- WITH THIS EXTRA SETTING AT 99 USER
REM -- IN THE DOS COMMAND NET SHARE AND THE COMPUTER NAME IN THE FORMAT SHOWN IT REPORT THE SHARE HAS MAXIMUM USERS - NOT ANY LIMIT
REM ---------------------------------------------------------------------------
ECHO ON
REM NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /USERS:99 /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /USERS:99 /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /USERS:99 /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /USERS:99 /UNLIMITED /CACHE:NONE
@ECHO OFF
goto :EOF

::
:3-LINDA-PC
REM ------------------------------------------------------------------------
ECHO ON
REM NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
@ECHO OFF
goto :EOF

::
:4-ASUS-GL522VW
REM ------------------------------------------------------------------------
ECHO ON
REM NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_REMOTE_1_SAMSUNG_4TB"="H:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_05_NVR_1000_DRIVE"="G:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_07_USB32_GB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_05_1SAMSUNG_4TB"="G:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_08_1_ELEMENTS_4TB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_1_SAMSUNG_4TB_C"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_05_1_SAMSUNG_4TB_D"="G:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_08_1_SAMSUNG_4TB_E"="H:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_10_1_ELEMENTS_4TB_HUBIC" /DELETE
REM NET SHARE "%NAME_VAR%_10_1_ELEMENT_4TB_HUBIC"="I:\1" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_10_1_ELEMENT_4TB"="I:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_09_2ELEMENTS4TB"="J:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_10_1_CAMEREA_MC_CARD"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

REM IF SHARE PATH NOT DELETE THEN MOST LIKELY THE DRIVE NAME TO THEM IS NOT THERE
REM USE THEN
REM SUBST G: E:\


@ECHO OFF
goto :EOF

::
:5-ASUS-P2520LA
REM ------------------------------------------------------------------------
ECHO ON
REM NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_REMOTE_2_SAMSUNG_4TB"="H:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
@ECHO OFF
goto :EOF

::
:7-ASUS-GL522VW
REM ------------------------------------------------------------------------
ECHO ON
REM NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

REM NET SHARE "%NAME_VAR%_04_HUBIC_2TB"="M:\2" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

REM NET SHARE "%NAME_VAR%_04_1_SAMSUNG_4TB_C"="K:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_05_1_SAMSUNG_4TB_D"="L:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_08_1_SAMSUNG_4TB_E"="M:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

REM NET SHARE "%NAME_VAR%_09_2_SAMSUNG_4TB_C"="K:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
rem NET SHARE "%NAME_VAR%_10_2_SAMSUNG_4TB_D"="O:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_11_2_SAMSUNG_4TB_E"="M:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE


NET SHARE "%NAME_VAR%_10_1_SAMSUNG_4TB_C" /DELETE
NET SHARE "%NAME_VAR%_10_1_SAMSUNG_4TB_D" /DELETE
NET SHARE "%NAME_VAR%_10_1_SAMSUNG_4TB_E" /DELETE
NET SHARE "%NAME_VAR%_20_2_SAMSUNG_4TB_C" /DELETE
NET SHARE "%NAME_VAR%_20_2_SAMSUNG_4TB_D" /DELETE
NET SHARE "%NAME_VAR%_20_2_SAMSUNG_4TB_E" /DELETE
NET SHARE "7_ASUS_GL522VW_34_4_SAMSUNG_2TB_CLOUD_2TB" /DELETE

@REM __ NET SHARE HAVE TO BE DELETE BEFORE MAKE AGAIN IF MAKE A CHANGE TO DRIVE LETTER USER

NET SHARE "%NAME_VAR%_10_1_SAMSUNG_4TB_C"="L:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_10_1_SAMSUNG_4TB_D"="M:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_10_1_SAMSUNG_4TB_E"="N:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_20_2_SAMSUNG_4TB_C"="P:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_20_2_SAMSUNG_4TB_D"="Q:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_20_2_SAMSUNG_4TB_E"="R:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

NET SHARE "%NAME_VAR%_30_4_SAMSUNG_4TB_HUBIC"="O:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

NET SHARE "%NAME_VAR%_40_1_SAMSUNG_2TB_CLOUD_2TB"="K:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_40_2_SAMSUNG_2TB_CLOUD_2TB"="S:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

NET SHARE "%NAME_VAR%_80_3_SAMSUNG_4TB_D"="T:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_80_4_SAMSUNG_4TB_D"="U:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE


REM NET SHARE "%NAME_VAR%_04_1_SAMSUNG_4TB_C" /DELETE
REM NET SHARE "%NAME_VAR%_05_1_SAMSUNG_4TB_D" /DELETE
REM NET SHARE "%NAME_VAR%_08_1_SAMSUNG_4TB_E" /DELETE

REM NET SHARE "%NAME_VAR%_28_4_SAMSUNG_4TB_HUBIC" /DELETE

@ECHO OFF
goto :EOF

::
:8-MSI-GP62M-7RD
REM ------------------------------------------------------------------------
ECHO ON
REM NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

REM NET SHARE "%NAME_VAR%_05_1_SAMSUNG_4TB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

NET SHARE "%NAME_VAR%_04_4_SAMSUNG_4TB_HUBIC"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_05_2_SAMSUNG_4TB_D"="G:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_08_2_SAMSUNG_4TB_E"="H:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

NET SHARE "%NAME_VAR%_10_4_SAMSUNG_4TB"="I:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE




@ECHO OFF
goto :EOF




REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------

GOTO ENDE
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------


REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
REM THE SETTER OF BLOCKED NET SHARE TO OUTSIDE WORLD

REM D:\>NET SHARE

REM Share name   Resource                        Remark

REM -------------------------------------------------------------------------------
REM IPC$                                         Remote IPC
REM G$           G:\                             Default share
REM E$           E:\                             Default share
REM ADMIN$       C:\Windows                      Remote Admin
REM C$           C:\                             Default share
REM D$           D:\                             Default share
REM F$           F:\                             Default share
REM Users        C:\Users

REM 7_ASUS_GL522VW_28_MY_DOCS_#_01_My_Documents_#_DATA$
             REM D:\VB6\VB-NT\00_Best_VB_01\Clip_Type\TEXTBOX BIGGER
                                             REM Caching disabled
REM 7_ASUS_GL522VW_21_Clipboard_Logger_#_DATA$
             REM D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA
                                             REM Caching disabled
REM 7_ASUS_GL522VW_22_MY_DOCS_#_01_My_Documents_#_DATA$
             REM D:\# MY DOCS\# 01 My Documents  
			 REM Caching disabled
REM 7_ASUS_GL522VW_23_MY_DOCS_#_01_My_Documents_#_DATA$
             REM D:\VIDEO\NOT\ME                 Caching disabled
REM 7_ASUS_GL522VW_23_VIDEO_NOT_ME$
             REM D:\VIDEO\NOT\ME                 Caching disabled
REM 7_ASUS_GL522VW_24_BT_Cloud_Sync_OLD$
             REM D:\BT Cloud Sync OLD            Caching disabled
REM 7_ASUS_GL522VW_25_0_00_Art_Loggers_#_APP_AND_SCREEN$
             REM D:\0 00 Art Loggers\# APP AND SCREEN
                                             REM Caching disabled
REM 7_ASUS_GL522VW_25_MY_DOCS_#_01_My_Documents_#_DATA$
             REM D:\0 00 Art Loggers\# APP AND SCREEN
                                             REM Caching disabled
REM 7_ASUS_GL522VW_26_MY_DOCS_#_01_My_Documents_#_DATA$
             REM D:\0 00 ART LOGGERS - WEBCAM    Caching disabled
REM 7_ASUS_GL522VW_27_VB_NT_00_Best_VB_01_Clip_Type_TEXTBOX$
             REM D:\VB6\VB-NT\00_Best_VB_01\Clip_Type\TEXTBOX
                                             REM Caching disabled
REM 7_ASUS_GL522VW_28_VB_NT_00_Best_VB_01_Clip_Type_TEXTBOX_BIGGER$
             REM D:\VB6\VB-NT\00_Best_VB_01\Clip_Type\TEXTBOX BIGGER
                                             REM Caching disabled
REM 7_ASUS_GL522VW_29_01_My_Documents_00_BlueTooth_SMS's_From_CLipType$
             REM D:\# MY DOCS\# 01 My Documents\00 BlueTooth SMS's From CLipType
                                             REM Caching disabled
REM 7_ASUS_GL522VW_29_MY_DOCS_#_01_My_Documents_#_DATA$
             REM D:\# MY DOCS\# 01 My Documents\00 BlueTooth SMS's From CLipType
                                             REM Caching disabled

											 REM 7_ASUS_GL522VW_01_C_DRIVE
             REM C:\                             Caching disabled
REM 7_ASUS_GL522VW_02_D_DRIVE
             REM D:\                             Caching disabled
REM 7_ASUS_GL522VW_03_FAT32_4GB
             REM E:\                             Caching disabled

REM The command completed successfully.
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------

:VERY_ENDER_BLOCK
:ENDE




