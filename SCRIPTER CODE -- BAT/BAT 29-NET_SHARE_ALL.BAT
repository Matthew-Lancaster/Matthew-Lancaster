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

REM -----------------------------------------------------------------
REM -- UPDATE ON 
REM -- TINY THING AND TIDY AFTER QUICKER
REM -----------------------------------------------------------------
REM -- Thu 08-Oct-2020 19:29:04
REM -- Thu 08-Oct-2020 22:59:00 -- 03 HOUR 29 MINUTE
REM -----------------------------------------------------------------


REM ------------------------------------------------------------------------
REM NET SHARE "USERS"="C:\USERS" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM USERS WAS SHARED SUCCESSFULLY.
REM ------------------------------------------------------------------------

REM ------------------------------------------------------------------------
REM [/CACHE:MANUAL | DOCUMENTS | PROGRAMS | BRANCHCACHE | NONE]
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
	@REM :: CHECK FOR ADMIN PRIVILEGES AGAIN
	>NUL 2>&1 "%SYSTEMROOT%\SYSTEM32\CACLS.EXE" "%SYSTEMROOT%\SYSTEM32\CONFIG\SYSTEM"
	@IF '%ERRORLEVEL%' GTR '2' SET NOT_GOT_ADMIN=TRUE
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

@REM :: CHECK FOR ADMIN PRIVILEGES
>NUL 2>&1 "%SYSTEMROOT%\SYSTEM32\CACLS.EXE" "%SYSTEMROOT%\SYSTEM32\CONFIG\SYSTEM"
@IF '%ERRORLEVEL%' GTR '2' (
	@REM GET ADMIN PRIVILEGES
	@ECHO Set UAC = CreateObject^("Shell.Application"^) >> "%TEMP%\getadmin.vbs"
	@ECHO UAC.ShellExecute "%~s0", "/ADMIN", "", "runas", 1 >> "%TEMP%\getadmin.vbs"

	"%TEMP%\GETADMIN.VBS"
	@DEL "%TEMP%\GETADMIN.VBS"
	GOTO ENDE
	@EXIT /B
)
@ECHO GOT ADMIN PRIVILEGES


ECHO -----------------------------------------------------------------------
ECHO -----------------------------------------------------------------------
ECHO -----------------------------------------------------------------------
ECHO -----------------------------------------------------------------------

REM ------------------------------------------------------------------------
REM GOTO JUMP_OVER
REM ------------------------------------------------------------------------
REM -- WIN XP DON'T APPEAR TO USE COMPUTER NAME AT ALL
REM -- AND STILL WORK WITH IT IN COMMAND LINE
REM -- BUT AND THEN NOT HARMFUL TO KEEP IT IN THE COMMAND LINE
REM ------------------------------------------------------------------------
@D:
@CD\


@REM ---------------------------------------------------------------
@ECHO OFF
@REM ---------------------------------------------------------------
SET WINXP__=NOT
SET WIN07__=NOT
SET WIN10__=NOT
SET WIN11__=NOT
VER | FIND "5.1." > NUL
IF %ERRORLEVEL% EQU 0 SET WINXP__=YES
VER | FIND "6.1." > NUL
IF %ERRORLEVEL% EQU 0 (
	SET WIN07__=YES
	SET WIN10__=YES
)
VER | FIND "10.0.19045" > NUL
IF %ERRORLEVEL% EQU 0 (
	SET WIN10__=YES
	SET WIN07__=YES
)
VER | FIND "10.0.26200" > NUL
IF %ERRORLEVEL% EQU 0 (
	SET WIN11__=YES
)
IF %COMPUTERNAME%==9-ASUS-G815LM   SET WIN11__=YES

@REM ---------------------------------------------------------------

@REM ---------------------------------------------------------------

IF "%WIN07__%"=="YES" SET SET WINXP__=NOT

IF "%WIN10__%"=="YES" SET SET WINXP__=NOT
IF "%WIN10__%"=="YES" SET WIN07__=NOT

IF "%WIN11__%"=="YES" SET WINXP__=NOT
IF "%WIN11__%"=="YES" SET WIN07__=NOT
IF "%WIN11__%"=="YES" SET WIN10__=YES


IF "%WINXP__%"=="YES" ECHO WINDOWS XP
IF "%WIN07__%"=="YES" ECHO WINDOWS 7
IF "%WIN10__%"=="YES" ECHO WINDOWS 10
IF "%WIN11__%"=="YES" ECHO WINDOWS 11



REM SETLOCAL ENABLEDELAYEDEXPANSION

REM SET "PATTERN=-"
REM SET "REPLACE=_"

REM Set "COMPUTERNAME_DASH_CONVERT_VAR=%COMPUTERNAME%"
REM REM %%~a"
REM Ren "%%a" "!File:%Pattern%=%Replace%!"
	
REM -- ----
REM -- CMD WINDOWS REPLACE SPACES WITH UNDERSCORES - STACK OVERFLOW
REM -- http://stackoverflow.com/questions/20791264/cmd-windows-replace-spaces-with-underscores/37853017#37853017
REM -- ----	
SETLOCAL ENABLEDELAYEDEXPANSION
SET COMPUTERNAME_DASH_CONVERT_VAR=%COMPUTERNAME%
SET COMPUTERNAME_DASH_CONVERT_VAR=%COMPUTERNAME:-=_%
SET NAME_VAR=%COMPUTERNAME:-=_%

CALL :UPCase NAME_VAR

IF %COMPUTERNAME%==1-ASUS-X5DIJ     SET NAME_VAR=1_Asus_X5DIJ
IF %COMPUTERNAME%==2_ASUS-EEE       SET NAME_VAR=2_Asus_EEE
IF %COMPUTERNAME%==3-LINDA-PC       SET NAME_VAR=3_LINDA_PC
IF %COMPUTERNAME%==4-ASUS-GL522VW   SET NAME_VAR=4_Asus_Gl522Vw
IF %COMPUTERNAME%==5-ASUS-P2520LA   SET NAME_VAR=5_Asus_P2520la
IF %COMPUTERNAME%==7-ASUS-GL522VW   SET NAME_VAR=7_Asus_Gl522Vw
IF %COMPUTERNAME%==8-MSI-GP62M-7RD  SET NAME_VAR=8_Msi_Gp62m_7rd
IF %COMPUTERNAME%==9-ASUS-G815LM    SET NAME_VAR=9_Asus_G815LM




REM ------------------------------------------------------------------------
REM ---- GLOBAL DELETE
CALL :COMPUTER_NAME_COMMON_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
REM IF %COMPUTERNAME%==1-ASUS-X5DIJ     CALL :1-ASUS-X5DIJ_NET_SHARE_DELETE
REM IF %COMPUTERNAME%==2-ASUS-EEE       CALL :2-ASUS-EEE_NET_SHARE_DELETE
REM IF %COMPUTERNAME%==3-LINDA-PC       CALL :3-LINDA-PC_NET_SHARE_DELETE
REM IF %COMPUTERNAME%==4-ASUS-GL522VW   CALL :4-ASUS-GL522VW_NET_SHARE_DELETE
REM IF %COMPUTERNAME%==5-ASUS-P2520LA   CALL :5-ASUS-P2520LA_NET_SHARE_DELETE
REM IF %COMPUTERNAME%==7-ASUS-GL522VW   CALL :7-ASUS-GL522VW_NET_SHARE_DELETE
IF %COMPUTERNAME%==8-MSI-GP62M-7RD  CALL :8-MSI-GP62M-7RD_NET_SHARE_DELETE
IF %COMPUTERNAME%==9-ASUS-G815LM    CALL :9-ASUS-G815LM_NET_SHARE_DELETE


REM ------------------------------------------------------------------------
:JUMP_OVER
REM ------------------------------------------------------------------------

REM ------------------------------------------------------------------------
REM -- SET THE NET SHARE SET
REM ------------------------------------------------------------------------
CALL :COMPUTER_NAME_COMMON_NET_SHARE_CREATE
REM ------------------------------------------------------------------------
IF %COMPUTERNAME%==2-ASUS-EEE       CALL :2-ASUS-EEE
IF %COMPUTERNAME%==1-ASUS-X5DIJ     CALL :1-ASUS-X5DIJ
IF %COMPUTERNAME%==3-LINDA-PC       CALL :3-LINDA-PC
IF %COMPUTERNAME%==4-ASUS-GL522VW   CALL :4-ASUS-GL522VW
IF %COMPUTERNAME%==5-ASUS-P2520LA   CALL :5-ASUS-P2520LA
IF %COMPUTERNAME%==7-ASUS-GL522VW   CALL :7-ASUS-GL522VW
IF %COMPUTERNAME%==8-MSI-GP62M-7RD  CALL :8-MSI-GP62M-7RD
IF %COMPUTERNAME%==9-ASUS-G815LM    CALL :9-ASUS-G815LM
REM ------------------------------------------------------------------------

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

REM ------------------------------------------------------------------------
REM -- A NETWORK DRIVE WILL NOT DELETE UNLESS IT CONNECT AR
REM -- AND THEN SUBST SOME CRAP THERE AND DO AGAIN
REM ------------------------------------------------------------------------

::
:COMPUTER_NAME_COMMON_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
ECHO OFF

REM IS CODE TO DELETE ALL NETWORK SHARE BAT CMD

REM ------------------------------------------
REM -- HERE ONLY WORK FOR MAPPED NETWORK DRIVE
REM ------------------------------------------
REM NET USE * /DELETE



NET SHARE "USERS" \\%COMPUTERNAME% /DELETE



REM SET FILE_NAME="%~dp0%~n0.PS1"
SET FILE_NAME="C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 05 - DELETE ALL NET-SHARE.PS1"

START "" /WAIT /B "powershell" -executionpolicy bypass -File %FILE_NAME%

REM START "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 05 - DELETE ALL NET-SHARE.BAT"


GOTO SKIP_HERE


setlocal EnableDelayedExpansion

REM DON'T WORK
@ECHO OFF
@FOR %%G IN (C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (
	IF EXIST %%G:\. (
		SET "TARGET_="
		SET TARGET_=%%G:
		CALL :IsDeviceReady !TARGET_! isready_
		IF !isready_!==true (
			REM ECHO DELETE NETWORK DRIVE !TARGET_!
			ECHO ----------------------------------------------------
			ECHO NET USE !TARGET_! /DELETE
			ECHO ----------------------------------------------------
			net use !TARGET_! /delete

		)
	)
)





REM -- THIS OPEN DOES HER
REM ----------------------------------------
REM Deleting all network shares from a Windows system in Batch - Stack Overflow
REM ----------------------------------------
REM https://stackoverflow.com/questions/40142759/deleting-all-network-shares-from-a-windows-system-in-batch
REM ----------------------------------------


REM -- DON'T WORK WINDOWS 11
IF "%WIN11__%"=="YES" GOTO SKIP_HERE
REM 'WMIC' is not recognized as an internal or external command,
REM IF "%WIN11__%"=="YES" GOTO SKIP_HERE

GOTO SKIP_HERE
REM -- POWERSHELL DO THE JOB NOW


@Echo Off
For /F "Tokens=2* Delims==;" %%A In ('WMIC Class StdRegProv Call EnumValues^
 "&H80000002"^, "System\CurrentControlSet\Services\LanmanServer\Shares"
 ^|Find "sNames"') Do Set "_=%%A"
For %%A In (%_:~2,-1%) Do (
ECHO %%A
REM Net Share %%A /Delete
)


:SKIP_HERE


REM PAUSE


REM SOD
REM setlocal DisableDelayedExpansion
REM for /F "usebackq delims=" %%S in ("%APPDATA%\shares2.txt") do (
    REM set "tempy=%%S"
    REM setlocal EnableDelayedExpansion
    REM for /F "tokens=1 delims=|" %%N in (^""!tempy:    =|!"^") do (
        REM endlocal
        REM REM net share "%%~N" /delete
		REM ECHO HERE %%~N
		REM ECHO %TEMPY%
		
    REM )
REM )
REM endlocal





REM SOD
REM ----------------------------------------
REM windows - Script to disconnect mapped network location - Super User
REM ----------------------------------------
REM https://superuser.com/questions/595412/script-to-disconnect-mapped-network-location
REM ----------------------------------------
REM ECHO ON
REM for /f "tokens=1,2,3,4" %%D in ('net share ^| findSTRING "C:") do (
    
	REM ECHO %%D
	REM ECHO %%E
	
    REM if %%D==Z: (
        REM REM
    REM ) else (
        REM REM
    REM )
 
REM )
REM PAUSE

@ECHO OFF


@REM --------------------------------------------

REM NET SHARE "%NAME_VAR%_01_C_Drive" /DELETE
REM NET SHARE "%NAME_VAR%_02_D_Drive" /DELETE
REM NET SHARE "%NAME_VAR%_03_E_Drive" /DELETE
REM NET SHARE "%NAME_VAR%_03_fat32_4gb" /DELETE

REM NET SHARE "MY MUSIC" \\%COMPUTERNAME% /DELETE
REM NET SHARE "MUSIC" \\%COMPUTERNAME% /DELETE
REM NET SHARE "PICTURES" \\%COMPUTERNAME% /DELETE
REM NET SHARE "MY PICTURES" \\%COMPUTERNAME% /DELETE
REM NET SHARE "MY VIDEO" \\%COMPUTERNAME% /DELETE
REM NET SHARE "USERS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "SHAREDDOCS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "C$" \\%COMPUTERNAME% /DELETE
REM NET SHARE "D$" \\%COMPUTERNAME% /DELETE
REM NET SHARE "E$" \\%COMPUTERNAME% /DELETE
REM NET SHARE "F$" \\%COMPUTERNAME% /DELETE
REM NET SHARE "G$" \\%COMPUTERNAME% /DELETE

REM NET SHARE "PRINT$" \\%COMPUTERNAME% /DELETE
REM NET SHARE "Admin$" \\%COMPUTERNAME% /DELETE
REM NET SHARE Admin$ /DELETE
REM NET SHARE "IPC$" /DELETE
REM ------------------------------------------------------------------------
REM IF SET HERE TO DELETE AND PUT BACK HAVE TO REBOOT MACHINE TO GAIN ACCESS
REM GAIN ACCESS THE REMOTE CALLER COMPUTER WONT HAVE ANY UNTIL REBOOT
REM ------------------------------------------------------------------------
REM --------------------------------------------------------------------
REM DIFFICULT TO DO DELETE IF USER ARE CONNECTED THEN ASK A QUESTION Y/N
REM CAN'T SEND INPUT TO ANSWER
REM --------------------------------------------------------------------
REM NOT HAVE TO INCLUDE COMPUTER NAME WHEN DELETER
REM --------------------------------------------------------------------
REM NET SHARE "%NAME_VAR%_09_2_ELEMENTS_4TB" /DELETE
REM NET SHARE "4_asus_gl522vw_08_1ELEMENTS4TB" \\%COMPUTERNAME% /DELETE
REM NET SHARE I:\ /DELETE
REM NET SHARE "%NAME_VAR%_04_2_Samsung_4TB_C" /DELETE

REM --------------------------------------------------------------------


GOTO :EOF
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------




:COMPUTER_NAME_COMMON_NET_SHARE_CREATE
REM ---------------------------------------------------------------------------
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

REM ---------------------------------------------------------------------------
REM WIN07__ ____ ARE LINK WITH AR ____ WIN10__
REM ---------------------------------------------------------------------------

IF "%WINXP__%"=="YES" GOTO WIN_XP_SET
IF "%WIN10__%"=="YES" GOTO WIN_10_SET
IF "%WIN11__%"=="YES" GOTO WIN_10_SET

GOTO END_IT
:WIN_XP_SET
	REM NET SHARE IPC$
	REM NET SHARE ADMIN$ /USERS:99 /UNLIMITED /CACHE:NONE
	NET SHARE "%NAME_VAR%_C_Drive"="C:" /USERS:99 /UNLIMITED /CACHE:NONE
	NET SHARE "%NAME_VAR%_D_Drive"="D:" /USERS:99 /UNLIMITED /CACHE:NONE
	NET SHARE "%NAME_VAR%_E_Drive"="E:" /USERS:99  /UNLIMITED /CACHE:NONE
	IF "%WIN10__%"=="YES" GOTO WIN_10_SET
GOTO END_IT

:WIN_10_SET
	
	REM NET SHARE IPC$
	REM NET SHARE ADMIN$ /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	
	NET SHARE "%NAME_VAR%_C_Drive"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	NET SHARE "%NAME_VAR%_D_Drive"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	NET SHARE "%NAME_VAR%_E_Drive"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	NET SHARE "Users"="C:\Users" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

	
	REM -------------------------------------------------------------
	REM Step 2: Add switches to your script
	REM You can put switches to your script for different options, such as:
	REM /delete - to delete a mapped drive
	REM /persistent:yes/no - to reconnect the drive upon login
	REM -------------------------------------------------------------

	REM NET USE W: \\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE /persistent:yes
	REM NET USE X: \\1-ASUS-X5DIJ\1_ASUS_X5DIJ_02_D_DRIVE /persistent:yes
	REM NET USE Y: \\1-ASUS-X5DIJ\1_ASUS_X5DIJ_03_FAT32_4GB /persistent:yes

	REM NET USE W: \\2-ASUS-EEE\2_ASUS_EEE_01_C_DRIVE /persistent:yes
	REM NET USE X: \\2-ASUS-EEE\2_ASUS_EEE_02_D_DRIVE /persistent:yes
	REM NET USE Y: \\2-ASUS-EEE\2_ASUS_EEE_03_FAT32_4GB /persistent:yes
	
	REM NET USE W: \\5-ASUS-P2520LA\5_ASUS_P2520LA_01_C_DRIVE /persistent:yes
	REM NET USE X: \\5-ASUS-P2520LA\5_ASUS_P2520LA_02_D_DRIVE /persistent:yes
	REM NET USE Y: \\5-ASUS-P2520LA\5_ASUS_P2520LA_03_FAT32_4GB /persistent:yes
	
	

	REM USERS        C:\USERS

	
	REM NET SHARE "CLIPBOARD_LOGGER_#_DATA$"="D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "ART_LOGGERS_-_WEBCAM$"="D:\0 00 ART LOGGERS - WEBCAM" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "ART_LOGGERS_#_APP_AND_SCREEN$"="D:\0 00 Art Loggers\# APP AND SCREEN" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "MY_DOCS_#_01_MY_DOCUMENTS$"="D:\# MY DOCS\# 01 My Documents" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "VB_NT_00_BEST_VB_01_CLIP_TYPE_TEXTBOX$"="D:\VB6\VB-NT\00_Best_VB_01\Clip_Type\TEXTBOX" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "VB_NT_00_BEST_VB_01_CLIP_TYPE_TEXTBOX_BIGGER$"="D:\VB6\VB-NT\00_Best_VB_01\Clip_Type\TEXTBOX BIGGER" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
	REM NET SHARE "MY_DOCUMENTS_00_BLUETOOTH_SMS'S_FROM_CLIPTYPE$"="D:\# MY DOCS\# 01 My Documents\00 BlueTooth SMS's From CLipType" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
:END_IT
GOTO :EOF





::
:9-ASUS-G815LM_NET_SHARE_DELETE

REM ------------------------------------------------------------------------
REM NET SHARE "D DRIVE" \\%COMPUTERNAME% /DELETE
REM NET SHARE "START MENU" \\%COMPUTERNAME% /DELETE
REM NET SHARE "STARTUP" \\%COMPUTERNAME% /DELETE
REM NET SHARE "# 01 MY DOCUMENTS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "MUSIC" \\%COMPUTERNAME% /DELETE
REM NET SHARE "PICTURES" \\%COMPUTERNAME% /DELETE
REM NET SHARE "USERS" \\%COMPUTERNAME% /DELETE

REM NET SHARE "SHAREDDOCS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "USERS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "ProgramData" C:\ProgramData /DELETE

REM NET SHARE "9_ASUS_G815LM_01_C_DRIVE" \\%COMPUTERNAME% /DELETE
REM NET SHARE "9_ASUS_G815LM_02_D_DRIVE" \\%COMPUTERNAME% /DELETE
REM NET SHARE "9_ASUS_G815LM_03_FAT32_4GB" \\%COMPUTERNAME% /DELETE

REM NET SHARE "%NAME_VAR%_10_1_ELEMENT_4TB" \\%COMPUTERNAME% /DELETE
REM NET SHARE "%NAME_VAR%_04_1_Samsung_4TB_C" \\%COMPUTERNAME% /DELETE
REM NET SHARE "%NAME_VAR%_05_1_Samsung_4TB_D" \\%COMPUTERNAME% /DELETE

REM MUSIC        D:\ON-LINE-SYNC\ONEDRIVE\MUSIC
REM PICTURES     D:\ON-LINE-SYNC\ONEDRIVE\PICTURES
REM USERS        C:\USERS

GOTO :EOF
REM ------------------------------------------------------------------------





REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
REM -- THE DELETE NET SHARE SET 1 TO 7 ROUTINE 1 MISSING ---- 2ND OF 2
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------

::
:1-ASUS-X5DIJ_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
REM NET SHARE "SHAREDDOCS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "PRINTER" \\%COMPUTERNAME% /DELETE
REM NET SHARE "PRINTER4" \\%COMPUTERNAME% /DELETE
GOTO :EOF
REM ------------------------------------------------------------------------

::
:2-ASUS-EEE_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
REM NET SHARE "ASUS-EEE E-WIN 7" \\%COMPUTERNAME% /DELETE
REM NET SHARE "ASUS-EEE E-WIN 7" \\%COMPUTERNAME% /DELETE
REM NET SHARE "ASUS-EEE WD-55GB" \\%COMPUTERNAME% /DELETE
REM NET SHARE "SHAREDDOCS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "PRINTER2" \\%COMPUTERNAME% /DELETE
REM NET SHARE "PRINTER4" \\%COMPUTERNAME% /DELETE
GOTO :EOF
REM ------------------------------------------------------------------------

::
:3-LINDA-PC_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
REM ECHO ON
REM NET SHARE "MY MUSIC" \\%COMPUTERNAME% /DELETE
REM NET SHARE "MY MUSIC2" \\%COMPUTERNAME% /DELETE
REM NET SHARE "MY PICTURES" \\%COMPUTERNAME% /DELETE
REM NET SHARE "PICTURES2" \\%COMPUTERNAME% /DELETE
REM NET SHARE "MY VIDEO" \\%COMPUTERNAME% /DELETE
REM NET SHARE "MY VIDEO2" \\%COMPUTERNAME% /DELETE
REM NET SHARE "SHAREDDOCS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "USERS" \\%COMPUTERNAME% /DELETE

REM MY MUSIC     D:\# MY DOCS\# 01 MY DOCUMENTS\MY MUSIC
REM MY MUSIC2    D:\# MY DOCS COMMON\MY MUSIC
REM MY PICTURES  D:\# MY DOCS\# 01 MY DOCUMENTS\MY PICTURES
REM MY PICTURES2 D:\# MY DOCS COMMON\MY PICTU...
REM MY VIDEO     D:\# MY DOCS\# 01 MY DOCUMENTS\MY VIDEO
REM MY VIDEO2    D:\# MY DOCS COMMON\MY VIDEO
REM SHAREDDOCS   C:\USERS\PUBLIC\DOCUMENTS
REM USERS        C:\USERS

GOTO :EOF
REM ------------------------------------------------------------------------

::
:4-ASUS-GL522VW_NET_SHARE_DELETE

@ECHO OFF
REM ------------------------------------------------------------------------

REM NET SHARE "SHAREDDOCS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "USERS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "ProgramData" C:\ProgramData /DELETE

REM NET SHARE "%NAME_VAR%_05_NVR_1000_Drive" \\%COMPUTERNAME% /DELETE

REM NET SHARE "%NAME_VAR%_10_1_ELEMENT_4TB" \\%COMPUTERNAME% /DELETE
REM NET SHARE "%NAME_VAR%_04_1_Samsung_4TB_C" \\%COMPUTERNAME% /DELETE
REM NET SHARE "%NAME_VAR%_05_1_Samsung_4TB_D" \\%COMPUTERNAME% /DELETE


REM NET SHARE "4_asus_gl522vw_44_ADATA_HD710_PRO_4TB" \\%COMPUTERNAME% /DELETE
REM NET SHARE "4_asus_gl522vw_40_ADATA_HD710_PRO_4TB" \\%COMPUTERNAME% /DELETE

REM NET SHARE "4_asus_gl522vw_82_2_MC_CARD_CAMERA" \\%COMPUTERNAME% /DELETE
REM NET SHARE "4_asus_gl522vw_82_1_MC_CARD_CAMERA" \\%COMPUTERNAME% /DELETE

REM NET SHARE "V:\" \\%COMPUTERNAME% /DELETE
REM NET SHARE "V:\" /DELETE
REM NET SHARE "V:" /DELETE
REM NET SHARE "V:" \\%COMPUTERNAME% /DELETE

REM NET SHARE "4_asus_gl522vw_82_2_MC_CARD_CAMERA" Y:\ /DELETE

REM SUBST G: D:\

REM SHAREDDOCS   C:\USERS\PUBLIC\DOCUMENTS
REM USERS        C:\USERS

GOTO :EOF
REM ------------------------------------------------------------------------

::
:5-ASUS-P2520LA_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
REM -- ECHO ON
REM -- BOTH ON K DRIVE AND DRIVE DON'T EXIST -- AND BUT WON'T DELETE
REM -- USE WORK-AROUND AND SUBST WON'T DO IT
REM -- CREATE A PARTITION WITH PART-WIZARD 
REM -- AND AND NOT AT ROOT THE PATH HAS TO BE MADE FOR IT ABLE TO DELETE
REM -- ECHO OFF

REM NET SHARE "5-ASUS-P2520L-2SAMASUNG4TB" \\%COMPUTERNAME% /DELETE
REM NET SHARE "#C" \\%COMPUTERNAME% /DELETE
REM NET SHARE "MUSIC" \\%COMPUTERNAME% /DELETE
REM NET SHARE "PICTURES" \\%COMPUTERNAME% /DELETE
REM NET SHARE "USERS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "SHAREDDOCS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "01 BANGING TUNES" \\%COMPUTERNAME% /DELETE
REM NET SHARE "2015 2016" \\%COMPUTERNAME% /DELETE
REM NET SHARE "5-ASUS-P2520LA_REMOTE_2-SamSung_4TB" \\%COMPUTERNAME% /DELETE
REM NET SHARE "5-ASUS-P2520LA_REMOTE_2-SamSung_4TB" \\%COMPUTERNAME% /DELETE

REM #C                                   K:\#\#C
REM 01 BANGING TUNES                     D:\0 00 MUSIC ---\04 MY MUSIC\01 BANGING TUNES
REM 2015 2016                            D:\DSC\2015 2016
REM 5-ASUS-P2520L FAT32-4TB              F:\
REM 5-ASUS-P2520L-2SAMASUNG4TB           K:\
REM 5-ASUS-P2520LA_01_C-DRIVE            C:\
REM 5-ASUS-P2520LA_02_D-DRIVE            D:\
REM 5-ASUS-P2520LA_03_FAT32              E:\
REM 5-ASUS-P2520LA_REMOTE_2-SamSung_4TB  H:\
REM SHAREDDOCS                           C:\USERS\PUBLIC\DOCUMENTS

GOTO :EOF
REM ------------------------------------------------------------------------

::
:7-ASUS-GL522VW_NET_SHARE_DELETE
REM ------------------------------------------------------------------------

REM NET SHARE "MUSIC" \\%COMPUTERNAME% /DELETE
REM NET SHARE "PICTURES" \\%COMPUTERNAME% /DELETE
REM NET SHARE "USERS" \\%COMPUTERNAME% /DELETE

REM 7-ASUS-GL522VW_01_C-DRIVE        C:\
REM 7-ASUS-GL522VW_02_D-DRIVE        D:\
REM 7-ASUS-GL522VW_03_FAT32          E:\
REM MUSIC        D:\ON-LINE-SYNC\ONEDRIVE\MUSIC
REM PICTURES     D:\ON-LINE-SYNC\ONEDRIVE\PICTURES
REM USERS        C:\USERS

GOTO :EOF
REM ------------------------------------------------------------------------

::
:8-MSI-GP62M-7RD_NET_SHARE_DELETE
REM ------------------------------------------------------------------------

@ECHO OFF
REM NET SHARE "D DRIVE" \\%COMPUTERNAME% /DELETE
REM NET SHARE "START MENU" \\%COMPUTERNAME% /DELETE
REM NET SHARE "STARTUP" \\%COMPUTERNAME% /DELETE
REM NET SHARE "# 01 MY DOCUMENTS" \\%COMPUTERNAME% /DELETE
REM NET SHARE "MUSIC" \\%COMPUTERNAME% /DELETE
REM NET SHARE "PICTURES" \\%COMPUTERNAME% /DELETE
REM NET SHARE "USERS" \\%COMPUTERNAME% /DELETE

REM NET SHARE "IPC$" \\%COMPUTERNAME% /DELETE

REM MUSIC        D:\ON-LINE-SYNC\ONEDRIVE\MUSIC
REM PICTURES     D:\ON-LINE-SYNC\ONEDRIVE\PICTURES
REM USERS        C:\USERS

GOTO :EOF
REM ------------------------------------------------------------------------


REM ------------------------------------------------------------------------
REM ---- CREATE __ 2ND
REM ------------------------------------------------------------------------








::
:2-ASUS-EEE
REM ---------------------------------------------------------------------------
REM -- WIN XP DON'T USE THIS ____ /GRANT:EVERYONE,FULL
REM -- BUT AND DOES USE THIS ____ /USERS:99 
REM -- IN THE GUI SETTING IT DON'T APPEAR TO USE UNLIMITED ACCESS IN ALL VERSION WIN AND SET MAX AT 20 USER
REM -- WITH THIS EXTRA SETTING AT 99 USER
REM -- IN THE DOS COMMAND NET SHARE AND THE COMPUTER NAME IN THE FORMAT SHOWN IT REPORT THE SHARE HAS MAXIMUM USERS - NOT ANY LIMIT
REM ---------------------------------------------------------------------------
REM NET SHARE "%NAME_VAR%_01_C_Drive"="C:" /USERS:99 /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_02_D_Drive"="D:" /USERS:99 /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_03_E_Drive"="E:" /USERS:99  /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /USERS:99 /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_MC_SD_16GB"="H:" /USERS:99  /UNLIMITED /CACHE:NONE
REM ---------------------------------------------------------------------------
REM NET SHARE "%NAME_VAR%_WD-55GB"  /USERS:99  /UNLIMITED /CACHE:NONE
GOTO :EOF

::
:1-ASUS-X5DIJ
REM ---------------------------------------------------------------------------
REM -- WIN XP DON'T USE THIS ____ /GRANT:EVERYONE,FULL
REM -- BUT AND DOES USE THIS ____ /USERS:99 
REM -- IN THE GUI SETTING IT DON'T APPEAR TO USE UNLIMITED ACCESS IN ALL VERSION WIN AND SET MAX AT 20 USER
REM -- WITH THIS EXTRA SETTING AT 99 USER
REM -- IN THE DOS COMMAND NET SHARE AND THE COMPUTER NAME IN THE FORMAT SHOWN IT REPORT THE SHARE HAS MAXIMUM USERS - NOT ANY LIMIT
REM ---------------------------------------------------------------------------
REM NET SHARE "%NAME_VAR%_01_C_Drive"="C:" /USERS:99 /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_02_D_Drive"="D:" /USERS:99 /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_03_E_Drive"="E:" /USERS:99 /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /USERS:99 /UNLIMITED /CACHE:NONE
GOTO :EOF

::
:3-LINDA-PC
REM ------------------------------------------------------------------------
ECHO ON
REM NET SHARE "%NAME_VAR%_01_C_Drive"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_02_D_Drive"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_03_E_Drive"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
@ECHO OFF
GOTO :EOF

::
:4-ASUS-GL522VW

@ECHO OFF
REM ------------------------------------------------------------------------

REM NET SHARE "%NAME_VAR%_04_1_Samsung_4TB_C"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_05_1_Samsung_4TB_D"="G:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

REM NET SHARE "%NAME_VAR%_10_1_ELEMENT_4TB"="H:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

REM -----------------------------------------------------------------
REM -----------------------------------------------------------------
REM 20-MAY-2026 23:15:01 WED
REM IF USE A MIXTURE OF LCASE UCASE THEN THE CORRECT CASE WILL FOLLOW
REM -----------------------------------------------------------------

ECHO ON

NET SHARE "%NAME_VAR%_I_Samsung_5TB_C"="I:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_J_Samsung_5TB_D"="J:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_K_Samsung_5TB_E"="K:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

NET SHARE "%NAME_VAR%_O_Samsung_4TB"="O:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_P_Samsung_5TB"="P:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

NET SHARE "%NAME_VAR%_U_Samsung_5TB"="u:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_V_Samsung_5TB"="v:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

@ECHO OFF
GOTO :EOF

::
:5-ASUS-P2520LA
REM ------------------------------------------------------------------------
@ECHO OFF
REM NET SHARE "%NAME_VAR%_01_C_Drive"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_02_D_Drive"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_03_E_Drive"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_REMOTE_2_Samsung_4TB"="H:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
@ECHO OFF
GOTO :EOF

::
:7-ASUS-GL522VW
REM ------------------------------------------------------------------------
@ECHO OFF
REM NET SHARE "%NAME_VAR%_10_1_Samsung_4TB_C" /DELETE
REM NET SHARE "%NAME_VAR%_10_1_Samsung_4TB_D" /DELETE
REM NET SHARE "%NAME_VAR%_10_1_Samsung_4TB_E" /DELETE
REM NET SHARE "%NAME_VAR%_20_2_Samsung_4TB_C" /DELETE
REM NET SHARE "%NAME_VAR%_20_2_Samsung_4TB_D" /DELETE
REM NET SHARE "%NAME_VAR%_20_2_Samsung_4TB_E" /DELETE
REM NET SHARE "7_asus_gl522vw_34_4_Samsung_2TB_CLOUD_2TB" /DELETE

REM NET SHARE "%NAME_VAR%_10_1_Samsung_4TB_C"="L:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_10_1_Samsung_4TB_D"="M:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_10_1_Samsung_4TB_E"="N:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_20_2_Samsung_4TB_C"="P:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_20_2_Samsung_4TB_D"="Q:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_20_2_Samsung_4TB_E"="R:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_30_4_Samsung_4TB_HUBIC"="O:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_40_1_Samsung_2TB_CLOUD_2TB"="K:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_40_2_Samsung_2TB_CLOUD_2TB"="S:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_80_3_Samsung_4TB_D"="T:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_80_4_Samsung_4TB_D"="U:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE


REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_HUBIC_2TB"="M:\2" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_1_Samsung_4TB_C"="K:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_05_1_Samsung_4TB_D"="L:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_08_1_Samsung_4TB_E"="M:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_09_2_Samsung_4TB_C"="K:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
rem NET SHARE "%NAME_VAR%_10_2_Samsung_4TB_D"="O:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_11_2_Samsung_4TB_E"="M:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
@REM __ NET SHARE HAVE TO BE DELETE BEFORE MAKE AGAIN IF MAKE A CHANGE TO DRIVE LETTER USER
REM NET SHARE "%NAME_VAR%_04_1_Samsung_4TB_C" /DELETE
REM NET SHARE "%NAME_VAR%_05_1_Samsung_4TB_D" /DELETE
REM NET SHARE "%NAME_VAR%_08_1_Samsung_4TB_E" /DELETE
REM NET SHARE "%NAME_VAR%_28_4_Samsung_4TB_HUBIC" /DELETE

@ECHO OFF
GOTO :EOF

::
:8-MSI-GP62M-7RD
REM ------------------------------------------------------------------------
@ECHO OFF
REM NET SHARE "%NAME_VAR%_04_4_Samsung_4TB_HUBIC"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_05_2_Samsung_4TB_D"="G:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_08_2_Samsung_4TB_E"="H:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_10_4_Samsung_4TB"="I:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

REM ---------------------------------------------------
REM -- HERE IS BETTER TO HAVE NETWORK SHARE FOR MC-CARD
REM -- ONE HAS BEEN ALL TIME
REM -- 3 COMPUTER HAVE SHARE
REM -- AND MC-CARD IN ONE CASE IS ONLY EVER ONE
REM -- STILL REQUIRE 3 SHARE WHEN MC-CARD NOT THERE
REM -- OR MISTAKE HAPPEN
REM ---------------------------------------------------
REM -- 08-OCT-2020 22:12:41 THU
REM ---------------------------------------------------
REM -- HERE ONE LITTLE NET SHARE SET AND COUPLE TIDY ONE COMPUTER
REM -- DAY WATER MOTION FRICTION MODE
REM ---------------------------------------------------
REM -- AND NEAR 3 HOUR TO ONE TINY SMALL WORK 
REM -- I MIGHT AS WELL GONE OUT SOMEWHERE THAN ENJOY COMPUTER
REM ---------------------------------------------------
REM -- THU 08-OCT-2020 19:29:04
REM -- THU 08-OCT-2020 22:15:00 -- 2 HOUR 45 MINUTE
REM ---------------------------------------------------
  
REM NET SHARE "%NAME_VAR%_01_C_Drive"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_02_D_Drive"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_03_E_Drive"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_04_USB32_GB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM NET SHARE "%NAME_VAR%_05_1_Samsung_4TB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE


@ECHO OFF
GOTO :EOF


::
:9-ASUS-G815LM
REM ------------------------------------------------------------------------
@ECHO OFF


@ECHO OFF
GOTO :EOF


REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
GOTO ENDE
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------

REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
REM THE SETTER OF BLOCKED NET SHARE TO OUTSIDE OUTDOOR
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------

D:\>NET SHARE

SHARE NAME   RESOURCE                        REMARK
-------------------------------------------------------------------------------
IPC$                                         Remote IPC
G$           G:\                             Default share
E$           E:\                             Default share
ADMIN$       C:\Windows                      Remote Admin
C$           C:\                             Default share
D$           D:\                             Default share
F$           F:\                             Default share
Users        C:\Users



REM THE COMMAND COMPLETED SUCCESSFULLY.
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------

:VERY_ENDER_BLOCK
:ENDE

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
EXIT


@ECHO OFF
SET String=Hello, how are you ?
SET String
CALL :TCase String
SET String
CALL :UpCase String
SET String
CALL :LoCase String
SET String
GOTO:EOF

:LoCase
:: Subroutine to convert a variable VALUE to all lower case.
:: The argument for this subroutine is the variable NAME.
FOR %%i IN ("A=a" "B=b" "C=c" "D=d" "E=e" "F=f" "G=g" "H=h" "I=i" "J=j" "K=k" "L=l" "M=m" "N=n" "O=o" "P=p" "Q=q" "R=r" "S=s" "T=t" "U=u" "V=v" "W=w" "X=x" "Y=y" "Z=z") DO CALL SET "%1=%%%1:%%~i%%"
GOTO:EOF

:UpCase
:: Subroutine to convert a variable VALUE to all UPPER CASE.
:: The argument for this subroutine is the variable NAME.
FOR %%i IN ("a=A" "b=B" "c=C" "d=D" "e=E" "f=F" "g=G" "h=H" "i=I" "j=J" "k=K" "l=L" "m=M" "n=N" "o=O" "p=P" "q=Q" "r=R" "s=S" "t=T" "u=U" "v=V" "w=W" "x=X" "y=Y" "z=Z") DO CALL SET "%1=%%%1:%%~i%%"
GOTO:EOF

:TCase
:: Subroutine to convert a variable VALUE to Title Case.
:: The argument for this subroutine is the variable NAME.
FOR %%i IN (" a= A" " b= B" " c= C" " d= D" " e= E" " f= F" " g= G" " h= H" " i= I" " j= J" " k= K" " l= L" " m= M" " n= N" " o= O" " p= P" " q= Q" " r= R" " s= S" " t= T" " u= U" " v= V" " w= W" " x= X" " y= Y" " z= Z") DO CALL SET "%1=%%%1:%%~i%%"
GOTO:EOF
REM ----------------------------------------
REM Batch files - How To ... Convert Strings To Upper Case Or Lower Case
REM ----------------------------------------
REM https://www.robvanderwoude.com/battech_convertcase.php
REM ----------------------------------------




::
:IsDeviceReady
setlocal
set "ready_="
set ready_=true
dir "%~1" > nul 2>&1
if %errorlevel% NEQ 0 set ready_=false
endlocal & set "%2=%ready_%" & goto :EOF


:ENDE