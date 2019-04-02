@ECHO OFF
REM ------------------------------------------------------------------------

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

ECHO -----------------------------------------------------------------------
ECHO -----------------------------------------------------------------------
ECHO -----------------------------------------------------------------------
ECHO -----------------------------------------------------------------------

REM ------------------------------------------------------------------------
REM GOTO JUMP_OVER
REM ------------------------------------------------------------------------
call :COMPUTER_NAME_COMMON_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
IF %COMPUTERNAME%==1-ASUS-EEE      call :1-ASUS-EEE_NET_SHARE_DELETE
IF %COMPUTERNAME%==2-ASUS-X5DIJ    call :2-ASUS-X5DIJ_NET_SHARE_DELETE
IF %COMPUTERNAME%==3-LINDA-PC      call :3-LINDA-PC_NET_SHARE_DELETE
IF %COMPUTERNAME%==4-ASUS-GL522VW  call :4-ASUS-GL522VW_NET_SHARE_DELETE
IF %COMPUTERNAME%==5-ASUS-P2520LA  call :5-ASUS-P2520LA_NET_SHARE_DELETE
IF %COMPUTERNAME%==7-ASUS-GL522VW  call :7-ASUS-GL522VW_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
:JUMP_OVER
REM ------------------------------------------------------------------------


REM ------------------------------------------------------------------------
REM -- SET THE NET SHARE SET
REM ------------------------------------------------------------------------
IF %COMPUTERNAME%==1-ASUS-EEE      call :1-ASUS-EEE
IF %COMPUTERNAME%==2-ASUS-X5DIJ    call :2-ASUS-X5DIJ
IF %COMPUTERNAME%==3-LINDA-PC      call :3-LINDA-PC
IF %COMPUTERNAME%==4-ASUS-GL522VW  call :4-ASUS-GL522VW
IF %COMPUTERNAME%==5-ASUS-P2520LA  call :5-ASUS-P2520LA
IF %COMPUTERNAME%==7-ASUS-GL522VW  call :7-ASUS-GL522VW
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

::
:1-ASUS-EEE
REM ------------------------------------------------------------------------
SET "NAME_VAR="
SET NAME_VAR=1_ASUS_EEE
ECHO ON
NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM E OR F
REM NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_WD-55GB" \\%COMPUTERNAME% /DELETE
@ECHO OFF
goto :EOF
::
:2-ASUS-X5DIJ
REM ------------------------------------------------------------------------
SET "NAME_VAR="
SET NAME_VAR=2_ASUS_X5DIJ
ECHO ON
NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
REM E OR F
REM NET SHARE "%NAME_VAR%_03_FAT32_4GB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
@ECHO OFF
goto :EOF
::
:3-LINDA-PC
REM ------------------------------------------------------------------------
SET "NAME_VAR="
SET NAME_VAR=3_LINDA_PC
ECHO ON
NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_03_FAT32_4GB"="F:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
@ECHO OFF
goto :EOF
::
:4-ASUS-GL522VW
REM ------------------------------------------------------------------------
SET "NAME_VAR="
SET NAME_VAR=4_ASUS_GL522VW
ECHO ON
NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_04_NVR_1000_DRIVE"="G:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE


@ECHO OFF
goto :EOF
::
:5-ASUS-P2520LA
REM ------------------------------------------------------------------------
SET "NAME_VAR="
SET NAME_VAR=5_ASUS_P2520LA
ECHO ON
NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_04_REMOTE_2_SAMSUNG_4TB"="H:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
@ECHO OFF
goto :EOF
::
:7-ASUS-GL522VW
REM ------------------------------------------------------------------------

SET "NAME_VAR="
SET NAME_VAR=7_ASUS_GL522VW
ECHO ON
NET SHARE "%NAME_VAR%_01_C_DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_02_D_DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
@ECHO OFF
goto :EOF






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
NET SHARE "%NAME_VAR%_01_C-DRIVE"="C:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_02_D-DRIVE"="D:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE
NET SHARE "%NAME_VAR%_03_FAT32_4GB"="E:" /GRANT:EVERYONE,FULL /UNLIMITED /CACHE:NONE

NET SHARE "My Music" \\%COMPUTERNAME% /DELETE
NET SHARE "Music" \\%COMPUTERNAME% /DELETE
NET SHARE "Pictures" \\%COMPUTERNAME% /DELETE
NET SHARE "My Pictures" \\%COMPUTERNAME% /DELETE
NET SHARE "My Video" \\%COMPUTERNAME% /DELETE
NET SHARE "Users" \\%COMPUTERNAME% /DELETE
NET SHARE "SharedDocs" \\%COMPUTERNAME% /DELETE
goto :EOF

REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
REM -- THE DELETE NET SHARE SET 1 TO 7 ROUTINE 1 MISSING ---- 2ND OF 2
REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------
::
:1-ASUS-EEE_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
NET SHARE "asus-eee c" \\%COMPUTERNAME% /DELETE
NET SHARE "asus-eee d" \\%COMPUTERNAME% /DELETE
NET SHARE "asus-eee e-win 7" \\%COMPUTERNAME% /DELETE
NET SHARE "asus-eee e-win 7" \\%COMPUTERNAME% /DELETE
NET SHARE "asus-eee g-fat32" \\%COMPUTERNAME% /DELETE
NET SHARE "asus-eee wd-55gb" \\%COMPUTERNAME% /DELETE
NET SHARE "SharedDocs" \\%COMPUTERNAME% /DELETE
NET SHARE "Printer2" \\%COMPUTERNAME% /DELETE
NET SHARE "Printer4" \\%COMPUTERNAME% /DELETE
goto :EOF
REM ------------------------------------------------------------------------
::
:2-ASUS-X5DIJ_NET_SHARE_DELETE
REM ------------------------------------------------------------------------
NET SHARE "asus-x5dij c" \\%COMPUTERNAME% /DELETE
NET SHARE "asus-x5dij d" \\%COMPUTERNAME% /DELETE
NET SHARE "asus-x5dij fat32" \\%COMPUTERNAME% /DELETE
NET SHARE "SharedDocs" \\%COMPUTERNAME% /DELETE
REM NET SHARE "Printer" \\%COMPUTERNAME% /DELETE
REM NET SHARE "Printer4" \\%COMPUTERNAME% /DELETE
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
NET SHARE "_01_C-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "_02_D-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "_03_FAT32_4GB" \\%COMPUTERNAME% /DELETE
NET SHARE "3_LINDA_PC_01_C-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "3_LINDA_PC_02_D-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "3_LINDA_PC_03_FAT32_4GB" \\%COMPUTERNAME% /DELETE
NET SHARE "3_LINDA_PC_C" \\%COMPUTERNAME% /DELETE
NET SHARE "3_LINDA_PC_D" \\%COMPUTERNAME% /DELETE
NET SHARE "3_LINDA_PC_FAT32_4GB" \\%COMPUTERNAME% /DELETE

REM My Music     D:\# MY DOCS\# 01 My Documents\My Music
REM My Music2    D:\# MY DOCS COMMON\My Music
REM My Pictures  D:\# MY DOCS\# 01 My Documents\My Pictures
REM My Pictures2 D:\# MY DOCS COMMON\My Pictu...
REM My Video     D:\# MY DOCS\# 01 My Documents\My Video
REM My Video2    D:\# MY DOCS COMMON\My Video
REM SharedDocs   C:\USERS\PUBLIC\DOCUMENTS
REM Users        C:\Users
REM ------------------------------------------------------------------------
REM @ECHO OFF
goto :EOF
REM ------------------------------------------------------------------------
::
:4-ASUS-GL522VW_NET_SHARE_DELETE
REM ------------------------------------------------------------------------

NET SHARE "2015 2016" \\%COMPUTERNAME% /DELETE
NET SHARE "4-ASUS-GL522VW_01_C-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "4-ASUS-GL522VW_02_D-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "4-ASUS-GL522VW_03_FAT32" \\%COMPUTERNAME% /DELETE
NET SHARE "4-ASUS-GL522VW_04_NVR-1000-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "SharedDocs" \\%COMPUTERNAME% /DELETE
NET SHARE "Users" \\%COMPUTERNAME% /DELETE
NET SHARE "4-ASUS-GL522VW_01_C-DRIVE" \\%COMPUTERNAME% /DELETE

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
REM -- PAUSE
REM -- ECHO OFF

NET SHARE "Music" \\%COMPUTERNAME% /DELETE
NET SHARE "Pictures" \\%COMPUTERNAME% /DELETE
NET SHARE "Users" \\%COMPUTERNAME% /DELETE
NET SHARE "01 Banging Tunes" \\%COMPUTERNAME% /DELETE
NET SHARE "2015 2016" \\%COMPUTERNAME% /DELETE
NET SHARE "5_ASUS-P2520LA_01-C-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "5_ASUS-P2520LA_02-D-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "_01_C-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "_02_D-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "_03_FAT32_4GB" \\%COMPUTERNAME% /DELETE
NET SHARE "5_ASUS-P2520LA_03_FAT32" \\%COMPUTERNAME% /DELETE
NET SHARE "5-ASUS-P2520LA_01_C-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "5-ASUS-P2520LA_02_D-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "5-ASUS-P2520LA_03_FAT32" \\%COMPUTERNAME% /DELETE
NET SHARE "5-ASUS-P2520LA_REMOTE_2-SAMSUNG_4TB" \\%COMPUTERNAME% /DELETE
NET SHARE "5-ASUS-P2520LA_REMOTE_2-SAMSUNG_4TB" \\%COMPUTERNAME% /DELETE
NET SHARE "5-ASUS-P2520L FAT32-4TB" \\%COMPUTERNAME% /DELETE
NET SHARE "5-ASUS-P2520L FAT32-4TB" \\%COMPUTERNAME% /DELETE
NET SHARE "SharedDocs" \\%COMPUTERNAME% /DELETE
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
NET SHARE "_01_C-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "_02_D-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "_03_FAT32_4GB" \\%COMPUTERNAME% /DELETE
NET SHARE "7_ASUS_GL522VW_01_C-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "7_ASUS_GL522VW_02_D-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "7-ASUS-GL522VW_01_C-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "7-ASUS-GL522VW_02_D-DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "7-ASUS-GL522VW_01_C_DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "7-ASUS-GL522VW_02_D_DRIVE" \\%COMPUTERNAME% /DELETE
NET SHARE "7-ASUS-GL522VW_03_FAT32" \\%COMPUTERNAME% /DELETE
NET SHARE "Music" \\%COMPUTERNAME% /DELETE
NET SHARE "Pictures" \\%COMPUTERNAME% /DELETE
ECHO ON
REM -- MIGHT BE PROBLEM ACCESS DENIED
NET SHARE "Users" \\%COMPUTERNAME% /DELETE
REM PAUSE
ECHO OFF

REM 7-ASUS-GL522VW_01_C-DRIVE        C:\
REM 7-ASUS-GL522VW_02_D-DRIVE        D:\
REM 7-ASUS-GL522VW_03_FAT32          E:\
REM Music        D:\ON-LINE-SYNC\OneDrive\Music
REM Pictures     D:\ON-LINE-SYNC\OneDrive\Pictures
REM Users        C:\Users
goto :EOF
REM ------------------------------------------------------------------------

REM ------------------------------------------------------------------------
REM ------------------------------------------------------------------------

ENDE: