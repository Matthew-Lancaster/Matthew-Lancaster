@ECHO OFF

REM FOUND 2015
REM https://peter.hahndorf.eu/blog/post/2011/12/23/ChangingTheLocationOfWindowsSpecialFoldersOnTheCommandLine

REM NOTEPAD++ CHANGE HERE
REM https://peter.hahndorf.eu/blog/post/2011/12/23/ChangingTheLocationOfWindowsSpecialFoldersOnTheCommandLine

rem IF VALUENAME IS DEFAULT THEN YOU SET /v "" Default Sets All Values as Said
rem   /v       The value name, under the selected Key, to add.
rem   /ve      adds an empty value name (Default) for the key.

REM WELL MOST SOURCE INCLUDING MICROSOFT SAID ADD A CAROT ^ IN FRONT OF % BUT DOUBLE %% IS ANSWER
REM WELL THAT SOURCE WAS FOR THE WINDOWS 7 UPWARDS

REM AND YOU CAN HAVE SOMETHING LIKE THIS 
REM REG add HKCU\SOFTWARE\Classes\sourcecode\shell\open\command /v "" /t REG_SZ /d "\"C:\Program Files\Vim\vim72\gvim.exe\" --remote-tab-silent \"%%1\"" /f
REM http://vim.wikia.com/wiki/Windows_file_associations

REM -- ONE GOOD HELPFUL NOTE IS IF YOU EVER CHANGE YOU COMPUTER NAME
REM -- YOU GOT TO SET THE OWNERSHIP TO YOUR USERS FROM ADJUSTED TO DEFAULT ADMINISTRATORS
REM -- BE GOING IN SAFE-MODE AND THAT WAY

REM -- A NOTE FOR THE DAY - WHEN I APPLY THESE SETTING IN SAFE MODE
REM -- AND RELOGON TO DO FOR ALL USER
REM -- WON'T ACCEPT PASSWORD TO GET IN UNLESS REBOOT FIRST 

ECHO #### ------------------------------------------------ ####
ECHO #### -- MATT.LAN@BTINTERNET.COM JANUARY 201SIXTEEN -- ####
ECHO #### -- SAFE GUARD ONLY RUN ON KNOWN COMPUTER ------- ####
ECHO #### -- AND ONLY WINDOWS XP 32 BIT ------------------ ####
ECHO #### -- READY TO GO --------------------------------- ####
ECHO #### ------------------------------------------------ ####


@REM ECHO TEST YOU CAN SET VARIABLES PROPER
@ECHO OFF
REM ---
SET DRIVEPATH=Z:
REM --- SET HARD-CODED DEFAULT OR FIND OUR PATH LATER


@REM -- FIND OUR OWN PROGRAM FOR REFERENCE FOR THE MAIN DRIVE-PATH WE USE LIKE GRUB4DOS
@REM -- MAKE SURE FIRST FIND IS ONE IS WHAT WANT AND NOT ANOTHER LATER ON THE DRIVE
@REM ------------------

@ECHO OFF


REM MY ENVIROMENT VAR CHECK OVER --- NOKIA IS THE FRONT
REM C:\Program Files\PC Connectivity Solution\;%SystemRoot%\system32;%SystemRoot%;%SystemRoot%\System32\Wbem;C:\Program Files\Common Files\Ulead Systems\MPEG;C:\WINDOWS\system32\WindowsPowerShell\v1.0;C:\Program Files\Microsoft ASP.NET\ASP.NET Web Pages\v1.0\;C:\Program Files\QuickTime\QTSystem\;C:\Program Files\Common Files\Acronis\SnapAPI\

REM C:\Program Files\PC Connectivity Solution\;%SystemRoot%\system32;%SystemRoot%;%SystemRoot%\System32\Wbem;C:\Program Files\QuickTime\QTSystem\;C:\Program Files\Common Files\Acronis\SnapAPI\;C:\BAT_WIN_XP


REM -----------------------------------------------------------------
REM -----------------------------------------------------------------
REM ----------- FIND OUR DRIVE FAT32 OF LINK AND SHORTCUT
REM -----------------------------------------------------------------
REM -----------------------------------------------------------------

FOR %%G IN (A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@IF EXIST "%%G:\REG KEY SETTINGS\# Windows XP REG SCRIPT\# TWEAK REG WIN XP - NEW USER SETUP - SPECIAL FOLDERS.BAT" SET DRIVEPATH=%%G:)

ECHO OFF

FOR %%G IN (A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@IF EXIST "%%G:\REG KEY SETTINGS\# Windows XP REG SCRIPT\# TWEAK REG WIN XP - NEW USER SETUP - SPECIAL FOLDERS.BAT" ECHO #### -- SCAN FOUND DRIVEPATH UNIQUE FILE IS ON ----------- DRIVE %%G -- ####)

@ECHO #### -- --------------------------------------------------------    -- ####
@ECHO #### -- DOUBLE CHECK THIS RESULT IS NOT MORE THAN ONE

@ECHO #### -- --------------------------------------------------------    -- ####

@ECHO #### -- SET THE RELOCATABLE DRIVE OF OUR *DRIVEPATH* VARIABLE -- %DRIVEPATH% -- ####
@ECHO #### -- --------------------------------------------------------    -- ####

CD \
@ECHO OFF

@SET VARLATCH=0

@REM OS=Windows_NT

@IF NOT "%OS%"=="Windows_NT" ECHO NOT A MATCHING WINDOWS XP OS DETECTED -- %OS% -- GOTO END 
@IF NOT "%OS%"=="Windows_NT" GOTO END_SUB
ECHO #### -- WINDOWS XP OPERATING SYSTEM DETECTED OKAY -- %OS%


@IF NOT "%PROCESSOR_ARCHITECTURE%"=="x86" ECHO NOT A MATCHING 32 BIT COMPUTER -- %PROCESSOR_ARCHITECTURE% -- GOTO END
@IF NOT "%PROCESSOR_ARCHITECTURE%"=="x86" GOTO END_SUB
ECHO #### -- 32 BIT COMPUTER DETECTED OKAY -- %PROCESSOR_ARCHITECTURE%

IF "%COMPUTERNAME%" == "1-ASUS-EEE" SET VARLATCH=1
IF "%COMPUTERNAME%" == "2-ASUS-X5DIJ" SET VARLATCH=1

IF %VARLATCH% == 0 GOTO NOT_A_MATCH_RESULT

IF "%USERNAME%"=="MATT 01" SET VARLATCH=2
IF "%USERNAME%"=="MATT 02" SET VARLATCH=2
IF "%USERNAME%"=="MATT 03" SET VARLATCH=2
IF "%USERNAME%"=="MATT 03 LIMITED" SET VARLATCH=2
IF "%USERNAME%"=="MATT 04" SET VARLATCH=2

:NOT_A_MATCH_RESULT

@REM TEST TOGGLE
@REM @SET VARLATCH=0

REM ECHO. = GET A BLANK LINE
REM @IF %VARLATCH% == 1 ECHO.
REM @IF %VARLATCH% == 2 ECHO. 

ECHO #### -- COMPUTER NAME = %COMPUTERNAME% -- USER NAME = %USERNAME%


@IF NOT %VARLATCH% == 0 ECHO #### ------------------------------------------------------------------- ####


@IF %VARLATCH% == 1 ECHO #### -- MATCHING COMPUTER DETECTED - BUT NOT USER NAME

@IF %VARLATCH% == 2 ECHO #### -- MATCHING COMPUTER AND USER NAME DETECTED SUCCESS RESULT OKAY --- ####

@IF %VARLATCH% == 0 ECHO #### -- NOT A MATCHING COMPUTER NAME DETECTED -- GOTO END 
@IF %VARLATCH% == 0 GOTO END_SUB
@IF %VARLATCH% == 1 GOTO END_SUB

@REM GOT A MATCHING COMPUTER AND USER ADJUST TO 1 TO GO
@SET VARLATCH=1

@ECHO #### ------------------------------------------------------------------- ####
@ECHO #### -- MAKE SURE YOU WANT TO GO ON THIS CODE *CTRL C* IF YOU DON'T ---- ####
@ECHO #### ------------------------------------------------------------------- ####
@ECHO.
@PAUSE

@SET VARSTRING=-------------------------------------------------------------------------------
@REM ##################
@REM -------- SET NOTEPAD++ ----------
@ECHO %VARSTRING%
@ECHO %VARSTRING%
@ECHO ---------- SET NOTEPAD++ FOR EDIT FILES

@ECHO OFF

@REM -- MIGHT AS WELL - SET THE DRIVE AND PATH FIRST AS IT IS DONE LATER

@%DRIVEPATH%
@CD \



reg.exe ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "C:\Program Files\Notepad++\notepad++.exe %%1"

REM ---------------------------------
REM -------- RESTORE ORGINAL --------
REM reg.exe ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE %%1"
REM ---------------------------------
REM ##################

@ECHO ON
@ECHO.
@ECHO %VARSTRING%
@ECHO ---------- SET THE SPECIAL SYSTEM FOLDERS

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Favorites /t REG_EXPAND_SZ /d "%DRIVEPATH%\01 Favs"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Favorites /t REG_SZ /d "%DRIVEPATH%\01 Favs" 
@ECHO.
@ECHO %VARSTRING%

@REM THE COMMONS ARE SET IN HK LOCAL MACHINE
@REM AND OTHER CURRENT USER

reg.exe ADD "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Favorites" /t REG_EXPAND_SZ /d "%DRIVEPATH%\01 Favs Common"
reg.exe ADD "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Common Favorites" /t REG_SZ /d "%DRIVEPATH%\01 Favs Common"
@ECHO.
@ECHO %VARSTRING%

@REM ---------------------
@REM -- SET DESKTOP FOLDER
@REM - ########################## ADJUST USER VARIABLE 1 OF 3
@REM ---------------------

@ECHO AFTER DESKTOP SET TO A NEW FOLDER -- YOU DON'T NEED REQUIRE LOG OFF - ON -- ONLY F5 REFRESH

@REM -- SINGLE COMMON METHOD
@REM reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Desktop /t REG_EXPAND_SZ /d "%DRIVEPATH%\\01 Desktop"
@REM reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Desktop /t REG_SZ /d "%DRIVEPATH%\\01 Desktop" 
@REM @ECHO.
@REM @ECHO %VARSTRING%

%DRIVEPATH%
CD \

@IF NOT EXIST "%DRIVEPATH%\01 Desktop\" MKDIR "%DRIVEPATH%\01 Desktop\"
@IF NOT EXIST "%DRIVEPATH%\01 Desktop\#_%COMPUTERNAME%\" MKDIR "%DRIVEPATH%\01 Desktop\#_%COMPUTERNAME%\"

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Desktop" /t REG_EXPAND_SZ /d "%DRIVEPATH%\01 Desktop\#_%COMPUTERNAME%"

@REM G:\01 Desktop\#_1-ASUS-EEE

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Desktop" /t REG_SZ /d "%DRIVEPATH%\01 Desktop\#_%COMPUTERNAME%" 
@ECHO.
@ECHO %VARSTRING%



reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Personal /t REG_EXPAND_SZ /d "D:\# MY DOCS\# 01 My Documents"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Personal /t REG_SZ /d "D:\# MY DOCS\# 01 My Documents" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v SendTo /t REG_EXPAND_SZ /d "%DRIVEPATH%\01 SendTo"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v SendTo /t REG_SZ /d "%DRIVEPATH%\01 SendTo" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Start Menu" /t REG_EXPAND_SZ /d "%DRIVEPATH%\01 Start Menu"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Start Menu" /t REG_SZ /d "%DRIVEPATH%\01 Start Menu" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Templates /t REG_EXPAND_SZ /d "%DRIVEPATH%\01 Templates"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Templates /t REG_SZ /d "%DRIVEPATH%\01 Templates" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Programs /t REG_EXPAND_SZ /d "%DRIVEPATH%\01 Start Menu\Programs"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Programs /t REG_SZ /d "%DRIVEPATH%\01 Start Menu\Programs" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Startup /t REG_EXPAND_SZ /d "%DRIVEPATH%\01 Start Menu\Programs\Startup
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Startup /t REG_SZ /d "%DRIVEPATH%\01 Start Menu\Programs\Startup" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "My Pictures" /t REG_EXPAND_SZ /d "D:\# MY DOCS\# 01 Shared Documents\My Pictures"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "My Pictures" /t REG_SZ /d "D:\# MY DOCS\# 01 Shared Documents\My Pictures" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "My Music" /t REG_EXPAND_SZ /d "D:\# MY DOCS\# 01 Shared Documents\My Music"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "My Music" /t REG_SZ /d "D:\# MY DOCS\# 01 Shared Documents\My Music" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "My Video" /t REG_EXPAND_SZ /d "D:\# MY DOCS\# 01 My Documents\My Video"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "My Video" /t REG_SZ /d "D:\# MY DOCS\# 01 My Documents\My Video" 
@ECHO.
@ECHO %VARSTRING%

@REM -- SET RECENT FOLDER
@REM - ########################## ADJUST USER VARIABLE 1 OF 3
%DRIVEPATH%
CD \

@IF NOT EXIST "%DRIVEPATH%\01 Recent\" MKDIR "%DRIVEPATH%\01 Recent\" 
@REM IF NOT EXIST "%DRIVEPATH%\01 Recent\#_%COMPUTERNAME%" MKDIR "%DRIVEPATH%\01 Recent\#_%COMPUTERNAME%"
@IF NOT EXIST "%DRIVEPATH%\01 Recent\#_%COMPUTERNAME%_%USERNAME%" MKDIR "%DRIVEPATH%\01 Recent\#_%COMPUTERNAME%\_%USERNAME%"

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Recent" /t REG_EXPAND_SZ /d "%DRIVEPATH%\01 Recent\#_%COMPUTERNAME%_%USERNAME%"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Recent" /t REG_SZ /d "%DRIVEPATH%\01 Recent\#_%COMPUTERNAME%_%USERNAME%" 
@ECHO.
@ECHO %VARSTRING%


@REM - AFTER THAT YOU SEE SETTING FAV SPECIAL FOLDERS REQUIRES A DUPLCATE USE OF 
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
@REM - ----------------------------------------


@REM - IF I COULD GRAB THE LOCAL USER NAME BETTER
@REM - BETTER STILL DO ANOTHER VERSION IN VBS - COMPATIABLE IN WINDOWS WITHOUT VISUAL BASIC LANGUAGE IDE STUDIO INTERGRATED DEVELOPMENT INVIROMENT
@REM - STUDIO AUDIO VIDEO IMAGE
@REM - BETTER AGAIN COMPLETE IT IN THIS BATCH

@REM - ########################## ADJUST USER VARIABLE 2 OF 3
@REM - SET TEMP FOLDER IN THE SPECIAL FOLDERS

@REM "C:\TEMP\#_TEMP_MATT_02"

@REM -- SHOW WHAT DRIVE SETTING ARE MADE TO -- NOT NECCESSARY REALLY
@REM -- DONT THINK WE NEED IT HERE
@REM - #############################################################


%HOMEDRIVE%
CD \

@IF NOT EXIST "%HOMEDRIVE%\TEMP\" MKDIR "%HOMEDRIVE%\TEMP\" 
@IF NOT EXIST "%HOMEDRIVE%\TEMP\-%COMPUTERNAME%" MKDIR "%HOMEDRIVE%\TEMP\-%COMPUTERNAME%"


@ECHO OFF
REM @IF NOT EXIST "C:\TEMP\#_%COMPUTERNAME%\%USERNAME%" MKDIR "C:\TEMP\#_%COMPUTERNAME%\%USERNAME%"
REM -- THIS IS THE TEMP FOLDER FOR ALL WINDOWS USERS SO CAN'T HAVE TEMP USERNAME
REM -- REBOOT IS REQUIRED TO SET

REM NOTE TEST 1 -- THE TEMP FOLDER DON'T LIKE UNDERSCORE'S IN PATH NAME WITH INSTALL SOME PROGRAMS
REM --------------- SEEMS TO STRIP IT OUT LIKE DON'T WANT IT AND FSAIL CAN'T FIND INSTALL PATH WANTED
REM -- REMOVE THE HASH ALSO I THINK
REM -- CAN CONFIRM DON'T LIKE THE HASH
REM -- NOW SAME TO DO WITH HISTORY AND RECENT - THINK
@ECHO ON

reg.exe ADD "HKCU\Environment" /f /v "TEMP" /t REG_MULTI_SZ /d "%HOMEDRIVE%\TEMP\-%COMPUTERNAME%"
reg.exe ADD "HKCU\Environment" /f /v "TMP" /t REG_MULTI_SZ /d "%HOMEDRIVE%\TEMP\-%COMPUTERNAME%"
reg.exe ADD "HKU\.DEFAULT\Environment" /f /v "TEMP" /t REG_MULTI_SZ /d "%HOMEDRIVE%\TEMP\-%COMPUTERNAME%"
reg.exe ADD "HKU\.DEFAULT\Environment" /f /v "TMP" /t REG_MULTI_SZ /d "%HOMEDRIVE%\TEMP\-%COMPUTERNAME%"
reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /f /v "TEMP" /t REG_MULTI_SZ /d "%HOMEDRIVE%\TEMP\-%COMPUTERNAME%" 
reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /f /v "TMP" /t REG_MULTI_SZ /d "%HOMEDRIVE%\TEMP\-%COMPUTERNAME%"




@REM -- SHOW WHAT DRIVE SETTING ARE MADE TO -- NOT NECCESSARY REALLY
@REM - #############################################################
%DRIVEPATH%
CD \

@ECHO.
@ECHO %VARSTRING%


@REM - SET THE COMMON START MENU PROGRAMS - FOLDER ONLY FOUND TO BE SETTABLE IN TWEAK XP PRO
reg.exe ADD "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Programs" /t REG_EXPAND_SZ /d "%DRIVEPATH%\01 Start Menu Common\Programs" 
reg.exe ADD "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Common Programs" /f /v "Common Programs" /t REG_SZ /d "%DRIVEPATH%\01 Start Menu Common\Programs" 
@ECHO.
@ECHO %VARSTRING%



@REM - WELL I DIDN'T KNOW HISTORY AND RECENT WERE DIFFERENT - RECENT DOCUMENT AND INTERNET EXPLORER
@REM - ########################## ADJUST USER VARIABLE 3 OF 3
@REM - SET INTERNET EXPLORER HISTORY - CAREFUL TO USE A COMPUTER NAME AND USER NAME - WHEN SYNC-ER

%DRIVEPATH%
CD \

@IF NOT EXIST "%DRIVEPATH%\01 History\" MKDIR "%DRIVEPATH%\01 History\" 
@REM IF NOT EXIST "%DRIVEPATH%\01 History\#_%COMPUTERNAME%" MKDIR "%DRIVEPATH%\01 History\#_%COMPUTERNAME%"
@IF NOT EXIST "%DRIVEPATH%\01 History\#_%COMPUTERNAME%_%USERNAME%" MKDIR "%DRIVEPATH%\01 History\#_%COMPUTERNAME%_%USERNAME%"

reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "History" /t REG_EXPAND_SZ /d "%DRIVEPATH%\01 History\#_%COMPUTERNAME%_%USERNAME%"
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "History" /t REG_SZ /d "%DRIVEPATH%\01 History\#_%COMPUTERNAME%_%USERNAME%" 
@ECHO.
@ECHO %VARSTRING%

@REM - IT'S UP TO YOU IF YOU WANT - RECENT / HISTORY
@REM - CAN'T THINK OF ANOTHER ONE - I MIGHT DO THAT ONE LATER - DONE NOW...
@REM - ADMIN TOOLS - AND I KNOW ABOUT THIS FROM WEB BEFORE 
@REM - IT UPDATED ITS REG ENTRY JUST BY MOVE IT - LIKE MANY SPECIAL SYSTEM FOLDERS

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

@REM 84.59259796	Tweak-xp.exe:5920	SetValue	HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Common Programs	SUCCESS	"%DRIVEPATH%\01 Start Menu Common\Programs"	
@REM 84.59375763	Tweak-xp.exe:5920	SetValue	HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Common Programs	SUCCESS	"%DRIVEPATH%\01 Start Menu Common\Programs"	

@REM 204.59127808	Tweak-xp.exe:8772	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\History	SUCCESS	"%DRIVEPATH%\01 History\ASUS-EEE\MATT 02"	
@REM 204.59255981	Tweak-xp.exe:8772	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\History	SUCCESS	"%DRIVEPATH%\01 History\ASUS-EEE\MATT 02"	



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

@ECHO DO A EVERY RUN -- BACKUP OF WINDOWS ACTIVATION FILE
@ECHO DESTINATION -- "%DRIVEPATH%\01 WINDOWS XP Activation Backup\%VARDATE%_%COMPUTERNAME%_%USERNAME%\wpa.dbl"
@ECHO DON'T BOTHER WITH OFFICE XP OR ...
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
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" /f /v "NoSharedDocuments" /t REG_DWORD /d 1 
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
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CaretWidth" /t REG_DWORD /d 2
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


@REM -- SHOW WHAT DRIVE SETTING ARE MADE TO -- NOT NECCESSARY REALLY
@REM - #############################################################
%HOMEDRIVE%
@REM %DRIVEPATH%
CD \


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

reg.exe ADD "HKCU\Software\Microsoft\IntelliPoint\Scrolling" /f /v "ScrollLines" /t REG_DWORD /d 3
reg.exe ADD "HKCU\Software\Microsoft\IntelliPoint\Scrolling" /f /v "AccelerationCurve" /t REG_DWORD /d 6

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
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "CursorSize" /t REG_DWORD /d 50
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "HistoryBufferSize" /t REG_DWORD /d 90
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "InsertMode" /t REG_DWORD /d 1
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "NumberOfHistoryBuffers" /t REG_DWORD /d 4


:: 655294544 equals 9999 lines in the GUI
reg.exe add "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /v ScreenBufferSize /t REG_DWORD /d 655294544 /f

:: We also need to change the Window Height, 3276880 = 50 lines
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
@REM -- TweakUI.exe:17492	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ThumbnailQuality	SUCCESS	0x5F	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ThumbnailSize	SUCCESS	0x64	

@REM -- SIZE 64 = 100 PIXELS OF HEX YES
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer" /f /v "ThumbnailQuality" /t REG_DWORD /d 95
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer" /f /v "ThumbnailSize" /t REG_DWORD /d 100
@ECHO.
@ECHO %VARSTRING%
@REM -


@REM -- WELL BEEN SETTING ALL THESE TODAY 11-12-2015 FOUND NEW ICON WITH REGMON TO JUMP TO REGISTRY
@REM -- MAKES EASIER -- AND REGMON NOT CRASH ON MY LOW END ASUS EEE 
@REM -- EXPLORER CUSTOMIZE
@REM -- SET FOLDERS TO REMEMBER 
@REM -- TweakUI.exe:17492	SetValue	HKCU\Software\Microsoft\Windows\ShellNoRoam\Bag MRU Size	SUCCESS	0x1F4	
@REM -- TweakUI.exe:17492	SetValue	HKCU\Software\Microsoft\Windows\Shell\Bag MRU Size	SUCCESS	0x1F4	
@REM -- SEEMS TO BE A FAULT WITH ANOTHER ALMOST SAME SETTING HERE -- BagMRU Size
@REM -- HEX 1F4 = 500
reg.exe ADD "HKCU\Software\Microsoft\Windows\ShellNoRoam" /f /v "Bag MRU Size" /t REG_DWORD /d 500
@REM -- SEEMS TO BE A FAULT WITH ANOTHER ALMOST SAME SETTING HERE -- BagMRU Size
reg.exe ADD "HKCU\Software\Microsoft\Windows\Shell" /f /v "Bag MRU Size" /t REG_DWORD /d 500
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- EXPLORER CUSTOMIZE
@REM -- SLIDE SHOW DELAY
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ShellImageView" /f /v "Timeout" /t REG_DWORD /d 1
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- EXPLORER -- NEED SOME WORK - FEW CHECK-BOXES
@REM -- BIT HARD WORK THIS ONE -- DO LATER LAST ONE TO DO AS WELL
@rem ------------------------------------------------------------

@REM -- COMMON DIALOGUES
@REM -- TweakUI.exe:14396	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoComplete\AutoComplete In File Dialog	SUCCESS	0x1	
@REM -- TweakUI.exe:14396	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\comdlg32\NoFileMru	SUCCESS	0x0	
@REM -- TweakUI.exe:14396	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\comdlg32\NoBackButton	SUCCESS	0x0	
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoComplete" /f /v "AutoComplete" /t REG_DWORD /d 1
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\comdlg32" /f /v "NoFileMru" /t REG_DWORD /d 0
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\comdlg32" /f /v "NoBackButton" /t REG_DWORD /d 0
@ECHO.
@ECHO %VARSTRING%
@REM -


@REM -- TASKBAR AND STARTMENU
@REM -- TweakUI.exe:14396	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoUserNameInStartMenu	SUCCESS	01 00 00 00 	
@REM -- TweakUI.exe:14396	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoTrayItemsDisplay	SUCCESS	01 00 00 00 	
@REM -- TweakUI.exe:14396	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoLowDiskSpaceChecks	SUCCESS	0x1	
@REM -- TweakUI.exe:14396	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\EnableBalloonTips	SUCCESS	0x1	
@REM -- TweakUI.exe:14396	SetValue	

@REM SHOW CHECKED ENABLED = 1
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" /f /v "NoUserNameInStartMenu" /t REG_BINARY /d "00000000"
@REM SHOW CHECKED ENABLED = 1
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" /f /v "NoTrayItemsDisplay" /t REG_BINARY /d "00000000"

@REM WARN WHEN LOW = CHECKED ENABLED = 0
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" /f /v "NoLowDiskSpaceChecks" /t REG_DWORD /d 0

reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /f /v "EnableBalloonTips" /t REG_DWORD /d 1
@ECHO.
@ECHO %VARSTRING%
@REM -


@REM -- TASKBAR AND STARTMENU - GROUPING
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /f /v "TaskbarGroupSize" /t REG_DWORD /d 0
@ECHO.
@ECHO %VARSTRING%
@REM -

@REM -- DESKTOP ICONS -- NEED SOME WORK - FEW CHECK-BOXES
@REM -- CAN'T BE DONE ALL IN SERIAL NUMBERS
@REM -- USE TWEAK UI


@REM -- DESKTOP FIRST ICON -- NEED SOME WORK - FEW CHECK-BOXES
@REM -- WONT SET THIS ONE PROPER -- USES SERIAL NUMBER
@REM -- HKEY_CLASSES_ROOT\CLSID\{450D8FBA-AD25-11D0-98A8-0800361B1103}\SortOrderIndex = DWORD 54 FOR MY COMPUTER FIRST -- MY COMPUTER OR MY DOCUMENTS 54 OR 48 HEX

reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" /f /v "NoDriveTypeAutoRun" /t REG_DWORD /d 95
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

@REM 0 OTB ART & VB GRAPHICS & DEMO'S
@REM 0 OTB VB EXE
@REM 1 OTB 1
@REM 1 OTB GAMES
@REM 2 OTB 2
@REM 3 OTB 3
@REM 4 OTB 4
@REM 5 OTB 5

reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\0 OTB ART & VB GRAPHICS & DEMO'S" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\0 OTB ART & VB GRAPHICS & DEMO'S"
reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\0 OTB VB EXE" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\0 OTB VB EXE"
reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\1 OTB 1" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\1 OTB 1"
reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\1 OTB GAMES" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\1 OTB GAMES"
reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\2 OTB 2" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\2 OTB 2"
reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\3 OTB 3" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\3 OTB 3"
reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\4 OTB 4" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\4 OTB 4"
reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\5 OTB 5" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\5 OTB 5"


@REM reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\1 OTB" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\1 OTB"
@REM reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\2 OTB" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\2 OTB"
@REM reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\3 OTB" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\3 OTB"
@REM reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\4 OTB" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\4 OTB"
@REM reg.exe ADD "HKCU\Software\Microsoft\Office\9.0\Shortcut Bar\Toolbars\5 OTB" /f /v "Path" /t REG_SZ /d "%DRIVEPATH%\00 Office ToolBar\5 OTB"


@REM -- 7.88941622	Tweak-xp.exe:19696	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Run\TransparentIcons	SUCCESS	""C:\Program Files\Tweak-XP Pro\tranicon.exe" -ex"	

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
 
ECHO ####--------------------------------------------------------------------####
ECHO REGISTRY CHANGES GOT MADE AND 4 DIRECTORYS MADE IF THEY DIDN'T EXIST ALREADY
ECHO ####--------------------------------------------------------------------####

REM ------------------------------------------------
REM -- NEED TO SET THIS PROPER IF SYS IS NOT C DRIVE
REM -- WHAT EVER THE SYS OPP PATH IS FOR TEMP

ECHO 1 OF 3 --- "C:\TEMP\-%COMPUTERNAME%"
ECHO 2 OF 3 --- "%DRIVEPATH%\01 Recent\#_%COMPUTERNAME%_%USERNAME%"
ECHO 3 OF 3 --- "%DRIVEPATH%\01 History\#_%COMPUTERNAME%_%USERNAME%"
ECHO 4 OF 4 --- "%DRIVEPATH%\01 WINDOWS XP Activation Backup\%VARDATE%_%COMPUTERNAME%_%USERNAME%"

:END_SUB

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
REM reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "My Video" /t REG_EXPAND_SZ /d "M:\Media\Videos"
REM reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "My Video" /t REG_SZ /d "M:\Media\Videos"
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
171.67987061	Tweak-xp.exe:86024	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\StartMenuFavorites	SUCCESS	0x0	

---------------------
---------------------
Count = 087 -- Sun 23-Aug-2015 01:28:09
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
205.61871338	Tweak-xp.exe:86024	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_ShowMyPics	SUCCESS	0x0	

---------------------
---------------------
Count = 088 -- Sun 23-Aug-2015 01:28:23
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
223.71253967	Tweak-xp.exe:86024	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_ShowMyMusic	SUCCESS	0x0	

---------------------
---------------------
Count = 089 -- Sun 23-Aug-2015 01:30:04
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
43.50534821	Tweak-xp.exe:86024	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\StartMenuFavorites	SUCCESS	0x0	

---------------------
---------------------
Count = 090 -- Sun 23-Aug-2015 01:30:35
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
77.38096619	Tweak-xp.exe:86024	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_NotifyNewApps	SUCCESS	0x0	

---------------------
---------------------
Count = 091 -- Sun 23-Aug-2015 01:31:23
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
105.33364105	Tweak-xp.exe:86024	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\StartMenuAdminTools	SUCCESS	"YES"	

---------------------
---------------------
Count = 092 -- Sun 23-Aug-2015 01:32:49
---------------------
Registry Monitor - Sysinternals: www.sysinternals.com
---------------------
148.41154480	Tweak-xp.exe:86024	OpenKey	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\MenuOrder\Start Menu2\	SUCCESS	Access: 0x1 	

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
