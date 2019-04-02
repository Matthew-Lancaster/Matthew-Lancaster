@ECHO OFF

REM FOUND 2015
REM REM https://peter.hahndorf.eu/blog/post/2011/12/23/ChangingTheLocationOfWindowsSpecialFoldersOnTheCommandLine

REM NOTEPAD++ CHANGE HERE
REM REM https://peter.hahndorf.eu/blog/post/2011/12/23/ChangingTheLocationOfWindowsSpecialFoldersOnTheCommandLine

rem IF VALUENAME IS DEFAULT THEN YOU SET /v "" Default Sets All Values as Said
rem   /v       The value name, under the selected Key, to add.
rem   /ve      adds an empty value name (Default) for the key.

REM WELL MOST SOURCE INCLUDING MICROSOFT SAID ADD A CAROT ^ IN FRONT OF % BUT DOUBLE %% IS ANSWER
REM WELL THAT SOURCE WAS FOR THE WINDOWS 7 UPWARDS

REM AND YOU CAN HAVE SOMETHING LIKE THIS 
REM reg add HKCU\SOFTWARE\Classes\sourcecode\shell\open\command /v "" /t REG_SZ /d "\"C:\Program Files\Vim\vim72\gvim.exe\" --remote-tab-silent \"%%1\"" /f
REM http://vim.wikia.com/wiki/Windows_file_associations


ECHO ###---------------------------------####
ECHO -- SAFE GUARD ONLY RUN ON KNOWN COMPUTER
ECHO -- AND ONLY WINDOWS XP -------------####
ECHO -- READY TO GO ---------------------####
CD \
@ECHO ON
SET VARLATCH=0
IF "%COMPUTERNAME%" == "ASUS-EEE" SET VARLATCH=1
@IF %VARLATCH% == 0 ECHO OFF 
IF "%COMPUTERNAME%" == "ASUS-BIGGER" SET VARLATCH=1
@IF %VARLATCH% == 0 ECHO OFF 
IF "%USERNAME%"=="MATT 01" SET VARLATCH=2
@IF %VARLATCH% == 0 ECHO OFF 
IF "%USERNAME%"=="MATT 02" SET VARLATCH=2
@IF %VARLATCH% == 0 ECHO OFF 
IF "%USERNAME%"=="MATT 03" SET VARLATCH=2
@IF %VARLATCH% == 0 ECHO OFF 
IF "%USERNAME%"=="MATT 04" SET VARLATCH=2

@REM TEST TOGGLE
@REM @SET VARLATCH=0

@IF %VARLATCH% == 1 ECHO.
@IF %VARLATCH% == 2 ECHO. 
@IF %VARLATCH% == 1 ECHO MATCHING COMPUTER DETECTED - BUT NOT USER  
@IF %VARLATCH% == 2 ECHO MATCHING COMPUTER AND USER DETECTED SUCCESS RESULT

@IF %VARLATCH% == 0 GOTO END_SUB
@IF %VARLATCH% == 1 GOTO END_SUB

@REM GOT A MATCHING COMPUTER AND USER ADJUST TO 1 TO GO
@SET VARLATCH=1

@ECHO ####-------------------------------------------------####
@ECHO -- SURE YOU WANT TO GO ON THIS CODE *CTRL C* IF YOU DON'T
@ECHO ####-------------------------------------------------####
@ECHO.
@PAUSE


@SET VARSTRING=-------------------------------------------------------------------------------
@REM ##################
@REM -------- SET NOTEPAD++ ----------
@ECHO %VARSTRING%
@ECHO %VARSTRING%
@ECHO ---------- SET NOTEPAD++ FOR EDIT FILES

reg.exe ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "C:\Program Files\Notepad++\notepad++.exe %%1"

@ECHO OFF

REM ---------------------------------
REM -------- RESTORE ORGINAL --------
REM reg.exe ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE %%1"
REM ---------------------------------
REM ##################

@ECHO ON
@ECHO.
@ECHO %VARSTRING%
@ECHO ---------- SET THE SPECIAL SYSTEM FOLDERS

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Favorites /t REG_EXPAND_SZ /d "E:\01 Favs"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Favorites /t REG_SZ /d "E:\01 Favs" 
@ECHO.
@ECHO %VARSTRING%

@REM THE COMMONS ARE SET IN HK LOCAL MACHINE
@REM AND OTHER CURRENT USER

reg.exe ADD "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Favorites" /t REG_EXPAND_SZ /d "E:\01 Favs Common"
reg.exe ADD "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Common Favorites" /t REG_SZ /d "E:\01 Favs Common"
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Desktop /t REG_EXPAND_SZ /d "E:\\01 Desktop"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Desktop /t REG_SZ /d "E:\\01 Desktop" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Personal /t REG_EXPAND_SZ /d "D:\\# MY DOCS\\# 01 My Documents"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Personal /t REG_SZ /d "D:\\# MY DOCS\\# 01 My Documents" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v SendTo /t REG_EXPAND_SZ /d "E:\\01 SendTo"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v SendTo /t REG_SZ /d "E:\\01 SendTo" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Start Menu" /t REG_EXPAND_SZ /d "E:\\01 Start Menu"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Start Menu" /t REG_SZ /d "E:\\01 Start Menu" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Templates /t REG_EXPAND_SZ /d "E:\\01 Templates"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Templates /t REG_SZ /d "E:\\01 Templates" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Programs /t REG_EXPAND_SZ /d "E:\\01 Start Menu\\Programs"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Programs /t REG_SZ /d "E:\\01 Start Menu\\Programs" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v Startup /t REG_EXPAND_SZ /d "E:\\01 Start Menu\\Programs\\Startup
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v Startup /t REG_SZ /d "E:\\01 Start Menu\\Programs\\Startup" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "My Pictures" /t REG_EXPAND_SZ /d "D:\\# MY DOCS\\# 01 Shared Documents\\My Pictures"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "My Pictures" /t REG_SZ /d "D:\\# MY DOCS\\# 01 Shared Documents\\My Pictures" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "My Music" /t REG_EXPAND_SZ /d "D:\\# MY DOCS\\# 01 Shared Documents\\My Music"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "My Music" /t REG_SZ /d "D:\\# MY DOCS\\# 01 Shared Documents\\My Music" 
@ECHO.
@ECHO %VARSTRING%

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "My Video" /t REG_EXPAND_SZ /d "D:\\# MY DOCS\\# 01 My Documents\\My Video"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "My Video" /t REG_SZ /d "D:\\# MY DOCS\\# 01 My Documents\\My Video" 
@ECHO.
@ECHO %VARSTRING%

@REM -- SET RECENT FOLDER
@REM - ########################## ADJUST USER VARIABLE 1 OF 3
@E:
@CD \
@IF NOT EXIST "E:\01 Recent\" MKDIR "E:\01 Recent\" 
@IF NOT EXIST "E:\01 Recent\#_%COMPUTERNAME%" MKDIR "E:\01 Recent\#_%COMPUTERNAME%"
@IF NOT EXIST "E:\01 Recent\#_%COMPUTERNAME%\%USERNAME%" MKDIR "E:\01 Recent\#_%COMPUTERNAME%\""%USERNAME%"

reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Recent" /t REG_EXPAND_SZ /d "E:\01 Recent\#_%COMPUTERNAME%\%USERNAME%"
reg.exe ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "Recent" /t REG_SZ /d "E:\01 Recent\#_%COMPUTERNAME%\%USERNAME%" 
@ECHO.
@ECHO %VARSTRING%


@REM - AFTER THAT YOU SEE SETTING FAV SPECIAL FOLDERS REQUIRES A DUPLCATE USE OF REG_EXPAND_SZ AS WELL AS REG_SZ OR REG_MULTI_SZ
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

@C:
@CD \

@IF NOT EXIST "C:\TEMP\" MKDIR "C:\TEMP\" 
@IF NOT EXIST "C:\TEMP\#_%COMPUTERNAME%" MKDIR "C:\TEMP\#_%COMPUTERNAME%"
REM @IF NOT EXIST "C:\TEMP\#_%COMPUTERNAME%\%USERNAME%" MKDIR "C:\TEMP\#_%COMPUTERNAME%\%USERNAME%"
REM -- THIS IS THE TEMP FOLDER FOR ALL WINDOWS USERS SO CAN'T HAVE TEMP USERNAME
REM -- REBOOT IS REQUIRED TO SET

REM -- ANOTHER LAST CONCLUTION PROBLEM
REM -- CHKNTFS /C OPTION DOES NOT GUARANTEE THE DRIVE WILL BE CHECK AT BOOT
REM -- TRUE BY THE DOCUMENTATION HELP OF COMMAND ALSO I BELIEVE
REM -- EVEN WHEN CHKDSK READ ONLY FOUND FAULTS
REM -- SO RUN CHKDSK /F WHILE STILL BOOTED IF LOCKED AND THEN WILL ASK TO DO AT REBOOT

REM -- ADD THE OPTION TO ASK QUESTION AT END MAYBE - MIGHT DO
REM -- I DON'T WORRY ABOUT BEING ON DRIVE WHEN CHKDSK UNLOCKING THE DRIVE UNLOCKING 
REM -- THAT IS A MINOR PROBLEM 


reg.exe ADD "HKCU\Environment" /f /v "TEMP" /t REG_MULTI_SZ /d "C:\TEMP\#_%COMPUTERNAME%"
reg.exe ADD "HKCU\Environment" /f /v "TMP" /t REG_MULTI_SZ /d "C:\TEMP\#_%COMPUTERNAME%"
reg.exe ADD "HKU\.DEFAULT\Environment" /f /v "TEMP" /t REG_MULTI_SZ /d "C:\TEMP\#_%COMPUTERNAME%"
reg.exe ADD "HKU\.DEFAULT\Environment" /f /v "TMP" /t REG_MULTI_SZ /d "C:\TEMP\#_%COMPUTERNAME%"
reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /f /v "TEMP" /t REG_MULTI_SZ /d "C:\TEMP\#_%COMPUTERNAME%" 
reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /f /v "TMP" /t REG_MULTI_SZ /d "C:\TEMP\#_%COMPUTERNAME%"





@ECHO.
@ECHO %VARSTRING%


@REM - SET THE COMMON START MENU PROGRAMS - FOLDER ONLY FOUND TO BE SETTABLE IN TWEAK XP PRO
reg.exe ADD "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "Common Programs" /t REG_EXPAND_SZ /d "E:\01 Start Menu Common\Programs" 
reg.exe ADD "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Common Programs" /f /v "Common Programs" /t REG_SZ /d "E:\01 Start Menu Common\Programs" 
@ECHO.
@ECHO %VARSTRING%



@REM - WELL I DID'NT KNOW HISTORY AND RECENT WERE DIFFERENT - RECENT DOCUMENT AND INTERNET EXPLORER
@REM - ########################## ADJUST USER VARIABLE 3 OF 3
@REM - SET INTERNET EXPLORER HISTORY - CAREFUL TO USE A COMPUTER NAME AND USER NAME - WHEN SYNC-ER

@E:
@CD \
@IF NOT EXIST "E:\01 History\" MKDIR "E:\01 History\" 
@IF NOT EXIST "E:\01 History\#_%COMPUTERNAME%" MKDIR "E:\01 History\#_%COMPUTERNAME%"
@IF NOT EXIST "E:\01 History\#_%COMPUTERNAME%\%USERNAME%" MKDIR "E:\01 History\#_%COMPUTERNAME%\%USERNAME%"

reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /f /v "History" /t REG_EXPAND_SZ /d "E:\01 History\#_%COMPUTERNAME%\%USERNAME%"
reg.exe ADD "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /f /v "History" /t REG_SZ /d "E:\01 History\#_%COMPUTERNAME%\%USERNAME%" 
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

@REM 84.59259796	Tweak-xp.exe:5920	SetValue	HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Common Programs	SUCCESS	"E:\01 Start Menu Common\Programs"	
@REM 84.59375763	Tweak-xp.exe:5920	SetValue	HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Common Programs	SUCCESS	"E:\01 Start Menu Common\Programs"	

@REM 204.59127808	Tweak-xp.exe:8772	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\History	SUCCESS	"E:\01 History\ASUS-EEE\MATT 02"	
@REM 204.59255981	Tweak-xp.exe:8772	SetValue	HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\History	SUCCESS	"E:\01 History\ASUS-EEE\MATT 02"	



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


@REM TWEAK XP PRO LETS GO
@REM STARTING FROM THE TOP
@REM ---- CACHE OPTIMIZATION ----
@REM SYTEMS WITH RAM OVER 512 MB TWEAK ENABLE TO FORCE THE CORE TO BE KEPT IN MEMORY AND NOT PAGED TO DISK 
@REM #01
@REM - ENABLE CORE SYSTEM PERFORMANCE 

reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /f /v "DisablePagingExecutive" /t REG_DWORD /d 1 
@ECHO.
@ECHO %VARSTRING%


@REM - DISABLE CORE SYSTEM PERFORMANCE 
@REM reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /f /v "DisablePagingExecutive" /t REG_DWORD /d 0 

@REM #02
@REM INCREASE NTFS PEFORMANCE BY DISABLING THE LAST ACCESS TIME STAMP

reg.exe ADD "HKLM\System\CurrentControlSet\Control\FileSystem" /f /v "NtfsDisableLastAccessUpdate" /t REG_DWORD /d 1 
@ECHO.
@ECHO %VARSTRING%

@REM - SET REG_DWORD /d 0 -- TO UNDO


@REM #03 --------------
@REM ENABLE BOOT DEFRAG FOR FASTER/QUICKER BOOTING
@REM BOOT TIME DEFRAGGER = YES

reg.exe ADD "HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction" /f /v "\Enable" /t REG_SZ /d "Y" 
reg.exe ADD "HKLM\SOFTWARE\Microsoft\Dfrg" /f /v "Enable" /t REG_SZ /d "Y" 
@ECHO.
@ECHO %VARSTRING%

@ECHO OFF
@ECHO.

@REM DISABLE - DELETE - NOT TESTED YET - PUT /f BACK BEFORE /v FOR FORCE DELETE WHEN DONE
@REM reg.exe DELETE "HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction" /v "Enable" 
@REM reg.exe DELETE "HKLM\SOFTWARE\Microsoft\Dfrg" /v "Enable" 

@REM -------------

REM - #04 YOU MIGHT WANT 
REM - OPTIMIZE WINDOWS XP PREFETCHER SETTINGS - MINE IS ENABLED

REM -------------

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
ECHO REGISTRY CHANGES GOT MADE AND 3 DIRECTORYS MADE IF THEY DIDN'T EXIST ALREADY
ECHO ####--------------------------------------------------------------------####

ECHO 1 OF 3 --- "C:\TEMP\#_%COMPUTERNAME%\%USERNAME%"
ECHO 2 OF 3 --- "E:\01 Recent\#_%COMPUTERNAME%\%USERNAME%"
ECHO 3 OF 3 --- "E:\01 History\#_%COMPUTERNAME%\%USERNAME%"

:END_SUB

@ECHO OFF
@IF %VARLATCH% == 0 ECHO #####-----------------------------------------------------####
@IF %VARLATCH% == 0 ECHO NOTHING WAS DONE - NOT A RECOGNISED COMPUTER NAME OR USER NAME
@ECHO ####-----------------------------------------------------------------------####
@ECHO --- READ ME - BATCH FILE FOR - DIY INFO  ----------------------------------####
@ECHO --- FINISHED - CAREFULLY CHECK SUCCESS RESULTS BEFORE HITT KEY END THIS PROGRAM
@PAUSE
@EXIT



REM '----------------------------
REM '-Summary
REM '----------------------------
REM '-Full file path: E:\REG KEY SETTINGS\ASUS EEE NEW USERS SETTING\SPECIAL FOLDERS.BAT
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
Full file path: E:\REG KEY SETTINGS\ASUS EEE NEW USERS SETTING\SPECIAL FOLDERS.BAT
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
