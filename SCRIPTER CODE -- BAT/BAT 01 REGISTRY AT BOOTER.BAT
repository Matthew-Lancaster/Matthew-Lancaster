@ECHO OFF

@REM ----------------------------------------------------------------
@REM __ COMMENT __ @REMMED OUT BUT WONDER ABOUT HERE SEARCH TERM
@REM ----------------------------------------------------------------
@REM WINDOWS 10
@REM DISABLED WINDOWS DEFENDER
@REM WINDOWS 07
@REM WINDOWS XP
@REM ----------------------------------------------------------------

@REM WORK 
@REM ----------------------------------------------------------------
@REM MADE WORK FOR XP MINOR EFFORT
@REM @SETLOCAL EnableExtensions
@REM ---------------------------------------
@REM FROM    [ Sat 15:23:50 Pm_28 Apr 2018 ]
@REM TO      [ Sat 16:07:48 Pm_28 Apr 2018 ] 1 HOUR-
@REM ----------------------------------------------------------------

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


@REM ----------------------------------------------------------------
@ECHO ____ BAT 01 REGISTRY AT BOOTER.BAT ____ 
TITLE %~f0
@CD\
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM __ SET ADMIN A SAFE DOUBLE CHECKER WAY NOT A LOOPY ONE
@REM __ IF CAN'T GET IT TO GO ADMIN DOESN'T RSI INJURY ITSELF AND YOU
@REM ----------------------------------------------------------------

@ECHO OFF

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



@REM ----------------------------------------------------------------
@REM ____ THERE IS A FAULT WITH VISUAL BASIC CAN'T RUN ANYMORE
@REM ____ IT KNOCKS OUT THE SCHEDULER 
@REM ____ AND LEFT WITH HARD WORK TO REINSTALL THE OPPO SYSTEM AND TRY NUMEROUS MEASURE
@REM ____ TALK ABOUT LEAVING A LAST LEGACY FOR CHILDREN TECHNOLOGY IS SO CUTTING EDGE
@REM ____ CAN'T EVEN RUN VBS VB SCRIPTS SOON
@REM ____ PASSING ON TO GENERATIONS DOESN'T HAPPEN WITH MICROSOFT
@REM ----------------------------------------------------------------


@REM __ C:\Program Files\# NO INSTALL REQUIRED\Reg Converter\RegConvert.exe

@REM ----------------------------------------------------------------
@REM -- WHERE OPERATION GET EXECUTED IN BRACKETS FOR IF STATEMENT 
@REM -- THEY ARE EXECUTED AND THE AFTER LEFT THE BRACKET ENCLOSED 
@REM -- SEE THE RESULT FROM OUTPUT OF COMMAND 
@REM ----------------------------------------------------------------

@SETLOCAL DisableDelayedExpansion
@SETLOCAL EnableExtensions

@REM ----------------------------------------------------------------
@REM -- JUMP OVER -- THESE COMMAND CAUSE CMD TO BOMB OUT IN 32 BIT XP
@REM ---------------------------------------------------------------
@ECHO OFF
@REM ---------------------------------------------------------------
ver | find "5.1." > nul
@IF %errorlevel% EQU 0 SET WINXP__=YES
@IF %errorlevel% EQU 0 SET VARLATCH2=WINXP
ver | find "6.1." > nul
@IF %errorlevel% EQU 0 SET WIN07__=YES
@IF %errorlevel% EQU 0 SET VARLATCH2=WIN7
ver | find "10.0." > nul
@IF %errorlevel% EQU 0 SET WIN10__=YES
@IF %errorlevel% EQU 0 SET VARLATCH2=WIN10
@REM ---------------------------------------------------------------

REM IF "%COMPUTERNAME%" == "1-ASUS-X5DIJ" SET VARLATCH2=WINXP
REM IF "%COMPUTERNAME%" == "2-ASUS-EEE" SET VARLATCH2=WINXP
REM IF "%COMPUTERNAME%" == "3-LINDA-PC" SET VARLATCH2=WIN7
REM IF "%COMPUTERNAME%" == "4-ASUS-GL522VW" SET VARLATCH2=WIN10
REM IF "%COMPUTERNAME%" == "5-ASUS-PL2520LA" SET VARLATCH2=WIN7
REM IF "%COMPUTERNAME%" == "7-ASUS-GL522VW" SET VARLATCH2=WIN10
REM IF "%COMPUTERNAME%" == "8-MSI-GP62M-7RD" SET VARLATCH2=WIN10

:X86_WIN_XP_32_BIT_COMPUTER

@REM ---------------------------------------------------------------
@IF "%WINXP__%"=="YES" @ECHO WINDOWS XP
@IF "%WIN07__%"=="YES" @ECHO WINDOWS 7
@IF "%WIN10__%"=="YES" @ECHO WINDOWS 10
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@IF %COMPUTERNAME%==2-ASUS-EEE       SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==1-ASUS-X5DIJ     SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==3-LINDA-PC       SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==4-ASUS-GL522VW   SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==5-ASUS-P2520LA   SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==7-ASUS-GL522VW   SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==8-MSI-GP62M-7RD  SET MATTHEW_S_COMPUTER__=YES
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
REM SET
REM PAUSE
@REM ----------------------------------------------------------------
@REM LOOK AT SET COMMAND __ BOTH REQUIRED SOME COMPUTER SET VER AS x86 
@REM ON %PROCESSOR_ARCHITECTURE%
@REM BUT ONLY IS SET FROM BATCH FILE NOT A CMD CONSOLE WINDOW
@REM ----------------------------------------------------------------
@IF %PROCESSOR_ARCHITECTURE%==AMD64  SET COMPUTER_64BIT__=YES

@IF %COMPUTERNAME%==3-LINDA-PC       SET COMPUTER_64BIT__=YES
@IF %COMPUTERNAME%==4-ASUS-GL522VW   SET COMPUTER_64BIT__=YES
@IF %COMPUTERNAME%==5-ASUS-P2520LA   SET COMPUTER_64BIT__=YES
@IF %COMPUTERNAME%==7-ASUS-GL522VW   SET COMPUTER_64BIT__=YES
@IF %COMPUTERNAME%==8-MSI-GP62M-7RD  SET COMPUTER_64BIT__=YES


@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM PROBLEM READ THIS VAR FROM WINDOWS XP BOMBS OUT BATCH SCRIPT
@REM ----------------------------------------------------------------
@REM @IF %PROCESSOR_ARCHITEW6432%==AMD64  SET COMPUTER_64BIT__=YES
@REM ----------------------------------------------------------------
@ECHO ON

@REM -----------------
@REM -- PAUSE
@REM -----------------
@REM -- MAP OF CONTENT
@REM -----------------
@REM ----------------------------------------------------------------
@REM --- 01 OF 02
@REM ----------------------------------------------------------------
@REM --- E:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\# TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT
@REM ----------------------------------------------------------------
@REM --- 02 OF 02 __ HERE
@REM ----------------------------------------------------------------
@REM --- E:\01 Start Menu\#_7-ASUS-GL522VW\Programs\Startup-01-Delayed\7-ASUS-GL522VW--MATT 01\BAT 03 REGISTRY AT BOOTER.BAT
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM NERO WANTS THIS @REMOVED FROM STARTUP
Reg.exe DELETE "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v "FlickrUploadr" /f
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM -- PAUL BETTS AUTHOR
@REM Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v "FlickrUploadr" /t REG_SZ /d "\"C:\Users\MATT 04\AppData\Local\FlickrUploadrWindows\Update.exe\" --processStart Flickr.exe" /f
@REM ----------------------------------------------------------------


@REM ----------------------------------------------------------------
@REM 64 BIT
@REM ----------------------------------------------------------------
@IF NOT "%COMPUTER_64BIT__%"=="YES" GOTO NOT_64BIT

Reg.exe add "HKCR\Directory\shell\HashMyFiles" /v "LegacyDisable" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Directory\shell\HashMyFiles\command" /ve /t REG_SZ /d "\"C:\PStart\Progs\0_Nirsoft_Package\NirSoft\x64\HashMyFiles.exe\" /file \"%%1\"" /f

Reg.exe add "HKCU\Software\Classes\*\shell\HashMyFiles\command" /ve /t REG_SZ /d "\"C:\PStart\Progs\0_Nirsoft_Package\NirSoft\x64\HashMyFiles.exe\" /file \"%%1\"" /f

@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
:NOT_64BIT

@REM -- SETUP FOR ALL USER
@REM ----------------------------------------------------------------
@REM schtasks /Change /S %COMPUTERNAME% /TN "Adobe Acrobat Update Task" /DISABLE
@REM ----------------------------------------------------------------
@REM schtasks /Change /S %COMPUTERNAME% /TN "DropboxUpdateTaskMachineUA" /ST 00:00 /DU 48:01 /RI 2880
@REM schtasks /Change /S %COMPUTERNAME% /TN "DropboxUpdateTaskMachineCore" /SD 01/01/2010 /ST 00:00 /DU 49:00 /RI 2880

REM @IF NOT "%WIN10__%"=="YES" GOTO NOT_WINDOWS_10
	REM @ECHO WINDOWS 10
	REM @REM Reg.exe ADD " HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell" /f /v "{374DE290-123F-4565-9164-39C4925E467B}" /t REG_SZ /d "C:\Downloads"
	REM Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v "{374DE290-123F-4565-9164-39C4925E467B}" /t REG_EXPAND_SZ /d "C:\Downloads" /f
	REM @REM %USERPROFILE%\Downloads = "%%USERPROFILE%%\Downloads"
REM :NOT_WINDOWS_10

@IF NOT "%WIN10__%"=="YES" GOTO NOT_WINDOWS_10
@ECHO WINDOWS 10
@REM __ SET THE RIBBON AT TOP OF EXPLORER SHOW HER
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Ribbon" /v "MinimizedStateTabletModeOff" /t REG_DWORD /d "0" /f

@REM ---------------------------------------------------------------
@REM __ IN WINDOWS 10 ULTIMATE TWEAKER HAS SETTING
@REM __ This PC
@REM __ 
@REM __ Add/remove useful utilities like Power Options, Recycle Bin, Programs and Features, User Accounts etc. in “This PC”.
@REM __ BUT I CAN'T GET THE REGISTRY KEY SETTING FROM IT
@REM __ WHAT HAPPEN IN EXPLORER THERE ARE THESE SETTING ABOVE FOLDERS IN EXPLORER
@REM __ AND YOU HAVE TO UNCHECK THEM IN ULTIMATE TWEAKER
@REM ---------------------------------------------------------------




:NOT_WINDOWS_10

@REM ----------------------------------------------------------------
@REM -- SET THE AUTO BOOT PROGRAM DELAY_LOADER IN REGISTRY QUICKER BOOTER
@REM -- MORE INFO HERE 
@REM -- E:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\# TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT
@REM ----------------------------------------------------------------
@REM -- CHECK 64 BIT BE
@REM -- DELAYED_LOADER.exe
@REM ----------------------------------------------------------------
@IF NOT "%WIN10__%"=="YES" GOTO NOT_WINDOWS_10
	@ECHO WINDOWS 10
	@REM ----------------------------------------------------------------
	@REM __ NOT WORKER
	@REM __ reg.exe ADD "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" /f /v "DELAY_LOADER_ALL_USER" /t REG_SZ /d "D:\VB6\VB-NT\00_Best_VB_01\DELAYED_LOADER\DELAYED_LOADER.exe"
:NOT_WINDOWS_10
@IF NOT "%WINXP__%"=="YES"  GOTO NOT_WINDOWS_XP
	@ECHO WINDOWS XP
	@REM ----------------------------------------------------------------
	@REM __ WORKER
	@REM __ reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /f /v "DELAY_LOADER_ALL_USER" /t REG_SZ /d ""D:\VB6\VB-NT\00_Best_VB_01\DELAYED_LOADER\DELAYED_LOADER.exe""
:NOT_WINDOWS_XP




@REM IS A UNICODE JOB __ USE TYPE COMMAND SAVE OUTPUT TO A FILE
@SET "VALUE_RETURN="
schtasks /QUERY /TN "SYSTEM LOGON.BAT" >NUL 2>%TEMP%\_TEST_2.TXT
@FOR /F "tokens=1,2 delims= " %%A IN ('TYPE %TEMP%\_TEST_2.TXT') DO @IF "%%A"=="ERROR:" SET VALUE_RETURN=YES_REQUIRE_SETTING
@DEL %TEMP%\_TEST_2.TXT

@IF "%VALUE_RETURN%"=="" GOTO NOT_HERE
schtasks /CREATE /RU ADMINISTRATORS /TN "\SYSTEM LOGON.BAT" /TR "C:\22_LOGON.BAT" /SC ONLOGON /RL HIGHEST /F
@REM Schtasks /DELETE /S %COMPUTERNAME% /TN "\SYSTEM LOGON.BAT"

:NOT_HERE



@REM ---------------------------------------------------
@REM -- ONLY TO RUN HERE IN BRACKET @IF MATTHEW_S MACHINE
@REM -- OWN CUSTOM CODE ON RIGHT CLICKER
@REM -- ONE THING TO NOTE RIGHT CLICKER PASSES THE FILE OVER 
@REM -- LIKE SEND_TO BUT DOES IT IN THE SHORT NAME FORMAT
@REM -- LITTLE CONFUSER FOR A BIT 
@REM ---------------------------------------------------
@REM ---------------------------------------------------
@IF NOT "%MATTHEW_S_COMPUTER__%"=="YES" GOTO NOT_MATTHEW_S_MACHINE
	@ECHO MATTHEW_S_COMPUTER__

	@REM MOVING THINGS AROUND AND OLDER CODE LIKE IT THERE
	@REM THIS DONT WORK TO CALL PROGRAM EXE
	@REM mklink /j "C:\Program Files\# NO INSTALL REQUIRED" "C:\PStart\# NOT INSTALL REQUIRED"
	@REM RMDIR TO @REMOVE NOT DEL
	
	@REM MKDIR "C:\PROGRAM FILES\# NO INSTALL REQUIRED"

	Reg.exe add "HKCR\Folder\shell\MenuText" /ve /t REG_SZ /d "Open a Command Prompt Here" /f
	Reg.exe add "HKCR\Folder\shell\MenuText\Command" /ve /t REG_SZ /d "cmd.exe /K chdir /D \"%%1\"" /f
	

	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	@REM __ TO SET MY PERSONAL PROGRAM COMMAND IN THE CONTEXT MENU FOR 
	@REM _________________________________________________________________
	@REM __ OPEN A FILE IN EXPLORER
	@REM _________________________________________________________________
	@REM __ IN-CASE WANT SOME RENAMING FIRST
	@REM __ DELETE AND RECREATE
	@REM _________________________________________________________________
	
	reg.exe DELETE "HKLM\Software\Classes\Drive\shell\Multi -------- Drive XXXXXXXXXXXXXXXXXXXXX" /f
	reg.exe DELETE "HKLM\Software\Classes\Directory\shell\Multi -------- Directory XXXXXXXXXXXXX" /f
	Reg.exe DELETE "HKLM\Software\Classes\Directory\Background\shell\Multi -------- XXXXXXXXXXXX" /f

	SET NAME_V1=DATE MODIFY ____
	SET NAME_V2=%NAME_V1% FOLDER
	SET NAME_V3=D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe
	reg.exe DELETE "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%" /f
	reg.exe DELETE "HKLM\Software\Classes\Folder\shell\%NAME_V2%" /f
	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%" /ve /t REG_SZ /d "%NAME_V1%" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%" /v "icon" /t REG_SZ /d "%NAME_V3%" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%\Command" /ve /t REG_SZ /d "\"%NAME_V3%\" \"%%1\"" /f
	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\Software\Classes\Folder\shell\%NAME_V2%" /ve /t REG_SZ /d "%NAME_V2%" /f
	Reg.exe add "HKLM\Software\Classes\Folder\shell\%NAME_V2%" /v "icon" /t REG_SZ /d "%NAME_V3%" /f
	Reg.exe add "HKLM\Software\Classes\Folder\shell\%NAME_V2%\Command" /ve /t REG_SZ /d "\"%NAME_V3%\" \"%%1\"" /f
	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\Software\Classes\Directory\Background\shell\%NAME_V1%" /ve /t REG_SZ /d "%NAME_V1%" /f
	Reg.exe add "HKLM\Software\Classes\Directory\Background\shell\%NAME_V1%" /v "icon" /t REG_SZ /d "%NAME_V3%" /f
	Reg.exe add "HKLM\Software\Classes\Directory\Background\shell\%NAME_V1%\Command" /ve /t REG_SZ /d "\"%NAME_V3%\" \"%%1\"" /f
	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	
	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	SET NAME_V1=DATE MODIFY ____
	SET NAME_V2=%NAME_V1% FOLDER
	SET NAME_V3=D:\VB6\VB-NT\00_Send_To\Send To TouchDate\#Send To TouchDate.exe
	reg.exe DELETE "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%" /f
	reg.exe DELETE "HKLM\Software\Classes\Folder\shell\%NAME_V2%" /f
	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%" /ve /t REG_SZ /d "%NAME_V1%" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%" /v "icon" /t REG_SZ /d "%NAME_V3%" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%\Command" /ve /t REG_SZ /d "\"%NAME_V3%\" \"%%1\"" /f
	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\Software\Classes\Folder\shell\%NAME_V2% Folder" /ve /t REG_SZ /d "%NAME_V2%" /f
	Reg.exe add "HKLM\Software\Classes\Folder\shell\%NAME_V2% Folder" /v "icon" /t REG_SZ /d "%NAME_V3%" /f
	Reg.exe add "HKLM\Software\Classes\Folder\shell\%NAME_V2% Folder\Command" /ve /t REG_SZ /d "\"%NAME_V3%\" \"%%1\"" /f
	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\Software\Classes\Directory\Background\shell\%NAME_V1%" /ve /t REG_SZ /d "%NAME_V1%" /f
	Reg.exe add "HKLM\Software\Classes\Directory\Background\shell\%NAME_V1%" /v "icon" /t REG_SZ /d "%NAME_V3%" /f
	Reg.exe add "HKLM\Software\Classes\Directory\Background\shell\%NAME_V1%\Command" /ve /t REG_SZ /d "\"%NAME_V3%\" \"%%1\"" /f
	REM --------------------------------------------------------------
	REM --------------------------------------------------------------




	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	REM ONE MORE FOR GOOD MEASURE
	REM NOT SURE WHAT ONE IS PUTTING IN ENTRY FOR CMD.COM BUT FORGETTING ICON
	REM WHEN ICON IS PUT IN FROM SYSTEM32 IT CHANGES IT TO THIS ONE C:\Windows\SysWow64\cmd.exe
	REM -- HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\background\shell\cmd\command
	Reg.exe add "HKLM\SOFTWARE\Classes\Directory\background\shell\cmd" /v "icon" /t REG_SZ /d "C:\Windows\SysWow64\cmd.exe" /f
	
	REM ANOTHER FEW BUT SOME BLOCKED LIKE CMD.EXE
	REM -------------------------------------------------------------
	Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\shell\Browse with &IrfanView" /v "icon" /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe" /f
	
	Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\shell\Browse with XnView" /v "icon" /t REG_SZ /d "C:\Program Files (x86)\XnView\xnview.exe" /f
	
	Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\shell\Powershell" /v "icon" /t REG_SZ /d "C:\Windows\SysWow64\WindowsPowerShell\v1.0\powershell.exe" /f
	
	REM AND MORE IN FOLDER
	REM -------------------------------------------------------------
	Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\*\shell\filealyzer2" /v "icon" /t REG_SZ /d "C:\Program Files (x86)\Safer Networking\FileAlyzer 2\FileAlyzer2.exe" /f
	
	Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\shell\HashMyFiles" /v "icon" /t REG_SZ /d "C:\PStart\Progs\0_Nirsoft_Package\NirSoft\x64\HashMyFiles.exe" /f
	
	rem can't write to powershell in drive
	
	Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Drive\shell\Browse with &IrfanView" /v "icon" /t REG_SZ /d "C:\Program Files (x86)\IrfanView\i_view32.exe" /f
	
	Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Drive\shell\Browse with XnView" /v "icon" /t REG_SZ /d "C:\Program Files (x86)\XnView\xnview.exe" /f
	
	REM ADD ICON TO THIS ONE DIDN'T DO ANYTHING _ SEEN YET _ TAKEN OUT
	REM HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Folder\shell\explore
	REM -------------------------------------------------------------
	REM THAT MUCKED UP MY DESKTOP ICONS NOW GONE BLACK _ TRY REBOOT
	EM DESKTOP CAME BACK AFTER WORKING ON SOMETHING OTHER
	REM -------------------------------------------------------------
	
	
	@REM ______________________________________________________________________
	@REM __ OPEN IN ____ Directory and FOLDER APPEAR TO BE SLIGHTLY THE SAME AT FIRST LOOK
	@REM ______________________________________________________________________
	@REM __ IS NOT MUCH LEARN FOR FOLDER SO MIGHT BE BETTER LEARN DIRECTORY
	@REM __ MOSTLY HOW EXPLORER USE HER GUESS
	@REM ______________________________________________________________________
	@REM __ THE DIRECTORY HAPPEN WHEN FILE OF /SELECT EXPLORER LIKE RIGHT HAND PANE
	@REM __ AND NOT ON ROOT  THINKER SORTING
	@REM __ WHEN QUOTE AROUND LONG STRING WITH OTHER QUOTER USE SLASH IN FRONT OF
	@REM __ THE ONE INSIDE LIKE HERE (" /" ")
	@REM __ AND THE % IN BATCH FILE HAS TO BE DOUBLE TO MAKE SHOW AS ONE
	@REM ______________________________________________________________________

	Reg.exe add "HKLM\Software\Classes\Directory\shell\Multi ---- Directory" /ve /t REG_SZ /d "Multi ---- Directory" /f
	Reg.exe add "HKLM\Software\Classes\Directory\shell\Multi ---- Directory" /v "icon" /t REG_SZ /d "D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe" /f
	Reg.exe add "HKLM\Software\Classes\Directory\shell\Multi ---- Directory\Command" /ve /t REG_SZ /d "\"D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe\" \"%%1\"" /f

	Reg.exe add "HKLM\Software\Classes\Drive\shell\Multi ---- Drive" /ve /t REG_SZ /d "Multi ---- Drive" /f
	Reg.exe add "HKLM\Software\Classes\Drive\shell\Multi ---- Drive" /v "icon" /t REG_SZ /d "D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe" /f
	Reg.exe add "HKLM\Software\Classes\Drive\shell\Multi ---- Drive\Command" /ve /t REG_SZ /d "\"D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe\" \"%%1\"" /f

	@REM __ HKLM\SOFTWARE\Classes\*
	@REM __ HKLM\Software\Classes\Folder\shell\yourappname\command
	@REM __ HKLM\Software\Classes\Directory\shell\yourappname\command
	@REM __ HKLM\Software\Classes\Drive\shell\yourappname\command
	
	
	
	
	@REM -- MY PERSONAL PREF SOME APP LIKE GOOD SYNC TAKE A WHILE TO WIND IT IN
	@REM -- TURN OFF AUTOMATIC APPLICATIONS TERMINATION AT SHUTDOWN
	@REM -- ULTIMATE TWEAKER
	Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /v "AllowBlockingAppsAtShutdown" /t REG_DWORD /d "1" /f
	
	@REM -- THATS ALL OF ULTIMATE TWEAKER WINDOWS 10 __ DONE
:NOT_MATTHEW_S_MACHINE


@GOTO JUMPER_WRAP_AROUND_TEXT_WRAPPER

@REM ________________________________________________________________
@REM ________________________________________________________________
@REM THE CODE I USE FOR DEBUGGING
@REM REGISTRY JUMPER
@REM HERE ARE
@REM C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\REGJUMP_CLIP_BOARD.BAT

@REM REQUIRED ALSO IS THE EXE COMMAND USER
@REM regjump.exe
@REM FROM SysinternalsSuite

@REM ________________________________________________________________
@REM IS A GOOD JUMP TO NEXT LEARN
@REM HAS A TRIM EFFECT 
@REM NOT ALLOW ADDITION SPACER AT FRONT BUT ALLOW ANYTHING JUNK AFTER
@REM COPY AND PASTE A REG STRING AND JUMP TO
@REM ________________________________________________________________
@REM BEGINNER WITHER
@REM ________________________________________________________________
@REM HKLM
@REM HKCU
@REM HKEY_CLASSES_ROOT
@REM ----------------------------------------------------------------
@REM AND A FEW OTHER
@REM TO RECOGNISE @IF WANT TO
@REM ----------------------------------------------------------------
@REM I USE A FRONT END IN VB TO RECOGNISE THE REG ENTRY A SELECTOR THAT USER
@REM AND HTMP WEB PAGES
@REM AND EXPLORER FILE NAME
@REM AND OTHER
@REM ________________________________________________________________

@REM ________________________________________________________________
@REM OFF TOPIC SEMI LEARN SWITCH SUBJECT TO DO WITH REGISTRY
@REM AND THE STARTUP CONFIG SETTING
@REM FOR SYNCRONISOR THE INI FILE OF ALL PROGRAM FILE FOLDER
@REM ------------------------------------------------------------------
@REM ANOTHER GOOD ONE TO SPY ON PROGRAM AS THEY USER ITEM FILE OPEN CLOSE
@REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\ProcessActivityView.exe
@REM ________________________________________________________________


@REM ________________________________________________________________
@REM ________________________________________________________________
@ECHO OFF
@REM "regedit jump to"
@"C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\regjump.exe" -c | findstr /I "Complete." | findstr /I "Not a valid registry path"
@IF %errorlevel% == 0 (
    @ECHO OKAY REGISTRY PATH ----
) else (
    @ECHO.
    @ECHO ---------------------------------------
    @ECHO ERROR REGISTRY PATH                   -
    @ECHO ---------------------------------------
    @ECHO.
    "C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\regjump.exe" -c
    @ECHO.
    @ECHO ---------------------------------------
    @ECHO.
    PAUSE
)
@REM @ECHO DONE ----
@REM PAUSE

@REM ________________________________________________________________
@REM ________________________________________________________________

:JUMPER_WRAP_AROUND_TEXT_WRAPPER




@REM NERO_TUNEITUP
@REM NERO_TuneItUp
@REM NEROTuneItUp
@REM STARTED WORKING HERE
@REM AND FIRST OFF __ CHROME IS NOT IMPORTANT TO BE STARTED AT STARTUP
@REM CREDIT DUE TO THIS PAGE
@REM ----
@REM NT's FOR /F command: tokens and delims
@REM http://www.robvanderwoude.com/ntfortokens.php
@REM ----
@REM THIS IS A NICE CODE TO DETECT LINES AT A TIME FOR A SEARCH STRING 
@REM THAT IS IN THE TOKEN 2ND SLOT SPACE SEPARATOR ASK FOR A QUERY AND RETURNS ITSELF 
@REM WITH HERE REG_SZ IN THE 2ND SEGMENT REG_SZ THE 1ST HAS WHAT USER IS ASK WANTING REQUESTER 
@REM ONLY THAT MIDDLE LINE WILL DO 
@REM THE 1ST IS THE QUERY HKCU\SOFTWARE\ SECOND IS THE INFO WANTING IF ONE OCCURRENCE 3RD IS END OF SEARCH


@SET "VALUE_RETURN="
@FOR /F "tokens=1,2 delims= " %%A IN ('REG.EXE QUERY HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /S /F GoogleChromeAutoLaunch_ /T REG_SZ') DO @IF "%%B"=="REG_SZ" SET VALUE_RETURN=%%A
@if "%VALUE_RETURN%"=="" GOTO NOT_HERE
reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v "%VALUE_RETURN%" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /f
:NOT_HERE

@REM rem ------------------------------------------------------------------
@REM rem nero tuneitup recommend NOT to run Google at start
@REM rem ------------------------------------------------------------------
@REM Reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v "GoogleChromeAutoLaunch_122671188F3391FD20739CBFFB5392E6" /f
@REM Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /f

@REM rem ------------------------------------------------------------------
@REM rem nero tuneitup recommend NOT to run FREEMAKE UPDATE at start
@REM rem ------------------------------------------------------------------

@SET "VALUE_RETURN="
@FOR /F "tokens=1,2 delims= " %%A IN ('REG.EXE QUERY HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run /S /F ProductUpdater /T REG_SZ') DO @IF "%%B"=="REG_SZ" SET VALUE_RETURN=%%A
@if "%VALUE_RETURN%"=="" GOTO NOT_HERE
reg.exe delete "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" /v "%VALUE_RETURN%" /f
Reg.exe add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" /f
:NOT_HERE


@REM rem ------------------------------------------------------------------
@REM rem nero tuneitup recommend NOT to run VideoDownloaderUltimate at start
@REM rem ------------------------------------------------------------------

@SET "VALUE_RETURN="
@FOR /F "tokens=1,2 delims= " %%A IN ('REG.EXE QUERY HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /S /F VideoDownloaderUltimate /T REG_SZ') DO @IF "%%B"=="REG_SZ" SET VALUE_RETURN=%%A
@if "%VALUE_RETURN%"=="" GOTO NOT_HERE
reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v "%VALUE_RETURN%" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /f
:NOT_HERE




@IF NOT EXIST "C:\Program Files\Notepad++\Notepad++.exe" (
	@IF NOT EXIST "C:\Program Files (x86)\Notepad++\Notepad++.exe" (
		GOTO DONT_WANT_NOTEPAD_CHANGER
	)
)


@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM NOTEPAD++ CONTEXT MENU -----------------------------------------
@REM ----------------------------------------------------------------
@REM __ #01_ CREATE AND DELETE THE REG_KEY WITH NOTEPAD++ _ SET ADMINISTRATOR SAVE.BAT

@ECHO.
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO.
@IF EXIST "C:\Program Files\Notepad++\Notepad++.exe" (
	@ECHO NOTEPAD++ SETTING __ FOR 64 BIT NOTEPAD++ NAMELY FOR COMPUTER NAME __ 8-MSI-GP62M-7RD
	@ECHO.
)
@IF NOT EXIST "C:\Program Files\Notepad++\Notepad++.exe" (
	@ECHO NOTEPAD++ SETTING __ FOR 32 BIT NOTEPAD++  __ USUAL SETTING 
	@ECHO.
)

@ECHO 64 BIT NOTEPAD++ ABLE HOLD LARGER FILE BUT PROBLEM ACCESS EXAMPLE TXT DOCUMENT IN CERTAIN 64 BIT FOLDER DUPE OF ANOTHER 32 BIT ESPCIALLY IN WINDOWS FOLDER _ SOLUTION WAS FOUND
@ECHO.
@ECHO ALSO 32 BIT OS SYSTEM IS SAME FOLDER \PROGRAM FILES\
@ECHO.
@ECHO ALSO 64 BIT NOTEPAD++ LESS PLUGINS
@ECHO.
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO.


@ECHO OFF
@REM CMD/C IS USED BECAUSE ONLY INTERNAL COMMANDS CAN BE CARRIED OUT BY USE IF STATEMENT AND INCLUDE Parenthesis ()
@REM IGNOR THAT
@REM THE PROBLEM OF CHO SOME Parenthesis () WITHIN THE Parenthesis () ALREADY BY IF INCLUDE ^ EXAMPLE Program Files ^(x86^)
@REM AND THEN STAGER TO ANOTHER PROBLEM 
@REM EASIER WITH A JUMP GOTO

REM SET PROGRAM_FILES_32_BIT=C:\Program Files (x86)
REM SET PROGRAM_FILES_64_BIT=C:\Program Files

REM @IF NOT EXIST "C:\Program Files\Notepad++\Notepad++.exe" (
	REM @ECHO NOTEPAD++ 32 BIT __ C:\Program Files ^(x86^)\
	REM @ECHO.
	REM Reg.exe add "HKCR\*\shell\OpenWithNotepad" /v "icon" /t REG_SZ /d "C:\Program Files (x86)\Notepad++\Notepad++.exe" /f
	REM Reg.exe add "HKCR\*\shell\OpenWithNotepad\Command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\Notepad++\Notepad++.exe\" \"%%1\"" /f
REM )

REM PAUSE
	
REM @IF EXIST "C:\Program Files\Notepad++\Notepad++.exe" (
	REM @ECHO NOTEPAD++ 64 OR 32 BIT __ C:\Program Files\
	REM @ECHO.
	REM CMD /C "Reg.exe add "HKCR\*\shell\OpenWithNotepad" /v "icon" /t REG_SZ /d "C:\Program Files\Notepad++\Notepad++.exe" /f"
	REM CMD /C "Reg.exe add "HKCR\*\shell\OpenWithNotepad\Command" /ve /t REG_SZ /d "\"C:\Program Files\Notepad++\Notepad++.exe\" \"%%1\"" /f"
REM )
@ECHO ON

@IF EXIST "C:\Program Files\Notepad++\Notepad++.exe"       GOTO PROGRAM_FILES_64_OR_32_BIT
@IF EXIST "C:\Program Files (x86)\Notepad++\Notepad++.exe" GOTO PROGRAM_FILES_32_BIT

GOTO DONT_WANT_NOTEPAD_CHANGER

:PROGRAM_FILES_64_OR_32_BIT
@ECHO NOTEPAD++ 64 OR 32 BIT __ C:\Program Files\
@ECHO.
Reg.exe add "HKCR\*\shell\OpenWithNotepad" /ve /t REG_SZ /d "Edit with &Notepad++" /f
Reg.exe add "HKCR\*\shell\OpenWithNotepad" /v "icon" /t REG_SZ /d "C:\Program Files\Notepad++\Notepad++.exe" /f
Reg.exe add "HKCR\*\shell\OpenWithNotepad\Command" /ve /t REG_SZ /d "\"C:\Program Files\Notepad++\Notepad++.exe\" \"%%1\"" /f

GOTO END_OF_NOTEPAD_ROUTINE

:PROGRAM_FILES_32_BIT
@ECHO NOTEPAD++ 32 BIT __ C:\Program Files (x86)\
@ECHO.
Reg.exe add "HKCR\*\shell\OpenWithNotepad" /ve /t REG_SZ /d "Edit with &Notepad++" /f
Reg.exe add "HKCR\*\shell\OpenWithNotepad" /v "icon" /t REG_SZ /d "C:\Program Files (x86)\Notepad++\Notepad++.exe" /f
Reg.exe add "HKCR\*\shell\OpenWithNotepad\Command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\Notepad++\Notepad++.exe\" \"%%1\"" /f

@ECHO OFF
@REM Reg.exe DELETE "HKCR\*\shell\OpenWithNotepad" /f

@REM ---------------------------------------------------------------------
@REM NOTEPAD++ CONTEXT MENU ITEM WORKS BETTER AT HKLM\ RATHER THAN HKCR\
@REM SOMETHING ON MY COMPUTER NEW SETTER HAD BOTH
@REM SOMETIME ON A  @REMOTE DRIVE VIA A LINK NOTEPAD++ WOULDN'T OPEN TO IT 
@REM HKLM DOES IT BETTER
@REM OUT WITH THE OLD IN WITH THE NEW
@REM SLIGHTLY DIFFERENT PATH BUT BESIDES REPLACE A BIT TREE STRUCTURE THE SAME
@REM ---------------------------------------------------------------------

@REM ---------------------------------------------------------------------
@REM NICE IDEA TO USE HKLM INSTEAD HKCR BUT DISAPPEARED AFTER A BIT
@REM KEEP BOTH GOING FOR A WHILE
@REM ---------------------------------------------------------------------
@ECHO ON


@IF EXIST "C:\Program Files\Notepad++\Notepad++.exe"       GOTO PROGRAM_FILES_64_OR_32_BIT
@IF EXIST "C:\Program Files (x86)\Notepad++\Notepad++.exe" GOTO PROGRAM_FILES_32_BIT

GOTO DONT_WANT_NOTEPAD_CHANGER

:PROGRAM_FILES_64_OR_32_BIT
	@ECHO NOTEPAD++ 64 OR 32 BIT __ C:\Program Files\
	@ECHO.
	Reg.exe add "HKLM\SOFTWARE\Classes\CLSID\*\shell\OpenWithNotepad" /ve /t REG_SZ /d "Edit with &Notepad++" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\CLSID\*\shell\OpenWithNotepad" /v "icon" /t REG_SZ /d "C:\Program Files\Notepad++\Notepad++.exe" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\CLSID\*\shell\OpenWithNotepad\Command" /ve /t REG_SZ /d "\"C:\Program Files\Notepad++\Notepad++.exe\" \"%%1\"" /f
GOTO END_OF_NOTEPAD_ROUTINE

:PROGRAM_FILES_32_BIT
	@ECHO NOTEPAD++ 32 BIT __ C:\Program Files (x86)\
	@ECHO.
	Reg.exe add "HKLM\SOFTWARE\Classes\CLSID\*\shell\OpenWithNotepad" /ve /t REG_SZ /d "Edit with &Notepad++" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\CLSID\*\shell\OpenWithNotepad" /v "icon" /t REG_SZ /d "C:\Program Files (x86)\Notepad++\Notepad++.exe" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\CLSID\*\shell\OpenWithNotepad\Command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\Notepad++\Notepad++.exe\" \"%%1\"" /f

:END_OF_NOTEPAD_ROUTINE

@for /f "delims=" %%a in ('reg query "HKLM\SOFTWARE\Classes\CLSID" /s /f "ANotepad++64" /t REG_SZ') do @echo %%a|findstr /v "REG_SZ" | findstr /v /c:"End of search:" >nul && reg.exe delete "%%a" /f

@for /f "delims=" %%a in ('reg query "HKCR\CLSID" /s /f "ANotepad++64" /t REG_SZ') do @echo %%a|findstr /v "REG_SZ" | findstr /v /c:"End of search:" >nul && reg.exe delete "%%a" /f

:DONT_WANT_NOTEPAD_CHANGER

@ECHO.
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO.


@REM PAUSE
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM CANNOT SEEM TO GET EDIT TO CALL UP NOTEPAD FOR EDIT A BATCH FILE
@REM HAPPEN AFTER COMMAND SET ABOVE NOTEPAD++
@REM REBOOT REQUIRED HERE TO SEE CHANGES
@REM ----------------------------------------------------------------
@REM ;File association fix for [.BAT] - Windows 7
@REM ;http://www.winhelponline.com/blog
@IF NOT "%WIN07__%"=="YES" GOTO NOT_WINDOWS_07
	@ECHO WINDOWS 07
	Reg.exe add "HKCR\.BAT" /ve /t REG_SZ /d "batfile" /f
	Reg.exe add "HKCR\.BAT\PersistentHandler" /ve /t REG_SZ /d "{5e941d80-bf96-11cd-b579-08002b30bfeb}" /f
	Reg.exe add "HKCR\batfile" /ve /t REG_SZ /d "Windows Batch File" /f
	Reg.exe add "HKCR\batfile" /v "EditFlags" /t REG_BINARY /d "30040000" /f
	Reg.exe add "HKCR\batfile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\acppage.dll,-6002" /f
	Reg.exe add "HKCR\batfile\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\imageres.dll,-68" /f
	Reg.exe add "HKCR\batfile\shell\edit\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE %%1" /f
	Reg.exe add "HKCR\batfile\shell\open" /v "EditFlags" /t REG_BINARY /d "00000000" /f
	Reg.exe add "HKCR\batfile\shell\open\command" /ve /t REG_SZ /d "\"%%1\" %%*" /f
	Reg.exe add "HKCR\batfile\shell\print\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE /p %%1" /f
	Reg.exe add "HKCR\batfile\shell\runas" /v "HasLUAShield" /t REG_SZ /d "" /f
	Reg.exe add "HKCR\batfile\shell\runas\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\cmd.exe /C \"%%1\" %%*" /f
	Reg.exe add "HKCR\batfile\shell\runasuser" /ve /t REG_SZ /d "@shell32.dll,-50944" /f
	Reg.exe add "HKCR\batfile\shell\runasuser" /v "Extended" /t REG_SZ /d "" /f
	Reg.exe add "HKCR\batfile\shell\runasuser" /v "SuppressionPolicyEx" /t REG_SZ /d "{F211AA05-D4DF-4370-A2A0-9F19C09756A7}" /f
	Reg.exe add "HKCR\batfile\shell\runasuser\command" /v "DelegateExecute" /t REG_SZ /d "{ea72d00e-4960-42fa-ba92-7792a7944c1d}" /f
	Reg.exe add "HKCR\batfile\ShellEx\ContextMenuHandlers" /ve /t REG_SZ /d "Compatibility" /f
	Reg.exe add "HKCR\batfile\ShellEx\ContextMenuHandlers\Compatibility" /ve /t REG_SZ /d "{1d27f844-3a1f-4410-85ac-14651078412d}" /f
	Reg.exe add "HKCR\batfile\ShellEx\DropHandler" /ve /t REG_SZ /d "{86C86720-42A0-1069-A2E8-08002B30309D}" /f
	Reg.exe add "HKCR\batfile\ShellEx\PropertySheetHandlers\PifProps" /ve /t REG_SZ /d "{86F19A00-42A0-1069-A2E9-08002B30309D}" /f
	Reg.exe add "HKCR\batfile\ShellEx\PropertySheetHandlers\ShimLayer Property Page" /ve /t REG_SZ /d "{513D916F-2A8E-4F51-AEAB-0CBC76FB1AF8}" /f
	Reg.exe add "HKCR\batfile\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}" /ve /t REG_SZ /d "{1531d583-8375-4d3f-b5fb-d23bbd169f22}" /f
	Reg.exe delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.BAT\UserChoice" /f
	Reg.exe add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.BAT\OpenWithProgids" /v "batfile" /t REG_NONE /d "" /f
	@REM ----------------------------------------------------------------
	@REM ----------------------------------------------------------------
:NOT_WINDOWS_07

@IF NOT "%WIN10__%"=="YES" GOTO NOT_WINDOWS_10
	@ECHO WINDOWS 10
	rem ------------------------------------------
	Reg.exe add "HKCR\.bat" /ve /t REG_SZ /d "batfile" /f
	Reg.exe add "HKCR\.bat\PersistentHandler" /ve /t REG_SZ /d "{5e941d80-bf96-11cd-b579-08002b30bfeb}" /f
	Reg.exe add "HKCR\batfile" /ve /t REG_SZ /d "Windows Batch File" /f
	Reg.exe add "HKCR\batfile" /v "EditFlags" /t REG_BINARY /d "30040000" /f
	Reg.exe add "HKCR\batfile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\acppage.dll,-6002" /f
	Reg.exe add "HKCR\batfile\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\imageres.dll,-68" /f
	Reg.exe add "HKCR\batfile\shell\edit\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE %%1" /f
	Reg.exe add "HKCR\batfile\shell\open" /v "EditFlags" /t REG_BINARY /d "00000000" /f
	Reg.exe add "HKCR\batfile\shell\open\command" /ve /t REG_SZ /d "\"%%1\" %%*" /f
	Reg.exe add "HKCR\batfile\shell\print\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE /p %%1" /f
	Reg.exe add "HKCR\batfile\shell\runas" /v "HasLUAShield" /t REG_SZ /d "" /f
	Reg.exe add "HKCR\batfile\shell\runas\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\cmd.exe /C \"%%1\" %%*" /f
	Reg.exe add "HKCR\batfile\shell\runasuser" /ve /t REG_SZ /d "@shell32.dll,-50944" /f
	Reg.exe add "HKCR\batfile\shell\runasuser" /v "Extended" /t REG_SZ /d "" /f
	Reg.exe add "HKCR\batfile\shell\runasuser" /v "SuppressionPolicyEx" /t REG_SZ /d "{F211AA05-D4DF-4370-A2A0-9F19C09756A7}" /f
	Reg.exe add "HKCR\batfile\shell\runasuser\command" /v "DelegateExecute" /t REG_SZ /d "{ea72d00e-4960-42fa-ba92-7792a7944c1d}" /f
	Reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.bat" /f
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.bat\OpenWithProgids" /v "batfile" /t REG_NONE /d "" /f

	@REM ----------------------------------------------------------------
	@REM ----------------------------------------------------------------
:NOT_WINDOWS_10




@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM __ STOP THE _ GET ASK QUESTION TO RUN ADMIN PROGRAMS
@REM __ DOUBLE GOOD EFFECT SETTING TO NAUGHT EVEN IN HOME EDITION RUNS ALL CODE AS ADMIN AND ALSO WITHOUT QUESTION
@REM ----------------------------------------------------------------
@REM ----
@REM Disable the User Account Control by Registry - Google Search
@REM https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB744GB744&ei=lkceWqDkC-Pm_QaIirWwCg&q=Disable+the+User+Account+Control+by+Registry&oq=Disable+the+User+Account+Control+by+Registry&gs_l=psy-ab.3..35i39k1.1797.2326.0.2866.2.2.0.0.0.0.69.127.2.2.0....0...1c.1.64.psy-ab..0.2.126....0.JCMJZtWI4Z0
@REM ----
@REM ----
@REM TechNet Registry Key to Disable UAC
@REM https://gallery.technet.microsoft.com/Registry-Key-to-Disable-UAC-45d0df25
@REM ----
@REM ----
@REM How To (Really) Completely Disable UAC on Windows 7
@REM https://www.howtogeek.com/howto/4820/how-to-really-completely-disable-uac-on-windows-7/
@REM ----

@REM _ OR BY CONTROL PANEL __ USER SETTING
@REM Click on the “Change User Account Control Settings” link. To turn off UAC, move the slider to the Never notify position, and then click OK.

@REM REG_DWORD DEFAULT SET TO 1

@REM ----------------------------------------------------------------
@REM FUNNY SITUATION AS THE SLIDER DOES NOT WORK THE QUESTION NEVER GET ASK IF WANT TO ELEVATE 
@REM IF IN A USER THAT HAS STANDARD PERMISSION
@REM IF RUNNING A SCRIPT TO ELEVATE IT WON'T KNOW SUCCESS AND GO IN A LOOP UNLESS MORE 
@REM WORK DONE
@REM BUT IT SAVES HAVING TO SET ELEVATED PERMISSION FOR EVERY PROGRAM WANTS IT
@REM ----------------------------------------------------------------
@REM NEW 2018 CHANGE BEEN SET FROM UAC OFF TO UAC ON 
@REM CAN'T OPEN WINDOWS STORE IF USER ACCOUNT CONTROL SWITCH IS SET OFF DISABLED
@REM REBOOT REQUIRED ON CHANGES
@REM THERE ARE TWO SETTING FOR THIS IN HERE .BAT FILE ONE BELOW BEEN REMMER
@REM AND ALSO HAS TWO REGISTRY SETTING ARE MADE FOR THIS ONE
@REM ----------------------------------------------------------------

@IF %COMPUTERNAME%==7-ASUS-GL522VW  GOTO JUMP_NOT_HERE
@REM 2018 MAY ---------------------------------------------------------------------
@REM THIS COMPUTER HAS DEVELOPED A FUNNY PROBLEM WHERE IT DOESN'T LIKE THIS SETTING
@REM AND MICROSFOT SHOP WON'T OPEN IF SET IT
@REM ------------------------------------------------------------------------------

Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v "NoDispCPL" /t REG_DWORD /d "0" /f

@REM __ DISABLE UAC USER ACCOUNT CONTROL
reg.exe ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /f /v "EnableLUA" /t REG_DWORD /d 0

@REM __ ENABLE UAC USER ACCOUNT CONTROL
REM reg.exe ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /f /v "EnableLUA" /t REG_DWORD /d 1

:JUMP_NOT_HERE

@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------


@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM TAKE THE SHORTCUT APPENDED AWAY FROM COPY TO A SHORTCUT
@REM DEFAULT ENABLED VALUE "19000000"
@REM ----------------------------------------------------------------
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /v "Link" /t REG_BINARY /d "00000000" /f
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM SHOW WINDOWS VERSION ON DESKTOP
@REM ----------------------------------------------------------------
Reg.exe add "HKCU\Control Panel\Desktop" /v "PaintDesktopVersion" /t REG_DWORD /d "1" /f
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------


@IF NOT "%WIN10__%"=="YES" GOTO NOT_WINDOWS_10
	@ECHO WINDOWS 10

	@REM ----------------------------------------------------------------
	@REM ----------------------------------------------------------------
	@REM SHOW SECONDS IN THE RIGHT HAND CORNER SYSTEM CLOCK
	@REM AND DEFAULT SET A COUPLE OPTION FOR CLOCK TO MAKE SURE SHOWING
	@REM ULTIMATE TWEAKER
	@REM ----------------------------------------------------------------
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "ShowSecondsInSystemClock" /t REG_DWORD /d "1" /f
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "HideClock" /t REG_DWORD /d "0" /f
	@REM ----------------------------------------------------------------
	@REM ----------------------------------------------------------------

	@REM ----------------------------------------------------------------
	@REM ----------------------------------------------------------------
	@REM VerboseStatus = ENABLED GIVE INFO ABOUT SHUTDOWN AND STARTUP MESSENGER
	@REM ULTIMATE TWEAKER
	@REM ----------------------------------------------------------------
	Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v "VerboseStatus" /t REG_DWORD /d "1" /f


	@REM -- ENABLED DETECTION OF APPLICATION INSTALLATION AND PROMPT FOR ELEVATION
	@REM -- SETTING GOOD TO HAVE AS DEFAULT	
	@REM ULTIMATE TWEAKER
	Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v "EnableInstallerDetection" /t REG_DWORD /d "1" /f

	@REM -- NOT_ENABLE ADMIN APPROVAL MODE FOR BUILT IN ADMINISTRATOR ACCOUNT
	@REM -- SETTING GOOD TO HAVE AS DEFAULT		
	@REM ULTIMATE TWEAKER
	Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v "FilterAdministratorToken" /t REG_DWORD /d "1" /f
	@REM ------------------------------
	@REM THIS TURNED OUT TO BE NOT A GOOD IDEA
	@REM IT STOPS THE CALCULATOR WORKING
	@REM AND MUST BE SET TO ONE -- 1 -- /d "1"
	@REM ------------------------------
	@REM HERE IS A WEBSITE EXPLAINS WHY
	@REM ----
	@REM Fix: Apps and Start menu do not open under Administrator account in Windows 10
	@REM https://winaero.com/blog/fix-apps-and-start-menu-do-not-open-under-administrator-account-in-windows-10/
	@REM ----
	@REM SNIPPET ----------------------

	@REM Fix: Apps and Start menu do not open under Administrator account in Windows 10
	@REM Users who have enabled the disabled-by-default Administrator account in Windows 10 face the issue that Universal apps like Photos, Calculator do not work. Also, the new Start menu does not open. Any XAML/Modern app code fails to work. Here is a workaround to fix this issue.

	@REM As you might be knowing, when you enable the hidden Administrator account in Windows 10, it will operate without UAC enabled. Since all Modern/Universal apps like Photos, Calculator require UAC to be enabled, they won't work! The worst thing is that the Start menu also won't work, since it is also a Modern app in Windows 10.
	@REM If you need to use the Administrator account but want to keep apps and the Start menu working, you need to apply a Registry tweak.
	
	@REM You might need to restart Windows 10 for the changes to take effect.
	@REM ------------------------------------------------------------------
	@REM ------------------------------------------------------------------
	
	@REM ------------------------------------------------------------------
	@REM -- DISABLED USER ACCOUNT CONTROL UAC
	@REM -- TWO SETTINGS ONE OF THEM IS OBVIOUS
	@REM ULTIMATE TWEAKER
	@REM Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v "NoDispCPL" /t REG_DWORD /d "0" /f
	
	@REM THIS IS SET ABOVE AND DISABLED USER ACCOUNT CONTROL IS NOT WANTED ANYMORE
	@REM AS WINDOWS SOTRE WON;T OPEN IN A ADMIN ACCOUNT
	@REM Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v "EnableLUA" /t REG_DWORD /d "0" /f
	@REM ------------------------------------------------------------------
	
	
	@REM DISABLED WINDOWS DEFENDER
	@REM ULTIMATE TWEAKER
	Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender" /v "DisableAntispyware" /t REG_DWORD /d "1" /f

	@REM REGISTRY DISABLED WINDOWS UPDATE
	@REM ULTIMATE TWEAKER
	@REM Reg.exe add "HKLM\SYSTEM\ControlSet001\Services\wuauserv" /v "Start" /t REG_DWORD /d "4" /f

	@REM -- DISABLED THE BING SEARCH WHEN IN START MENU AND TYPING WEB SEARCH RESULT GET IN THE WAY
	@REM --
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /v "BingSearchEnabled" /t REG_DWORD /d "0" /f

	
	@REM ----------------------------------------------------------------
	@REM -- SOME ENABLED DISABLE APPS SHARING INFO
	@REM -- NOT SURE @IF THIS IS WINDOWS STORE APP I THINK IT TRUE
	@REM -- NOT SURE ABOUT SETTING @IF RELEVANT TO MY MACHINE ON MY WITH CODE NUMBER BEEN CHECKED OUT UP ON
	@REM --     ENABLED = "Allow"
	@REM -- NOT ENABLED = "Deny"
	@REM --
	@REM -- APPS MESSAGES
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\{992AFA70-6F47-4148-B3E9-3003349C1548}" /v "Value" /t REG_SZ /d "Allow" /f
	@REM -- APPS MICROPHONE
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\{2EEF81BE-33FA-4800-9670-1CD474972C3F}" /v "Value" /t REG_SZ /d "Allow" /f
	@REM -- APPS CAMERA
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\{E5323777-F976-4f5b-9B55-B94699C46E44}" /v "Value" /t REG_SZ /d "Allow" /f
	@REM -- APPS CALENDAR
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\{D89823BA-7180-4B81-B50C-7E471E6121A3}" /v "Value" /t REG_SZ /d "Allow" /f
	
	@REM -- APPS LOCATION
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}" /v "Value" /t REG_SZ /d "Allow" /f
	@REM ----------------------------------------------------------------

	@REM ----------------------------------------------------------------
	@REM -- DISABLE STEPS RECORDER
	Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\AppCompat" /v "DisableUAR" /t REG_DWORD /d "0" /f
	@REM ----------------------------------------------------------------

	@REM ----------------------------------------------------------------
	@REM -- DISABLE FEEDBACK RECORDER
	@REM -- GOING TO DEFAULT ENABLED AND ANSWER IS A PAIR
	@REM Reg.exe delete "HKCU\SOFTWARE\Microsoft\Siuf\Rules" /v "NumberOfSIUFInPeriod" /f
	@REM Reg.exe add "HKCU\SOFTWARE\Microsoft\Siuf\Rules" /f
	@REM -------------------
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Siuf\Rules" /v "NumberOfSIUFInPeriod" /t REG_DWORD /d "0" /f
	@REM -------------------
	@REM ----------------------------------------------------------------

	@REM ----------------------------------------------------------------
	@REM -- STOP WINDOWS UPDATE FROM BEING SHARED OVER INTERNET
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization" /v "SystemSettingsDownloadMode" /t REG_DWORD /d "0" /f
	Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config" /v "DoDownloadMode" /t REG_DWORD /d "0" /f
	@REM ----------------------------------------------------------------

	@REM ----------------------------------------------------------------
	@REM -- DISABLE TELEMETRY
	@REM -- DATA COLLECTION CAN;T BE STOPPED COMPLETELY HOWEVER TWEAKER CAN DISABLE IT TO SOME EXTENT -- SOME SETTING AREN'T APPLICABLE TO WINDOWS HOME USER
	@REM -- IT DETECT DIAGNOSTIC AND USAGE DATA AND SENDS IT TO MICROSOFT
	@REM --
	@REM -- DISABLE TELEMETRY
	@REM -- INITIAL DEFAULT VALUE ARE 2 FOR EACH PAIR
	Reg.exe add "HKLM\SYSTEM\ControlSet001\Services\DiagTrack" /v "Start" /t REG_DWORD /d "4" /f
	Reg.exe add "HKLM\SYSTEM\ControlSet001\Services\dmwappushservice" /v "Start" /t REG_DWORD /d "4" /f
	@REM ----------------------------------------------------------------
	
	@REM ----------------------------------------------------------------
	@REM -- DISABLE ANY TYPE OF SYSTEM BIOMETRICS FOR THIS DEVICE
	Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Biometrics" /v "Enabled" /t REG_DWORD /d "0" /f

	@REM ----------------------------------------------------------------
	@REM -- COLLECTION
	@REM -- PUT ALL THESE PROGRAMS ON THE CONTEXT MENU DESKTOP AT THE BOTTOM AND OF RIGHT CLICK
	@REM -- Windows Update
	@REM -- Control Panel
	@REM -- Create Restore Point
	@REM -- Calculator
	@REM -- Command Prompt
	@REM -- Magnifier
	@REM -- Paint
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Windows Update" /v "Icon" /t REG_SZ /d "%%SystemRoot%%\\System32\\shell32.dll,-47" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Windows Update" /v "Position" /t REG_SZ /d "Bottom" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Windows Update\Command" /ve /t REG_SZ /d "explorer ms-settings:windowsupdate" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Control Panel" /v "Icon" /t REG_SZ /d "control.exe" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Control Panel" /v "Position" /t REG_SZ /d "Bottom" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Control Panel\Command" /ve /t REG_SZ /d "control.exe" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Create Restore Point" /v "Icon" /t REG_SZ /d "SystemPropertiesProtection.exe" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Create Restore Point" /v "Position" /t REG_SZ /d "Bottom" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Create Restore Point\Command" /ve /t REG_SZ /d "WScript C:\RestorePoint.vbs" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Calculator" /v "Icon" /t REG_SZ /d "calc.exe,0" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Calculator" /v "Position" /t REG_SZ /d "Bottom" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Calculator\Command" /ve /t REG_SZ /d "calc.exe" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Command Prompt" /v "Icon" /t REG_SZ /d "cmd.exe,0" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Command Prompt" /v "Position" /t REG_SZ /d "Bottom" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Command Prompt\Command" /ve /t REG_SZ /d "cmd.exe" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Magnifier" /v "Icon" /t REG_SZ /d "magnify.exe,0" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Magnifier" /v "Position" /t REG_SZ /d "Bottom" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Magnifier\Command" /ve /t REG_SZ /d "magnify.exe" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Paint" /v "Icon" /t REG_SZ /d "mspaint.exe,0" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Paint" /v "Position" /t REG_SZ /d "Bottom" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Paint\Command" /ve /t REG_SZ /d "mspaint.exe" /f
	@REM ----------------------------------------------------------------


	@REM ----------------------------------------------------------------
	@REM ADD C DRIVE AND D DRIVE
	@REM THOUGHT THE OTHER CONTEXT MENU TAB SETTING WHICH IS STILL SAME DESKTOP RIGHT CLICK
	@REM ----------------------------------------------------------------
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\OS (C:)\Command" /ve /t REG_SZ /d "explorer.exe \"C:\\\"" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\OS (C:)" /v "Icon" /t REG_SZ /d "C:\Windows\System32\Shell32.dll,8" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\OS (C:)" /v "Position" /t REG_SZ /d "Top" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Data (D:)\Command" /ve /t REG_SZ /d "explorer.exe \"D:\\\"" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Data (D:)" /v "Icon" /t REG_SZ /d "C:\Windows\System32\Shell32.dll,8" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Data (D:)" /v "Position" /t REG_SZ /d "Top" /f
	@REM ----------------------------------------------------------------
	
	@REM -- ADD MICROSOFT STORE
	@REM ----------------------------------------------------------------
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Store\Command" /ve /t REG_SZ /d "explorer.exe ms-windows-store://" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\DesktopBackground\Shell\Store" /v "Position" /t REG_SZ /d "Bottom" /f
	@REM ----------------------------------------------------------------

	@REM ----------------------------------------------------------------
	@REM -- ULTIMATE OPEN COMMAND WINDOW IS DONE BUT DON'T SHOW YOU HOW TO DOER IT 
	@REM -- HERE FROM MY FILE COLLECTION
	@REM -- Right_Click_Only_with_Open_Command_Window_Here.reg
	@REM -- http://www.sevenforums.com
	@REM -- http://www.sevenforums.com/tutorials/134840-open-command-window-here-add-@REMove-shift-right-click.html
	@REM -- 
	Reg.exe add "HKCR\Directory\Background\shell\cmd" /ve /t REG_SZ /d "@shell32.dll,-8506" /f
	Reg.exe delete "HKCR\Directory\Background\shell\cmd" /v "Extended" /f
	Reg.exe add "HKCR\Directory\Background\shell\cmd" /v "NoWorkingDirectory" /t REG_SZ /d "" /f
	Reg.exe add "HKCR\Directory\Background\shell\cmd\command" /ve /t REG_SZ /d "cmd.exe /s /k pushd \"%%V\"" /f
	
	@REM TWO GROUP OF 4 BELOW NOT WORKING ACCESS ERROR
	@REM Reg.exe add "HKCR\Directory\shell\cmd" /ve /t REG_SZ /d "@shell32.dll,-8506" /f
	@REM Reg.exe delete "HKCR\Directory\shell\cmd" /v "Extended" /f
	@REM Reg.exe add "HKCR\Directory\shell\cmd" /v "NoWorkingDirectory" /t REG_SZ /d "" /f
	@REM Reg.exe add "HKCR\Directory\shell\cmd\command" /ve /t REG_SZ /d "cmd.exe /s /k pushd \"%%V\"" /f
	
	@REM Reg.exe add "HKCR\Drive\shell\cmd" /ve /t REG_SZ /d "@shell32.dll,-8506" /f
	@REM Reg.exe delete "HKCR\Drive\shell\cmd" /v "Extended" /f
	@REM Reg.exe add "HKCR\Drive\shell\cmd" /v "NoWorkingDirectory" /t REG_SZ /d "" /f
	@REM Reg.exe add "HKCR\Drive\shell\cmd\command" /ve /t REG_SZ /d "cmd.exe /s /k pushd \"%%V\"" /f
	@REM ----------------------------------------------------------------
	@REM -- NOT ALL THIS ONE WORKS CAN'T GET ACCESS TO CLASSES ROOT
	@REM -- TIME TO @REMOVE IT 
	
	
	@REM ----------------------------------------------------------------
	@REM ENABLED FULL ROW SELECT WINDOWS 10 OR MAYBE ALSO 7
	@REM NOT SURE ABOUT CODE HASHING HERE 5C4F28B5 FOR THIS ONE 
	@REM BUT GONNA TRY ANYWAY 
	@REM THINKER ONLY REALLY WANT THE FIRST SETTING REG_DWORD /d "1"
	@REM CONVERSION FROM ULTIMATE WINDOWS TWEAKER AND NIRSOFT__RegFromApp__ AND REG CONVERTER V1.1
	@REM ----------------------------------------------------------------
	@REM ----------------------------------------------------------------
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "FullRowSelect" /t REG_DWORD /d "1" /f
	@REM ----------------------------------------------------------------

	@REM ----------------------------------------------------------------
	@REM - DON'T DISABLE HideDrivesWithNoMedia
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "HideDrivesWithNoMedia" /t REG_DWORD /d "0" /f
	@REM ----------------------------------------------------------------

	@REM ----------------------------------------------------------------
	@REM EXPLORER OPEN TO THIS PC RATHER THAN QUICK ACCESS
	@REM DEFAULT __ QUICK ACCESS = REG_DWORD /d "2"
	@REM HARDER WORK TO FIND THE DEFAULT EACH TIME REQUIRE REGEDIT OPEN AT THE PAGING
	@REM ----------------------------------------------------------------
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "LaunchTo" /t REG_DWORD /d "1" /f

	@REM ----------------------------------------------------------------
	@REM QUICK ACCESS SHOW RECENT FILES AND SHOW FREQUENT FOLDER OFF 
	@REM DEFAULT REG_DWORD /d "1" _ ON _ FOUND SYSTEM TO SHOW DEFAULT EASIER IN SETTING OF NIRSOFT
	@REM ----------------------------------------------------------------
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /v "ShowRecent" /t REG_DWORD /d "0" /f
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /v "ShowFrequent" /t REG_DWORD /d "0" /f	
	
	@REM ----------------------------------------------------------------
	@REM GROUP BY FILE SYSTEM NTFS
	@REM ----------------------------------------------------------------
	@REM Reg.exe add "HKCU\Software\Classes\Local Settings\MuiCache\ad7\52C64B7E" /v "LanguageList" /t REG_MULTI_SZ /d "en-US\0en" /f
	@REM ----------------------------------------------------------------

	
	@REM ----------------------------------------------------------------
	@REM 5 Windows 10 Registry Hacks To Optimize Windows 10 Experience
	@REM https://www.maketecheasier.com/windows-10-registry-hacks/
	@REM ----
	@REM ---- STARTUP HAS A DELAY AND SOMETIME TAKE FOREVER AFTER DEFAULT SETTING GONE
	@REM ---- HERE MAKE IT INSTANT
	Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Serialize" /v "StartupDelayInMSec" /t REG_DWORD /d "0" /f
	
	
	@REM ----------------------------------------------------------------
	@REM DONT LET WINDOWS UPDATE DO SHUTDOWN BUTTON POWER SWITCH UPDATES
	@REM SHUT DOWN WITHOUT UPDATE WITH THE POWER BUTTON
	@REM ----------------------------------------------------------------
	@REM ----
	@REM Stop Windows Update from Hijacking the Sleep/Shutdown Button
	@REM https://www.howtogeek.com/howto/windows-vista/stop-windows-update-from-hijacking-the-sleep-button/
	@REM ----
	@REM ----------------------------------------------------------------
	Reg.exe add "HKCU\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" /v "NoAUAsDefaultShutdownOption" /t REG_DWORD /d "0" /f

	@REM ----------------------------------------------------------------
	@REM -- DISABLE ALL LIVE TILES -- DEFAULT 1
	@REM ----------------------------------------------------------------
	Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "NoChangeStartMenu" /t REG_DWORD /d "0" /f
	
	@REM ----------------------------------------------------------------
	@REM -- DISABLE LOOK FOR APP IN THE STORE WHEN UNKNOWN FILE IS OPENED
	@REM ----------------------------------------------------------------
	Reg.exe add "HKCU\SOFTWARE\Policies\Microsoft\Windows\Explorer" /v "NoUseStoreOpenWith" /t REG_DWORD /d "1" /f
	
:NOT_WINDOWS_10


@REM ----------------------------------------------------------------
@REM -- AFTER NOTEPAD++ CONVERT CONTEXT MENU TO OPEN RATHER THAN NOTEPAD IN A WAY DOES BOTH BUT NOTEPAD ONE STOP WORKING FOR EDIT ON RUNAS __ SO CHANGE RUNAS EDIT TO DO TAKE_OWNERSHIP NOT SURE @IF WANTED MORE COMMAND ON RUNAS __ HERE EXAMPLE
@REM -- WELL AFTER THEM TWO EDIT SEEM TO COME BACK WORKING ON AGAIN
@REM ----------------------------------------------------------------



Reg.exe add "HKCR\*\shell\RunAs" /ve /t REG_SZ /d "Take Ownership" /f
Reg.exe add "HKCR\*\shell\RunAs" /v "NoWorkingDirectory" /t REG_SZ /d "" /f
Reg.exe add "HKCR\*\shell\RunAs\command" /ve /t REG_SZ /d "cmd.exe /c takeown /f \\\"%%1\\\" && icacls \\\"%%1\\\" /grant administrators:F" /f
Reg.exe add "HKCR\*\shell\RunAs\command" /v "IsolatedCommand" /t REG_SZ /d "cmd.exe /c takeown /f \\\"%%1\\\" && icacls \\\"%%1\\\" /grant administrators:F" /f

Reg.exe add "HKCR\Directory\shell\RunAs" /ve /t REG_SZ /d "Take Ownership" /f
Reg.exe add "HKCR\Directory\shell\RunAs" /v "NoWorkingDirectory" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Directory\shell\RunAs\command" /ve /t REG_SZ /d "cmd.exe /c takeown /f \\\"%%1\\\" && icacls \\\"%%1\\\" /grant administrators:F" /f
Reg.exe add "HKCR\Directory\shell\RunAs\command" /v "IsolatedCommand" /t REG_SZ /d "cmd.exe /c takeown /f \\\"%%1\\\" && icacls \\\"%%1\\\" /grant administrators:F" /f


@REM -- ONE BY ULTIMATE TWEAKER 
@REM -- @REM_OUT ONE ABOVE PERHAPS LEAVE IN BECAUSE MADE NOTEPAD WORK AGAIN

Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\RunAs" /ve /t REG_SZ /d "Take Ownership" /f
Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\RunAs" /v "HasLUAShield" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\RunAs" /v "NoWorkingDirectory" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\RunAs\command" /ve /t REG_SZ /d "cmd.exe /c takeown /f \"%%1\" && icacls \"%%1\" /grant administrators:F" /f
Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\RunAs\command" /v "IsolatedCommand" /t REG_SZ /d "cmd.exe /c takeown /f \"%%1\" && icacls \"%%1\" /grant administrators:F" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\RunAs" /ve /t REG_SZ /d "Take Ownership" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\RunAs" /v "HasLUAShield" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\RunAs" /v "NoWorkingDirectory" /t REG_SZ /d "" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\RunAs\command" /ve /t REG_SZ /d "cmd.exe /c takeown /f \"%%1\" && icacls \"%%1\" /grant administrators:F" /f
Reg.exe add "HKLM\SOFTWARE\Classes\Directory\shell\RunAs\command" /v "IsolatedCommand" /t REG_SZ /d "cmd.exe /c takeown /f \"%%1\" && icacls \"%%1\" /grant administrators:F" /f





@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "AlwaysShowMenus" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "Hidden" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "HideDrivesWithNoMedia" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "HideFileExt" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "HideMergeConflicts" /t REG_DWORD /d "0" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "SeparateProcess" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "PersistBrowsers" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "ShowEncryptCompressedColor" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "NavPaneExpandToCurrentFolder" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "NavPaneShowAllFolders" /t REG_DWORD /d "1" /f
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM DISPLAY FULL PATH IN TITLE BAR
@REM ----------------------------------------------------------------
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CabinetState" /v "FullPath" /t REG_DWORD /d "1" /f


@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM DON'T GROUP TASK-BAR BUTTONS
@REM ----------------------------------------------------------------
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "TaskbarGlomLevel" /t REG_DWORD /d "2" /f
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------


@REM -- KEEP YOU CLOCK MY FREQUENTER IN TIME BY POLLING MORE
@REM ----
@REM Change Internet Time Update interval in Windows 10 / 8 / 7
@REM http://www.thewindowsclub.com/change-internet-time-update-interval-windows
@REM ----
@REM ---- ALSO A CHECK AND WORK WITH XP
@REM ---- DEFAULT FOUND 604800 WITH XP
@REM EVERY HOUR 
@REM IN SECOND
@REM 60*60 = 3600

reg.exe ADD "HKLM\SYSTEM\ControlSet001\Services\W32Time\TimeProviders\NtpClient" /f /v "SpecialPollInterval" /t REG_DWORD /d 3600


@ECHO WINDOWS 10
@IF NOT "%WIN10__%"=="YES" GOTO NOT_WINDOWS_10
	@REM ----------------------------------------------------------------
	@REM -- How to Fix High CPU Usage by RuntimeBroker.exe on Windows 10 - Google Chrome
	@REM ----------------------------------------------------------------
	@REM reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Services\TimeBroker" /f /v "Start" /t REG_DWORD /d 3

	@REM -----------------------------------------
	@REM -- HERE WAS SETTER BUT DON'T WANT ANYMORE
	@REM reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Services\TimeBrokerSvc" /f /v "Start" /t REG_DWORD /d 4
	@REM -----------------------------------------


	@REM ----
	@REM What is Runtime Broker, and Why is it working so pathetically bad? - DonationCoder.com
	@REM http://www.donationcoder.com/forum/index.php?topic=42009.0
	@REM ----
	@REM Find the following entry:
	@REM [HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\TimeBroker] "Start"=dword:00000003
	@REM Change the 3 to a 4.
	@REM [HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\TimeBroker] "Start"=dword:00000004
	@REM 4 is Disabled, 3 is Manual and 2 is Automatic startup.
	@REM PAUSE
:NOT_WINDOWS_10

@REM ----------------------------------------------------------------
reg.exe ADD "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OptimalLayout" /f /v "EnableAutoLayout" /t REG_DWORD /d 0

reg.exe ADD "HKCU\Control Panel\Desktop\WindowMetrics" /f /v "MinAnimate" /t REG_SZ /d "0"

reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "PaintDesktopVersion" /t REG_DWORD /d 0
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CursorBlinkRate" /t REG_SZ /d "200"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "DragFullWindows" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "WheelScrollLines" /t REG_SZ /d "5"

reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseTrails" /t REG_SZ /d "4"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseSpeed" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseSensitivity" /t REG_SZ /d "4"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseHoverTime" /t REG_SZ /d "20" 
@REM -------------------------------------------------------------------
@REM __ MouseHoverTime SPEEDY TIME HERE BUT AND LIKE A DECAY TIME AWAY AGAIN SEEM TO GET STUCK OPEN
@REM -------------------------------------------------------------------





@REM -------- ################################# ---------------------------------------------------------------
@REM -------- NOT SET NOTEPAD++ AT MOMENT 64 BIT RELOCATE SORT TO DO
@REM -------- SET NOTEPAD++ --------------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM @ECHO ---------- SET NOTEPAD++ FOR EDIT FILES
@REM ----------------------------------------------------------------
@ECHO WINDOWS 10
@IF "%WIN10__%"=="YES" (
	@REM -- reg.exe ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "C:\Program Files\Notepad++\notepad++.exe %%1"
)
@REM ----------------------------------------------------------------
@REM @ECHO ---------- RESTORE ORIGINAL NOTEPAD ------------------------------------
@REM reg.exe ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE %%1"
@REM ----------------------------------------------------------------
@REM -------- ################################# ------------------------------------------------
@REM ----------------------------------------------------------------

@REM ------------------------------------------------------------------
@REM -- @ECHO http://www.optimizingpc.com/optimize/registrytweaks.html
@REM ------------------------------------------------------------------
@REM -- Disabling wrapping icon/shortcut labels
@REM -- Probably your desktop contains icons (shortcuts) with long names, which aren't shown 
@REM -- Ccompletely (the are wrapped......). With a registry tweak you can disable the title wrapping, 
@REM -- Create a String value IconTitleWrap with value 0 in the following registry key:
@REM -- HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics
@REM ------------------------------------------------------------------
reg.exe ADD "HKCU\Control Panel\Desktop\WindowMetrics" /f /v "IconTitleWrap" /t REG_SZ /d "0" 
@REM ------------------------------------------------------------------


@REM __ ------------------------------------------------------------------------------
@REM __ Registry Settings that can be Modified to Improve Operating System Performance
@REM __ https://msdn.microsoft.com/en-us/library/ee377058(v=bts.10).aspx
@REM __ ------------------------------------------------------------------------------
@REM ________________________________________________________________
@REM __ Increase space available for the master file table
@REM __ Add the NtfsMftZoneReservation entry to the registry to allow the master file table (MFT) to grow optimally. When you add this entry to the registry, the system reserves space on the volume for the master file table. @IF your NTFS volumes contain relatively few large files, set the value of this registry entry to 1 (the default). Typically you can set this entry to a value of 2 or 3 for volumes that contain a moderate numbers of files, and use a value of 4 (the maximum) @IF your volumes tend to contain a relatively large number of files.
@REM __ Ee377058.Important(en-US,BTS.10).gifImportant
@REM __ Test any settings greater than 2, because setting this entry to a value greater than 2 will cause the system to reserve a much larger portion of the disk for the master file table.
@REM __ NtfsMftZoneReservation
@REM __ Key:
@REM __ HKLM\SYSTEM\CurrentControlSet\Control\FileSystem
@REM __ Value:
@REM __ NtfsMftZoneReservation
@REM __ Data Type:
@REM __ REG_DWORD
@REM __ Range:
@REM __ 1 – 4
@REM __ Default value:
@REM __ 1
@REM __ Recommended value:
@REM __ 1 @IF volumes typically store fewer files.
@REM __ 2 or 3 @IF volumes typically store a moderate number of files.
@REM __ 4 @IF volumes typically store a large number of files.
@REM __ 
@REM __ Value exists by default?
@REM __ No, needs to be added.
@REM ________________________________________________________________
@REM - SET REG_DWORD /d 0 -- UNDO
@REM ________________________________________________________________
reg.exe ADD "HKCU\Control Panel\Desktop\WindowMetrics" /f /v "NtfsMftZoneReservation" /t REG_DWORD /d 4

reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Control\SessionManager\Executive" /f /v "AdditionalDelayedWorkerThreads" /t REG_DWORD /d 16
@REM __ DEFAULT 0 __ AdditionalDelayedWorkerThreads        __________________________________
reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Control\SessionManager\Executive" /f /v "AdditionalCriticalWorkerThreads" /t REG_DWORD /d 16
@REM __ DEFAULT 0 __ AdditionalCriticalWorkerThreads       __________________________________
reg.exe ADD "HKCU\Control Panel\Desktop\WindowMetrics" /f /v "NtfsMftZoneReservation" /t REG_DWORD /d 4
reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" /f /v "MaxWorkItems" /t REG_DWORD /d 8192
@REM __ DEFAULT Configured dynamically __ MaxWorkItems    __________________________________
reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" /f /v "MaxMpxCt" /t REG_DWORD /d 2048
@REM __ DEFAULT 50 __ MaxMpxCt                            __________________________________
reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters" /f /v "MaxCmds" /t REG_DWORD /d 2048
@REM __ DEFAULT 50 __ MaxCmds                             __________________________________
reg.exe ADD "HKLM\System\CurrentControlSet\Control\Session Manager\MemoryManagement" /f /v "LargeSystemCache" /t REG_DWORD /d 0
reg.exe ADD "HKLM\System\CurrentControlSet\Services\LanmanServer\Parameters" /f /v "Size" /t REG_DWORD /d 3
reg.exe ADD "HKLM\SYSTEM\CurrentControlSet\Control\FileSystem" /f /v "DontVerifyRandomDrivers" /t REG_DWORD /d 1
@REM __ DEFAULT 0 ___ DontVerifyRandomDrivers             __________________________________
@REM __ ------------------------------------------------------------------------------
@REM NERO TUNEITUP RECOMENDER SETTING
Reg.exe add "HKLM\SYSTEM\ControlSet001\Services\LanmanServer\Parameters" /v "AutoShareWks" /t REG_DWORD /d "0" /f
@REM OTHER SETTING FROM THE LAN NETWORK BY NERO_TUNEITUP
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "NoNetCrawling" /t REG_DWORD /d "1" /f
Reg.exe add "HKLM\SYSTEM\ControlSet001\Services\Tcpip\Parameters" /v "EnablePMTUBHDetect" /t REG_DWORD /d "1" /f
Reg.exe add "HKLM\SYSTEM\ControlSet001\Services\Tcpip\Parameters" /v "EnablePMTUDiscovery" /t REG_DWORD /d "1" /f
Reg.exe add "HKLM\SYSTEM\ControlSet001\Services\Tcpip\Parameters" /v "DefaultTTL" /t REG_DWORD /d "128" /f
Reg.exe add "HKLM\SYSTEM\ControlSet001\Services\Dnscache\Parameters" /v "MaxNegativeCacheTtl" /t REG_DWORD /d "0" /f
Reg.exe add "HKLM\SYSTEM\ControlSet001\Services\Dnscache\Parameters" /v "NetFailureCacheTime" /t REG_DWORD /d "0" /f
Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Psched" /v "NonBestEffortLimit" /t REG_DWORD /d "0" /f

@REM INTERNET EXPLORER NETWORK SETTING NERO_TUNEITUP WANT CHANGING ARE A FEW ABOUT 4 TO THE IDEAL
@REM IN A SCREEN SHOT HERE
@REM UNABLE TO EXTRACT THEM YET
@REM D:\0 00 ART LOGGERS\# APP AND SCREEN\8-MSI-GP62M-7RD\HOT_KEY_SCREEN_SHOT\Hot-Key-App-Shots\HotKey-Capture- 2017-12-07 00-28-30 Thu.jpg

@REM CHROME NETWORK SETTING NERO_TUNEITUP WANT CHANGING ARE A FEW ABOUT 4 TO THE IDEAL
@REM IN A SCREEN SHOT HERE
@REM UNABLE TO EXTRACT THEM YET
@REM D:\0 00 ART LOGGERS\# APP AND SCREEN\8-MSI-GP62M-7RD\HOT_KEY_SCREEN_SHOT\Hot-Key-App-Shots\HotKey-Capture- 2017-12-07 00-30-55 Thu.jpg

@REM FIREFOX NETWORK SETTING NERO_TUNEITUP WANT CHANGING ARE A FEW ABOUT 20 TO THE IDEAL
@REM IN A SCREEN SHOT HERE
@REM UNABLE TO EXTRACT THEM YET
@REM D:\0 00 ART LOGGERS\# APP AND SCREEN\8-MSI-GP62M-7RD\HOT_KEY_SCREEN_SHOT\Hot-Key-App-Shots\HotKey-Capture- 2017-12-07 00-22-46 Thu.jpg

@REM THESE 4 CONFIG SET MICROSOFT EDGE MAXIMUM BY NERO_TUNEITUP
Reg.exe add "HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppContainer\Storage\microsoft.microsoftedge_8wekyb3d8bbwe\MicrosoftEdge\FlipAhead" /v "FPEnabled" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppContainer\Storage\microsoft.microsoftedge_8wekyb3d8bbwe\MicrosoftEdge\PhishingFilter" /v "EnabledV9" /t REG_DWORD /d "1" /f
Reg.exe add "HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppContainer\Storage\microsoft.microsoftedge_8wekyb3d8bbwe\MicrosoftEdge\New Windows" /v "PopupMgr" /t REG_SZ /d "yes" /f
Reg.exe add "HKCU\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppContainer\Storage\microsoft.microsoftedge_8wekyb3d8bbwe\MicrosoftEdge\Main" /v "DoNotTrack" /t REG_DWORD /d "1" /f


@REM ALLOW ADVERT ID
@REM NERO TUNEITUP
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo" /v "Enabled" /t REG_DWORD /d "0" /f
@REM Disable PEEK Preview Desktop WHEN MOVE BOTTOM RIGHT CORNER
@REM NOT SURE SAW ANY DIFFERANCE YET STILL WORK TO MINIMIZE ALL APP-ING
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "DisablePreviewDesktop" /t REG_DWORD /d "1" /f

@REM SPEED UP THE ExtendedUIHoverTime WHEN DISPLAY A THUMBNAIL OVER THE TASK BAR ITEM AND MAKE QUICKER
@REM BUT HARDER TO CALL UP PROPERTIES IF MULTI LEVEL TASK BAR LIKE 3 LEVEL 
@REM ONLY THE ONE MOVED TO THE TOP
@REM NERO TUNEITUP
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v "ExtendedUIHoverTime" /t REG_DWORD /d "0" /f

@REM AutoRestartShell GOT THOWN IN
@REM NERO TUNEITUP
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v "AutoRestartShell" /t REG_DWORD /d "1" /f



@REM FIND ENABLE MOUSE POINTER SHADOWS LOOK A BIT BETTER



@IF "%WINXP__%"=="YES" GOTO NOT_WHEN_WIN_XP
@IF "%WIN07__%"=="YES"  GOTO NOT_WHEN_WIN_07

@REM __ ----
@REM __ Settings that can be Modified to Improve Network Performance
@REM __ https://msdn.microsoft.com/en-us/library/ee377084(v=bts.10).aspx
@REM __ ----
netsh int ipv4 set dynamicport tcp start=1025 num=64511  
reg.exe ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" /f /v "TcpTimedWaitDelay" /t REG_DWORD /d 30
@REM __ DEFAULT 120 __ TcpTimedWaitDelay                  __________________________________

:NOT_WHEN_WIN_07
:NOT_WHEN_WIN_XP


GOTO TEXT_JUMP_@REMMER_01
@REM ----
@REM General Guidelines for Improving Network Performance
@REM https://msdn.microsoft.com/en-us/library/ee377033(v=bts.10).aspx
@REM ----
@REM @REMove unnecessary network protocols
@REM Windows Server computers sometimes have more network services and protocols installed than are actually required. Each additional network client, service or protocol places additional overhead on system resources.
@REM In addition, each installed protocol generates network traffic. By @REMoving unnecessary network clients, services and protocols, system resources are made available for other processes, excess network traffic is avoided and the number of network bindings that must be negotiated is reduced to a minimum.
@REM To see the currently installed network clients, protocols and services, follow these steps:
@REM Click Start, and then click Control Panel.
@REM In Control Panel, do one of the following
@REM In Adjust your computer’s settings, set View by to Category, click Network and Internet, and then click Network and Sharing Center.
@REM In Adjust your computer’s settings, set View by to either Large icons or Small icons, and then click Network and Sharing Center.
@REM In the Tasks pane, click Change adapter settings.
@REM Right-click Local Area Connection (or the entry for your network connection), and then click Properties to display the properties dialog box for the network connection.
@REM To @REMove an unnecessary item, select it and click Uninstall. To disable an item, simply clear the checkbox associated with the item.
@REM @IF you are unsure about the effects of uninstalling an item for the connection, then disable the item rather than uninstalling it. Disabling items allows you to determine which services, protocols and clients are actually required on a system. When it has been determined that disabling an item has no adverse affect on the server, the item can then be uninstalled.
@REM In many cases, only the following three components are required for operation on a standard TCP/IP based network:
@REM Client for Microsoft Networks
@REM File and Printer Sharing for Microsoft Networks
@REM Internet Protocol (TCP/IP)
:TEXT_JUMP_@REMMER_01






@REM __ reg.exe ADD "" /f /v "" /t REG_DWORD /d 1
@REM __ reg.exe ADD "" /f /v "" /t REG_DWORD /d 1
@REM __ reg.exe ADD "" /f /v "" /t REG_DWORD /d 1
@REM __ reg.exe ADD "" /f /v "" /t REG_DWORD /d 1




reg.exe ADD "HKCU\Control Panel\Desktop\Sound" /f /v "Beep" /t REG_SZ /d "4" 
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "PaintDesktopVersion" /t REG_DWORD /d 0

@REM ------------------------------------
@REM -- JUMP TO REGISTRY KEY CHANGE SETTING BRING BACK HERE
@REM ------------------------------------
@REM -- Mouse
@REM ------------------------------------
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseTrails" /t REG_SZ /d "5"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseThreshold1" /t REG_SZ /d "6"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseThreshold2" /t REG_SZ /d "10"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseSpeed" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "MouseSensitivity" /t REG_SZ /d "4"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "SnapToDefaultButton" /t REG_SZ /d "0"
@REM ------------------------------------
@REM -- DOUBLE CLICK SPEED
@REM ------------------------------------
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "DoubleClickSpeed" /t REG_SZ /d "1000"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "DoubleClickHeight" /t REG_SZ /d "5"
reg.exe ADD "HKCU\Control Panel\Mouse" /f /v "DoubleClickWidth" /t REG_SZ /d "5"
@REM ------------------------------------
@REM -- WHAT IS Beep FOR CURRENTLY OFF NO
@REM -- AND ExtendedSounds CURRENTLY OFF NO
@REM ------------------------------------
@REM -- Desktop
@REM ------------------------------------
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "WheelScrollLines" /t REG_SZ /d "5"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "MenuShowDelay" /t REG_SZ /d "100"
@REM -- MenuShowDelay __ UP FROM 40 TO 100 TOO QUICK FOR THE WIN 10 TASKBAR
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "DragFullWindows" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CursorBlinkRate" /t REG_SZ /d "10"

reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "CaretWidth" /t REG_DWORD /d 3
@REM  __ REG_DWORD NOT AS MISTAKER REG_SZ 

@REM -------------------------------------------------
@REM ____ NOTER ______________________________________
@REM ____ CaretWidth _________ CURSOR TEXT SIZE MAX 50
@REM ____ CursorBlinkRate ____ MILLI SECOND 10
@REM -------------------------------------------------

@REM ------------------------------------------------------------------
@REM -- CAP LOCK WANTING
@REM -- ----
@REM -- Turn NUM LOCK on at start-up - Microsoft Community
@REM -- https://answers.microsoft.com/en-us/windows/forum/windows_7-hardware/turn-num-lock-on-at-start-up/b6f5c70d-a053-4311-946d-ba9e8879f644?page=4
@REM -- Keyboard Preference
@REM ------------------------------------------------------------------
@REM -- Keyboard
@REM -- 3 – Turn on Caps Lock and Num Lock
@REM -- 1 – Turn on Caps Lock
@REM ------------------------------------------------------------------
@REM -- BE FOR WHAT I WANT NUM_LOCK_OFF AND CAP_LOCK_ON
@REM ------------------------------------------------------------------
reg.exe ADD "HKCU\Control Panel\Keyboard" /f /v "InitialKeyboardIndicators" /t REG_SZ /d "1"
@REM ------------------------------------------------------------------
@REM -- KEYBOARD SPEED MAINLY DEFAULT REBOOT KEEP GETTING SET BACK NOT QUICKER
@REM ------------------------------------------------------------------
reg.exe ADD "HKCU\Control Panel\Keyboard" /f /v "KeyboardDelay" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\Keyboard" /f /v "KeyboardSpeed" /t REG_SZ /d "31"
@REM ------------------------------------------------------------------
@REM __ Keyboard "KeyboardSpeed" "1" TO "31" TOP SPEED
@REM __ IMPORTANT REPEAT SETTING AS RESET HAPPEN LESS QUICK
@REM __ MAIN REASON REG FILE
@REM ------------------------------------------------------------------



@REM __ Windows Registry Editor Version 5.00
@REM __ [HKEY_CURRENT_USER\Control Panel\Sound]
@REM ----------------------------------------------------------------
@REM -- COMPLETE SET SCRIPT ON MY COMPUTER WINDOWS XP PRO -- SUB KEYS EXIST
@REM ----------------------------------------------------------------
reg.exe ADD "HKCU\Control Panel\Sound" /f /v "Beep" /t REG_SZ /d "Yes"
reg.exe ADD "HKCU\Control Panel\Sound" /f /v "ExtendedSounds" /t REG_SZ /d "Yes"
@REM ----------------------------------------------------------------
@REM -- MAY WANT TO SET THIS ONE
@REM ----------------------------------------------------------------
@REM -- HKEY_CURRENT_USER\Control Panel\Accessibility
@REM ----------------------------------------------------------------
reg.exe ADD "HKCU\Control Panel\Accessibility\Blind Access" /f /v "On" /t REG_SZ /d "0"
reg.exe ADD "HKCU\Control Panel\Accessibility\Keyboard Preference" /f /v "On" /t REG_SZ /d "1"
reg.exe ADD "HKCU\Control Panel\Accessibility\ShowSounds" /f /v "On" /t REG_SZ /d "0"
@REM ----------------------------------------------------------------
@REM __ Keyboard Preference __ ENABLE HERE OR ACCESSIBILITY NOT ENABLED IN CONTROL PANEL TO SET 
@REM __ KEYBOARD BLINK RARE AND CURSOR SIZE
@REM __ @IF EVER FOUND REG TWEAK TURNER OFF
@REM ----------------------------------------------------------------


@REM ------------------------------------
@REM -- WIN 10 
@REM -- SET SLIDE SHOW QUICKER THAN MINUTE LEARN
@REM ------------------------------------
@IF NOT "%WIN10__%"=="YES" GOTO NOT_WINDOWS_10
	@ECHO WINDOWS 10

	@ECHO ____ WINDOW 10
	@REM reg.exe ADD "HKCU\Control Panel\Personalization\Desktop Slideshow" /f /v "Interval" /t REG_DWORD /d 10000
	@REM reg.exe ADD "HKCU\Control Panel\Personalization\Desktop Slideshow" /f /v "Shuffle" /t REG_DWORD /d 0
	@REM reg.exe ADD "HKCU\Control Panel\PowerCfg" /f /v "CurrentPowerPolicy" /t REG_SZ /d "0"
	@REM -- WIN 10 -- DEFAULT Wallpaper
	@REM __ NOT SURE @IF CHANGE
	@REM reg.exe ADD "HKCU\Control Panel\Desktop" /f /v "WallPaper" /t REG_SZ /d "c:\windows\web\wallpaper\windows\img0.jpg"
:NOT_WINDOWS_10

@REM ----------------------------------------------------------------
@REM -- SLIDE SHOW DELAY
@REM ----------------------------------------------------------------
@IF NOT "%WINXP__%"=="YES" GOTO NOT_WINDOWS_XP
	ECHO WINDOWS XP
	@REM reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellImageView" /f /v "Timeout" /t REG_DWORD /d 1
:NOT_WINDOWS_XP
@REM ----------------------------------------------------------------





@REM ------------------------------------
@REM __ GOTO JUMP
@REM -- HKEY_CURRENT_USER\Console
@REM __ Windows Registry Editor Version 5.00
@REM __ [HKEY_CURRENT_USER\Console]
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
@REM -- @ECHO ^%systemroot^%
@REM -- from within a batch file:
@REM -- @ECHO %%systemroot%%

@REM -- PROBLEM ALL DWORD ARE DEFAULT DECIMAL NOT HEX
@REM -- DOS ____
@REM __ THE CURSOR SIZE OF DOS CMD COMMAND WINDOW DIFFERENT TO SYSTEM GUI ONE
@REM __ ---------------------------------------------------------------------
@REM -- MEDIUM CURSOR SIZE __ MAX 50 MAYBE TRUE FOR SYSTEM GUI ONE
@REM reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "CursorSize" /t REG_DWORD /d 50
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "HistoryBufferSize" /t REG_DWORD /d 90
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "InsertMode" /t REG_DWORD /d 1
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "NumberOfHistoryBuffers" /t REG_DWORD /d 4

@REM __ :: 655294544 equals 9999 lines in the GUI
reg.exe add "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /v ScreenBufferSize /t REG_DWORD /d 655294544 /f

@REM __ :: We also need to change the Window Height, 3276880 = 50 lines
@REM __ reg.exe add "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /v WindowSize /t REG_DWORD /d 3276880 /f
@REM __ @IF YOU HAVE DEFAULT 25 LINES WONT BE A REGISTRY KEY GET DELETED

@REM reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "FaceName" /t REG_SZ /d "Lucida Console"
@REM reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "FontFamily" /t REG_DWORD /d 54
@REM reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "FontSize" /t REG_DWORD /d 1048576
@REM reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "FontWeight" /t REG_DWORD /d 700
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "HistoryNoDup" /t REG_DWORD /d 0
reg.exe ADD "HKCU\Console\%%SystemRoot%%_system32_cmd.exe" /f /v "QuickEdit" /t REG_DWORD /d 2048


@REM -- WELL BEEN SETTING ALL THESE TODAY 11-12-2015 FOUND NEW ICON WITH REGMON TO JUMP TO REGISTRY
@REM -- MAKES EASIER -- AND REGMON NOT CRASH ON MY LOW END ASUS-EEE OR SPEECH E-PC
@REM -- EXPLORER CUSTOMIZE
@REM -- SET FOLDERS TO @REMEMBER 
@REM ------------------------------------
@REM __ MASTER_BOOT_RECORD IMPROVE ADJUSTER
@REM ------------------------------------
@REM __ XP _ MAYBE
@IF NOT "%WINXP__%"=="YES" GOTO NOT_WINDOWS_XP
	@ECHO WINDOWS XP
	@REM -- TweakUI.exe:17492	SetValue	HKCU\SOFTWARE\Microsoft\Windows\ShellNoRoam\Bag MRU Size	SUCCESS	0x1F4	
	@REM -- TweakUI.exe:17492	SetValue	HKCU\SOFTWARE\Microsoft\Windows\Shell\Bag MRU Size	SUCCESS	0x1F4	

	@REM -- NOT SURE
	@REM -- SEEMS TO BE A FAULT WITH ANOTHER ALMOST SAME SETTING HERE -- BagMRU Size
	@REM -- HEX 1F4 = 500
	@REM -- reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\ShellNoRoam" /f /v "Bag MRU Size" /t REG_DWORD /d 500
	@REM -- SEEMS TO BE A FAULT WITH ANOTHER ALMOST SAME SETTING HERE -- BagMRU Size
	@REM -- reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\Shell" /f /v "Bag MRU Size" /t REG_DWORD /d 500
:NOT_WINDOWS_XP

@IF NOT "%WINXP__%"=="YES" GOTO NOT_WINDOWS_XP
	@ECHO WINDOWS XP
	@REM ------------------------------------
	@REM -- NOT WIN_10 IS CORRECT VERIFY
	@REM -- ONLY XP
	@REM -- SIZE 64 = 100 PIXELS OF HEX YES
	@ECHO WINDOWS XP

	@REM -- -------------------------------
	reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /f /v "ThumbnailQuality" /t REG_DWORD /d 95
	reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /f /v "ThumbnailSize" /t REG_DWORD /d 100
	@REM -- EXPLORER CUSTOMIZE
	@REM -- SLIDE SHOW DELAY
	@REM -- ------------------
	@REM -- NOT WIN_10 IS CORRECT VERIFY
	reg.exe ADD "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellImageView" /f /v "Timeout" /t REG_DWORD /d 1
	@ECHO.
:NOT_WINDOWS_XP



@REM ----------------------------------------------------------------
@REM -- DON'T REQUIRE 2 AND I USE DEFRAGGLER PAID FOR
@REM -- HAVE A CONFLICT BECAUSE PIREA(SP) BOOT DEFRAGGLER DOES MY SYSTEM BOOT
@REM -- BOOT TIME DEFRAGGER = YES
@REM ----------------------------------------------------------------
@REM reg.exe ADD "HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction\Enable" /f /v "SUCCESS /t REG_SZ /d
@REM ----------------------------------------------------------------
@REM -- ENABLE BOOT DEFRAG FOR QUICKER BOOTING
@REM ----------------------------------------------------------------
@REM reg.exe ADD "HKLM\SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction" /f /v "Enable" /t REG_SZ /d "Y" 
@REM reg.exe ADD "HKLM\SOFTWARE\Microsoft\Dfrg" /f /v "Enable" /t REG_SZ /d "Y" 
@REM ----------------------------------------------------------------


@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM -- ----
@REM -- 01 OF 02 FIXER
@REM -- ----
@REM -- Microsoft ActiveX Common Control MSCOMCTL.OCX Security Update Problem with Registration Affects Treeview and ListBox in Microsoft Access, Office, VB6
@REM -- http://www.fmsinc.com/microsoftaccess/controls/mscomctl/
@REM -- ----
@REM -- 02 OF 02 MODERN FIXER
@REM -- ----
@REM -- MS12-060: Description of the security update for Office 2010: August 14, 2012
@REM -- https://support.microsoft.com/en-gb/help/2597986/ms12-060-description-of-the-security-update-for-office-2010-august-14,-2012
@REM -- ----
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- WINDOWS UPDATE TO NOT RUN AUTO BUT RUN MANUAL WHEN WANT WORKING NOT DO STOP SERVICE
@REM ----------------------------------------------------------------
@REM IF "%VARLATCH2%" == "WIN10" reg.exe ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" /f /v "AU" /t REG_DWORD /d 1
@REM ----------------------------------------------------------------

@REM Turn off the Error Dialog Via the Registry
@REM --------------------------------
IF "%VARLATCH2%" == "WIN10" reg.exe ADD "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\Windows Error Reporting" /f /v "DontShowUI" /t REG_DWORD /d 1
@REM --------------------------------

REM REG ADD "HKCR\exefile\shell\Run with Above Normal Priority" /ve /f
REM REG ADD "HKCR\exefile\shell\Run with Above Normal Priority\Command" /ve /t REG_SZ /d "cmd.exe /c start \"Run with Above Normal Priority\" /AboveNormal \"%%1\"" /f

REM --
REM GOT TO HAVE THIS RUNNING AS DEPENDENCY  
REM REMOTE HAVE DEPENDENCY ON IT OR NOT HAVE ANY HOMEGROUP OUTSIDE WORLD
REM --------------------------------------------------------------------
ECHO OFF
IF %VARLATCH2% == WIN10 (
	ECHO WIN 10 _ WIN SEARCH STARTER
	SC CONFIG WSearch START= AUTO
	NET START WSearch
)
IF %VARLATCH2% == WIN07 (
	ECHO WIN 07 _ WIN SEARCH STARTER
	SC CONFIG WSearch START= AUTO
	NET START WSearch
)
IF %COMPUTERNAME% == 3-LINDA-PC (
	ECHO 3-LINDA-PC _ WIN SEARCH DISABLED
	SC CONFIG WSearch START= DISABLED
	NET STOP WSearch
)
IF %VARLATCH2% == WINXP (
	ECHO WIN XP _ WIN SEARCH DISABLED
	SC CONFIG WSearch START= DISABLED
	NET STOP WSearch
)
ECHO ON
REM TASKKILL /F /FI "SERVICES eq WSearch"
REM --


@REM -- 2ND USEFUL ONE TO HAVE ABOUT WINDOWS 10 UPDATE
@REM -- BUT PROBLEM YOU CAN'T STOP AND UPDATE ONLY AFTER IT BEEN UPDATE
@REM -- ----
@REM -- ----------------------------------------------------------------------------------
@REM -- https://support.microsoft.com/en-us/kb/3073930?utm_source=twitter#bookmark-1607
@REM -- https://support.microsoft.com/en-us/kb/3073930?utm_source=twitter#bookmark-1607
@REM -- ----------------------------------------------------------------------------------

@REM -- ----------------------------------------------------------------------------------
@rem -- MICROSOFT ACCOUNT -- VS -- LOCAL USER
@REM -- ----------------------------------------------------------------------------------
@REM -- Form FindWindow ---
@REM -- How do I switch back to a local account sign-in for Windows 10? - Ask Leo! - Google Chrome
@REM -- ---------------------
@REM -- ----
@REM -- Make Windows 10/8 Show All User Accounts on Login Screen | Password Recovery
@REM -- http://www.top-password.com/blog/make-windows-10-8-show-all-user-accounts-on-login-screen/
@REM -- --------
@REM -- How do I switch back to a local account sign-in for Windows 10? - Ask Leo!
@REM -- https://askleo.com/how-do-i-switch-back-to-a-local-account-sign-in-for-windows-10/
@REM -- ----
@REM -- ---------------------
@REM -- ---------------------
@REM -- Count = 113 -- Sun 21-Aug-2016 22:31:38
@REM -- ---------------------
@REM -- Form FindWindow ---
@REM -- HOW TO CHANGE MICROSOFT EMAIL ACCOUNT USE AS LOGON SCREEN BUT STILL KEEP LOGIN - Google @REM @REM -- Search - Google Chrome
@REM -- ---------------------
@REM -- ----
@REM -- WIN10 MORE USER ON LOGON SCREEN - Google Search
@REM -- https://www.google.co.uk/search?q=WIN10+MORE+USER+ON+LOGON+SCREEN&oq=WIN10+MORE+USER+ON+LOGON+SCREEN&aqs=chrome..69i57j69i59.79818j0j7&sourceid=chrome&ie=UTF-8
@REM -- --------
@REM -- HOW TO CHANGE MICROSOFT EMAIL ACCOUNT USE AS LOGON SCREEN BUT STILL KEEP LOGIN - Google Search
@REM -- https://www.google.co.uk/search?q=HOW+TO+CHANGE+MICROSOFT+EMAIL+ACCOUNT+USE+AS+LOGON+SCREEN+BUT+STILL+KEEP+LOGIN&oq=HOW+TO+CHANGE+MICROSOFT+EMAIL+ACCOUNT+USE+AS+LOGON+SCREEN+BUT+STILL+KEEP+LOGIN&aqs=chrome..69i57.33651j0j7&sourceid=chrome&ie=UTF-8
@REM -- ----

@REM -- WHILE THE ENDING THIS SUBJECT BOOT UP AND SPECIAL SYSTEM FOLDER
@REM -- ----------------------------------------------------------------------------------
@REM -- 1.. DON'T FORGET AFTER THE CHANGE ABOVE USER BOOT IN WITH
@REM -- PUT AUTO BOOT IN BACK ON WITH HERE 
@REM -- ----------------------------------------------------------------------------------
@REM -- %windir%\system32\control.exe userpasswords2
@REM -- ----------------------------------------------------------------------------------

@REM -- 2..  NONE WILL AUTO BOOT IF THE SHORT-CUT IS SET ADMINISTRATOR
@rem -- YOU HAVE TO READ MORE THAN HALF OF THIS 
@REM -- COMPUTER ILLITERATE FOR NEWBIE-ISM TO -- FIND -- HUNT -- THE PRIZE Talk About Teaser Suspense
@REM -- IN FOLDER START-UP -- NONE WILL AUTO BOOT IF THE SHORT-CUT IS SET ADMINISTRATOR -- BACK TO FRONT
@REM -- ----
@REM -- How to Access the Windows 10 Startup Folder
@REM -- https://www.tekrevue.com/tip/windows-10-startup-folder/
@REM -- ----

@REM -- 3.. IF YOU GOT ANY VB6 BOOT DELAY LOADER -- BOOT STRAPPER STORM TROOPER 
@REM -- AND USE THE GETSHORTCUT LINK 
@REM -- GETSHORTLINK
@REM -- AND WITH IT
@REM -- GetLinkInfo
@REM -- THAT IS USED WITH ---- sheltarg.dll ---- FS.fileexists("sheltarg.dll") = False CURRENT WORKING DIRECTORY
@REM -- ----
@REM -- IF YOU ASK IT NICELY TO RETURN YOU THE LINK PATH AND IT A PROGRAM IN \PROGRAM FILES\ IT WILL 
@REM -- DEFAULT TO THINK ALL IS IN THE PROGRAM FILES (X86)\ -- 32 BIT FOLDER
@REM -- UNLESS I USE IT CORRECT HERE EXAMPLE
@REM -- lRet = GetLinkInfo(Form1.hWnd, sTemp, sLinkPath, sLinkDescript)
@REM -- I SUBMIT THE Form1.hWnd HWND OF MY PROGRAM CALLING FROM 
@REM -- AND VB6 IS ALL 32 BIT 
@REM -- WORK AROUND IT -- STINKER -- DETECT PRIORITY TO A 64 BIT IF EXIST IN 64 BIT FOLDER \PROGRAM FILES\ 
@REM -- AND 64 BIT COMPUTER
@REM -- ----





@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------


GOTO JUMPER_ENDER
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------

@ECHO ON

reg QUERY hkcr\typelib\{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}\2.0
@IF %errorlevel% EQU 0 (
	@ECHO THE REGISTRY KEY EXIST AND THEN DO ACTION
	@ECHO REGISTER THE DLL TAKE TIME AFTER SO SKILL THIS PART @IF DONE

	SET INFO1_1=^%systemroot^%\system32\regsvr32 /u mscomctl.ocx /s __ UNREGISTRER SILENTLY MODE
	SET INFO2_1=^%systemroot^%\system32\regsvr32 mscomctl.ocx /s _____ REGISTRER SILENTLY MODE
	SET INFO1_2=^%systemroot^%\sysWOW64\regsvr32 /u mscomctl.ocx /s __ UNREGISTRER SILENTLY MODE
	SET INFO2_2=^%systemroot^%\sysWOW64\regsvr32 mscomctl.ocx /s _____ REGISTRER SILENTLY MODE

	Reg delete hkcr\typelib\{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}\2.0 /f
	@IF exist %systemroot%\SysWOW64\cscript.exe goto 64 
	@ECHO %INFO1_1%
	%systemroot%\system32\regsvr32 /u mscomctl.ocx /s
	@ECHO %INFO2_1%
	%systemroot%\system32\regsvr32 mscomctl.ocx /s
	@GOTO exit-ER
	:64 
	@ECHO %INFO1_2%
	%systemroot%\sysWOW64\regsvr32 /u mscomctl.ocx /s
	@ECHO %INFO2_2%
	%systemroot%\sysWOW64\regsvr32 mscomctl.ocx /s
	@GOTO exit-ER

	@REM -- To run the regsvr32 commands silently, add a /s switch to each of the regsvr32 commands in the previous example. For example:
	@REM %systemroot%\system32\regsvr32 /u mscomctl.ocx /s

	:exit-ER
)

@REM ----------------------------------------------------------------



@REM PAUSE



@REM ----------------------------------------------------------------
@REM -- NOTE DUMPING
@REM ----------------------------------------------------------------
@REM -- PREVIOUS VERSION MICROSOFT KEYBOARD NOT GOODER
@REM ----------------------------------------------------------------
@REM reg.exe ADD "HKCU\SOFTWARE\Microsoft\IntelliPoint\Scrolling" /f /v "ScrollLines" /t REG_DWORD /d 5
@REM reg.exe ADD "HKCU\SOFTWARE\Microsoft\IntelliPoint\Scrolling" /f /v "AccelerationCurve" /t REG_DWORD /d 5
@REM ----------------------------------------------------------------

:JUMPER_ENDER

@REM PAUSE
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@ECHO OFF

@IF NOT "%WINXP__%"=="YES" GOTO NOT_WINDOWS_XP
@ECHO WINDOWS XP
@REM 40 SECONDS
FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
FOR /L %%# IN (40,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)


:NOT_WINDOWS_XP
@IF "%WINXP__%"=="YES" GOTO WINDOWS_XP
@ECHO WINDOWS 7 TO WINDOWS 10
TIMEOUT /T 40

:WINDOWS_XP

:VERY_ENDER_BLOCK
:ENDE

EXIT /B

EXIT
