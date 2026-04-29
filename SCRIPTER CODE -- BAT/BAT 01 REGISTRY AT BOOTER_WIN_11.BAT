

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

REM IF '%1%'=='/ADMIN' (
	REM @REM :: Check for ADMIN Privileges Again
	REM >nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
	REM @IF '%errorlevel%' GTR '2' SET NOT_GOT_ADMIN=TRUE
REM )

REM @IF "%NOT_GOT_ADMIN%"=="TRUE" (
	REM ECHO AN ATTEMPT WAS MADE TO SET ADMIN BUT NOT SUCCESSFUL
	REM ECHO YOUR USER MUST BE A STANDARD USER OR SOMETHING LIKE HER
	REM ECHO SAVES GET STUCK IN A LOOP ASKING OR ALSO DELAY REQUIRING
	REM ECHO END
	REM PAUSE
	REM GOTO ENDE
	REM EXIT /B 1
REM )

REM @REM :: Check for ADMIN Privileges
REM >nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
REM @IF '%errorlevel%' GTR '2' (
	REM @REM Get ADMIN Privileges
	REM @ECHO Set UAC = CreateObject^("Shell.Application"^) >> "%temp%\getadmin.vbs"
	REM @ECHO UAC.ShellExecute "%~s0", "/ADMIN", "", "runas", 1 >> "%temp%\getadmin.vbs"

	REM "%temp%\getadmin.vbs"
	REM @DEL "%temp%\getadmin.vbs"
	REM GOTO ENDE
	REM @EXIT /B
REM )

REM @ECHO Got ADMIN Privileges



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
@IF %COMPUTERNAME%==1-ASUS-X5DIJ     SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==2-ASUS-EEE       SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==3-LINDA-PC       SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==4-ASUS-GL522VW   SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==5-ASUS-P2520LA   SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==7-ASUS-GL522VW   SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==8-MSI-GP62M-7RD  SET MATTHEW_S_COMPUTER__=YES
@IF %COMPUTERNAME%==9-ASUS-G815LM    SET MATTHEW_S_COMPUTER__=YES
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
@IF %COMPUTERNAME%==9-ASUS-G815LM    SET COMPUTER_64BIT__=YES


@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM PROBLEM READ THIS VAR FROM WINDOWS XP BOMBS OUT BATCH SCRIPT
@REM ----------------------------------------------------------------
@REM @IF %PROCESSOR_ARCHITEW6432%==AMD64  SET COMPUTER_64BIT__=YES
@REM ----------------------------------------------------------------
@ECHO ON





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




@REM ------------------------------------------------------------------
@REM -- DISABLED USER ACCOUNT CONTROL UAC
@REM -- TWO SETTINGS ONE OF THEM IS OBVIOUS
@REM ULTIMATE TWEAKER
@REM Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v "NoDispCPL" /t REG_DWORD /d "0" /f

@REM THIS IS SET ABOVE AND DISABLED USER ACCOUNT CONTROL IS NOT WANTED ANYMORE
@REM AS WINDOWS SOTRE WON;T OPEN IN A ADMIN ACCOUNT
REM --------------------------------------------------------------------
REM -- WANTED IN WINDOWS 11 CAN'T INSTALL WINAMP LATEST VERSION PUT A BLOCK 
REM -- AN ADMINISTRATOR HAS BLOCKED YOU FROM RUNNING THIS APP
REM -- RESTART COMPUTER REQUIRE FOR THIS ONE
REM -- CHANGE THE REG_DWORD VALUE FROM 1 TO 0
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v "EnableLUA" /t REG_DWORD /d "0" /f
@REM ------------------------------------------------------------------




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

REM ----------------------------------------------------
REM -- HERE COMES IN AS THE BOTTOM ENTRY IN CONTEXT MENU
REM -- AND IS THE VAILD ONE -- THE TOP ONE IN CONTEXT MENU ISN'T WORKING
REM -- 2026 04 25
REM ----------------------------------------------------

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

:END_OF_NOTEPAD_ROUTINE

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

REM -- IF ABOVE OPTION NOT SET THEN DONT RUN ANY 
REM --------------------------------------------
GOTO DONT_WANT_NOTEPAD_CHANGER

:PROGRAM_FILES_64_OR_32_BIT
	@ECHO NOTEPAD++ 64 OR 32 BIT __ C:\Program Files\
	@ECHO.
	Reg.exe add "HKLM\SOFTWARE\Classes\CLSID\*\shell\OpenWithNotepad" /ve /t REG_SZ /d "Edit with &Notepad++" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\CLSID\*\shell\OpenWithNotepad" /v "icon" /t REG_SZ /d "C:\Program Files\Notepad++\Notepad++.exe" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\CLSID\*\shell\OpenWithNotepad\Command" /ve /t REG_SZ /d "\"C:\Program Files\Notepad++\Notepad++.exe\" \"%%1\"" /f
	
GOTO END_OF_NOTEPAD_ROUTINE


REM ----------------------------------------------------------------------------------------------------------
REM -- THIS ENTRY ISN'T WORKING -- CAN'T SEE IN CONTEXT MENU -- MIGHT BE WORKING SOMEWHERE FORGOTEN HOW WORKER
REM -- 2026-04-25
REM -- THIS ENTRY IS NOT GETTNING SET BY THIS CODE MAYBE PERMISSION
REM -- BUT WRONG ENTRY UNDER COMMAND "C:\Program Files (x86)\Notepad++\Notepad++.exe" "%%1"
REM -- SHOULD BE %1
REM -- CAN'T BE ASSED MUCKING AROUND WITH ME
REM ----------------------------------------------------------------------------------------------------------
:PROGRAM_FILES_32_BIT
	@ECHO NOTEPAD++ 32 BIT __ C:\Program Files (x86)\
	@ECHO.
	Reg.exe add "HKLM\SOFTWARE\Classes\CLSID\*\shell\OpenWithNotepad" /ve /t REG_SZ /d "Edit with &Notepad++" /f
	REM -- ICON
	Reg.exe add "HKLM\SOFTWARE\Classes\CLSID\*\shell\OpenWithNotepad" /v "icon" /t REG_SZ /d "C:\Program Files (x86)\Notepad++\Notepad++.exe" /f
	REM -- COMMAND
	Reg.exe add "HKLM\SOFTWARE\Classes\CLSID\*\shell\OpenWithNotepad\Command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\Notepad++\Notepad++.exe\" \"%%1\"" /f


:END_OF_NOTEPAD_ROUTINE





for /f "delims=" %%a in ('reg query "HKLM\SOFTWARE\Classes\CLSID" /s /f "ANotepad++64" /t REG_SZ') do @echo %%a|findstr /v "REG_SZ" | findstr /v /c:"End of search:" >nul && reg.exe delete "%%a" /f

for /f "delims=" %%a in ('reg query "HKCR\CLSID" /s /f "ANotepad++64" /t REG_SZ') do @echo %%a|findstr /v "REG_SZ" | findstr /v /c:"End of search:" >nul && reg.exe delete "%%a" /f



@ECHO.
@ECHO ---------------------------------------------------------------
@ECHO ---------------------------------------------------------------
@ECHO.






@REM ---------------- WORKS INSTANTLY NOT REQUIRE TO REBOOT OR LOG OFF
@REM ---------------- YES THAT WORKING FOR WINDOWS 11
@REM @ECHO ---------- SET NOTEPAD++ FOR EDIT AUTOHOTKEYS AHK FILES
IF "%WIN10__%"=="YES" (
	reg.exe ADD "HKEY_CLASSES_ROOT\AutoHotkeyScript\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "C:\Program Files (x86)\Notepad++\notepad++.exe %%1"
)

REM Windows Registry Editor Version 5.00 
REM [HKEY_CLASSES_ROOT\AutoHotkeyScript\Shell\Edit\Command]
REM @="\"C:\\Program Files\\Notepad++\\notepad++.exe\" \"%1\""
REM ----------------------------------------
REM Configure Autohotkey to edit scripts with Notepad++ - Stack Overflow
REM ----------------------------------------
REM https://stackoverflow.com/questions/45552141/configure-autohotkey-to-edit-scripts-with-notepad
REM ----------------------------------------



@REM ----------------------------------------------------------------
@REM @ECHO ---------- RESTORE ORIGINAL NOTEPAD FOR TXT ------------------------------------
@REM reg.exe ADD "HKEY_CLASSES_ROOT\AutoHotkeyScript\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE %%1"
@REM ----------------------------------------------------------------
@REM -------- ################################# ------------------------------------------------
@REM ----------------------------------------------------------------



@REM -------- ################################# ---------------------------------------------------------------
@REM -------- NOT SET NOTEPAD++ AT MOMENT 64 BIT RELOCATE SORT TO DO
@REM -------- SET NOTEPAD++ --------------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM @ECHO ---------- SET NOTEPAD++ FOR EDIT BATCH FILES
@REM ----------------------------------------------------------------
@ECHO WINDOWS 10 -- HERE WORKS FOR WINDOWS 11 OPEN BATCH FILE TO NOTEPAD++ WHEN EDIT

IF "%WIN10__%"=="YES" (
	reg.exe ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "C:\Program Files (x86)\Notepad++\notepad++.exe %%1"
)

@REM ----------------------------------------------------------------
@REM @ECHO ---------- RESTORE ORIGINAL NOTEPAD FOR BAT ------------------------------------
REM reg.exe ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE %%1"
@REM ----------------------------------------------------------------
@REM -------- ################################# ------------------------------------------------
@REM ----------------------------------------------------------------



:DONT_WANT_NOTEPAD_CHANGER





	
REM PAUSE









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

	REM Reg.exe add "HKCR\Folder\shell\MenuText" /ve /t REG_SZ /d "Open a Command Prompt Here" /f
	REM Reg.exe add "HKCR\Folder\shell\MenuText\Command" /ve /t REG_SZ /d "cmd.exe /K chdir /D \"%%1\"" /f
	

	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	@REM __ TO SET MY PERSONAL PROGRAM COMMAND IN THE CONTEXT MENU FOR 
	@REM _________________________________________________________________
	@REM __ OPEN A FILE IN EXPLORER
	@REM _________________________________________________________________
	@REM __ IN-CASE WANT SOME RENAMING FIRST
	@REM __ DELETE AND RECREATE
	@REM _________________________________________________________________
	
	REM reg.exe DELETE "HKLM\Software\Classes\Drive\shell\Multi -------- Drive XXXXXXXXXXXXXXXXXXXXX" /f
	REM reg.exe DELETE "HKLM\Software\Classes\Directory\shell\Multi -------- Directory XXXXXXXXXXXXX" /f
	REM Reg.exe DELETE "HKLM\Software\Classes\Directory\Background\shell\Multi -------- XXXXXXXXXXXX" /f

	SET NAME_V1=MULTI MENU ____
	SET NAME_V2=%NAME_V1% FOLDER
	SET NAME_V3=%NAME_V1% DRIVE
	SET NAME_V4=D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe
	
	reg.exe DELETE "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%" /f
	reg.exe DELETE "HKLM\Software\Classes\Folder\shell\%NAME_V2%" /f
	REM --------------------------------------------------------------
	REM -- MULTI MENU ____
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%" /ve /t REG_SZ /d "%NAME_V1%" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%" /v "icon" /t REG_SZ /d "%NAME_V4%" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%\Command" /ve /t REG_SZ /d "\"%NAME_V4%\" \"%%1\"" /f
	REM --------------------------------------------------------------
	REM -- MULTI MENU ____
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\Software\Classes\Folder\shell\%NAME_V2%" /ve /t REG_SZ /d "%NAME_V2%" /f
	Reg.exe add "HKLM\Software\Classes\Folder\shell\%NAME_V2%" /v "icon" /t REG_SZ /d "%NAME_V4%" /f
	Reg.exe add "HKLM\Software\Classes\Folder\shell\%NAME_V2%\Command" /ve /t REG_SZ /d "\"%NAME_V4%\" \"%%1\"" /f
	REM --------------------------------------------------------------
	REM -- MULTI MENU ____
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\Software\Classes\Directory\Background\shell\%NAME_V1%" /ve /t REG_SZ /d "%NAME_V1%" /f
	Reg.exe add "HKLM\Software\Classes\Directory\Background\shell\%NAME_V1%" /v "icon" /t REG_SZ /d "%NAME_V4%" /f
	Reg.exe add "HKLM\Software\Classes\Directory\Background\shell\%NAME_V1%\Command" /ve /t REG_SZ /d "\"%NAME_V4%\" \"%%1\"" /f
	REM --------------------------------------------------------------
	REM -- MULTI MENU ____
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\Software\Classes\Drive\shell\%NAME_V3%" /ve /t REG_SZ /d "%NAME_V3%" /f
	Reg.exe add "HKLM\Software\Classes\Drive\shell\%NAME_V3%" /v "icon" /t REG_SZ /d "%NAME_V4%" /f
	Reg.exe add "HKLM\Software\Classes\Drive\shell\%NAME_V3%\Command" /ve /t REG_SZ /d "\"%NAME_V4%\" \"%%1\"" /f

	
	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	SET NAME_V1=DATE MODIFY ____
	SET NAME_V2=%NAME_V1% FOLDER
	SET NAME_V4=D:\VB6\VB-NT\00_Send_To\Send To MODIFY DATE\#SEND TO MODIFY DATE.EXE
	reg.exe DELETE "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%" /f
	reg.exe DELETE "HKLM\Software\Classes\Folder\shell\%NAME_V2%" /f
	REM --------------------------------------------------------------
	REM -- DATE MODIFY ____
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%" /ve /t REG_SZ /d "%NAME_V1%" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%" /v "icon" /t REG_SZ /d "%NAME_V4%" /f
	Reg.exe add "HKLM\SOFTWARE\Classes\*\shell\%NAME_V1%\Command" /ve /t REG_SZ /d "\"%NAME_V4%\" \"%%1\"" /f
	REM --------------------------------------------------------------
	REM -- DATE MODIFY ____
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\Software\Classes\Folder\shell\%NAME_V2% Folder" /ve /t REG_SZ /d "%NAME_V2%" /f
	Reg.exe add "HKLM\Software\Classes\Folder\shell\%NAME_V2% Folder" /v "icon" /t REG_SZ /d "%NAME_V4%" /f
	Reg.exe add "HKLM\Software\Classes\Folder\shell\%NAME_V2% Folder\Command" /ve /t REG_SZ /d "\"%NAME_V4%\" \"%%1\"" /f
	REM --------------------------------------------------------------
	REM -- DATE MODIFY ____
	REM --------------------------------------------------------------
	Reg.exe add "HKLM\Software\Classes\Directory\Background\shell\%NAME_V1%" /ve /t REG_SZ /d "%NAME_V1%" /f
	Reg.exe add "HKLM\Software\Classes\Directory\Background\shell\%NAME_V1%" /v "icon" /t REG_SZ /d "%NAME_V4%" /f
	Reg.exe add "HKLM\Software\Classes\Directory\Background\shell\%NAME_V1%\Command" /ve /t REG_SZ /d "\"%NAME_V4%\" \"%%1\"" /f
	REM --------------------------------------------------------------
	REM --------------------------------------------------------------

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

	@REM __ HKLM\SOFTWARE\Classes\*
	@REM __ HKLM\Software\Classes\Folder\shell\yourappname\command
	@REM __ HKLM\Software\Classes\Directory\shell\yourappname\command
	@REM __ HKLM\Software\Classes\Drive\shell\yourappname\command
	



	REM --------------------------------------------------------------
	REM --------------------------------------------------------------
	REM ONE MORE FOR GOOD MEASURE
	REM NOT SURE WHAT ONE IS PUTTING IN ENTRY FOR CMD.COM BUT FORGETTING ICON
	REM WHEN ICON IS PUT IN FROM SYSTEM32 IT CHANGES IT TO THIS ONE C:\Windows\SysWow64\cmd.exe
	REM -- HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\background\shell\cmd\command
	REM Reg.exe add "HKLM\SOFTWARE\Classes\Directory\background\shell\cmd" /v "icon" /t REG_SZ /d "C:\Windows\SysWow64\cmd.exe" /f
	
	REM ANOTHER FEW BUT SOME BLOCKED LIKE CMD.EXE
	REM -------------------------------------------------------------
	Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\shell\Browse with &IrfanView" /v "icon" /t REG_SZ /d "C:\Program Files\IrfanView\i_view32.exe" /f
	
	REM Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\shell\Browse with XnView" /v "icon" /t REG_SZ /d "C:\Program Files\XnView\xnview.exe" /f
	

	
	Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Drive\shell\Browse with &IrfanView" /v "icon" /t REG_SZ /d "C:\Program Files\IrfanView\i_view32.exe" /f
	
	REM Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Drive\shell\Browse with XnView" /v "icon" /t REG_SZ /d "C:\Program Files\XnView\xnview.exe" /f
	
	
	REM Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\shell\Powershell" /v "icon" /t REG_SZ /d "C:\Windows\SysWow64\WindowsPowerShell\v1.0\powershell.exe" /f
	REM ----------------------------------
	rem can't write to powershell in drive
	
	REM AND MORE IN FOLDER
	REM -------------------------------------------------------------
	REM Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\*\shell\filealyzer2" /v "icon" /t REG_SZ /d "C:\Program Files (x86)\Safer Networking\FileAlyzer 2\FileAlyzer2.exe" /f
	
	
	REM SETTING -- HERE DON'T ADD ICON WINDOWS 11
	Reg.exe add "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\shell\HashMyFiles" /v "icon" /t REG_SZ /d "C:\PStart\Progs\0_Nirsoft_Package\NirSoft\x64\hashmyfiles.exe" /f

	REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\x64\hashmyfiles.exe
	
	
	
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
	
	
	REM PAUSE
	
	REM ADD ICON TO THIS ONE DIDN'T DO ANYTHING _ SEEN YET _ TAKEN OUT
	REM HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Folder\shell\explore
	REM -------------------------------------------------------------
	REM THAT MUCKED UP MY DESKTOP ICONS NOW GONE BLACK _ TRY REBOOT
	REM DESKTOP CAME BACK AFTER WORKING ON SOMETHING OTHER
	REM -------------------------------------------------------------
	
	
	
	
	
	REM CHECK HERE
	REM ----------
	@REM -- MY PERSONAL PREF SOME APP LIKE GOOD SYNC TAKE A WHILE TO WIND IT IN
	@REM -- TURN OFF AUTOMATIC APPLICATIONS TERMINATION AT SHUTDOWN
	@REM -- ULTIMATE TWEAKER
	
	REM REMOVER 2026 04 25
	REM ------------------
	
	REM Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /v "AllowBlockingAppsAtShutdown" /t REG_DWORD /d "1" /f
	
	@REM -- THATS ALL OF ULTIMATE TWEAKER WINDOWS 10 __ DONE
	
	
	
	
REM REMOVER 2026-04-25 ABORT PUT BACK ONLY WORRY ABOUT WINDOWS 11
REM -------------------------------------------------------------
REM GOTO JUMP_HERE
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
	REM REMOVE WIN 11 2026
	REM Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v "EnableInstallerDetection" /t REG_DWORD /d "1" /f

	@REM -- NOT_ENABLE ADMIN APPROVAL MODE FOR BUILT IN ADMINISTRATOR ACCOUNT
	@REM -- SETTING GOOD TO HAVE AS DEFAULT		
	@REM ULTIMATE TWEAKER
	REM REMOVE WIN 11 2026
	REM Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v "FilterAdministratorToken" /t REG_DWORD /d "1" /f
	
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
	REM REMOVE WIN 11 2026
	REM Reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows Defender" /v "DisableAntispyware" /t REG_DWORD /d "1" /f

	@REM REGISTRY DISABLED WINDOWS UPDATE
	@REM ULTIMATE TWEAKER
	@REM Reg.exe add "HKLM\SYSTEM\ControlSet001\Services\wuauserv" /v "Start" /t REG_DWORD /d "4" /f

	@REM -- DISABLED THE BING SEARCH WHEN IN START MENU AND TYPING WEB SEARCH RESULT GET IN THE WAY
	@REM --
	REM REMOVE WIN 11 2026
	REM Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" /v "BingSearchEnabled" /t REG_DWORD /d "0" /f

	
	@REM ----------------------------------------------------------------
	@REM -- SOME ENABLED DISABLE APPS SHARING INFO
	@REM -- NOT SURE @IF THIS IS WINDOWS STORE APP I THINK IT TRUE
	@REM -- NOT SURE ABOUT SETTING @IF RELEVANT TO MY MACHINE ON MY WITH CODE NUMBER BEEN CHECKED OUT UP ON
	@REM --     ENABLED = "Allow"
	@REM -- NOT ENABLED = "Deny"
	@REM --
	@REM -- APPS MESSAGES
	REM REMOVE BLOCK WIN 11 2026
	REM ------------------------
	REM Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\{992AFA70-6F47-4148-B3E9-3003349C1548}" /v "Value" /t REG_SZ /d "Allow" /f
	@REM -- APPS MICROPHONE
	REM Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\{2EEF81BE-33FA-4800-9670-1CD474972C3F}" /v "Value" /t REG_SZ /d "Allow" /f
	@REM -- APPS CAMERA
	REM Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\{E5323777-F976-4f5b-9B55-B94699C46E44}" /v "Value" /t REG_SZ /d "Allow" /f
	@REM -- APPS CALENDAR
	REM Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\{D89823BA-7180-4B81-B50C-7E471E6121A3}" /v "Value" /t REG_SZ /d "Allow" /f
	
	@REM -- APPS LOCATION
	REM Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\DeviceAccess\Global\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}" /v "Value" /t REG_SZ /d "Allow" /f
	@REM ----------------------------------------------------------------

REM JUMP PUT 2026-04-25
:JUMP_HERE
	
	
	
	
	
:NOT_MATTHEW_S_MACHINE


PAUSE






