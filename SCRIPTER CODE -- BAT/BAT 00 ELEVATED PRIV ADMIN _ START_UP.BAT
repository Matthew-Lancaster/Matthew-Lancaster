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

@REM __ FROM ---- Mon 28-May-2018 13:53:21
@REM __ TO   ---- Mon 28-May-2018 14:31:43 __ 40 MINUTE
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


pushd "%cd%"

cd /d "%~dp0"

CD \

REM ------------------------------------------------------------------------------
REM __ THIS HAS PROBLEM WHEN EXECUTING A PROGRAM THROUGHT NOTEPAD++ SHELLEXE LEARN
REM __ IT SET THE ENVIROMENT VARIABLES TO EMULATE WINDOWS XP HERE
REM ------------------------------------------------------------------------------
REM __ FIND WORKS OKAY IN WINDOWS XP AFTER ALL
REM ------------------------------------------------------------------------------
@REM @IF %PROCESSOR_ARCHITECTURE%==x86  SET WINXP__=YES
@REM @IF %PROCESSOR_ARCHITECTURE%==x86  GOTO X86_WIN_XP_32_BIT_COMPUTER

@REM -- JUMP OVER -- THESE COMMAND CAUSE CMD TO BOMB OUT IN 32 BIT XP

ver | find "5.1." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WINXP
IF %errorlevel% EQU 0 SET WINXP__=YES
ver | find "6.1." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WIN07
ver | find "10.0." > nul
IF %errorlevel% EQU 0 SET VARLATCH2=WIN10


@ECHO OFF
ECHO.
@ECHO PROGRAM RUNNER NAME____
ECHO.
ECHO %~nx0
ECHO.
@ECHO PATH NAME______________
ECHO.
ECHO %~dp0
ECHO.



REM TIMEOUT /T 2

REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\nircmd sendkeypress rwin+"m"



REM CURRENT USER
Reg.exe add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /F /v "STARTUP_COMMON_04_ALL" /t REG_SZ /d "\"%~f0\"
ECHO.


REM ALL USERS
REM Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /F /v "STARTUP_COMMON_04_ALL" /t REG_SZ /d "\"E:\01 Start Menu\#_0_STARTUP_COMMON_04_ALL\BAT 00 ELEVATED PRIV ADMIN _ START_UP.BAT\"

REM START "" /WAIT "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 21-AutoRun.ahk"

REM PAUSE

ECHO RUN_NORTON_ONCE_AT_BOOT_WITH_MIDNIGHT^>"C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2_RUN_NORTON_ONCE_AT_BOOT_AR_#NFS_EX_%A_ComputerName%.TXT"


ECHO RUN_NORTON_ONCE_AT_BOOT_WITH_MIDNIGHT>"C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_2_RUN_NORTON_ONCE_AT_BOOT_AR_#NFS_EX_%A_ComputerName%.TXT"

ECHO.


START "" "C:\Program Files\AutoHotkey\AutoHotkey.exe" "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 21-AutoRun.ahk"
ECHO.

SET X_TIME=24

@IF NOT "%WINXP__%"=="YES" GOTO NOT_WINDOWS_XP
@REM PAUSE
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@ECHO OFF

@ECHO WINDOWS XP
@REM 40 SECONDS



FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
FOR /L %%# IN (%X_TIME%,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)

:NOT_WINDOWS_XP
@IF "%WINXP__%"=="YES" GOTO WINDOWS_XP
@ECHO WINDOWS 7-UP TO WINDOWS 10
TIMEOUT /T %X_TIME%

:WINDOWS_XP

:VERY_ENDER_BLOCK

:ENDE

EXIT /B

EXIT





REM PAUSE

REM TIMEOUT /T 10


REM %~F0
REM D:\VB6\VB-NT\00_Best_VB_01\DELAYED_LOADER\SYSTEM_DELAYED_LOADER.exe

REM ECHO RUN GoodSync\GoodSync-v10.exe
REM --------------------------------------------------------------------------------

REM START "" /WAIT /B "C:\Program Files\Siber Systems\GoodSync\GoodSync-v10.exe"

REM TIMEOUT /T 1

REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\nircmd sendkeypress rwin+"m"

REM ECHO RUN CLIPBOARD LOGGER
REM --------------------------------------------------------------------------------


REM IF %VARLATCH2% == WINXP GOTO NOT_HERE
REM IF "%COMPUTERNAME%" == "3-LINDA-PC" GOTO NOT_HERE
REM IF NOT EXIST=="D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe" GOTO NOT_HERE_2
REM START "" /WAIT /B "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe"
REM :NOT_HERE_2

REM IF NOT EXIST=="D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe" GOTO NOT_HERE_2
REM START "" /WAIT /B "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
REM :NOT_HERE_2
REM :NOT_HERE

REM HAS ERROR LOAD FROM NORMAL BOOT
REM MEANT TO BE ELEVATED
REM WHAT'S IT DO NOT SURE
REM MOUSE WORK MORE BETTER MORE POKEY WITH THAT ON
REM IF %VARLATCH2% == WINXP GOTO NOT_HERE

REM IF NOT EXIST=="C:\Program Files\Logitech\SetPointP\SetPoint.exe" GOTO NOT_HERE
REM START "" /WAIT /B "C:\Program Files\Logitech\SetPointP\SetPoint.exe" "/launchGaming"
REM :NOT_HERE
REM DONE BY AHK


REM START "" "C:\Program Files (x86)\Siber Systems\AI RoboForm\RoboTaskBarIcon.exe"

REM IF %VARLATCH2% == WINXP GOTO NOT_HERE
REM IF %VARLATCH2% == WIN07 GOTO NOT_HERE

REM __ FROM __ HKCU\Software\Microsoft\Windows\CurrentVersion\Run
REM START "" "C:\Program Files (x86)\Google\Drive\googledrivesync.exe" /autostart
REM DONE BY AHK

REM START "" "C:\Program Files\Siber Systems\GoodSync\GoodSync-v10.exe" /min
REM DONE BY AHK

REM START "" "C:\Program Files (x86)\Nokia\Nokia Suite\NokiaSuite.exe" -tray
REM DONE BY AHK

REM START "" "C:\Users\MATT 01\AppData\Local\Microsoft\OneDrive\OneDrive.exe" /background
REM DONE BY AHK

REM START "" "C:\Program Files (x86)\Pushbullet\pushbullet.exe" -show false
REM DONE BY AHK

REM START "" "C:\Program Files\OVH\hubiC\hubiC.exe"
REM DONE BY AHK

REM __ FROM __ HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Run
REM START "" "C:\Program Files (x86)\Dropbox\Client\Dropbox.exe" /systemstartup
REM DONE BY AHK

REM START "" "C:\Program Files (x86)\QNAP\Qfinder\QfinderPro.exe" /min /auto
REM DONE BY AHK

REM START "" "C:\Program Files (x86)\WordWeb\wweb32.exe" -startup
REM DONE BY AHK

REM START "" "C:\Program Files (x86)\LastPass\lastapp_x64.exe"
REM DONE BY AHK

REM __ FROM __ HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run
REM START "" "C:\Program Files\Box\Box Sync\BoxSync.exe" -m
REM DONE BY AHK

REM START "" "%ProgramFiles%\Windows Defender\MSASCuiL.exe"
REM PROBLEM

REM :NOT_HERE

REM __ FROM __ HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run
REM DONE BY AHK
REM START "" "C:\Program Files\Logitech\SetPointP\SetPoint.exe" /launchGaming
REM START "" "C:\Program Files\Logitech\LogiOptions\LogiOptions.exe" /noui


REM START "" /MIN "C:\PStart\Progs\#_PortableApps\Start.exe"


REM ECHO MINIMIZE	
REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\nircmd sendkeypress rwin+"m"

REM ECHO TASKKILL ELITE*
REM --------------------------------------------------------------------------------
REM TIMEOUT /T 1
REM TASKKILL /F /IM ELITE* /T

REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\nircmd sendkeypress rwin+"m"

REM START "" "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 01-F10 __ HOTKEY __ PRINT SCREEN.ahk"	

REM START "" "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 02-SAVE AS KEY ENTER ON.ahk"

REM START "" "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 04-SET CAP NUM & SCROLL TO LIKING - ONCE.ahk"	

REM START "" "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 10-READ MOUSE CURSOR ICON STATE AND BEEPER WHEN NOT BUSY HOUR GLASS OVER.ahk"

REM START "" "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 14-Brightnes With Dimmer.ahk"

REM START "" "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 15-BLOCK KEY CTRL ENTER ON WEB PAGES MYSMS.ahk"

REM START "" "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 17-SWAP ALL ENTER KEY TO CTRL ENTER CHROME.ahk"

REM START "" "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk"


REM START "" /WAIT /B "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 11-F9 __ HOTKEY __ DELETE FILTER MACRO FOR _GOODSYNC CTRL_F9 PAUSE.ahk"
REM START "" /WAIT /B "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 13-LOGITECH_SPEAKER_AUDIO 1Hz.ahk"




REM A STARTUP TIME FOR GOODSYNC AND REQUIRE DELAY HERE MINIMUM = 10 NOT PROPER
REM --------------------------------------------------------------------------

REM CALL TIMEOUT /T 1
REM TASKKILL /F /IM ELITE* /T	

REM REQUIRE ALREADY RUNNNER DETECTION
REM PUT BACK IN SHORTCUT
REM start "" "C:\Program Files\OVH\hubiC\hubiC.exe"

REM CMD /C "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01 REGISTRY AT BOOTER.BAT"

REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\nircmd sendkeypress rwin+"m"

REM TIMEOUT /T 1

REM C:\PStart\Progs\0_Nirsoft_Package\NirSoft\nircmd sendkeypress rwin+"m"

REM start "" "D:\VB6\VB-NT\00_Best_VB_01\DELAYED_LOADER\SYSTEM_DELAYED_LOADER.exe"



REM START "" /MIN "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 21-NORTON UPDATE and QUICK SCAN.BAT"

REM __ EDIT THE FONT OF %COMSPEC% HERE BY DOING THIS
REM __ START "" "%COMSPEC%"

REM START "" /MIN "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01 BOOT KILLER.BAT"

REM START "" /MIN "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01 REGISTRY AT BOOTER.BAT"


REM __ CHECK AT START FOR ONE VARIABLE AND RUN THIS ONE IF NEEDER
REM IF NOT EXIST "E:\01 Start Menu" (
	REM ECHO RUN THE CODE __ # TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT
	REM START "" /MIN "E:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\# TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT"
REM )



REM GOTO NOT_HERE
REM IF NOT %VARLATCH2% == WIN10 GOTO NOT_HERE
REM IF NOT EXIST=="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 12-KEEP WINDOWS UPDATE STOPPED.BAT" GOTO NOT_HERE
REM START "" /WAIT /B "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 12-KEEP WINDOWS UPDATE STOPPED.BAT"
REM :NOT_HERE



REM EXIT 

REM CMD /C "E:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\# TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT"



REM HELP ABOUT MINIMIZE ALL WINDOW AND THRON TO MIN AND RESTORE D RATHER THAN M MINIMIZE

REM ----
REM batch script minimize all windows - Google Search
REM https://www.google.co.uk/search?q=batch+script+minimize+all+windows&oq=MINIMIZE+ALL+WINDOWS+BATCH+&aqs=chrome.2.0j69i57j0.15690j0j4&sourceid=chrome&ie=UTF-8
REM --------
REM minimize all windows batch file - Stack Overflow
REM https://stackoverflow.com/questions/25596736/minimize-all-windows-batch-file
REM --------
REM How to minimize all desktop windows via scripts ?
REM https://www.experts-exchange.com/questions/22154358/How-to-minimize-all-desktop-windows-via-scripts.html
REM ----




REM START "" /WAIT /B "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\TASKKILLER ELITESPY.BAT"
REM START "" /WAIT /B "TASKKILL" "/F /IM ELITE* /T"

REM CALL TASKKILL /F /IM CMD* /T


REM _ REFERENCES
REM --------------------------------------------------------------------------------------
REM accolade 22 November 2011 20:38:16
REM Here's a one-liner that doesn't require additional tools:
REM ----
REM How to determine in batch-file, if user is administrator - Microsoft - Windows 2000/NT
REM http://www.tomshardware.co.uk/forum/169421-46-determine-batch-file-user-administrator
REM ----

REM source: http://stackoverflow.com/q/4054937

REM Here is how I use it, in a routine that elevates a batch script itself:
REM :: Ensure ADMIN Privileges

REM :: adaptation of <a href="https://sites.google.com/site/eneerge/home/BatchGotAdmin" rel="nofollow" target="_blank">https://sites.google.com/site/eneerge/home/BatchGotAdmi...</a> and <a href="http://stackoverflow.com/q/4054937" rel="nofollow" target="_blank">http://stackoverflow.com/q/4054937</a>
REM --------------------------------------------------------------------------------------

REM ----
REM cmd - What does it mean by command cd /d %~dp0 in Windows - Stack Overflow
REM http://stackoverflow.com/questions/18309941/what-does-it-mean-by-command-cd-d-dp0-in-windows
REM ----
REM There are three parts:
REM 1..
REM cd -- This is change directory command.
REM 2..
REM /d -- This switch makes cd change both drive and directory at once. Without it you would have to do cd %~d0 & cd %~p0.
REM 3..
REM %~dp0 -- This can be dissected further into three parts:
REM %0 -- This represents zeroth parameter of your batch script. It expands into the name of the batch file itself.
REM %~0 -- The ~ there strips double quotes (") around the expanded argument.
REM %dp0 -- The d and p there are modifiers of the expansion. The d forces addition of a drive letter and the p adds full path.


REM ----
REM batch file - What does %~dp0 mean, and how does it work? - Stack Overflow
REM https://stackoverflow.com/questions/5034076/what-does-dp0-mean-and-how-does-it-work
REM ----

REM Calling

REM for /?
REM in the command-line gives help about this syntax (which can be used outside FOR, too, this is just the place where help can be found).

REM In addition, substitution of FOR variable references has been enhanced. You can now use the following optional syntax:

REM %~I         - expands %I removing any surrounding quotes (")
REM %~fI        - expands %I to a fully qualified path name
REM %~dI        - expands %I to a drive letter only
REM %~pI        - expands %I to a path only
REM %~nI        - expands %I to a file name only
REM %~xI        - expands %I to a file extension only
REM %~sI        - expanded path contains short names only
REM %~aI        - expands %I to file attributes of file
REM %~tI        - expands %I to date/time of file
REM %~zI        - expands %I to size of file
REM %~$PATH:I   - searches the directories listed in the PATH
               REM environment variable and expands %I to the
               REM fully qualified name of the first one found.
               REM If the environment variable name is not
               REM defined or the file is not found by the
               REM search, then this modifier expands to the
               REM empty string
REM The modifiers can be combined to get compound results:

REM %~dpI       - expands %I to a drive letter and path only
REM %~nxI       - expands %I to a file name and extension only
REM %~fsI       - expands %I to a full path name with short names only
REM %~dp$PATH:I - searches the directories listed in the PATH
               REM environment variable for %I and expands to the
               REM drive letter and path of the first one found.
REM %~ftzaI     - expands %I to a DIR like output line
REM In the above examples %I and PATH can be replaced by other valid values. The %~ syntax is terminated by a valid FOR variable name. Picking upper case variable names like %I makes it more readable and avoids confusion with the modifiers, which are not case sensitive.




