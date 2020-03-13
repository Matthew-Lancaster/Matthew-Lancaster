@ECHO OFF
REM =================================================================
REM __ C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.BAT
REM __ C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1
REM __ 
REM __ PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.BAT
REM __ PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1
REM 
REM __ BY Matthew __ Matt.Lan@Btinternet.com
REM =================================================================
REM
REM    FIRST RUN AT HER
REM -- Thu 27-Feb-2020 10:06:55
REM -- Thu 27-Feb-2020 15:34:00 -- 5 HOUR 27 MINUTE
REM
REM -- Thu 27-Feb-2020 20:18:14
REM -- Thu 27-Feb-2020 22:20:00 -- 02 HOUR 02 MINUTE
REM -- Thu 27-Feb-2020 22:20:00 -- 07 HOUR 29 MINUTE -- TOTAL ADDITION OF TWO TIME PART 
REM -- Thu 27-Feb-2020 22:20:00 -- 12 HOUR 13 MINUTE -- TOTAL TIME WITH AR ANOTHER THING
REM
REM    SECOND RUN AT HER -- TWO DAY LATER -- SEEM LIKE ONE ANOTHER ALL DAY THING
REM -- Sat 29-Feb-2020 16:15:57
REM -- Sat 29-Feb-2020 19:04:00 -- 2 HOUR 48 MINUTE
REM =================================================================

REM -----------------------------------------------------------------
REM -----------------------------------------------------------------
REM -- MORE INFO AT THE CORRESPOND POWERSHELL SCRIPTOR
REM -----------------------------------------------------------------
REM -----------------------------------------------------------------

CD %~dp0

REM -----------------------------------------------------------------
REM -- HAVE CODE HASSLE -- POWERSHELL WHEN RUN NOT REPORT LOCATION 
REM -- OF SCRIPT RUN FOLDER
REM -- SET FROM BAT FOR NOW
REM -----------------------------------------------------------------
REM -- OKAY SORTER NOW POWERSHELL FIND OWN PATH FROM IT WORKER
REM -----------------------------------------------------------------
 
SET FILE_NAME="%~dp0%~n0.PS1"
START "" /MIN "powershell"  -noexit -executionpolicy bypass -File %FILE_NAME%

REM -----------------------------------------------------------------
REM -- Sat 29-Feb-2020 15:48:00
REM -- Sat 29-Feb-2020 16:45:33 -- 58 MINUTE
REM -----------------------------------------------------------------


REM -------------------------------------------------------------------------------
REM MAKE AN AUTOHOTKEYS SCRIPT AND PUT HERE
REM AND THEN WHEN SAVE PRESS F5 LIKE TO RUNNER
REM BUT INSTEAD WILL SAVE AND PRESS OCROSPND SAME FILE BAT WITH F5 AND RUN THAT BIT
REM -------------------------------------------------------------------------------
REM #IfWinActive .PS1 - Notepad++ ahk_class Notepad++
REM {
REM F5:: 
	REM SENDINPUT ^S
	REM SOUNDBEEP 1000,100
	REM RETURN
REM }
REM #ifwinactive
REM -------------------------------------------------------------------------------

REM -----------------------------------------------------------------
REM ---- FILENAME WITHOUT EXTENSION
REM ECHO %~n0
REM ---- FILENAME WITH   EXTENSION
REM ECHO %~n0%~x0
REM ---- FULL SCRIPT PATH
REM ECHO %~f0
REM ---- FULL SCRIPT PATH AND FILENAME WITHOUT EXTENSION
REM ECHO %~dp0%~n0
REM -----------------------------------------------------------------

REM -----------------------------------------------------------------
REM %~dp0 -- This can be dissected further into three parts:
REM %0    -- This represents zeroth parameter of your batch script. It expands into the name of the batch file itself.
REM %~0   -- The ~ there strips double quotes (") around the expanded argument.
REM %dp0  -- The d and p there are modifiers of the expansion. The d forces addition of a drive letter and the p adds full path.
REM %TEMP%
REM %~f0
REM %~dpnx0
REM %~dp0
REM -----------------------------------------------------------------

REM -----------------------------------------------------------------
REM -- PowerShell.exe -Command "& {Start-Process PowerShell.exe -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0%~n0.PS1""'}"
REM -----------------------------------------------------------------

REM -----------------------------------------------------------------
REM powershell -noexit -executionpolicy bypass -File "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 03 - MICROSOFT STORE THIS APP CAN'T OPEN _ FIX.PS1"

REM powershell -ExecutionPolicy Unrestricted Add-AppxPackage -DisableDevelopmentMode -Register $Env:SystemRoot\WinStore\AppxManifest.xml

REM powershell -noexit -executionpolicy bypass -File "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 03 - MICROSOFT STORE THIS APP CAN'T OPEN _ FIX.PS1"

REM C:\Users\MATT 03\AppData\Local\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe
REM -----------------------------------------------------------------


