@ECHO OFF
REM -------------------------------------------------------------------------------
REM -- C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\
REM -- \PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.BAT
REM -------------------------------------------------------------------------------
REM -- WORK TODAY
REM -- OPERATION WILL SCAN ALL NETWORK SHARE PATH
REM -- USER TO SUPPLY COMPUTER SYSTEM NAME TO ARRAY
REM -------------------------------------------------------------------------------
REM -- OPERATION USE CODE TO RUN WITH FILELOCATOR
REM -- LOOK DOCUMENT FOR FILELOCATOR THEY NOT ANY MORE SUPPLY 
REM -- FILELOCATOR NETWORK VERSION
REM -- C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe
REM -- ONLY ONE APP WHEN PRO - AR 
REM -- IT GOT GOOD SCANNER
REM -- AND THEN -- NETWORK WE WANT
REM -- HERE WILL SCRIPT PROVIDE IN TEXT FILE AND DISPLAY
REM -- ALL NETWORK LOCATION 
REM -- MAIN SOURCE INITIAL
REM -- HERE CREDIT
REM -------------------------------------------------------------------------------
REM -- View all shares on remote machine with PowerShell
REM -- https://www.itprotoday.com/powershell/view-all-shares-remote-machine-powershell
REM -------------------------------------------------------------------------------
REM -- AND 40 PAGE URL OPEN MINUS CLICK BACK TRAVELER -- FOR MY FIRST LEARN THIS STYLER
REM -- INCLUDE 
REM --  1.. ADD STRING
REM --  2.. ARRAY 
REM --  3.. SET A VARIABLE
REM --  4.. WRITE VARIABLE TO SCREEN
REM --  5.. CARRIAGE RETURN LINE FED IN A STRING
REM --  *.. SEARCH AR STRING
REM --  7.. WITH USER __ IF STATEMENT COMMAND CONDITION
REM --  8.. GET THE RUN SCRIPT PATH NAME AND ALSO REMOVE EXTENSION
REM --  8.. CHECK A VARIABLE STRING EQUAL NOTHING
REM --  9.. GET THE RUN SCRIPT PATH NAME AND ALSO REMOVE EXTENSION
REM -- 10.. AUTOHOTKEYS SCRIPT MAKE THE RUN OF POWER-SHELL EASY -- FROM NOTEPAD++
REM -- 11.. REVERSE THE ARRAY
REM -------------------------------------------------------------------------------
REM -- Thu 27-Feb-2020 10:06:55
REM -- Thu 27-Feb-2020 15:34:00 -- 5 HOUR 27 MINUTE
REM -- AND SECOND SESSION LATER SAME SHIFT
REM -- Thu 27-Feb-2020 20:18:14
REM -- Thu 27-Feb-2020 22:20:00 -- 02 HOUR 02 MINUTE
REM -- Thu 27-Feb-2020 22:20:00 -- 07 HOUR 29 MINUTE -- TOTAL ADDITION OF TWO TIME PART 
REM -- Thu 27-Feb-2020 22:20:00 -- 12 HOUR 13 MINUTE -- TOTAL TIME WITH AR ANOTHER THING
REM -------------------------------------------------------------------------------
REM -- HERE IS THE LINE TO USER FOR FILELOCATOR HAS AN EQUAL
REM =C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__REMNFS_EX__.TXT
REM -------------------------------------------------------------------------------
REM -- HERE IS THE SCRIPT I USER FOR FILELOCATOR -- FOR USE SCRIPT MODE
REM -- AND TO FIND ALL GOODSYNC STUFF LEFT OVER SAVE WHAT HAVE YOU
REM -- WHILE IT THERE ALSO SCAN THE TEMP FOLDER -- WHICH MINE ARE NAME AFTER COMPUTER NAME SO HELPER
REM C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 68-FILE LOCATOR -- SCRIPT - GOODSYNC _GSDATA_.VBS
REM -- IT IS ON MY GITHUB -- VBS FOLDER
REM -------------------------------------------------------------------------------


PowerShell.exe -Command "& {Start-Process PowerShell.exe -ArgumentList '-ExecutionPolicy Bypass -File ""%~dp0%~n0.PS1""'}"






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
REM powershell -noexit -executionpolicy bypass -File "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 03 - MICROSOFT STORE THIS APP CAN'T OPEN _ FIX.PS1"

REM powershell -ExecutionPolicy Unrestricted Add-AppxPackage -DisableDevelopmentMode -Register $Env:SystemRoot\WinStore\AppxManifest.xml

REM powershell -noexit -executionpolicy bypass -File "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 03 - MICROSOFT STORE THIS APP CAN'T OPEN _ FIX.PS1"

REM C:\Users\MATT 03\AppData\Local\Packages\Microsoft.WindowsStore_8wekyb3d8bbwe
REM -----------------------------------------------------------------


REM -----------------------------------------------------------------
REM -----------------------------------------------------------------
REM -----------------------------------------------------------------
REM -----------------------------------------------------------------
REM -- 40 PAGE URL OPEN MINUS CLICK BACK TRAVELER -- FOR MY FIRST LEARN THIS STYLER
REM -- 40 URL OF REFERENCE SURF DONE
REM -----------------------------------------------------------------

REM -----------------------------------------------------------------
REM ----
REM windows - List mapped network drives from the command line to text file - Super User
REM https://superuser.com/questions/135752/list-mapped-network-drives-from-the-command-line-to-text-file
REM --------
REM windows xp - How do I get list of all network drives mapped to my system? - Super User
REM https://superuser.com/questions/248783/how-do-i-get-list-of-all-network-drives-mapped-to-my-system?rq=1
REM --------
REM CMD DIRECTORY NAME WITHOUT EXTENSION - Google Search
REM https://www.google.co.uk/search?q=CMD+DIRECTORY+NAME+WITHOUT+EXTENSION
REM --------
REM PowerShell Get List Of Folders Shared - Super User
REM https://superuser.com/questions/769679/powershell-get-list-of-folders-shared
REM --------
REM powershell - Get a list of computers that are mapped to a server - Super User
REM https://superuser.com/questions/1203629/get-a-list-of-computers-that-are-mapped-to-a-server
REM --------
REM PoshTip #35 – How To List Your Windows Mapped Drives? – www.get-cmd.com
REM http://get-cmd.com/?p=3720
REM --------
REM Get-SmbShareAccess
REM https://docs.microsoft.com/en-us/powershell/module/smbshare/get-smbshareaccess?view=win10-ps
REM --------
REM View all shares on remote machine with PowerShell
REM https://www.itprotoday.com/powershell/view-all-shares-remote-machine-powershell
REM --------
REM How to Use a Batch File to Make PowerShell Scripts Easier to Run
REM https://www.howtogeek.com/204088/how-to-use-a-batch-file-to-make-powershell-scripts-easier-to-run/
REM --------
REM Page Not Found - Daniel Schroeder’s Programming Blog
REM https://blog.danskingdom.com/tag/batch-file/
REM --------
REM Write-Output
REM https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-output?view=powershell-7
REM --------
REM Write-Debug
REM https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-debug?view=powershell-7
REM --------
REM Write-Verbose
REM https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-verbose?view=powershell-7
REM --------
REM POWERSHELL COMMAND OUTPUT TO A VARIABLE - Google Search
REM https://www.google.co.uk/search?safe=active&sxsrf=ALeKk00OGJ0ERg-dYPYYZXIu_seng9ZlFA:1582806982256&q=POWERSHELL+COMMAND+OUTPUT+TO+A+VARIABLE&spell=1&sa=X&ved=2ahUKEwjXjrO_3_HnAhWDYcAKHabPCpsQBSgAegQINRAq&biw=1536&bih=638&dpr=1.25
REM --------
REM Automatically assign command output to a variable in Powershell - Stack Overflow
REM https://stackoverflow.com/questions/49179697/automatically-assign-command-output-to-a-variable-in-powershell
REM --------
REM How do I capture the output into a variable from an external process in PowerShell? - Stack Overflow
REM https://stackoverflow.com/questions/8097354/how-do-i-capture-the-output-into-a-variable-from-an-external-process-in-powershe/35980675
REM --------
REM Show-Command
REM https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/show-command?view=powershell-7
REM --------
REM How do I add a newline to command output in PowerShell? - Stack Overflow
REM https://stackoverflow.com/questions/1639291/how-do-i-add-a-newline-to-command-output-in-powershell
REM --------
REM Out-File: Using PowerShell to Redirect Output to File
REM https://adamtheautomator.com/out-file-powershell-redirect-file/
REM --------
REM Get name of the PowerShell script file inside the script - Thomas Maurer
REM https://www.thomasmaurer.ch/2015/03/get-name-of-the-powershell-script-file-inside-the-script/
REM --------
REM PowerTip: Use PowerShell to Get Computer Name | Scripting Blog
REM https://devblogs.microsoft.com/scripting/powertip-use-powershell-to-get-computer-name/
REM --------
REM PowerTip: New Lines with PowerShell | Scripting Blog
REM https://devblogs.microsoft.com/scripting/powertip-new-lines-with-powershell/
REM --------
REM How do I concatenate strings and variables in PowerShell? - Stack Overflow
REM https://stackoverflow.com/questions/15113413/how-do-i-concatenate-strings-and-variables-in-powershell
REM --------
REM PowerShell - How to format Write-Host with multiple colors - Evotec
REM https://evotec.xyz/powershell-how-to-format-powershell-write-host-with-multiple-colors/
REM --------
REM Writing new lines to a text file in PowerShell - Stack Overflow
REM https://stackoverflow.com/questions/17434151/writing-new-lines-to-a-text-file-in-powershell
REM --------
REM How do I get a file path without extension in PowerShell? - Stack Overflow
REM https://stackoverflow.com/questions/14770587/how-do-i-get-a-file-path-without-extension-in-powershell
REM --------
REM Out-File
REM https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/out-file?view=powershell-7
REM --------
REM Out-Null
REM https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/out-null?view=powershell-7
REM --------
REM PowerShell Variables and Arrays
REM https://blog.netwrix.com/2018/10/04/powershell-variables-and-arrays/
REM --------
REM How do I concatenate strings and variables in PowerShell? - Stack Overflow
REM https://stackoverflow.com/questions/15113413/how-do-i-concatenate-strings-and-variables-in-powershell
REM --------
REM Parsing Text with PowerShell (1/3) | PowerShell
REM https://devblogs.microsoft.com/powershell/parsing-text-with-powershell-1-3/
REM --------
REM Terminating a script in PowerShell - Stack Overflow
REM https://stackoverflow.com/questions/2022326/terminating-a-script-in-powershell
REM --------
REM powershell - Iterating through a variable line by line - Stack Overflow
REM https://stackoverflow.com/questions/31563848/iterating-through-a-variable-line-by-line
REM --------
REM ConvertFrom-String
REM https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/convertfrom-string?view=powershell-5.1
REM --------
REM PowerShell -contains -like
REM https://www.itprotoday.com/powershell/powershell-contains
REM --------
REM Is there a PowerShell "string does not contain" cmdlet or syntax? - Stack Overflow
REM https://stackoverflow.com/questions/74957/is-there-a-powershell-string-does-not-contain-cmdlet-or-syntax
REM --------
REM PowerShell -contains -like
REM https://www.itprotoday.com/powershell/powershell-contains
REM --------
REM In Powershell, how to set variable values within a function and have that value available in the parent scope? - Stack Overflow
REM https://stackoverflow.com/questions/27847809/in-powershell-how-to-set-variable-values-within-a-function-and-have-that-value
REM --------
REM Powershell: check variable for null - Thomas Maurer
REM https://www.thomasmaurer.ch/2010/07/powershell-check-variable-for-null/
REM --------
REM Trim() TrimStart() TrimEnd() - PowerShell - SS64.com
REM https://ss64.com/ps/trim.html
REM ----
REM -----------------------------------------------------------------
