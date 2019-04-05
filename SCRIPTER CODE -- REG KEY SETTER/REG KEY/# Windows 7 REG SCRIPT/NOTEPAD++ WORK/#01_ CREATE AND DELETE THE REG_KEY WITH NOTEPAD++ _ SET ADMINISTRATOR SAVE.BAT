@echo off
@CD\

REM ------------------------------------------------------------------------------
REM E:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\NOTEPAD++ WORK\#01_ CREATE AND DELETE THE REG_KEY WITH NOTEPAD++ _ SET ADMINISTRATOR SAVE.BAT
REM ------------------------------------------------------------------------------
REM #01_ CREATE AND DELETE THE REG_KEY WITH NOTEPAD++ _ SET ADMINISTRATOR SAVE.BAT
REM ------------------------------------------------------------------------------
REM [ Monday 18:25:30 Pm_18 September 2017 ] The UK
REM [ Monday 17:25:30 Pm_18 September 2017 ] BST BDT DST DLS 1 Hour Ahead UTC+1 CET
REM ------------------------------------------------------------------------------
ECHO MATTHEW LANCASTER
ECHO MATT.LAN@BTINTERNET.COM

@echo off
REM :: Check for ADMIN Privileges

>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

if '%errorlevel%' GTR '2' (
	REM Get ADMIN Privileges
	echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
	echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
	"%temp%\getadmin.vbs"
	del "%temp%\getadmin.vbs"
	GOTO VERY_ENDER_BLOCK	
	exit /B
)

@ECHO Got ADMIN Privileges


REM --------------------------------------------------
REM HAS TO BE RUN AFTER EVERY NOTEPAD++ UPDATE
REM EDITED ON Fri 22 September 2017 07:30:38----------
REM --------------------------------------------------


REM -------------------------------------------------------------------------
REM __ This work Correctly Has Immediate Effect Learn
REM __ After the reg entry from this folder
REM __ One Key Has to Be Delete
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM __ This Code Does the Work
REM -------------------------------------------------------------------------
REM __ Is Another Test There Is Code to Save as Administrator in Notepad++
REM __ Yes and Also There are Two to Set the Txt as Notepad++ and Administrator Save
REM __ WITH TWO PLUG-IN TO SAVE THE WORK HERE __ TESTER
REM -------------------------------------------------------------------------
REM -------------------------------------------------------------------------
REM ONE KEY IS DELETER WITH 3 ITEM BE DELETER WITHT HE COPDE ITEM LIKE 
REM HERE ____ {B298D29A-A6ED-11DE-BA8C-A68E55D89593}
REM LIKE HERE
REM -------------------------------------------------------------------------
REM Windows Registry Editor Version 5.00

REM [HKEY_CLASSES_ROOT\CLSID\{B298D29A-A6ED-11DE-BA8C-A68E55D89593}]
REM @="ANotepad++64"

REM [HKEY_CLASSES_ROOT\CLSID\{B298D29A-A6ED-11DE-BA8C-A68E55D89593}\InprocServer32]
REM @="C:\\Program Files (x86)\\Notepad++\\NppShell_06.dll"
REM "ThreadingModel"="Apartment"

REM [HKEY_CLASSES_ROOT\CLSID\{B298D29A-A6ED-11DE-BA8C-A68E55D89593}\Settings]
REM "Title"="Edit with &Notepad++"
REM "Path"="C:\\Program Files (x86)\\Notepad++\\notepad++.exe"
REM "Custom"=""
REM "ShowIcon"=dword:00000001
REM "Dynamic"=dword:00000001
REM "Maxtext"=dword:00000019
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM WRITING YOUR OWN QUERY SEARCH AND AND DELETER ARE A BIT DEPENDANT ON THE JOB 
REM EXAMPLE LOOK FOR A CERTAIN AREA SAVE TIME SUB-BRANCHER
REM AND DEPEND IF LOOK FOR A STRING OR A VALUE
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
rem Edit with &Notepad++
rem ANotepad++64
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM A WEIRD MESSAGE AT END ABOUT SEARCH FOR NOTEPAD 
REM BUT OUT USEFULNESS __ UNABLE TO CODE TO NOT DISPLAY __ NOT ANY HARM
REM --
REM MESSENGER WEIRD ERROR SOMETIMES HAPPEN BUT NOT AFTER 1ST TIME TRIAL RUN
REM -------------------------------------------------------------------------
REM IT IS CORRECT ALL THREE 3 REG ENTRY HAVE TO BE DELETER THE ENTIRE TREE PATH 
REM FROM WHERE FOUND
REM AFTER TESTER FOUND MENU CONTEXT ITEM CAME BACK AGAIN AND HAD TO DELETE 
REM ENTIRE SMALL TREE IN QUESTION
REM SEARCH FIND DELETE DOES IT BY SEARCH AT MAXIMUM
REM THE WORD "ANotepad++64" __ MIGHT FIND THE TREE HAS ERROR DELETE 
REM FIRST PART BEFORE REST AND REST IS GONE MISSING AWOL
REM -------------------------------------------------------------------------
REM BEEN RIGOROUS TESTED BY PUT BACK AND DO AGAIN
REM INCLUDE SEARCH FOR INFINITE MAXIMUM
REM THE SEARCH FOR A DELETE IS ONLY DONE ON THE SUB-BRANCH WHERE EVENT IS TO HAPPEN THERE IS NOT 
REM A BIG DEAL ABOUT NOTEPAD++ THERE
REM -------------------------------------------------------------------------
REM FIND THIS PROGRAM __ BIGGER HELPER IN BATCH FILE FROM REG FILE CREATION
REM BLOGGER COMMENT SHOW GO BACK ONLY CREATED APR 2014 __ VERY HELPFUL INDEED

REM ----
REM ----
REM REG FILE TO BATCH FILE CONVERTER - Google Search
REM https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB744GB744&q=REG+FILE+TO+BATCH+FILE+CONVERTER&spell=1&sa=X&ved=0ahUKEwijx7WbnK_WAhWFUlAKHdO6AmgQvwUIJSgA&biw=1229&bih=523
REM ----
REM ----
REM How to convert a registry file to a batch file | Symantec Connect Community
REM https://www.symantec.com/connect/blogs/how-convert-registry-file-batch-file
REM ----
REM ----
REM Reg Converter
REM http://www.sordum.org/8478/reg-converter-v1-1/
REM ----



REM __ TO SET MY PERSONAL PROGRAM COMMAND IN THE CONTEXT MENU FOR 
REM _________________________________________________________________
REM __ OPEN A FILE IN EXPLORER
REM __ RUN AS ADMINISTRATOR TO WORK WITHOUT RESTRICTION ERROR WORKER
REM __ HAPPY REGISTRY WORKING
REM _________________________________________________________________

Echo on
@REM __ #01_ CREATE AND DELETE THE REG_KEY WITH NOTEPAD++ _ SET ADMINISTRATOR SAVE.BAT
Reg.exe add "HKCR\*\shell\OpenWithNotepad" /ve /t REG_SZ /d "Edit with &Notepad++" /f
Reg.exe add "HKCR\*\shell\OpenWithNotepad" /v "icon" /t REG_SZ /d "C:\Program Files (x86)\Notepad++\Notepad++.exe" /f
Reg.exe add "HKCR\*\shell\OpenWithNotepad\Command" /ve /t REG_SZ /d "\"C:\Program Files (x86)\Notepad++\Notepad++.exe\" \"%%1\"" /f

REM GOTO ENDER

Echo off
for /f "delims=" %%a in ('reg query "HKCR\CLSID" /s /f "ANotepad++64" /t REG_SZ') do echo %%a|findstr /v "REG_SZ" | findstr /v /c:"End of search:" >nul && reg.exe delete "%%a" /f

REM for /f "delims=" %%a in ('reg query "HKCR\CLSID" /s /f "ANotepad++64" /t REG_SZ') do echo %%a|findstr /v "REG_SZ" | findstr /v /c:"End of search:" >nul && reg.exe delete "%%a" /f

REM for /f "delims=" %%a in ('reg query "HKCR\CLSID" /s /f "Edit with &Notepad++" /t REG_SZ') do echo %%a|findstr /v "REG_SZ" | findstr /v /c:"End of search:" >nul && reg.exe delete "%%a" /f

REM -- E:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\NOTEPAD++ WORK\replace-notepad-with-npp.reg
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\notepad.exe" /v "Debugger" /t REG_SZ /d "\"C:\Program Files\Notepad++\Notepad++ParamProxy.exe\"" /f



GOTO ENDER
:ENDER

PAUSE

REM for /f "delims=" %%a in ('reg query "HKCR\CLSID" /s /f "Notepad+" /t REG_SZ') do echo %%a|findstr /v "REG_SZ" | findstr /v /c:"End of search:" >nul && ECHO test reg.exe delete "%%a" /f

REM for /f "delims=" %%a in ('reg query "%Key%" /s /f %STRING% /t REG_SZ') do (echo %%a | findstr /v "REG_SZ" | findstr /v /c:"End of search:" >nul && ECHO reg.exe delete "%%a" /f)


REM -------------------------------
REM REFERENCE SET _ END OF DOCUMENT 
REM -------------------------------
REM MY SEARCH
REM ----
REM batch find a registry entry and delete - Google Search
REM https://www.google.co.uk/search?q=batch+find+a+registry+entry+and+delete&rlz=1C1CHBD_en-GBGB744GB744&oq=batch+find+a+registry+entry+and+delete&aqs=chrome..69i57.9739j0j7&sourceid=chrome&ie=UTF-8
REM --------
REM reg query DELETE - Google Search
REM https://www.google.co.uk/search?q=reg+query+DELETE&rlz=1C1CHBD_en-GBGB744GB744&oq=reg+query+DELETE&aqs=chrome..69i57j69i60j0l4.5631j0j7&sourceid=chrome&ie=UTF-8
REM --------
REM batch pipe the error - Google Search
REM https://www.google.co.uk/search?q=batch+pipe+the+error&rlz=1C1CHBD_en-GBGB744GB744&oq=batch+pipe+the+error&aqs=chrome..69i57j0l4.7793j0j7&sourceid=chrome&ie=UTF-8
REM ----

REM RESULT SEARCH _ LAST ONE IN THE QUE
REM ----
REM [SOLVED] Query & Delete Registry Keys and/or Values w/ Batch Script? - Windows Forum - Spiceworks
REM https://community.spiceworks.com/topic/491050-query-delete-registry-keys-and-or-values-w-batch-script
REM --------
REM Windows Batch: reg query key value to a variable but do not display error if key doesn't exist - Stack Overflow
REM https://stackoverflow.com/questions/5369528/windows-batch-reg-query-key-value-to-a-variable-but-do-not-display-error-if-key
REM --------
REM batch file - Search for data value with Reg.exe and delete parent key - Stack Overflow
REM https://stackoverflow.com/questions/37557055/search-for-data-value-with-reg-exe-and-delete-parent-key
REM --------
REM batch file - Search and Delete registry keys based on ProductName value - Stack Overflow
REM https://stackoverflow.com/questions/17750053/search-and-delete-registry-keys-based-on-productname-value?rq=1
REM --------
REM registry - REG DELETE a value which contains quotes in it, inside a batch file? - Stack Overflow
REM https://stackoverflow.com/questions/13041700/reg-delete-a-value-which-contains-quotes-in-it-inside-a-batch-file?rq=1
REM --------
REM batch file - Search and Delete registry keys based on ProductName value - Stack Overflow
REM https://stackoverflow.com/questions/17750053/search-and-delete-registry-keys-based-on-productname-value?rq=1
REM --------
REM Redirecting Error Messages from Command Prompt: STDERR/STDOUT
REM https://support.microsoft.com/en-us/help/110930/redirecting-error-messages-from-command-prompt-stderr-stdout
REM ---- 