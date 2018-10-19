@ECHO OFF

@REM ----------------------------------------------------------------
@ECHO ____ BAT 23- ASSOCIATION EXTENSIONS.BAT ____ 
@REM ----------------------------------------------------------------
@CD\
TITLE %~f0

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



@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

@CD \
REM ECHO ON

:CUSTOM .AHK
:NOTEPAD++ FOR AUTOHOTKEYS
REM ORIGINAL
SET NOTE_ING=notepad.exe
REM MUST BE SMALL NOTEPAD.EXE FOR AHK
IF EXIST "C:\Program Files\Notepad++\notepad++.exe" SET NOTE_ING=\"C:\Program Files\Notepad++\notepad++.exe\"
IF EXIST "C:\Program Files (x86)\Notepad++\notepad++.exe" SET NOTE_ING=\"C:\Program Files (x86)\Notepad++\notepad++.exe\"

Reg.exe add "HKCR\AutoHotkeyScript\Shell\Edit" /ve /t REG_SZ /d "Edit Script" /f
Reg.exe add "HKCR\AutoHotkeyScript\Shell\Edit\Command" /ve /t REG_SZ /d "%NOTE_ING% \"%%1\"" /f


SET NOTE_ING=%%SystemRoot%%\system32\notepad.exe
IF EXIST "C:\Program Files\Notepad++\notepad++.exe" SET NOTE_ING=\"C:\Program Files\Notepad++\notepad++.exe\"
IF EXIST "C:\Program Files (x86)\Notepad++\notepad++.exe" SET NOTE_ING=\"C:\Program Files (x86)\Notepad++\notepad++.exe\"

:CUSTOM .CMD
Reg.exe add "HKCR\cmdfile\shell\edit\command" /ve /t REG_EXPAND_SZ /d "%NOTE_ING% \"%%1\"" /f
Reg.exe add "HKCR\cmdfile\shell\print\command" /ve /t REG_EXPAND_SZ /d "%NOTE_ING% /p \"%%1\"" /f

:CUSTOM .REG EDITOR
Reg.exe add "HKCR\regfile\shell\edit\command" /ve /t REG_EXPAND_SZ /d "%NOTE_ING% \"%%1\"" /f
Reg.exe add "HKCR\regfile\shell\print\command" /ve /t REG_EXPAND_SZ /d "%NOTE_ING% /p \"%%1\"" /f

:CUSTOM .VBS
Reg.exe add "HKCR\VBSFile\Shell\Edit\Command" /ve /t REG_EXPAND_SZ /d "%NOTE_ING% \"%%1\"" /f
Reg.exe add "HKCR\VBSFile\Shell\Print\Command" /ve /t REG_EXPAND_SZ /d "%NOTE_ING% /p \"%%1\"" /f



:CUSTOM .BAT
:NOTEPAD++ FOR AUTOHOTKEYS
SET NOTE_ING=%%SystemRoot%%\system32\notepad.exe
IF EXIST "C:\Program Files\Notepad++\notepad++.exe" SET NOTE_ING=\"C:\Program Files\Notepad++\notepad++.exe\"
IF EXIST "C:\Program Files (x86)\Notepad++\notepad++.exe" SET NOTE_ING=\"C:\Program Files (x86)\Notepad++\notepad++.exe\"

Reg.exe add "HKCR\.bat" /ve /t REG_SZ /d "batfile" /f
Reg.exe add "HKCR\.bat\PersistentHandler" /ve /t REG_SZ /d "{5e941d80-bf96-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\batfile" /ve /t REG_SZ /d "Windows Batch File" /f
Reg.exe add "HKCR\batfile" /v "EditFlags" /t REG_BINARY /d "30040000" /f
Reg.exe add "HKCR\batfile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\acppage.dll,-6002" /f
Reg.exe add "HKCR\batfile\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\imageres.dll,-68" /f
Reg.exe add "HKCR\batfile\shell\edit\command" /ve /t REG_EXPAND_SZ /d "%NOTE_ING% \"%%1\"" /f
Reg.exe add "HKCR\batfile\shell\open" /v "EditFlags" /t REG_BINARY /d "00000000" /f
Reg.exe add "HKCR\batfile\shell\open\command" /ve /t REG_SZ /d "\"%%1\" %%*" /f
Reg.exe add "HKCR\batfile\shell\print\command" /ve /t REG_EXPAND_SZ /d "%NOTE_ING% /p \"%%1\"" /f
Reg.exe add "HKCR\batfile\shell\runas" /v "HasLUAShield" /t REG_SZ /d "" /f
Reg.exe add "HKCR\batfile\shell\runas\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\cmd.exe /C \"%%1\" %%*" /f
Reg.exe add "HKCR\batfile\shell\runasuser" /ve /t REG_SZ /d "@shell32.dll,-50944" /f
Reg.exe add "HKCR\batfile\shell\runasuser" /v "Extended" /t REG_SZ /d "" /f
Reg.exe add "HKCR\batfile\shell\runasuser" /v "SuppressionPolicyEx" /t REG_SZ /d "{F211AA05-D4DF-4370-A2A0-9F19C09756A7}" /f
Reg.exe add "HKCR\batfile\shell\runasuser\command" /v "DelegateExecute" /t REG_SZ /d "{ea72d00e-4960-42fa-ba92-7792a7944c1d}" /f
Reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.bat" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.bat\OpenWithProgids" /v "batfile" /t REG_NONE /d "" /f

:CUSTOM TXT
:NOTEPAD++
SET NOTE_ING=%%SystemRoot%%\system32\notepad.exe
IF EXIST "C:\Program Files\Notepad++\notepad++.exe" SET NOTE_ING=\"C:\Program Files\Notepad++\notepad++.exe\"
IF EXIST "C:\Program Files (x86)\Notepad++\notepad++.exe" SET NOTE_ING=\"C:\Program Files (x86)\Notepad++\notepad++.exe\"

Reg.exe add "HKCR\.txt" /ve /t REG_SZ /d "txtfile" /f
Reg.exe add "HKCR\.txt" /v "Content Type" /t REG_SZ /d "text/plain" /f
Reg.exe add "HKCR\.txt" /v "PerceivedType" /t REG_SZ /d "text" /f
Reg.exe add "HKCR\.txt\PersistentHandler" /ve /t REG_SZ /d "{5e941d80-bf96-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\.txt\ShellNew" /v "ItemName" /t REG_EXPAND_SZ /d "@%NOTE_ING%,-470" /f
Reg.exe add "HKCR\.txt\ShellNew" /v "NullFile" /t REG_SZ /d "" /f
Reg.exe add "HKCR\txtfile" /ve /t REG_SZ /d "Text Document" /f
Reg.exe add "HKCR\txtfile" /v "EditFlags" /t REG_DWORD /d "2162688" /f

Reg.exe add "HKCR\txtfile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%NOTE_ING%,-469" /f
Reg.exe add "HKCR\txtfile\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%NOTE_ING%,-102" /f
REM __ DEFAULT ICON OF NOTEPAD++ = "C:\Program Files (x86)\Notepad++\Notepad++.exe",0
REM __ HKEY_CLASSES_ROOT\Notepad++_file\DefaultIcon

Reg.exe add "HKCR\txtfile\shell\open\command" /ve /t REG_EXPAND_SZ /d "%NOTE_ING% \"%%1\"" /f
Reg.exe add "HKCR\txtfile\shell\print\command" /ve /t REG_EXPAND_SZ /d "%NOTE_ING% /p \"%%1\"" /f
Reg.exe add "HKCR\txtfile\shell\printto\command" /ve /t REG_EXPAND_SZ /d "%NOTE_ING% /pt \"%%1\" \"%%2\" \"%%3\" \"%%4\"" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.txt\OpenWithProgids" /v "txtfile" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.txt\UserChoice" /v "Hash" /t REG_SZ /d "Ndb2OUV7duM=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.txt\UserChoice" /v "ProgId" /t REG_SZ /d "txtfile" /f

:VERY_ENDER_BLOCK


IF "%1"=="" PAUSE
IF "%1"=="/ADMIN" PAUSE

:ENDE
