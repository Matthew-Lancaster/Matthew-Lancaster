@ECHO OFF
@REM ----------------------------------------------------------------
@ECHO ____ BAT 22-FIX ALL EXTENSIONS.BAT ____ 
TITLE %~f0
@CD\
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

@REM ----
@REM How to change Registry Permissions with RegIni.exe (VBScript) – Decrypt my World
@REM https://blogs.msdn.microsoft.com/alejacma/2008/03/11/how-to-change-registry-permissions-with-regini-exe-vbscript/
@REM ----

@REM __ SET ADMIN A SAFE DOUBLE CHECKER WAY NOT A LOOPY ONE IF CAN'T GET IT TO GO ADMIN

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


@CD \


REM ORIGINAL 
REM GOTO JUMP_ER

ECHO _.BAT LINK

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

:JUMP_eR

ECHO _.EXE

Reg.exe add "HKCR\.exe" /ve /t REG_SZ /d "exefile" /f
Reg.exe add "HKCR\.exe" /v "Content Type" /t REG_SZ /d "application/x-msdownload" /f
Reg.exe add "HKCR\.exe\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\exefile" /ve /t REG_SZ /d "Application" /f
Reg.exe add "HKCR\exefile" /v "EditFlags" /t REG_BINARY /d "38070000" /f
Reg.exe add "HKCR\exefile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\shell32.dll,-10156" /f
Reg.exe add "HKCR\exefile\DefaultIcon" /ve /t REG_SZ /d "%%1" /f
Reg.exe add "HKCR\exefile\shell\open" /v "EditFlags" /t REG_BINARY /d "00000000" /f
Reg.exe add "HKCR\exefile\shell\open\command" /ve /t REG_SZ /d "\"%%1\" %%*" /f
Reg.exe add "HKCR\exefile\shell\open\command" /v "IsolatedCommand" /t REG_SZ /d "\"%%1\" %%*" /f
Reg.exe add "HKCR\exefile\shell\runas" /v "HasLUAShield" /t REG_SZ /d "" /f
Reg.exe add "HKCR\exefile\shell\runas\command" /ve /t REG_SZ /d "\"%%1\" %%*" /f
Reg.exe add "HKCR\exefile\shell\runas\command" /v "IsolatedCommand" /t REG_SZ /d "\"%%1\" %%*" /f
Reg.exe add "HKCR\exefile\shell\runasuser" /ve /t REG_SZ /d "@shell32.dll,-50944" /f
Reg.exe add "HKCR\exefile\shell\runasuser" /v "Extended" /t REG_SZ /d "" /f
Reg.exe add "HKCR\exefile\shell\runasuser" /v "SuppressionPolicyEx" /t REG_SZ /d "{F211AA05-D4DF-4370-A2A0-9F19C09756A7}" /f
Reg.exe add "HKCR\exefile\shell\runasuser\command" /v "DelegateExecute" /t REG_SZ /d "{ea72d00e-4960-42fa-ba92-7792a7944c1d}" /f
Reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.exe" /f


ECHO _.LNK LINK

Reg.exe add "HKCR\.lnk" /ve /t REG_SZ /d "lnkfile" /f
Reg.exe add "HKCR\.lnk\ShellNew" /v "Handler" /t REG_SZ /d "{ceefea1b-3e29-4ef1-b34c-fec79c4f70af}" /f
Reg.exe add "HKCR\.lnk\ShellNew" /v "IconPath" /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\shell32.dll,-16769" /f
Reg.exe add "HKCR\.lnk\ShellNew" /v "ItemName" /t REG_SZ /d "@shell32.dll,-30397" /f
Reg.exe add "HKCR\.lnk\ShellNew" /v "MenuText" /t REG_SZ /d "@shell32.dll,-30318" /f
Reg.exe add "HKCR\.lnk\ShellNew" /v "NullFile" /t REG_SZ /d "" /f
Reg.exe add "HKCR\.lnk\ShellNew\Config" /v "DontRename" /t REG_SZ /d "" /f
Reg.exe add "HKCR\lnkfile" /ve /t REG_SZ /d "Shortcut" /f
Reg.exe add "HKCR\lnkfile" /v "EditFlags" /t REG_DWORD /d "1" /f
Reg.exe add "HKCR\lnkfile" /v "FriendlyTypeName" /t REG_SZ /d "@shell32.dll,-4153" /f
Reg.exe add "HKCR\lnkfile" /v "IsShortcut" /t REG_SZ /d "" /f
Reg.exe add "HKCR\lnkfile" /v "NeverShowExt" /t REG_SZ /d "" /f
Reg.exe add "HKCR\lnkfile\CLSID" /ve /t REG_SZ /d "{00021401-0000-0000-C000-000000000046}" /f
Reg.exe add "HKCR\lnkfile\shellex\ContextMenuHandlers\OpenContainingFolderMenu" /ve /t REG_SZ /d "{37ea3a21-7493-4208-a011-7f9ea79ce9f5}" /f
Reg.exe add "HKCR\lnkfile\tabsets" /v "selection" /t REG_DWORD /d "1797" /f
Reg.exe delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked" /v "{00021401-0000-0000-C000-000000000046}" /f
Reg.exe delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked" /v "{85cfccaf-2d14-42b6-80b6-f40f65d016e7}" /f
Reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked" /f
Reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.lnk" /f


ECHO _.COM SAME AS CMD COMMAND.COM

Reg.exe add "HKCR\.com" /ve /t REG_SZ /d "comfile" /f
Reg.exe add "HKCR\.com\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\comfile" /ve /t REG_SZ /d "MS-DOS Application" /f
Reg.exe add "HKCR\comfile" /v "EditFlags" /t REG_BINARY /d "30000000" /f
Reg.exe add "HKCR\comfile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\shell32.dll,-8464" /f
Reg.exe add "HKCR\comfile\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\shell32.dll,2" /f
Reg.exe add "HKCR\comfile\shell\open" /v "EditFlags" /t REG_DWORD /d "0" /f
Reg.exe add "HKCR\comfile\shell\open\command" /ve /t REG_SZ /d "\"%%1\" %%*" /f
Reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.com" /f


ECHO _.CMD COMMAND.COM

Reg.exe add "HKCR\.cmd" /ve /t REG_SZ /d "cmdfile" /f
Reg.exe add "HKCR\.cmd\PersistentHandler" /ve /t REG_SZ /d "{5e941d80-bf96-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\cmdfile" /ve /t REG_SZ /d "Windows Command Script" /f
Reg.exe add "HKCR\cmdfile" /v "EditFlags" /t REG_BINARY /d "30040000" /f
Reg.exe add "HKCR\cmdfile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\acppage.dll,-6003" /f
Reg.exe add "HKCR\cmdfile\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\imageres.dll,-68" /f
Reg.exe add "HKCR\cmdfile\shell\edit\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE %%1" /f
Reg.exe add "HKCR\cmdfile\shell\open" /v "EditFlags" /t REG_BINARY /d "00000000" /f
Reg.exe add "HKCR\cmdfile\shell\open\command" /ve /t REG_SZ /d "\"%%1\" %%*" /f
Reg.exe add "HKCR\cmdfile\shell\print\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE /p %%1" /f
Reg.exe add "HKCR\cmdfile\shell\runas" /v "HasLUAShield" /t REG_SZ /d "" /f
Reg.exe add "HKCR\cmdfile\shell\runas\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\cmd.exe /C \"%%1\" %%*" /f
Reg.exe add "HKCR\cmdfile\shell\runasuser" /ve /t REG_SZ /d "@shell32.dll,-50944" /f
Reg.exe add "HKCR\cmdfile\shell\runasuser" /v "Extended" /t REG_SZ /d "" /f
Reg.exe add "HKCR\cmdfile\shell\runasuser" /v "SuppressionPolicyEx" /t REG_SZ /d "{F211AA05-D4DF-4370-A2A0-9F19C09756A7}" /f
Reg.exe add "HKCR\cmdfile\shell\runasuser\command" /v "DelegateExecute" /t REG_SZ /d "{ea72d00e-4960-42fa-ba92-7792a7944c1d}" /f
Reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.cmd" /f


ECHO REG FILES REGIRSTRY

Reg.exe add "HKCR\.reg" /ve /t REG_SZ /d "regfile" /f
Reg.exe add "HKCR\.reg\PersistentHandler" /ve /t REG_SZ /d "{5e941d80-bf96-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\regfile" /ve /t REG_SZ /d "Registration Entries" /f
Reg.exe add "HKCR\regfile" /v "EditFlags" /t REG_DWORD /d "1048576" /f
Reg.exe add "HKCR\regfile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\regedit.exe,-309" /f
Reg.exe add "HKCR\regfile\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\regedit.exe,1" /f
Reg.exe add "HKCR\regfile\shell\edit\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\notepad.exe \"%%1\"" /f
Reg.exe add "HKCR\regfile\shell\open" /ve /t REG_SZ /d "Mer&ge" /f
Reg.exe add "HKCR\regfile\shell\open" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\regedit.exe,-310" /f
Reg.exe add "HKCR\regfile\shell\open\command" /ve /t REG_SZ /d "regedit.exe \"%%1\"" /f
Reg.exe add "HKCR\regfile\shell\print\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\notepad.exe /p \"%%1\"" /f
Reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.reg" /f

ECHO VBSCRIPT VBS

Reg.exe add "HKCR\.vbs" /ve /t REG_SZ /d "VBSFile" /f
Reg.exe add "HKCR\.vbs\PersistentHandler" /ve /t REG_SZ /d "{5e941d80-bf96-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\VBSFile" /ve /t REG_SZ /d "VBScript Script File" /f
Reg.exe add "HKCR\VBSFile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\wshext.dll,-4802" /f
Reg.exe add "HKCR\VBSFile\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\WScript.exe,2" /f
Reg.exe add "HKCR\VBSFile\ScriptEngine" /ve /t REG_SZ /d "VBScript" /f
Reg.exe add "HKCR\VBSFile\ScriptHostEncode" /ve /t REG_SZ /d "{85131631-480C-11D2-B1F9-00C04F86C324}" /f
Reg.exe add "HKCR\VBSFile\Shell" /ve /t REG_SZ /d "Open" /f
Reg.exe add "HKCR\VBSFile\Shell\Edit\Command" /ve /t REG_EXPAND_SZ /d "\"%%SystemRoot%%\System32\Notepad.exe\" %%1" /f
Reg.exe add "HKCR\VBSFile\Shell\Open\Command" /ve /t REG_EXPAND_SZ /d "\"%%SystemRoot%%\System32\WScript.exe\" \"%%1\" %%*" /f
Reg.exe add "HKCR\VBSFile\Shell\Open2" /ve /t REG_EXPAND_SZ /d "Open &with Command Prompt" /f
Reg.exe add "HKCR\VBSFile\Shell\Open2" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\wshext.dll,-4511" /f
Reg.exe add "HKCR\VBSFile\Shell\Open2\Command" /ve /t REG_EXPAND_SZ /d "\"%%SystemRoot%%\System32\CScript.exe\" \"%%1\" %%*" /f
Reg.exe add "HKCR\VBSFile\Shell\Print\Command" /ve /t REG_EXPAND_SZ /d "\"%%SystemRoot%%\System32\Notepad.exe\" /p %%1" /f
Reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.vbs" /f


ECHO TEXT TXT

Reg.exe add "HKCR\.txt" /ve /t REG_SZ /d "txtfile" /f
Reg.exe add "HKCR\.txt" /v "Content Type" /t REG_SZ /d "text/plain" /f
Reg.exe add "HKCR\.txt" /v "PerceivedType" /t REG_SZ /d "text" /f
Reg.exe add "HKCR\.txt\PersistentHandler" /ve /t REG_SZ /d "{5e941d80-bf96-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\.txt\ShellNew" /v "ItemName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\notepad.exe,-470" /f
Reg.exe add "HKCR\.txt\ShellNew" /v "NullFile" /t REG_SZ /d "" /f
Reg.exe add "HKCR\txtfile" /ve /t REG_SZ /d "Text Document" /f
Reg.exe add "HKCR\txtfile" /v "EditFlags" /t REG_DWORD /d "2162688" /f
Reg.exe add "HKCR\txtfile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\notepad.exe,-469" /f
Reg.exe add "HKCR\txtfile\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\imageres.dll,-102" /f
Reg.exe add "HKCR\txtfile\shell\open\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\NOTEPAD.EXE %%1" /f
Reg.exe add "HKCR\txtfile\shell\print\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\NOTEPAD.EXE /p %%1" /f
Reg.exe add "HKCR\txtfile\shell\printto\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\notepad.exe /pt \"%%1\" \"%%2\" \"%%3\" \"%%4\"" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.txt\OpenWithProgids" /v "txtfile" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.txt\UserChoice" /v "Hash" /t REG_SZ /d "Ndb2OUV7duM=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.txt\UserChoice" /v "ProgId" /t REG_SZ /d "txtfile" /f


ECHO ICON ICO

Reg.exe add "HKCR\.ico" /ve /t REG_SZ /d "icofile" /f
Reg.exe add "HKCR\.ico" /v "Content Type" /t REG_SZ /d "image/x-icon" /f
Reg.exe add "HKCR\.ico" /v "PerceivedType" /t REG_SZ /d "image" /f
Reg.exe add "HKCR\.ico\OpenWithProgids" /v "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.ico\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\icofile" /ve /t REG_SZ /d "Icon" /f
Reg.exe add "HKCR\icofile" /v "EditFlags" /t REG_DWORD /d "2097152" /f
Reg.exe add "HKCR\icofile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\shell32.dll,-30601" /f
Reg.exe add "HKCR\icofile" /v "ImageOptionFlags" /t REG_DWORD /d "0" /f
Reg.exe add "HKCR\icofile\DefaultIcon" /ve /t REG_SZ /d "%%1" /f
Reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.ico" /f


ECHO ISO IMAGE

Reg.exe add "HKCR\.img" /ve /t REG_SZ /d "Windows.IsoFile" /f
Reg.exe add "HKCR\.img\OpenWithProgids" /v "Windows.IsoFile" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Windows.IsoFile" /ve /t REG_SZ /d "Disc Image File" /f
Reg.exe add "HKCR\Windows.IsoFile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\shell32.dll,-30602" /f
Reg.exe add "HKCR\Windows.IsoFile\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\imageres.dll,-5205" /f
Reg.exe add "HKCR\Windows.IsoFile\shell" /ve /t REG_SZ /d "mount" /f
Reg.exe add "HKCR\Windows.IsoFile\shell\burn" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\isoburn.exe,-351" /f
Reg.exe add "HKCR\Windows.IsoFile\shell\burn\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\isoburn.exe \"%%1\"" /f
Reg.exe add "HKCR\Windows.IsoFile\shell\mount" /v "CommandStateSync" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Windows.IsoFile\shell\mount" /v "ExplorerCommandHandler" /t REG_SZ /d "{9ab3b1c9-3225-4bb4-93b6-bfb3c0d93743}" /f
Reg.exe add "HKCR\Windows.IsoFile\shell\mount" /v "MultiSelectModel" /t REG_SZ /d "Document" /f
Reg.exe add "HKCR\Windows.IsoFile\shell\mount\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\Explorer.exe" /f
Reg.exe add "HKCR\Windows.IsoFile\shell\mount\command" /v "DelegateExecute" /t REG_SZ /d "{9ab3b1c9-3225-4bb4-93b6-bfb3c0d93743}" /f
Reg.exe add "HKCR\Windows.IsoFile\tabsets" /v "selection" /t REG_DWORD /d "1796" /f
Reg.exe delete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.img" /f
Exit



ECHO FILE SYSTEM FOLDER


Reg.exe add "HKCR\Folder" /ve /t REG_SZ /d "Folder" /f
Reg.exe add "HKCR\Folder" /v "ContentViewModeForBrowse" /t REG_SZ /d "prop:~System.ItemNameDisplay;~System.LayoutPattern.PlaceHolder;~System.LayoutPattern.PlaceHolder;~System.LayoutPattern.PlaceHolder;System.DateModified" /f
Reg.exe add "HKCR\Folder" /v "ContentViewModeForSearch" /t REG_SZ /d "prop:~System.ItemNameDisplay;System.DateModified;~System.ItemFolderPathDisplay" /f
Reg.exe add "HKCR\Folder" /v "ContentViewModeLayoutPatternForBrowse" /t REG_SZ /d "delta" /f
Reg.exe add "HKCR\Folder" /v "ContentViewModeLayoutPatternForSearch" /t REG_SZ /d "alpha" /f
Reg.exe add "HKCR\Folder" /v "EditFlags" /t REG_BINARY /d "d2030000" /f
Reg.exe add "HKCR\Folder" /v "FullDetails" /t REG_SZ /d "prop:System.PropGroup.Description;System.ItemNameDisplay;System.ItemTypeText;System.Size;System.HomeGroupSharingStatus" /f
Reg.exe add "HKCR\Folder" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Folder" /v "ThumbnailCutoff" /t REG_DWORD /d "0" /f
Reg.exe add "HKCR\Folder" /v "TileInfo" /t REG_SZ /d "prop:System.Title;System.HomeGroupSharingStatus" /f
Reg.exe add "HKCR\Folder\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\shell32.dll,3" /f
Reg.exe add "HKCR\Folder\shell\explore" /v "LaunchExplorerFlags" /t REG_DWORD /d "24" /f
Reg.exe add "HKCR\Folder\shell\explore" /v "MultiSelectModel" /t REG_SZ /d "Document" /f
Reg.exe add "HKCR\Folder\shell\explore" /v "ProgrammaticAccessOnly" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Folder\shell\explore\command" /v "DelegateExecute" /t REG_SZ /d "{11dbb47c-a525-400b-9e80-a54615a090c0}" /f
Reg.exe add "HKCR\Folder\shell\open" /v "MultiSelectModel" /t REG_SZ /d "Document" /f
Reg.exe delete "HKCR\Folder\shell\open" /v "LegacyDisable" /f
Reg.exe add "HKCR\Folder\shell\open\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\Explorer.exe" /f
Reg.exe add "HKCR\Folder\shell\open\command" /v "DelegateExecute" /t REG_SZ /d "{11dbb47c-a525-400b-9e80-a54615a090c0}" /f
Reg.exe add "HKCR\Folder\shell\opennewprocess" /v "ExplorerHost" /t REG_SZ /d "{ceff45ee-c862-41de-aee2-a022c81eda92}" /f
Reg.exe add "HKCR\Folder\shell\opennewprocess" /v "Extended" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Folder\shell\opennewprocess" /v "LaunchExplorerFlags" /t REG_DWORD /d "3" /f
Reg.exe add "HKCR\Folder\shell\opennewprocess" /v "MUIVerb" /t REG_SZ /d "@shell32.dll,-8518" /f
Reg.exe add "HKCR\Folder\shell\opennewprocess" /v "MultiSelectModel" /t REG_SZ /d "Document" /f
Reg.exe delete "HKCR\Folder\shell\opennewprocess" /v "LegacyDisable" /f
Reg.exe add "HKCR\Folder\shell\opennewprocess\command" /v "DelegateExecute" /t REG_SZ /d "{11dbb47c-a525-400b-9e80-a54615a090c0}" /f
Reg.exe add "HKCR\Folder\shell\opennewwindow" /v "LaunchExplorerFlags" /t REG_DWORD /d "1" /f
Reg.exe add "HKCR\Folder\shell\opennewwindow" /v "MUIVerb" /t REG_SZ /d "@windows.storage.dll,-8517" /f
Reg.exe add "HKCR\Folder\shell\opennewwindow" /v "MultiSelectModel" /t REG_SZ /d "Document" /f
Reg.exe add "HKCR\Folder\shell\opennewwindow" /v "OnlyInBrowserWindow" /t REG_SZ /d "" /f
Reg.exe delete "HKCR\Folder\shell\opennewwindow" /v "LegacyDisable" /f
Reg.exe add "HKCR\Folder\shell\opennewwindow\command" /v "DelegateExecute" /t REG_SZ /d "{11dbb47c-a525-400b-9e80-a54615a090c0}" /f
Reg.exe add "HKCR\Folder\shell\pintohome" /v "AppliesTo" /t REG_SZ /d "System.ParsingName:<>\"::{679f85cb-0220-4080-b29b-5540cc05aab6}\" AND System.ParsingName:<>\"::{645FF040-5081-101B-9F08-00AA002F954E}\" AND System.IsFolder:=System.StructuredQueryType.Boolean#True" /f
Reg.exe add "HKCR\Folder\shell\pintohome" /v "MUIVerb" /t REG_SZ /d "@shell32.dll,-51377" /f
Reg.exe add "HKCR\Folder\shell\pintohome\command" /v "DelegateExecute" /t REG_SZ /d "{b455f46e-e4af-4035-b0a4-cf18d2f6f28e}" /f
Reg.exe add "HKCR\Folder\ShellNew" /v "Directory" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Folder\ShellNew" /v "IconPath" /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\shell32.dll,3" /f
Reg.exe add "HKCR\Folder\ShellNew" /v "ItemName" /t REG_SZ /d "@shell32.dll,-30396" /f
Reg.exe add "HKCR\Folder\ShellNew" /v "MenuText" /t REG_SZ /d "@shell32.dll,-30317" /f
Reg.exe add "HKCR\Folder\ShellNew" /v "NonLFNFileSpec" /t REG_SZ /d "@shell32.dll,-30319" /f
Reg.exe add "HKCR\Folder\ShellNew\Config" /v "AllDrives" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Folder\ShellNew\Config" /v "IsFolder" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Folder\ShellNew\Config" /v "NoExtension" /t REG_SZ /d "" /f

ECHO FILE SYSTEM DIRECTORY

Reg.exe add "HKCR\Directory" /ve /t REG_SZ /d "File Folder" /f
Reg.exe add "HKCR\Directory" /v "AlwaysShowExt" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Directory" /v "EditFlags" /t REG_BINARY /d "d2010000" /f
Reg.exe add "HKCR\Directory" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\windows.storage.dll,-10152" /f
Reg.exe add "HKCR\Directory" /v "FullDetails" /t REG_SZ /d "prop:System.PropGroup.Description;System.DateCreated;System.FileCount;System.TotalFileSize" /f
Reg.exe add "HKCR\Directory" /v "InfoTip" /t REG_SZ /d "prop:System.Comment;System.DateCreated" /f
Reg.exe add "HKCR\Directory" /v "NoRecentDocs" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Directory" /v "PreviewDetails" /t REG_SZ /d "prop:System.DateModified;*System.SharedWith;*System.OfflineAvailability;*System.OfflineStatus" /f
Reg.exe add "HKCR\Directory" /v "PreviewTitle" /t REG_SZ /d "prop:System.ItemNameDisplay;System.ItemTypeText" /f
Reg.exe add "HKCR\Directory\Background\shell\cmd" /ve /t REG_SZ /d "@shell32.dll,-8506" /f
Reg.exe add "HKCR\Directory\Background\shell\cmd" /v "Extended" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Directory\Background\shell\cmd" /v "NoWorkingDirectory" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Directory\Background\shell\cmd\command" /ve /t REG_SZ /d "cmd.exe /s /k pushd \"%%V\"" /f
Reg.exe add "HKCR\directory\Background\shellex\ContextMenuHandlers\New" /ve /t REG_SZ /d "{D969A300-E7FF-11d0-A93B-00A0C90F2719}" /f
Reg.exe add "HKCR\directory\Background\shellex\ContextMenuHandlers\Sharing" /ve /t REG_SZ /d "{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}" /f
Reg.exe add "HKCR\Directory\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\shell32.dll,3" /f
Reg.exe add "HKCR\Directory\shell" /ve /t REG_SZ /d "none" /f
Reg.exe add "HKCR\Directory\shell\cmd" /ve /t REG_SZ /d "@shell32.dll,-8506" /f
Reg.exe add "HKCR\Directory\shell\cmd" /v "Extended" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Directory\shell\cmd" /v "NoWorkingDirectory" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Directory\shell\cmd\command" /ve /t REG_SZ /d "cmd.exe /s /k pushd \"%%V\"" /f
Reg.exe add "HKCR\Directory\shell\find" /v "LegacyDisable" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Directory\shell\find" /v "SuppressionPolicy" /t REG_DWORD /d "128" /f
Reg.exe add "HKCR\Directory\shell\find\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\Explorer.exe" /f
Reg.exe add "HKCR\Directory\shell\find\command" /v "DelegateExecute" /t REG_SZ /d "{a015411a-f97d-4ef3-8425-8a38d022aebc}" /f


ECHO FILE SYSTEM DRIVE

Reg.exe add "HKCR\Drive" /ve /t REG_SZ /d "Drive" /f
Reg.exe add "HKCR\Drive" /v "ContentViewModeForBrowse" /t REG_SZ /d "prop:~System.ItemNameDisplay;~System.PercentFull;~System.LayoutPattern.PlaceHolder;~System.LayoutPattern.PlaceHolder;~System.Volume.FileSystem;~System.Computer.DecoratedFreeSpace" /f
Reg.exe add "HKCR\Drive" /v "ContentViewModeLayoutPatternForBrowse" /t REG_SZ /d "delta" /f
Reg.exe add "HKCR\Drive" /v "EditFlags" /t REG_BINARY /d "d2010000" /f
Reg.exe add "HKCR\Drive" /v "InfoTip" /t REG_SZ /d "prop:System.FreeSpace;System.Capacity" /f
Reg.exe add "HKCR\Drive" /v "PreviewDetails" /t REG_SZ /d "prop:*System.PercentFull;System.FreeSpace;System.Capacity;System.Volume.FileSystem;*System.Volume.BitLockerProtection" /f
Reg.exe add "HKCR\Drive" /v "PreviewTitle" /t REG_SZ /d "prop:System.ItemNameDisplay;System.ItemTypeText" /f
Reg.exe add "HKCR\Drive" /v "TileInfo" /t REG_SZ /d "prop:*System.PercentFull;System.Computer.DecoratedFreeSpace;System.Volume.FileSystem" /f
Reg.exe add "HKCR\Drive\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\shell32.dll,8" /f
Reg.exe add "HKCR\Drive\shell" /ve /t REG_SZ /d "none" /f
Reg.exe add "HKCR\Drive\shell\cmd" /ve /t REG_SZ /d "@shell32.dll,-8506" /f
Reg.exe add "HKCR\Drive\shell\cmd" /v "Extended" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Drive\shell\cmd" /v "NoWorkingDirectory" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Drive\shell\cmd\command" /ve /t REG_SZ /d "cmd.exe /s /k pushd \"%%V\"" /f
Reg.exe add "HKCR\Drive\shell\find" /v "LegacyDisable" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Drive\shell\find" /v "SuppressionPolicy" /t REG_DWORD /d "128" /f
Reg.exe add "HKCR\Drive\shell\find\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\Explorer.exe" /f
Reg.exe add "HKCR\Drive\shell\find\command" /v "DelegateExecute" /t REG_SZ /d "{a015411a-f97d-4ef3-8425-8a38d022aebc}" /f
Reg.exe add "HKCR\Drive\tabsets" /v "location" /t REG_DWORD /d "1795" /f
Reg.exe add "HKCR\Drive\tabsets" /v "selection" /t REG_DWORD /d "1795" /f




:BEGIN_IMAGE

SET IMAGE=%%systemroot%%\system32\mspaint.exe
IF EXIST "C:\Program Files (x86)\IrfanView\i_view32.exe" SET IMAGE="\"C:\Program Files (x86)\IrfanView\i_view32.exe\"

ECHO IMAGE BMP

Reg.exe add "HKCR\.bmp" /ve /t REG_SZ /d "Paint.Picture" /f
Reg.exe add "HKCR\.bmp" /v "Content Type" /t REG_SZ /d "image/bmp" /f
Reg.exe add "HKCR\.bmp" /v "PerceivedType" /t REG_SZ /d "image" /f

Reg.exe add "HKCR\.bmp\OpenWithList\MSPaint.exe" /ve /t REG_SZ /d "" /f

Reg.exe add "HKCR\.bmp\OpenWithProgids" /v "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.bmp\OpenWithProgids" /v "Paint.Picture" /t REG_SZ /d "" /f
Reg.exe add "HKCR\.bmp\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f

REM Reg.exe add "HKCR\.bmp\ShellNew" /v "ItemName" /t REG_EXPAND_SZ /d "@%IMAGE%,-59414" /f
REM ORIGINAL
Reg.exe add "HKCR\.bmp\ShellNew" /v "ItemName" /t REG_EXPAND_SZ /d "@%%systemroot%%\system32\mspaint.exe,-59414" /f

Reg.exe add "HKCR\.bmp\ShellNew" /v "NullFile" /t REG_SZ /d "" /f
Reg.exe add "HKCR\Paint.Picture" /ve /t REG_SZ /d "Bitmap Image" /f
Reg.exe add "HKCR\Paint.Picture" /v "EditFlags" /t REG_DWORD /d "2097152" /f
Reg.exe add "HKCR\Paint.Picture" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\shell32.dll,-30597" /f
Reg.exe add "HKCR\Paint.Picture" /v "ImageOptionFlags" /t REG_DWORD /d "1" /f
Reg.exe add "HKCR\Paint.Picture\CLSID" /ve /t REG_SZ /d "{D3E34B21-9D75-101A-8C3D-00AA001A1652}" /f
Reg.exe add "HKCR\Paint.Picture\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%systemroot%%\system32\imageres.dll,-70" /f
Reg.exe add "HKCR\Paint.Picture\Insertable" /ve /t REG_SZ /d "" /f

Reg.exe add "HKCR\Paint.Picture\protocol\StdFileEditing\server" /ve /t REG_EXPAND_SZ /d "%%systemroot%%\system32\mspaint.exe" /f

Reg.exe add "HKCR\Paint.Picture\protocol\StdFileEditing\verb\0" /ve /t REG_SZ /d "&Edit" /f
Reg.exe add "HKCR\Paint.Picture\protocol\StdFileEditing\verb\0" /v "LocalizedString" /t REG_EXPAND_SZ /d "@%%systemroot%%\system32\mspaint.exe,-59420" /f
Reg.exe add "HKCR\Paint.Picture\shell\edit\command" /ve /t REG_EXPAND_SZ /d "\"%%systemroot%%\system32\mspaint.exe\" \"%%1\"" /f

Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.bmp\OpenWithProgids" /v "Paint.Picture" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.bmp\UserChoice" /v "Hash" /t REG_SZ /d "QvKoQ4BvvU0=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.bmp\UserChoice" /v "ProgId" /t REG_SZ /d "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /f



ECHO IMAGE GIF

Reg.exe add "HKCR\.gif" /ve /t REG_SZ /d "giffile" /f
Reg.exe add "HKCR\.gif" /v "Content Type" /t REG_SZ /d "image/gif" /f
Reg.exe add "HKCR\.gif" /v "PerceivedType" /t REG_SZ /d "image" /f
Reg.exe add "HKCR\.gif\OpenWithProgids" /v "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.gif\OpenWithProgids" /v "giffile" /t REG_SZ /d "" /f
Reg.exe add "HKCR\.gif\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\giffile" /ve /t REG_SZ /d "GIF Image" /f
Reg.exe add "HKCR\giffile" /v "EditFlags" /t REG_DWORD /d "2097152" /f
Reg.exe add "HKCR\giffile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\shell32.dll,-30595" /f
Reg.exe add "HKCR\giffile" /v "ImageOptionFlags" /t REG_DWORD /d "0" /f
Reg.exe add "HKCR\giffile\CLSID" /ve /t REG_SZ /d "{25336920-03F9-11cf-8FD0-00AA00686F13}" /f
Reg.exe add "HKCR\giffile\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\imageres.dll,-71" /f
Reg.exe add "HKCR\giffile\shell\Open" /v "CommandId" /t REG_SZ /d "IE.File" /f

Reg.exe add "HKCR\giffile\shell\Open\command" /ve /t REG_SZ /d "\"C:\Program Files\Internet Explorer\iexplore.exe\" %%1" /f

Reg.exe add "HKCR\giffile\shell\Open\command" /v "DelegateExecute" /t REG_SZ /d "{17FE9752-0B5A-4665-84CD-569794602F5C}" /f
Reg.exe add "HKCR\giffile\shell\printto" /v "NeverDefault" /t REG_SZ /d "" /f
Reg.exe add "HKCR\giffile\shell\printto\command" /ve /t REG_EXPAND_SZ /d "\"%%SystemRoot%%\System32\rundll32.exe\" \"%%SystemRoot%%\System32\shimgvw.dll\",ImageView_PrintTo /pt \"%%1\" \"%%2\" \"%%3\" \"%%4\"" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.gif\OpenWithProgids" /v "giffile" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.gif\UserChoice" /v "Hash" /t REG_SZ /d "gk2WsPVVdP0=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.gif\UserChoice" /v "ProgId" /t REG_SZ /d "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /f

ECHO IMAGE JPG

Reg.exe add "HKCR\.jpe" /ve /t REG_SZ /d "jpegfile" /f
Reg.exe add "HKCR\.jpe" /v "Content Type" /t REG_SZ /d "image/jpeg" /f
Reg.exe add "HKCR\.jpe" /v "PerceivedType" /t REG_SZ /d "image" /f
Reg.exe add "HKCR\.jpe\OpenWithProgids" /v "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.jpe\OpenWithProgids" /v "jpegfile" /t REG_SZ /d "" /f
Reg.exe add "HKCR\.jpe\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\.jpeg" /ve /t REG_SZ /d "jpegfile" /f
Reg.exe add "HKCR\.jpeg" /v "Content Type" /t REG_SZ /d "image/jpeg" /f
Reg.exe add "HKCR\.jpeg" /v "PerceivedType" /t REG_SZ /d "image" /f
Reg.exe add "HKCR\.jpeg\OpenWithProgids" /v "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.jpeg\OpenWithProgids" /v "jpegfile" /t REG_SZ /d "" /f
Reg.exe add "HKCR\.jpeg\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\.jpg" /ve /t REG_SZ /d "jpegfile" /f
Reg.exe add "HKCR\.jpg" /v "Content Type" /t REG_SZ /d "image/jpeg" /f
Reg.exe add "HKCR\.jpg" /v "PerceivedType" /t REG_SZ /d "image" /f
Reg.exe add "HKCR\.jpg\OpenWithProgids" /v "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.jpg\OpenWithProgids" /v "jpegfile" /t REG_SZ /d "" /f
Reg.exe add "HKCR\.jpg\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\jpegfile" /ve /t REG_SZ /d "JPEG Image" /f
Reg.exe add "HKCR\jpegfile" /v "EditFlags" /t REG_DWORD /d "2162688" /f
Reg.exe add "HKCR\jpegfile" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\shell32.dll,-30596" /f
Reg.exe add "HKCR\jpegfile" /v "ImageOptionFlags" /t REG_DWORD /d "1" /f
Reg.exe add "HKCR\jpegfile\CLSID" /ve /t REG_SZ /d "{25336920-03F9-11cf-8FD0-00AA00686F13}" /f
Reg.exe add "HKCR\jpegfile\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\imageres.dll,-72" /f
Reg.exe add "HKCR\jpegfile\shell\printto" /v "NeverDefault" /t REG_SZ /d "" /f
Reg.exe add "HKCR\jpegfile\shell\printto\command" /ve /t REG_EXPAND_SZ /d "\"%%SystemRoot%%\System32\rundll32.exe\" \"%%SystemRoot%%\System32\shimgvw.dll\",ImageView_PrintTo /pt \"%%1\" \"%%2\" \"%%3\" \"%%4\"" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpe\OpenWithProgids" /v "jpegfile" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpe\UserChoice" /v "Hash" /t REG_SZ /d "xdV0ulvVB+U=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpe\UserChoice" /v "ProgId" /t REG_SZ /d "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpeg\OpenWithProgids" /v "jpegfile" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpeg\UserChoice" /v "Hash" /t REG_SZ /d "X3k+ZJEbhsw=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpeg\UserChoice" /v "ProgId" /t REG_SZ /d "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpg\OpenWithProgids" /v "jpegfile" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpg\UserChoice" /v "Hash" /t REG_SZ /d "4EEK1g9hlRI=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.jpg\UserChoice" /v "ProgId" /t REG_SZ /d "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /f


ECHO IMAGE TIF

Reg.exe add "HKCR\.tif" /ve /t REG_SZ /d "TIFImage.Document" /f
Reg.exe add "HKCR\.tif" /v "Content Type" /t REG_SZ /d "image/tiff" /f
Reg.exe add "HKCR\.tif" /v "PerceivedType" /t REG_SZ /d "image" /f
Reg.exe add "HKCR\.tif\OpenWithProgids" /v "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.tif\OpenWithProgids" /v "TIFImage.Document" /t REG_SZ /d "" /f
Reg.exe add "HKCR\.tiff" /ve /t REG_SZ /d "TIFImage.Document" /f
Reg.exe add "HKCR\.tiff" /v "Content Type" /t REG_SZ /d "image/tiff" /f
Reg.exe add "HKCR\.tiff" /v "PerceivedType" /t REG_SZ /d "image" /f
Reg.exe add "HKCR\.tiff\OpenWithProgids" /v "AppX43hnxtbyyps62jhe9sqpdzxn1790zetc" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.tiff\OpenWithProgids" /v "TIFImage.Document" /t REG_SZ /d "" /f
Reg.exe add "HKCR\TIFImage.Document" /v "EditFlags" /t REG_DWORD /d "2097152" /f
Reg.exe add "HKCR\TIFImage.Document" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\System32\shell32.dll,-30599" /f
Reg.exe add "HKCR\TIFImage.Document" /v "ImageOptionFlags" /t REG_DWORD /d "1" /f
Reg.exe add "HKCR\TIFImage.Document\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\imageres.dll,-122" /f
Reg.exe add "HKCR\TIFImage.Document\shell\open" /v "MuiVerb" /t REG_EXPAND_SZ /d "@%%ProgramFiles%%\Windows Photo Viewer\photoviewer.dll,-3043" /f
Reg.exe add "HKCR\TIFImage.Document\shell\open\command" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\rundll32.exe \"%%ProgramFiles%%\Windows Photo Viewer\PhotoViewer.dll\", ImageView_Fullscreen %%1" /f
Reg.exe add "HKCR\TIFImage.Document\shell\open\DropTarget" /v "Clsid" /t REG_SZ /d "{FFE2A43C-56B9-4bf5-9A79-CC6D4285608A}" /f
Reg.exe add "HKCR\TIFImage.Document\shell\printto" /v "NeverDefault" /t REG_SZ /d "" /f
Reg.exe add "HKCR\TIFImage.Document\shell\printto\command" /ve /t REG_EXPAND_SZ /d "\"%%SystemRoot%%\System32\rundll32.exe\" \"%%SystemRoot%%\System32\shimgvw.dll\",ImageView_PrintTo /pt \"%%1\" \"%%2\" \"%%3\" \"%%4\"" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tif\OpenWithProgids" /v "TIFImage.Document" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tif\UserChoice" /v "Hash" /t REG_SZ /d "u+wCl7N4kRg=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tif\UserChoice" /v "ProgId" /t REG_SZ /d "PhotoViewer.FileAssoc.Tiff" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tiff\OpenWithProgids" /v "TIFImage.Document" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tiff\UserChoice" /v "Hash" /t REG_SZ /d "yf9GjrAGsew=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.tiff\UserChoice" /v "ProgId" /t REG_SZ /d "PhotoViewer.FileAssoc.Tiff" /f

:END_IMAGE


:BEGIN_AUDIO
ECHO AUDIO MP3

SET AUDIO="\"%%ProgramFiles(x86)%%\Windows Media Player\wmplayer.exe\"
IF EXIST "C:\Program Files (x86)\foobar2000\foobar2000.exe" SET AUDIO=\"C:\Program Files (x86)\foobar2000\foobar2000.exe\"

ECHO %AUDIO%

Reg.exe add "HKCR\.mp3" /ve /t REG_SZ /d "WMP11.AssocFile.MP3" /f
Reg.exe add "HKCR\.mp3" /v "Content Type" /t REG_SZ /d "audio/mpeg" /f
Reg.exe add "HKCR\.mp3" /v "PerceivedType" /t REG_SZ /d "audio" /f
Reg.exe add "HKCR\.mp3\OpenWithProgids" /v "AppXqj98qxeaynz6dv4459ayz6bnqxbyaqcs" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.mp3\OpenWithProgids" /v "WMP11.AssocFile.MP3" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.mp3\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3" /ve /t REG_SZ /d "MP3 Format Sound" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3" /v "EditFlags" /t REG_BINARY /d "00003100" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9925" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3" /v "PreferExecuteOnMismatch" /t REG_DWORD /d "1" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\wmploc.dll,-732" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3\shell" /ve /t REG_SZ /d "Play" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3\shell\Enqueue" /ve /t REG_SZ /d "&Add to Windows Media Player list" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3\shell\Enqueue" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9800" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3\shell\Enqueue\command" /v "DelegateExecute" /t REG_SZ /d "{45597c98-80f6-4549-84ff-752cf55e2d29}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3\shell\open" /v "LegacyDisable" /t REG_SZ /d "" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3\shell\open\command" /ve /t REG_EXPAND_SZ /d "\"%%ProgramFiles(x86)%%\Windows Media Player\wmplayer.exe\" /prefetch:6 /Open \"%%L\"" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3\shell\Play" /ve /t REG_SZ /d "&Play" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3\shell\Play" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9991" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3\shell\Play\command" /ve /t REG_EXPAND_SZ /d "\"%%ProgramFiles(x86)%%\Windows Media Player\wmplayer.exe\" /prefetch:6 /Play \"%%L\"" /f
Reg.exe add "HKCR\WMP11.AssocFile.MP3\shell\Play\command" /v "DelegateExecute" /t REG_SZ /d "{ed1d0fdf-4414-470a-a56d-cfb68623fc58}" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.mp3\OpenWithProgids" /v "WMP11.AssocFile.MP3" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.mp3\UserChoice" /v "Hash" /t REG_SZ /d "yXSHbkJgtdE=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.mp3\UserChoice" /v "ProgId" /t REG_SZ /d "AppXqj98qxeaynz6dv4459ayz6bnqxbyaqcs" /f
ECHO AUDIO WMA

Reg.exe add "HKCR\.wma" /ve /t REG_SZ /d "WMP11.AssocFile.WMA" /f
Reg.exe add "HKCR\.wma" /v "Content Type" /t REG_SZ /d "audio/x-ms-wma" /f
Reg.exe add "HKCR\.wma" /v "PerceivedType" /t REG_SZ /d "audio" /f
Reg.exe add "HKCR\.wma\OpenWithProgids" /v "AppXqj98qxeaynz6dv4459ayz6bnqxbyaqcs" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.wma\OpenWithProgids" /v "WMP11.AssocFile.WMA" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.wma\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma" /ve /t REG_SZ /d "Windows Media Audio file" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma" /v "EditFlags" /t REG_BINARY /d "00003100" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9912" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma" /v "PreferExecuteOnMismatch" /t REG_DWORD /d "1" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\wmploc.dll,-735" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma\shell" /ve /t REG_SZ /d "Play" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma\shell\Enqueue" /ve /t REG_SZ /d "&Add to Windows Media Player list" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma\shell\Enqueue" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9800" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma\shell\Enqueue\command" /v "DelegateExecute" /t REG_SZ /d "{45597c98-80f6-4549-84ff-752cf55e2d29}" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma\shell\open" /v "LegacyDisable" /t REG_SZ /d "" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma\shell\open\command" /ve /t REG_EXPAND_SZ /d "\"%%ProgramFiles(x86)%%\Windows Media Player\wmplayer.exe\" /prefetch:5 /Open \"%%L\"" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma\shell\Play" /ve /t REG_SZ /d "&Play" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma\shell\Play" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9991" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma\shell\Play\command" /ve /t REG_EXPAND_SZ /d "\"%%ProgramFiles(x86)%%\Windows Media Player\wmplayer.exe\" /prefetch:5 /Play \"%%L\"" /f
Reg.exe add "HKCR\WMP11.AssocFile.wma\shell\Play\command" /v "DelegateExecute" /t REG_SZ /d "{ed1d0fdf-4414-470a-a56d-cfb68623fc58}" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.wma\OpenWithProgids" /v "WMP11.AssocFile.WMA" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.wma\UserChoice" /v "Hash" /t REG_SZ /d "W8GoejaVk8M=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.wma\UserChoice" /v "ProgId" /t REG_SZ /d "AppXqj98qxeaynz6dv4459ayz6bnqxbyaqcs" /f
:END_AUDIO




:BEGIN_VIDEO

REM WINDOWS VISTA BEST COULD FIND MATCH TO WINDOWS HARDCORE 10
REM ----------------------------------------------------------

SET MPC="\"%%ProgramFiles(x86)%%\Windows Media Player\wmplayer.exe\"
IF EXIST "C:\Program Files\MPC-HC\mpc-hc64.exe" SET MPC="\"C:\Program Files\MPC-HC\mpc-hc64.exe\"

ECHO VIDEO AVI

Reg.exe add "HKCR\.avi" /ve /t REG_SZ /d "WMP11.AssocFile.AVI" /f
Reg.exe add "HKCR\.avi" /v "Content Type" /t REG_SZ /d "video/avi" /f
Reg.exe add "HKCR\.avi" /v "PerceivedType" /t REG_SZ /d "video" /f
Reg.exe add "HKCR\.avi\OpenWithProgids" /v "AppX6eg8h5sxqq90pv53845wmnbewywdqq5h" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.avi\OpenWithProgids" /v "AppXk0g4vb8gvt7b93tg50ybcy892pge6jmt" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.avi\OpenWithProgids" /v "WMP11.AssocFile.AVI" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.avi\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI" /ve /t REG_SZ /d "Video Clip" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI" /v "EditFlags" /t REG_BINARY /d "00003100" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9905" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI" /v "PreferExecuteOnMismatch" /t REG_DWORD /d "1" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\wmploc.dll,-731" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI\shell" /ve /t REG_SZ /d "Play" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI\shell\Enqueue" /ve /t REG_SZ /d "&Add to Windows Media Player list" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI\shell\Enqueue" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9800" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI\shell\Enqueue\command" /v "DelegateExecute" /t REG_SZ /d "{45597c98-80f6-4549-84ff-752cf55e2d29}" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI\shell\open" /v "LegacyDisable" /t REG_SZ /d "" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI\shell\open\command" /ve /t REG_EXPAND_SZ /d "\"%%ProgramFiles(x86)%%\Windows Media Player\wmplayer.exe\" /prefetch:8 /Open \"%%L\"" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI\shell\Play" /ve /t REG_SZ /d "&Play" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI\shell\Play" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9991" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI\shell\Play\command" /ve /t REG_EXPAND_SZ /d "\"%%ProgramFiles(x86)%%\Windows Media Player\wmplayer.exe\" /prefetch:8 /Play \"%%L\"" /f
Reg.exe add "HKCR\WMP11.AssocFile.AVI\shell\Play\command" /v "DelegateExecute" /t REG_SZ /d "{ed1d0fdf-4414-470a-a56d-cfb68623fc58}" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.avi\OpenWithProgids" /v "WMP11.AssocFile.AVI" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.avi\UserChoice" /v "Hash" /t REG_SZ /d "233Vijon5zY=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.avi\UserChoice" /v "ProgId" /t REG_SZ /d "AppX6eg8h5sxqq90pv53845wmnbewywdqq5h" /f


ECHO VIDEO MPEG


Reg.exe add "HKCR\.MPEG" /v "PerceivedType" /t REG_SZ /d "video" /f
Reg.exe add "HKCR\.MPEG" /ve /t REG_SZ /d "WMP11.AssocFile.MPEG" /f
Reg.exe add "HKCR\.MPEG" /v "Content Type" /t REG_SZ /d "video/mpeg" /f
Reg.exe add "HKCR\.MPEG\OpenWithProgIds" /v "MPEGFile" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.MPEG\OpenWithProgIds" /v "WMP11.AssocFile.MPEG" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.MPEG\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\.MPEG\ShellEx\{BB2E617C-0920-11D1-9A0B-00C04FC2D6C1}" /ve /t REG_SZ /d "{c5a40261-cd64-4ccf-84cb-c394da41d590}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG" /ve /t REG_SZ /d "Movie Clip" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG" /v "EditFlags" /t REG_BINARY /d "00000100" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9902" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG" /v "PreferExecuteOnMismatch" /t REG_DWORD /d "1" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\DefaultIcon" /ve /t REG_SZ /d "C:\Windows\system32\wmploc.dll,-733" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell" /ve /t REG_SZ /d "play" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\open" /v "LegacyDisable" /t REG_SZ /d "" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\open\command" /ve /t REG_EXPAND_SZ /d "\"%%ProgramFiles%%\Windows Media Player\wmplayer.exe\" /prefetch:9 /Open \"%%L\"" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\open\DropTarget" /v "CLSID" /t REG_SZ /d "{CE3FB1D1-02AE-4a5f-A6E9-D9F1B4073E6C}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\open\DropTarget" /v "MPlayer2.Bak" /t REG_SZ /d "{CE3FB1D1-02AE-4A5F-A6E9- D9F1B4073E6C}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\play" /ve /t REG_SZ /d "&Play" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\play" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9991" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\play" /v "LegacyDisable" /t REG_SZ /d "" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\play\command" /ve /t REG_EXPAND_SZ /d "\"%%ProgramFiles%%\Windows Media Player\wmplayer.exe\" /prefetch:9 /Play \"%%L\"" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\play\DropTarget" /v "CLSID" /t REG_SZ /d "{CE3FB1D1-02AE-4a5f-A6E9-D9F1B4073E6C}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\play\DropTarget" /v "MPlayer2.Bak" /t REG_SZ /d "{CE3FB1D1-02AE-4A5F-A6E9- D9F1B4073E6C}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shellex\ContextMenuHandlers\WMPAddToPlaylist" /ve /t REG_SZ /d "{F1B9284F-E9DC-4e68-9D7E-42362A59F0FD}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shellex\ContextMenuHandlers\WMPPlayAsPlaylist" /ve /t REG_SZ /d "{CE3FB1D1-02AE-4a5f-A6E9-D9F1B4073E6C}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}" /ve /t REG_SZ /d "{031EE060-67BC-460d-8847-E4A7C5E45A27}" /f


ECHO VIDEO MPG

REM WINDOWS VISTA BEST COULD FIND MATCH TO WINDOWS HARDCORE 10
REM ----------------------------------------------------------
Reg.exe add "HKCR\.MPG" /v "PerceivedType" /t REG_SZ /d "video" /f
Reg.exe add "HKCR\.MPG" /ve /t REG_SZ /d "WMP11.AssocFile.MPEG" /f
Reg.exe add "HKCR\.MPG" /v "Content Type" /t REG_SZ /d "video/mpeg" /f
Reg.exe add "HKCR\.MPG\OpenWithProgIds" /v "MPEGFile" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.MPG\OpenWithProgIds" /v "WMP11.AssocFile.MPEG" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.MPG\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\.MPG\ShellEx\{BB2E617C-0920-11D1-9A0B-00C04FC2D6C1}" /ve /t REG_SZ /d "{c5a40261-cd64-4ccf-84cb-c394da41d590}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG" /ve /t REG_SZ /d "Movie Clip" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG" /v "EditFlags" /t REG_BINARY /d "00000100" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9902" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG" /v "PreferExecuteOnMismatch" /t REG_DWORD /d "1" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\DefaultIcon" /ve /t REG_SZ /d "C:\Windows\system32\wmploc.dll,-733" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell" /ve /t REG_SZ /d "play" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\open" /v "LegacyDisable" /t REG_SZ /d "" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\open\command" /ve /t REG_EXPAND_SZ /d "\"%%ProgramFiles%%\Windows Media Player\wmplayer.exe\" /prefetch:9 /Open \"%%L\"" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\open\DropTarget" /v "CLSID" /t REG_SZ /d "{CE3FB1D1-02AE-4a5f-A6E9-D9F1B4073E6C}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\open\DropTarget" /v "MPlayer2.Bak" /t REG_SZ /d "{CE3FB1D1-02AE-4A5F-A6E9- D9F1B4073E6C}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\play" /ve /t REG_SZ /d "&Play" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\play" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9991" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\play" /v "LegacyDisable" /t REG_SZ /d "" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\play\command" /ve /t REG_EXPAND_SZ /d "\"%%ProgramFiles%%\Windows Media Player\wmplayer.exe\" /prefetch:9 /Play \"%%L\"" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\play\DropTarget" /v "CLSID" /t REG_SZ /d "{CE3FB1D1-02AE-4a5f-A6E9-D9F1B4073E6C}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shell\play\DropTarget" /v "MPlayer2.Bak" /t REG_SZ /d "{CE3FB1D1-02AE-4A5F-A6E9- D9F1B4073E6C}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shellex\ContextMenuHandlers\WMPAddToPlaylist" /ve /t REG_SZ /d "{F1B9284F-E9DC-4e68-9D7E-42362A59F0FD}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shellex\ContextMenuHandlers\WMPPlayAsPlaylist" /ve /t REG_SZ /d "{CE3FB1D1-02AE-4a5f-A6E9-D9F1B4073E6C}" /f
Reg.exe add "HKCR\WMP11.AssocFile.MPEG\shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}" /ve /t REG_SZ /d "{031EE060-67BC-460d-8847-E4A7C5E45A27}" /f

ECHO VIDEO WMV

Reg.exe add "HKCR\.wmv" /ve /t REG_SZ /d "WMP11.AssocFile.WMV" /f
Reg.exe add "HKCR\.wmv" /v "Content Type" /t REG_SZ /d "video/x-ms-wmv" /f
Reg.exe add "HKCR\.wmv" /v "PerceivedType" /t REG_SZ /d "video" /f
Reg.exe add "HKCR\.wmv\OpenWithProgids" /v "AppX6eg8h5sxqq90pv53845wmnbewywdqq5h" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.wmv\OpenWithProgids" /v "AppXk0g4vb8gvt7b93tg50ybcy892pge6jmt" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.wmv\OpenWithProgids" /v "WMP11.AssocFile.WMV" /t REG_NONE /d "" /f
Reg.exe add "HKCR\.wmv\PersistentHandler" /ve /t REG_SZ /d "{098f2470-bae0-11cd-b579-08002b30bfeb}" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV" /ve /t REG_SZ /d "Windows Media Audio/Video file" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV" /v "EditFlags" /t REG_BINARY /d "00003100" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV" /v "FriendlyTypeName" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9914" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV" /v "PreferExecuteOnMismatch" /t REG_DWORD /d "1" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV\DefaultIcon" /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\system32\wmploc.dll,-736" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV\shell" /ve /t REG_SZ /d "Play" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV\shell\Enqueue" /ve /t REG_SZ /d "&Add to Windows Media Player list" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV\shell\Enqueue" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9800" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV\shell\Enqueue\command" /v "DelegateExecute" /t REG_SZ /d "{45597c98-80f6-4549-84ff-752cf55e2d29}" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV\shell\open" /v "LegacyDisable" /t REG_SZ /d "" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV\shell\open\command" /ve /t REG_EXPAND_SZ /d "\"%%ProgramFiles(x86)%%\Windows Media Player\wmplayer.exe\" /prefetch:7 /Open \"%%L\"" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV\shell\Play" /ve /t REG_SZ /d "&Play" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV\shell\Play" /v "MUIVerb" /t REG_EXPAND_SZ /d "@%%SystemRoot%%\system32\unregmp2.exe,-9991" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV\shell\Play\command" /ve /t REG_EXPAND_SZ /d "\"%%ProgramFiles(x86)%%\Windows Media Player\wmplayer.exe\" /prefetch:7 /Play \"%%L\"" /f
Reg.exe add "HKCR\WMP11.AssocFile.WMV\shell\Play\command" /v "DelegateExecute" /t REG_SZ /d "{ed1d0fdf-4414-470a-a56d-cfb68623fc58}" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.wmv\OpenWithProgids" /v "WMP11.AssocFile.WMV" /t REG_NONE /d "" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.wmv\UserChoice" /v "Hash" /t REG_SZ /d "wDsCuH9Cpvo=" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.wmv\UserChoice" /v "ProgId" /t REG_SZ /d "AppX6eg8h5sxqq90pv53845wmnbewywdqq5h" /f

:END_VIDEO



:VERY_ENDER_BLOCK

IF "%1"=="" PAUSE
IF "%1"=="/ADMIN" PAUSE

:ENDE

