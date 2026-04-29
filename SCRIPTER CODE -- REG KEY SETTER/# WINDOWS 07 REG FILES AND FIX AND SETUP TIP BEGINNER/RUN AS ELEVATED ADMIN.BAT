@ECHO OFF
REM accolade 22 November 2011 20:38:16
REM Here's a one-liner that doesn't require additional tools:
REM ----
REM How to determine in batch-file, if user is administrator - Microsoft - Windows 2000/NT
REM http://www.tomshardware.co.uk/forum/169421-46-determine-batch-file-user-administrator
REM ----

REM >nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"&&(

REM     echo admin...

REM )

REM source: http://stackoverflow.com/q/4054937


REM Here is how I use it, in a routine that elevates a batch script itself:
REM :: Ensure ADMIN Privileges

REM     :: adaptation of <a href="https://sites.google.com/site/eneerge/home/BatchGotAdmin" rel="nofollow" target="_blank">https://sites.google.com/site/eneerge/home/BatchGotAdmi...</a> and <a href="http://stackoverflow.com/q/4054937" rel="nofollow" target="_blank">http://stackoverflow.com/q/4054937</a>

@echo off

REM :: Check for ADMIN Privileges

>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

if '%errorlevel%' NEQ '0' (

	REM Get ADMIN Privileges

	echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"

	echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"

	"%temp%\getadmin.vbs"

	del "%temp%\getadmin.vbs"

	exit /B

) else (

	REM Got ADMIN Privileges
	
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
	
	pushd "%cd%"

	cd /d "%~dp0"

	@echo on

	CMD /C "E:\REG KEY SETTINGS\# Windows 7 REG SCRIPT\# TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT"
)
	