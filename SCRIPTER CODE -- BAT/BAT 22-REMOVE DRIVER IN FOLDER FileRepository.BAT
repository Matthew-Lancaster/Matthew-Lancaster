@ECHO off
REM __ Sat 14 October 2017 22:21:21-------
REM __ -----------------------------------
REM __ FIND IN THE PATH TO EXECUTE PNPUTIL
REM __ -----------------------------------
REM __ THIS IS NOT SOMETHING CCLEANER HELP WITH TAKEN UP SPACE CLEARING
REM __ ----------------------------------------------------------------
REM __ MANY OF THE DRIVER ARE REPEAT VIDEO CARD TAKING UP SPACE
REM __ ----------------------------------------------------------------

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

CD\
ECHO.

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


for /f "delims=" %%a in ('dir C:\Windows\winsxs\PNPUTIL.EXE /S /B') do (
 REM call set concat=%%concat%%%%a
 call set concat=%%a
)

echo %concat%

%concat% -E


ECHO.
ECHO AFTER PAUSE THE ABOVE DRIVERS WILL BE DELETER ARE YOU SURE
ECHO.

PAUSE

ECHO ON
@FOR /L %%i IN (1,1,140) DO @( 
    ECHO %concat% /d oem%%i.inf 
	REM %concat% /d oem%%i.inf 
)

ECHO OFF

REM __ MY COMPUTER HAS HAS UPTO oem104.inf WITH GAPPING

REM __-------------------------------------------------------------
REM __ PnPUtil.exe is in C:\Windows\System32 but thats just a virtual link
REM C:\Windows\system32\pnputil /e
REM C:\Windows\winsxs\amd64_microsoft-windows-pnputil_31bf3856ad364e35_10.0.15063.0_none_0e779bcaf5563fd6\pnputil.exe /e
REM __-------------------------------------------------------------

REM __ DOS HELP /?
REM __ -e ==> /enum-drivers
REM __ PNPUTIL.EXE -E

REM __ C:\Windows\System32\DriverStore\FileRepository
REM __ 15.6 GB TAKEN UP __ OVER SIZE REQUIRE DELETE SOME OR ALL ONLY BACKUP BUT SOME ARE IN USE
REM __ 1.68 GB (AFTER THE RUN)

REM __ DELETE WON'T DO ANY HARM AND IS OPTION TO DO BY FORCE WHICH I WONT REQUIRE
REM __ DELETE BY IN THE FOLDER MIGHT DO SYSTEM CRASH AS TALKER

:VERY_ENDER_BLOCK

PAUSE

:ENDE


REM ----
REM TAKING UP A LOT OF SPACE C:/Windows/System32/DriverStore/FileRepository/ - Google Search
REM https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB744GB744&q=TAKING+UP+A+LOT+OF+SPACE+C:/Windows/System32/DriverStore/FileRepository/&spell=1&sa=X&ved=0ahUKEwj3zunn8vDWAhVGSRoKHYnaBMYQvwUIJSgA&biw=1024&bih=434
REM ----
REM ----
REM BATCH COUNT TO 100 - Google Search
REM https://www.google.co.uk/search?q=BATCH+COUNT+TO+100&rlz=1C1CHBD_en-GBGB744GB744&oq=BATCH+COUNT+TO+100&aqs=chrome..69i57.8272j0j7&sourceid=chrome&ie=UTF-8
REM ----
REM ----
REM batch - For Loop counting from 1 to n in a windows bat script - Server Fault
REM https://serverfault.com/questions/59014/for-loop-counting-from-1-to-n-in-a-windows-bat-script
REM ----
REM ------------------------------
REM The 1,1,104 is decoded as:
REM (start,step,end)
REM ------------------------------
REM __ ALTERNATIVELY
REM for(int i = 0; i < 100; i++) {
   REM // do something
REM }

REM ----
REM Pnputil
REM https://technet.microsoft.com/en-us/library/ff800798(v=ws.11).aspx
REM --------
REM PNPUTIL.EXE RUN FROM BATCH FILE - Google Search
REM https://www.google.co.uk/search?q=PNPUTIL.EXE+RUN+FROM+BATCH+FILE&rlz=1C1CHBF_en-GBGB759GB759&oq=PNPUTIL.EXE+RUN+FROM+BATCH+FILE&aqs=chrome..69i57.8765j0j7&sourceid=chrome&ie=UTF-8
REM --------
REM Calling PNPUTIL.exe in a batch file with a Software Package | Symantec Connect Community
REM https://www.symantec.com/connect/forums/calling-pnputilexe-batch-file-software-package
REM ----
REM It seems like the Altiris Task agent has issues finding programs that are storded 
REM in the WinSXS store. PnPUtil.exe is in C:\Windows\System32 but thats just a virtual link 
REM to C:\Windows\winsxs\amd64_microsoft-windows-pnputil_31bf3856ad364e35_6.1.7600.16385_none_5958b438d6388d15.
REM If you use that path, the command runs correctly.
----
REM DIR AND OUTPUT TO VARIABLE - Google Search
REM https://www.google.co.uk/search?q=DIR+AND+OUTPUT+TO+VARIABLE&rlz=1C1CHBF_en-GBGB759GB759&oq=DIR+AND+OUTPUT+TO+VARIABLE&aqs=chrome..69i57j0l5.5371j0j7&sourceid=chrome&ie=UTF-8
REM --------
REM How can I save output of dir in a variable?
REM https://www.computerhope.com/forum/index.php?topic=65683.0
REM ----