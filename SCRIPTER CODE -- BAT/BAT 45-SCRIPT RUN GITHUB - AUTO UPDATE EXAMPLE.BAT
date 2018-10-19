@REM ----------------------------------------------------------------
@REM __ C:\SCRIPTOR\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AUTO UPDATE EXAMPLE.BAT
@REM __ 
@REM __ BAT 45-SCRIPT RUN GITHUB - AUTO UPDATE EXAMPLE.BAT
@REM __ 
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM __ FROM ---- Fri 05-Oct-2018 12:52:00 Created Time
@REM __ TO   ---- Sun 07-Oct-2018 04:48:00 Last Editor #1
@REM __ TO   ---- Fri 19-Oct-2018 17:30:00 Last Editor #2
@REM ----------------------------------------------------------------
@REM MY FIRST SET UP OF GITHUB 
@REM HAS TAKEN ALL-NIGHT NICE TO WAKE WITH A BRIGHT SUNNY DAY
@REM
@REM HERE THIS CODE PURELY AN EXAMPLE TO SHOW HOW YOU CAN PUT YOUR USERNAME AND PASSWORD IN TO 
@REM AUTO-UPDATER
@REM I READ IN THE ON-LINE HELP PAGES 
@REM THAT THIS METHOD DOES NOT STORE YOUR DETAILS OF PASSWORD IN ANY OF THE CONFIGURE FILE SET
@REM MANY PEOPLE WERE GOING ABOUT USING CLONE AS THE METHOD
@REM BUT SOMEBODY SAID WAS BETTER TO HAVE IT WITH A PULL COMMAND
@REM FETCH 
@REM I INCLUDE TIME NOT IN LOUSY FORMAT SO EACH UPDATED EVENT HAS A NICE TIME PRINT NEXT TO HER
@REM ----------------------------------------------------------------
@REM -- 001 SESSION
@REM ----------------------------------------------------------------
@REM __ FROM  Fri 05-Oct-2018 01:37:29
@REM __ TO    Fri 05-Oct-2018 13:04:00 __ 12 HOUR+- WHOLE THING OF GITHUB 
@REM 									  GET GOING BEGINNER
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- 002 SESSION
@REM ----------------------------------------------------------------
@REM __ FROM  Sun 07-Oct-2018 04:40:00
@REM __ TO    Sun 07-Oct-2018 04:52:20
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- 003 SESSION
@REM ----------------------------------------------------------------
@REM __ FROM  Fri 19-Oct-2018 15:05:42
@REM __ TO    Fri 19-Oct-2018 17:30:00
@REM ----------------------------------------------------------------


@REM ----------------------------------------------------------------
@REM -- Located in Sync Location 1.. 2..
@REM ----------------------------------------------------------------
@REM -- 1..
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 45-SCRIPT RUN GITHUB - AUTO UPDATE EXAMPLE.BAT RUNNER.BAT
@REM ----------------------------------------------------------------
@REM -- 2..
@REM -- C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AUTO UPDATE EXAMPLE.BAT
@REM ----------------------------------------------------------------

ECHO DON'T RUN THIS ONE NOT FOR MY USE JUST AN EXAMPLE ONE TO GIVE AWAY
PAUSE
GOTO ENDER
EXIT 

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

@ECHO ON

TITLE GIT_RUNNNER__GIT_SYNC

@SET PATH_GIT="C:\Program Files\Git\bin"

CD /D "C:\SCRIPTOR GITHUB"

%PATH_GIT%\git init
%PATH_GIT%\git add .

@SET Month_Word=%DATE:~3,2%
@SET map_1=01-Jan;02-Feb;03-Mar;04-Apr;05-May;06-Jun;07-Jul;08-Aug;09-Sep;10-Oct;11-Nov;12-Dec
@CALL SET Month_Word=%%map_1:*%Month_Word%-=%%
@SET Month_Word=%Month_Word:;=&rem.%

REM -----------------------------------------------------------------
REM ---- HERE FOUND BUGGY CODE WITH MY FIX __ ADD AN __ IF __ GTR
REM ---- DON'T INCLUDE DASH IN VARIABLE FOR WEEKD LIKE NOT WEEK_DAY
REM ---- DON'T INCLUDE DASH IN MAP VAIABLE EITHER
REM ----
REM scripting - Setting a windows batch file variable to the day of the week - Stack Overflow
REM https://stackoverflow.com/questions/11364147/setting-a-windows-batch-file-variable-to-the-day-of-the-week
REM -----------------------------------------------------------------
FOR /F "skip=1" %%A IN ('WMIC Path Win32_LocalTime Get DayOfWeek' ) DO ( IF %%A GTR 0 set DOW=0%%A) 
@SET Week_Day=%DOW%
@SET map_2=01-Mon;02-Tue;03-Wed;04-Thu;05-Fri;06-Sat;07-Sun
@CALL SET Week_Day=%%map_2:*%Week_Day%-=%%
@SET Week_Day=%Week_Day:;=&rem.%

REM EQU - equal
REM NEQ - not equal
REM LSS - less than
REM LEQ - less than or equal
REM GTR - greater than
REM GEQ - greater than or equal

@REM ----
@REM windows - Need leading zero for batch script using %time% variable - Server Fault
@REM https://serverfault.com/questions/147515/need-leading-zero-for-batch-script-using-time-variable
@REM ----
@SET Time_Pad_Single_Digit_Hourer=%TIME: =0%

@ECHO %DATE:~6,4%-%Month_Word%-%DATE:~0,2% -- %Time_Pad_Single_Digit_Hourer:~0,8% %Week_Day%

%PATH_GIT%\git commit -m "%DATE:~6,4%-%Month_Word%-%DATE:~0,2% -- %Time_Pad_Single_Digit_Hourer:~0,8% %Week_Day%"
%PATH_GIT%\git config user.email "matt.lan@btinternet.coml"
%PATH_GIT%\git config user.name "Matthew-Lancaster"
%PATH_GIT%\git pull https://<USERNAME:PASSWORD@github.com//matthew-lancaster/matthew-lancaster master
%PATH_GIT%\git push -u origin master

TIMEOUT /T 50

:ENDE
:ENDER
EXIT

@REM ----------------------------------------------------------------
@REM 
@REM THE FIRST PROBLEM I HAD TO GITHUB
@REM WAS MAKING A REPOSITORY ON THE WEB
@REM AND NOT CLONING IT TO THE DESKTOP
@REM WOULD OF MADE EASIER
@REM TRAWL ALONG THE DOCUMENTATION AGAIN WHEN ALL ELSE FAILED

@REM AND THEN FINE TUNING OF THIS FILE
@REM AND THAT WAS IT

@REM AND SYNC IS DONE FROM MY SCRIPT FOLDER VIA GOODSYNC

@REM A WELL WORTH WHILE ALLNIGHTER IT WAS TO I BEEN DREAMER OF THIS ONE FOR A LONG TIME
@REM 
@REM ----------------------------------------------------------------

