@REM ----------------------------------------------------------------
@REM __ C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB.BAT
@REM __ 
@REM __ BAT 45-SCRIPT RUN GITHUB.BAT
@REM __ 
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ FROM ---- Fri 05-Oct-2018 11:18:00 Created Time
@REM __ TO   ---- Fri 07-Oct-2018 04:45:37 Last Editor #1 1st Bash
@REM __ TO   ---- Fri 19-Oct-2018 19:58:00 Last Editor #2 Current
@REM ----------------------------------------------------------------

@REM -- SESSION COUNT NUMEROUS
@REM -- NUMBER BEGINNER GUESS CALLED 002
@REM -- EARLIER DATE FOUND QUICKER

@REM ----------------------------------------------------------------
@REM -- 002 SESSION
@REM ----------------------------------------------------------------
@REM __ FROM  Sun 07-Oct-2018 04:40:00
@REM __ TO    Sun 07-Oct-2018 04:52:20
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- 003 SESSION 
@REM -- COP CODE OVER AND INTRODUCE TO A BROTHER FRIEND TWIN FOR 
@REM -- RUN GOODSYNC COMMAND LINE
@REM -- LITTLE WORK DONE NONE CHANGES EXCEPT BROTHER ONE
@REM ----------------------------------------------------------------
@REM __ FROM  Fri 19-Oct-2018 15:05:42
@REM __ TO    Fri 19-Oct-2018 17:20:00
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- 004 SESSION 
@REM -- ADD CODE _ TIMEOUT WITH MINIMIZE
@REM -- ADD CODE _ PUT A LEADING PAD NAUGHT IN DATE FORMAT WHEN HOUR 
@REM -- SINGLE DIGIT FOR COMMIT LABEL
@REM -- ADD CODE _ TOOK SOME WORK - A NEW IMPROVEMENT TO DATE LABEL 
@REM -- WITH COMMIT HAVE WEEKDAY
@REM ----------------------------------------------------------------
@REM __ FROM  Fri 19-Oct-2018 19:06:19
@REM __ TO    Fri 19-Oct-2018 19:58:00
@REM ----------------------------------------------------------------


@REM ----------------------------------------------------------------
@REM -- Located in Sync Location 1.. 2..
@REM ----------------------------------------------------------------
@REM -- 1..
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 45-SCRIPT RUN GITHUB.BAT
@REM ----------------------------------------------------------------
@REM -- 2..
@REM -- C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB.BAT
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
@ECHO ---------------------------------------------------------------

@ECHO OFF

SET TITLE_VAR=GITHUB_RUNNNER__GITHUB_SYNC

TITLE %TITLE_VAR%

SET PATH_GIT="C:\Program Files\Git\bin"
 
CD /D "C:\SCRIPTOR GITHUB"

"C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk" /QUICK_INTRO_DUMMY_RUN

%PATH_GIT%\git init
%PATH_GIT%\git add .

@SET Month_Word=%DATE:~3,2%
@SET map_1=01-Jan;02-Feb;03-Mar;04-Apr;05-May;06-Jun;07-Jul;08-Aug;09-Sep;10-Oct;11-Nov;12-Dec
@CALL SET Month_Word=%%map_1:*%Month_Word%-=%%
@SET Month_Word=%Month_Word:;=&rem.%

REM -----------------------------------------------------------------
REM ---- HERE FOUND BUGGY CODE WITH MY FIX __ ADD AN __ IF __ GTR
REM ---- DON'T INCLUDE DASH IN VARIABLE FOR WEEK LIKE NOT WEEK_DAY
REM ---- DON'T INCLUDE DASH IN MAP VARIABLE EITHER
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

%PATH_GIT%\git commit -m "%DATE:~6,4%-%Month_Word%-%DATE:~0,2% -- %Time_Pad_Single_Digit_Hourer:~0,8% -- %Week_Day%"
%PATH_GIT%\git config user.email "matt.lan@btinternet.coml"
%PATH_GIT%\git config user.name "Matthew-Lancaster"
%PATH_GIT%\git pull https://Matthew-Lancaster:KAFMDUZf84@github.com//matthew-lancaster/matthew-lancaster master
%PATH_GIT%\git push -u origin master
@ECHO.
@ECHO ---------------------------------------------------------------

TIMEOUT /T 8

@ECHO ---------------------------------------------------------------
ECHO RUN COMMAND 
ECHO BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk /%TITLE_VAR%
@ECHO ---------------------------------------------------------------

"C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk" %TITLE_VAR%

TIMEOUT /T 120

EXIT

:ENDE
EXIT
