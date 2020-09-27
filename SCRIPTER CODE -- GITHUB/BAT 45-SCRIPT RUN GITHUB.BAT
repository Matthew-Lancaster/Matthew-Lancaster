@ECHO OFF
REM ----------------------------------------------------------------
REM __ C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB.BAT
REM __ 
REM __ BAT 45-SCRIPT RUN GITHUB.BAT
REM __ 
REM __ MATTHEW LANCASTER 
REM __ MATT.LAN@BTINTERNET.COM
REM ----------------------------------------------------------------
REM __ FROM ---- Fri 05-Oct-2018 11:18:00 Created Time
REM __ TO   ---- Fri 07-Oct-2018 04:45:37 Last Editor #1 1st Bash
REM __ TO   ---- Fri 19-Oct-2018 19:58:00 Last Editor #2 Current
REM ----------------------------------------------------------------

REM -- SESSION COUNT NUMEROUS
REM -- NUMBER BEGINNER GUESS CALLED 002
REM -- EARLIER DATE FOUND QUICKER

REM # ------------------------------------------------------------------
REM  Location Internet
REM --------------------------------------------------------------------
REM  ----
REM  BAT 45-SCRIPT RUN GITHUB.BAT
REM  ----
REM Matthew-Lancaster/BAT 45-SCRIPT RUN GITHUB.BAT at master · Matthew-Lancaster/Matthew-Lancaster
REM https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20GITHUB/BAT%2045-SCRIPT%20RUN%20GITHUB.BAT
REM  ----
REM # ------------------------------------------------------------------

REM ----------------------------------------------------------------
REM -- 002 SESSION
REM ----------------------------------------------------------------
REM __ FROM  Sun 07-Oct-2018 04:40:00
REM __ TO    Sun 07-Oct-2018 04:52:20
REM ----------------------------------------------------------------

REM ----------------------------------------------------------------
REM -- 003 SESSION 
REM -- COP CODE OVER AND INTRODUCE TO A BROTHER FRIEND TWIN FOR 
REM -- RUN GOODSYNC COMMAND LINE
REM -- LITTLE WORK DONE NONE CHANGES EXCEPT BROTHER ONE
REM ----------------------------------------------------------------
REM __ FROM  Fri 19-Oct-2018 15:05:42
REM __ TO    Fri 19-Oct-2018 17:20:00
REM ----------------------------------------------------------------

REM ----------------------------------------------------------------
REM -- 004 SESSION 
REM -- ADD CODE _ TIMEOUT WITH MINIMIZE
REM -- ADD CODE _ PUT A LEADING PAD NAUGHT IN DATE FORMAT WHEN HOUR 
REM -- SINGLE DIGIT FOR COMMIT LABEL
REM -- ADD CODE _ TOOK SOME WORK - A NEW IMPROVEMENT TO DATE LABEL 
REM -- WITH COMMIT HAVE WEEKDAY
REM ----------------------------------------------------------------
REM __ FROM  Fri 19-Oct-2018 19:06:19
REM __ TO    Fri 19-Oct-2018 19:58:00
REM ----------------------------------------------------------------

REM ----------------------------------------------------------------
REM -- 005 SESSION 
REM -- ADD CODE NEW ON
REM -- THIS ARE THE BATCH I USER TO RUN GITHUB COMMIT UPLOAD CHANGES
REM -- IT TIME TO HANDLE IT A DIFFERENT WAY
REM -- NOW ABLE TO PUBLISHER ON-LINE
REM ----------------------------------------------------------------
REM __ FROM  Sun 10-Mar-2019 19:26:15
REM __ TO    Sun 10-Mar-2019 20:10:00
REM ----------------------------------------------------------------

REM ----------------------------------------------------------------
REM -- SIX SESSION 
REM -- DATE FORMATTER IS READ FORM VBSCRIPT NOW AND ARGUMENT PASS 
REM -- ON COMMAND LINE PIPE
REM -- SEEM GETTING DAY OF WEEK IN WORD IS ONLY SOMETHING 
REM -- MS WINDOWS XP USED TO HAVE AND UNTIL THINGS CHANGED
REM -- NOT MUCH DOCUMENTATION AROUND FOR CORRECTOR THE ERROR
REM ----------------------------------------------------------------
REM __ FROM  Mon 18-Mar-2019 14:03:47
REM __ TO    Mon 18-Mar-2019 15:28:00 -- DURING ALL-NIGHTER SUN TO MON
REM ----------------------------------------------------------------

REM ----------------------------------------------------------------
REM -- Located in Sync Location 1.. 2..
REM ----------------------------------------------------------------
REM -- 1..
REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 45-SCRIPT RUN GITHUB.BAT
REM ----------------------------------------------------------------
REM -- 2..
REM -- C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB.BAT
REM ----------------------------------------------------------------

REM ----------------------------------------------------------------
REM __ SET ADMIN A SAFE DOUBLE CHECKER WAY NOT A LOOPY ONE
REM __ IF CAN'T GET IT TO GO ADMIN DOESN'T RSI INJURY ITSELF AND YOU
REM ----------------------------------------------------------------

REM -----------------------------------------------------------------
REM THE CODE
REM ON THE HOUR 
REM 1..
REM C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB.exe
REM THIS CODE WILL EITHER RUN -- 
REM 2..
REM GITHUB - ONLY 
REM WHICH IS HERE RUN
REM BAT 45-SCRIPT RUN GITHUB.BAT
REM ..
REM OR WILL RUN 
REM 3..
REM GOODSYNC SYNC SET AND THEN GITHUB
REM 4..
REM ON THE HOUR IT WILL BE BOTH
REM ..
REM IF IT RUN GOODSYNC SET JOB IT WILL RUN HERE ITSELF ME 
REM BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT
REM 5..
REM -----------------------------------------------------------------
REM ON THE HOUR 
REM 1..
REM C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB.exe
REM IS RUN BY HERE
REM Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk
REM SUBROUTINE 
REM GITHUB_MIDNIGHT_AND_MIDDAY_TIMER:
REM THIS WILL NOT ACTOR IF GOODSYNC2GO IS RUN BY THE D DRIVE PROCESS
REM -----------------------------------------------------------------
REM Mon 02-Sep-2019 11:45:40
REM -----------------------------------------------------------------



@ECHO OFF

ECHO %* | FIND "/ADMIN" >NUL
IF %errorlevel% EQU 0 SET ADMIN_REQUEST=/ADMIN

IF '%ADMIN_REQUEST%'=='/ADMIN' (
	REM :: Check for ADMIN Privileges Again
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

REM :: Check for ADMIN Privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
@IF '%errorlevel%' GTR '2' (
	REM Get ADMIN Privileges
	@ECHO Set UAC = CreateObject^("Shell.Application"^) >> "%temp%\getadmin.vbs"
	@ECHO UAC.ShellExecute "%~s0", "/ADMIN %*", "", "runas", 1 >> "%temp%\getadmin.vbs"

	"%temp%\getadmin.vbs"
	@DEL "%temp%\getadmin.vbs"
	GOTO ENDE
	@EXIT /B
)

IF '%ADMIN_REQUEST%'=='/ADMIN'     @ECHO Got ADMIN Privileges -- By Push For Them
IF NOT '%ADMIN_REQUEST%'=='/ADMIN' @ECHO Got ADMIN Privileges -- By Natural

@ECHO ---------------------------------------------------------------

@ECHO OFF

REM SET TITLE_VAR=GITHUB_RUNNNER__GITHUB_SYNC

SETLOCAL EnableDelayedExpansion
REM ---- FILENAME WITHOUT EXTENSION
@SET TITLE_VAR=%~n0
REM ---- FILENAME WITH   EXTENSION
@SET TITLE_VAR=%~n0%~x0
REM ---- FULL SCRIPT PATH
@SET TITLE_VAR=%~f0
SETLOCAL DisableDelayedExpansion

@TITLE %TITLE_VAR%

"C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk" %~n0
REM %TITLE_VAR%

SETLOCAL EnableDelayedExpansion

REM ----
REM How can I echo a newline in a batch file? - Stack Overflow
REM https://stackoverflow.com/questions/132799/how-can-i-echo-a-newline-in-a-batch-file
REM ----
(set \n=^
%=Do not remove this line=%
)

SETLOCAL DisableDelayedExpansion

SET PATH_GIT="C:\Program Files\Git\bin"
 
CD /D "C:\SCRIPTOR GITHUB"

"C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk" /QUICK_INTRO_DUMMY_RUN


SET FILENAME_VBS="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 41-GET DATE AND TIME FORMATTER.VBS"
for /f "tokens=*" %%a in ('cscript //nologo %FILENAME_VBS%') do (set "strVBScriptResults=%%a")

SET DATE_TIME_FORMATER=%strVBScriptResults%

ECHO %DATE_TIME_FORMATER%

REM -----------------------------------------------------------------
REM -- THE FILENAME HAS SPACE IN CHARACTER SO HERE 
REM -- CONVERT THE TEXT FILENAME 
REM -- TO DOS ALTERNATIVE SHORT-NAME 8x3
REM -- EXAMPLE OUTPUT
REM -----------------------------------------------------------------
REM -- C:\RF\7-ASUS-GL522VW\SAFE NOTE\GITHUB_NUMBER.txt
REM -- C:\RF\7-ASUS~1\SAFENO~1\GITHUB~1.TXT
REM -----------------------------------------------------------------
for %%I in ("C:\RF\7-ASUS-GL522VW\SAFE NOTE\GITHUB_NUMBER.txt") do set GITHUB_NUMBER_PATH=%%~sI

REM -----------------------------------------------------------------
REM -- GET CONTENT OF FILE AND PUT IN VARIABLE
for /f "delims=" %%x in (%GITHUB_NUMBER_PATH%) do set GITHUB_NUMBER=%%x
REM -----------------------------------------------------------------

%PATH_GIT%\git init
%PATH_GIT%\git add .
%PATH_GIT%\git commit -m "%DATE_TIME_FORMATER%"
%PATH_GIT%\git config user.email "matt.lan@btinternet.com"
%PATH_GIT%\git config user.name "Matthew-Lancaster"
%PATH_GIT%\git pull https://Matthew-Lancaster:%GITHUB_NUMBER%@github.com//matthew-lancaster/matthew-lancaster master
%PATH_GIT%\git push -u origin master
@ECHO.
@ECHO ---------------------------------------------------------------

TIMEOUT /T 8

REM @ECHO ---------------------------------------------------------------
REM ECHO RUN COMMAND 
REM ECHO BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk /%TITLE_VAR%
REM @ECHO ---------------------------------------------------------------

"C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AHK_MINIMIZE.ahk" %TITLE_VAR% %*

TIMEOUT /T 120

:ENDE
EXIT












REM -----------------------------------------------------------------
REM -----------------------------------------------------------------

REM -----------------------------------------------------------------
REM -----------------------------------------------------------------

REM -----------------------------------------------------------------
REM -----------------------------------------------------------------

REM -----------------------------------------------------------------
REM -----------------------------------------------------------------
REM REFERENCE SET THING USED IN PAST
REM 
REM -----------------------------------------------------------------
REM [ Monday 15:12:10 Pm_18 March 2019 ]
REM -----------------------------------------------------------------

REM SETLOCAL DisableDelayedExpansion

REM SETLOCAL EnableDelayedExpansion
REM FOR /F "skip=1" %%A IN ('WMIC OS GET LOCALDATETIME') DO (SET "t=%%A" & GOTO NEXT)
REM :NEXT
REM SET "s=%t:~12,2%"
REM SET "m=%t:~10,2%"
REM SET "h=%t:~8,2%"
REM SET "h_24=%t:~8,2%"
REM SET "d=%t:~6,2%"
REM SET "z=%t:~4,2%"
REM SET "y=%t:~0,4%"
REM IF !h! GTR 11 (SET /A "h-=12" & SET "ap=P" & IF "!h!"=="0" (SET "h=00") ELSE (IF !h! LEQ 9 (SET "h=0!h!"))) ELSE (SET "ap=A")
REM ----
REM batch file - Windows CMD - How to create time variable with AM PM - Stack Overflow
REM https://stackoverflow.com/questions/48750465/windows-cmd-how-to-create-time-variable-with-am-pm
REM ----

REM ----
REM batch - How to put DayOfWeek value into file name? - Server Fault
REM https://serverfault.com/questions/199810/how-to-put-dayofweek-value-into-file-name
REM ----
REM -- NOT WORKING ANY MORE MAYBE USED TO ONCE A TIMER
REM -----------------------------------------------------------------
REM -- USER VBSCRIPT INSTEAD
REM -- [ Monday 14:17:50 Pm_18 March 2019 ]
REM -----------------------------------------------------------------
REM SETLOCAL enabledelayedexpansion
REM SET /a DOW=0
REM FOR /F "skip=1" %%A IN ('WMIC Path Win32_LocalTime Get DayOfWeek' ) DO (
	REM IF !DOW! GTR 0 GOTO NEXT
    REM SET /a DOW+=1
REM ) 
REM :NEXT
REM ECHO %DOW%
REM FOR /F "skip=1 tokens=1-6" %%G IN ('WMIC Path Win32_LocalTime Get DayOfWeek^ ,Day^,Hour^,Minute^,Month^,Second^,Year /Format:table') DO (
   REM IF "%%~N"=="" goto s_done
      REM Set _yyyy=%%N
      REM Set _mm=00%%M
      REM Set _dd=00%%L
      REM Set _hour=00%%K
      REM SET _minute=00%%J
      REM SET _DOW=00%%I
REM )
REM :s_done
REM ECHO %_yyyy%
REM ECHO %_mm%
REM ECHO %_dd%
REM ECHO %_hour%
REM ECHO %_minute%
REM ECHO %_DOW%
REM SETLOCAL DisableDelayedExpansion

REM ECHO %DOW%

REM REM @SET MAP_2=01-SUN;02-MON;03-TUE;04-WED;05-THU;06-FRI;07-SAT
REM @SET MAP_2=01-Mon;02-Tue;03-Wed;04-Thu;05-Fri;06-Sat;07-Sun
REM @SET Week_Day=%DOW%
REM @CALL SET Week_Day=%%MAP_2:*%Week_Day%-=%%
REM @SET Week_Day=%Week_Day:;=&rem.%

REM ECHO %Week_Day%
REM PAUSE

REM @SET Month_Word=%z%
REM @SET map_1=01-Jan;02-Feb;03-Mar;04-Apr;05-May;06-Jun;07-Jul;08-Aug;09-Sep;10-Oct;11-Nov;12-Dec
REM @CALL SET Month_Word=%%map_1:*%Month_Word%-=%%
REM @SET Month_Word=%Month_Word:;=&rem.%


REM -----------------------------------------------------------------
REM -- ABOVE CODE
REM -- HERE FOUND BUGGY CODE WITH MY FIX __ ADD AN __ IF __ GTR
REM -- DON'T INCLUDE DASH IN VARIABLE FOR WEEK LIKE NOT WEEK_DAY
REM -- DON'T INCLUDE DASH IN MAP VARIABLE EITHER
REM ----
REM scripting - Setting a windows batch file variable to the day of the week - Stack Overflow
REM https://stackoverflow.com/questions/11364147/setting-a-windows-batch-file-variable-to-the-day-of-the-week
REM ----
REM -- THE DATE DAY OF WEEK WAS SOMETIME REPORT OF ERROR AN EXTRA DAY LIKE SUN ERROR AS MON
REM -- AFTER LOOK WEB REQUIRE -- IF %%A GTR 0 GOTO NEXT
REM -- OTHER PEOPLE WANT -- SETLOCAL enabledelayedexpansion
REM -- SO THEY DO PROCESS THE NUMBER INTO DAY OF WEEK TEXT ALL WITHIN THE FOR LOOP 
REM -- AND USER BLOCK OF IF STATEMENT RATHER THAN ONE LINER
REM -----------------------------------------------------------------

REM EQU - equal
REM NEQ - not equal
REM LSS - less than
REM LEQ - less than or equal
REM GTR - greater than
REM GEQ - greater than or equal

REM ----
REM windows - Need leading zero for batch script using %time% variable - Server Fault
REM https://serverfault.com/questions/147515/need-leading-zero-for-batch-script-using-time-variable
REM ----

REM SET DATE_TIME_FORMATER=%y%-%Month_Word%-%d% -- %h_24%:%m%:%s% - %h%:%m%:%s%-%ap%M - %Week_Day%

REM -----------------------------------------------------------------


REM SET Time_Pad_Single_Digit_Hourer=%TIME: =0%
REM SET FIRST_TIME_VAR_01=%DATE:~6,4%-%Month_Word%-%DATE:~0,2% -- %TIME_PAD_02%
REM SET FIRST_TIME_VAR_02=%FIRST_TIME_VAR_01% %Week_Day%
REM ECHO %FIRST_TIME_VAR_01%
REM ECHO %FIRST_TIME_VAR_02%

REM PAUSE
