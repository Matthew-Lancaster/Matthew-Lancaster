@CD\

@REM  ---------------------------------------------------------------
ver | find "10.0." > nul
@IF %errorlevel% EQU 0 @SET WIN10__=YES
ver | find "6.1." > nul
@IF %errorlevel% EQU 0 @SET WIN07__=YES
ver | find "5.1." > nul
@IF %errorlevel% EQU 0 @SET WINXP__=YES
@REM  ---------------------------------------------------------------

@REM  ---------------------------------------------------------------
@IF "%WINXP__%"=="YES" @ECHO WINDOWS XP
@IF "%WIN7___%"=="YES" @ECHO WINDOWS 7
@IF "%WIN10__%"=="YES" @ECHO WINDOWS 10
@REM  ---------------------------------------------------------------

@IF "%WINXP__%"=="YES"  GOTO VERY_ENDE
@IF "%WIN7___%"=="YES"  GOTO VERY_ENDE
REM @IF "%COMPUTERNAME%" == "7-ASUS-GL522VW" GOTO VERY_ENDE

@REM setlocal ENABLEEXTENSIONS

@REM -- [ Thursday 22:41:50 Pm_07 December 2017 ] The UK
@REM -- [ Thursday 22:41:50 Pm_07 December 2017 ] GMT UTC+0 WET
@REM -- --------------------
@REM -- COMMAND SET TO LEARN 
@REM -- --------------------
@REM -- 1.GET THE PATH OF CURRENT WORKING DIRECTORY
@REM -- 2.USE OF TYPE COMMAND CONVERT TEXT TO ASCII FROM UNICODE OUTPUT OF SOME COMMAND
@REM -- 3.USE TOKEN GET THE FIRST WORD OF A LINE TO CHECK WITHER IF STATEMENT
@REM -- 4.GET THE TASK SCHEDULER WORKING AND /QUERY FOR REQUIREMENT NOT SET TWO
@REM -- 5.SET THE TASK SCHEDULER WITH THE JOB OF THE NAME AND PATH OF THIS FILE
@REM -- 5.SET THE TASK SCHEDULER AS ADMINISTRATOR SOMETIME CAN STILL SEE JOB RUN IN VIEW IF NOT QUICKER
@REM -- 5.SET THE TASK SCHEDULER AS ADMINISTRATOR WITHOUT REQUIRE PASSWORD TO BE SETTER NOT VERY REVEALING DOCUMENTATION

@REM -- NET STOP WINDOWS UPDATE SERVICE AND DISABLE
NET STOP WUAUSERV
SC CONFIG WUAUSERV START= DISABLED

REM schtasks /DELETE /TN "z_KEEP_WINDOWS_UPDATE_STOPPED"

@REM schtasks /QUERY /TN "z_KEEP_WINDOWS_UPDATE_STOPPED"
@REM RETURN STRING WANTED "ERROR: The system cannot find the file specified."

REM schtasks /QUERY /TN "z_KEEP_WINDOWS_UPDATE_STOPPED" >NUL 2>"%~dp0SCHTasks_Output.TXT"
REM @FOR /F "tokens=1 delims= " %%A IN ('TYPE "%~dp0SCHTasks_Output.TXT"') DO @IF "%%A"=="ERROR:" SET Var_01=NOT_EXIST_SCHEDULE
REM DEL "%~dp0SCHTasks_Output.TXT"

REM @if "%Var_01%"=="" GOTO NOT_HERE
REM schtasks /CREATE /RU ADMINISTRATORS /TN ""z_KEEP_WINDOWS_UPDATE_STOPPED"" /TR "%~f0" /SC MINUTE /MO 20 /RL HIGHEST /F
REM :NOT_HERE


CMD /C CALL "%~dp0BAT 11-RESET WINDOWS UPDATE VER 10.BAT" /Q
CMD /C CALL "%~dp0BAT 14-Delete $Windows.~BT & $Windows.~WS Folders.BAT" /Q


@REM -- SOME COMMAND OUTPUT IN UNICODE FOR TEXT DETECTION HARDER 
@REM -- USE TYPE COMMAND TO CONVERT TO TEXT
@REM -- INFO SOURCE TO MAKE WORK
@REM ----
@REM Redirect output of batch command to an if statement (and string error?) - Stack Overflow
@REM https://stackoverflow.com/questions/8099411/redirect-output-of-batch-command-to-an-if-statement-and-string-error
@REM ----

@IF "%1"=="" (
	PAUSE
	REM TIMEOUT /T 40
)

:VERY_ENDE
EXIT /B
EXIT