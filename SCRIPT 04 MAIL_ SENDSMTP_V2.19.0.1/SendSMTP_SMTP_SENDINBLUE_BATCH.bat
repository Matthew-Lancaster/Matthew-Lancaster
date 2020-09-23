@SETLOCAL EnableDelayedExpansion
@ECHO OFF
SET DRIVER=%SystemDrive%
@CD\

REM ----
REM SendinBlue
REM https://account.sendinblue.com/pricing
REM ----

SET NAME_USE=SENDINBLUE

TITLE %~f0 __ %NAME_USE% EMAIL __
@REM ----------------------------------------------------------------

@ECHO ----------------------------------------------------------------
@ECHO %NAME_USE% __ EMAIL
@ECHO ----------------------------------------------------------------
@ECHO OFF

@SET VAR_DATE=%DATE:~6,4%-%DATE:~3,2%-%DATE:~0,2%--%TIME:~0,8%
REM __ ----------------------------------------------------------------
SET WORD_REPLACE_1=" "
SET WORD_REPLACE_2=0
SET WORD_REPLACE_1=%WORD_REPLACE_1:"=%
REM __ Command Quote Mark Stripper " "
CALL SET VAR_DATE=%%VAR_DATE:%WORD_REPLACE_1%=%WORD_REPLACE_2%%%
REM ALSO LIKE
REM set IP=%IP: =%

@ECHO %VAR_DATE%__%COMPUTERNAME%__%IP%
@ECHO ----------------------------------------------------------------

SET MAIL_MERGE1=ELASTICEMAIL

REM SET MAIL_MERGE2=SYSTEM

SET PATH_FILE_R="%DRIVER%\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP_SMTP_%NAME_USE%_*.TXT"

FOR /F "delims=" %%A in ('DIR "%PATH_FILE_R%" /a-d /b /s') DO (SET NAME_BODY="%%~nxA")

REM set %NAME_BODY%=%NAME_BODY: =%

ECHO %NAME_BODY%

SET NAME_BODY=%NAME_BODY:"=%

ECHO %NAME_BODY%
REM ECHO 2228

SET WORD_REPLACE_1=" "
SET WORD_REPLACE_2=0
SET WORD_REPLACE_1=%WORD_REPLACE_1:"=%
REM __ Command Quote Mark Stripper " "
CALL SET %NAME_BODY%=%%NAME_BODY:%WORD_REPLACE_1%=%WORD_REPLACE_2%%%
ECHO %NAME_BODY%

SET BODY_PATH_FILE_1="%DRIVER%\PStart\Progs\SendSMTP_v2.19.0.1\%NAME_BODY%"
ECHO %BODY_PATH_FILE_1%

REM -- SET SEND OUR OWN FILE SOURCE CODE __ SHOW OFF
REM -- IN CASE YOU WERE WONDERING
SET FILE_ATTACH1="%~f0"
SET FILE_ATTACH2="%FILE_ATTACH1:"=%.TXT"

REM ----
REM BATCH COMMAND GET NAME OF BATCH FILE RUNNING - Google Search
REM https://www.google.co.uk/webhp?sourceid=chrome-instant&rlz=1C1CHBD_en-GBGB721GB721&ion=1&espv=2&ie=UTF-8#q=BATCH+COMMAND+GET+NAME+OF+BATCH+FILE+RUNNING&*
REM ----
REM ----
REM windows - Finding out the file name of the running batch file - Stack Overflow
REM http://stackoverflow.com/questions/343518/finding-out-the-file-name-of-the-running-batch-file
REM ----
REM BOTH SAME
ECHO %~f0
@ECHO ----------------------------------------------------------------
REM --------------
REM ECHO "%~dpnx0"
REM --------------
REM pnx
REM --------------
REM PATH
REM NAME
REM EXTENSION
REM --------------

REM PAUSE

ECHO %MAIL_MERGE1%
ECHO %BODY_PATH_FILE_1%
ECHO %IP% 
ECHO %FILE_ATTACH2%

@ECHO ----------------------------------------------------------------
TITLE %~f0 __ %NAME_USE% EMAIL __ %VAR_DATE% -- %COMPUTERNAME% -- %IP%

SET SUBJECT_LINE_PASS_THROUGH=%VAR_DATE%
SET SUBJECT_LINE1="Matthew Lancaster_ _%NAME_USE%_ %VAR_DATE%"

SET HOST=smtp-relay.sendinblue.com
SET PORT=587
SET USER_ID=matt.lan@btinternet.com
SET AUTHENTICATION_METHOD=2
SET EMAIL_RECEIPENT="Matthew Lancaster <matt.lan@btinternet.com>"
SET EMAIL_RECEIPENT="RoidsRim <rub.rim@gmail.com>"

SET DRIVER_ME="%DRIVER%\PStart\Progs\SendSMTP_v2.19.0.1\SENDSMTP.EXE"
SET INI_NAME="%DRIVER%\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP_SMTP_%NAME_USE%.ini"
SET EMAIL_BY_ME="Matthew Lancaster <Matt.Lan@btinternet.com>"

ECHO EMAIL OKAY SENDER __ %MAIL_MERGE1%
ECHO NAME_BODY
ECHO %NAME_BODY%
ECHO BODY_PATH_FILE_1
ECHO %BODY_PATH_FILE_1%
SET PASS3=Dbxym3Sigu0g8lAGbjlQzx6DBmOfk9Jk90gOSNi4funuWWbotztv8xHc0GGwIIvG

START "" /WAIT /HIGH %DRIVER_ME% /I /HOST %HOST% /PORT %PORT% /USERID %USER_ID% /PASS3 %PASS3% /AUTH %AUTHENTICATION_METHOD% /S %INI_NAME% /TO %EMAIL_RECEIPENT% /BODYFILE %BODY_PATH_FILE_1% /SUBJECT %SUBJECT_LINE1% /TYPE PLAIN

rem START "" /WAIT /HIGH %DRIVER_ME% /I /HOST %HOST% /PORT %PORT% /USERID %USER_ID% /PASS3 %PASS3% /AUTH %AUTHENTICATION_METHOD% /S %INI_NAME% /TO %EMAIL_RECEIPENT% /BODYFILE %BODY_PATH_FILE_1% /SUBJECT %SUBJECT_LINE1% /TYPE PLAIN /FILES %FILE_ATTACH2%

GOTO ENDER

:SETVAR
SET /a varcounter=%varcounter% + 1
IF NOT {%1}=={} (
	REM ECHO NIC %varcounter% address is {%1}
	SET NIC%varcounter%=%1
	SET IP=%1 	
)
GOTO :eof

:ENDER

@ECHO OFF

REM -- CARET SYMBOL ARE USED IN FRONT THE ^& SIGN OR NOT DISPLAY THE AMPERSAND IN ECHO
REM ----
REM BATCH COMMAND FILE COUNT TO EXIT - Google Search
REM https://www.google.co.uk/webhp?sourceid=chrome-instant^&rlz=1C1CHBD_en-GBGB721GB721^&ion=1^&espv=2^&ie=UTF-8#q=BATCH+COMMAND+FILE+COUNT+TO+EXIT^&*
REM ----
REM Count down in a batch file - Programming (C^++, Delphi, VB/VBS, CMD/batch, etc.) - MSFN
REM http://www.msfn.org/board/topic/33520-count-down-in-a-batch-file/
REM ----
@ECHO ----------------------------
@ECHO INFORMATION AT LOCATION HERE ----------------------------
@ECHO __ %~f0 __

REM ___________________________
REM @ECHO %~f0 _ FROM SCHEDULER _VIA %SystemDrive%\22_LOGON.BAT ^& BY DELAY_LOADER.EXE VBCODE __ RUN ALL IN BATCH ABLE NOW BUT TIMING OF SYSTEM START MUST WAIT
@ECHO %~dp0IP_LOGGER.TXT
@ECHO %BODY_PATH_FILE_1:"=%
@ECHO %FILE_ATTACH2:"=%
@ECHO %FILE_ATTACH3:"=%
@ECHO -------------

@ECHO OFF

REM PAUSE

@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (10,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)
@ECHO.

EXIT