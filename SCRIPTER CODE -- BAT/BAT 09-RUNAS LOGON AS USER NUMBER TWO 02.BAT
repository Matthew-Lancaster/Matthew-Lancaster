@CD\
@ECHO ------------------------------------------------------
@TITLE %~f0

REM runas /user:"7-ASUS-GL522VW\MATT 02" "C:\Program Files\Siber Systems\GoodSync\GoodSync-v10.exe"
runas /user:"7-ASUS-GL522VW\MATT 02" "EXPLORER.EXE"
@ECHO OFF

REM runas /user:*7-ASUS-GL522VW\*account "MATT 02" name explorer.exe



REM ----
REM HOW TO LOGON AS ANOTHER USER BATCH COMMAND - Google Search
REM https://www.google.co.uk/search?q=HOW+TO+LOGON+AS+ANOTHER+USER+BATCH+COMMAND&rlz=1C1CHBD_en-GBGB721GB721&oq=HOW+TO+LOGON+AS+ANOTHER+USER+BATCH+COMMAND&aqs=chrome..69i57.8013j0j7&sourceid=chrome&ie=UTF-8
REM --------
REM Switch user accounts from the Command Prompt
REM http://lifehacker.com/290435/switch-user-accounts-from-the-command-prompt
REM ----
REM [ Thursday 18:08:20 Pm_11 May 2017 ] The UK
REM [ Thursday 17:08:20 Pm_11 May 2017 ] BST BDT DST DLS 1 Hr Ahead UTC+1 CET






@ECHO OFF
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (20,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)
@ECHO.
