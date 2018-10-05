@ECHO.
CD\

ECHO ____ TITLE.BAT ____ 
TITLE %~f0

@SETLOCAL ENABLEDELAYEDEXPANSION
SET "PATH_VAR="
rem :: delims is a TAB followed by a space

FOR /F "tokens=3* delims= " %%a in ('reg query "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v personal') do (
 SET PATH_VAR=%%a%%b
)

ECHO %PATH_VAR%

IF NOT EXIST "%PATH_VAR%\Duplicate Cleaner log.txt" GOTO ENDE

TYPE "%PATH_VAR%\Duplicate Cleaner log.txt" >> "%PATH_VAR%\Duplicate Cleaner log -- %COMPUTERNAME%.txt"

type nul >"%PATH_VAR%\Duplicate Cleaner log.txt"

DEL "%PATH_VAR%\Duplicate Cleaner log.txt"

@rem reg query "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v personal


@REM @SETLOCAL ENABLEDELAYEDEXPANSION
@REM SET "PATH_VAR="
@REM rem :: delims is a TAB followed by a space

@REM FOR /F "tokens=3* delims= " %%a in ('reg query "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v personal') do (
@REM  SET PATH_VAR=%%a%%b
@REM )

@REM Variable A contains all data before first space, B - rest part of ValueName (including further spaces), so appdir = ValueName

@REM how to get the path token delim with reg query when path has got spaces - Google Search
@REM https://www.google.co.uk/search?q=how+to+get+the+path+token+delim+with+reg+query+when+path+has+got+spaces&oq=how+to+get+the+path+token+delim+with+reg+query+when+path+has+got+spaces&aqs=chrome..69i57.32112j0j8&sourceid=chrome&ie=UTF-8

@REM How can I get the value of a registry key from within a batch script? - Stack Overflow
@REM https://stackoverflow.com/questions/445167/how-can-i-get-the-value-of-a-registry-key-from-within-a-batch-script


@REM PAUSE

:ENDE