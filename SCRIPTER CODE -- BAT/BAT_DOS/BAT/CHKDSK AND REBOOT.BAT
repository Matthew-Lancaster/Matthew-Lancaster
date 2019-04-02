@REM http://www.wikihow.com/Create-Options-or-Choices-in-a-Batch-File
@ECHO OFF

ECHO ---------------------------------
ECHO CHKDSK D:/X /F -- WILL BEGIN NEXT
ECHO ---------------------------------
ECHO BUT FIRST IF THE PROCESS BEGIN TO HAPPEN NORMAL 
ECHO AND NOT WITH A QUICK STOP BECUASE UNABLE TO LOCK DRIVE
ECHO IF THAT HAPPEN WHERE RUN BEGIN TO PROCESS
ECHO AND THEN YOU WANT TO TYPE Y FOR YES AND PRESS RETURN ENTER
ECHO AND THEN THAT WILL ENABLED A REBOOT AFTER THE JOB 
ECHO AND THEN THAT WILL ENABLE TO RUN RUN THE CHKDSK OF C: DRIVE AT REBOOT
ECHO WHEN YOU HAVE SET THAT TO DO BEFORE
ECHO ----
ECHO RISKY BECAUSE YOU WON'T SEE THE RESULT OF CHKDSK OF BIG D DRIVE DATA
ECHO ----
ECHO CHECK THE EVENT LOGG LIKE YOU DO AT REBOOT CHKDSK ANYWAY
ECHO --------------------------------------------------------




CHKDSK D:/X /F
REM DIR C:\*.* /S

@REM ECHO DO YOU WANT TO RUN REBOOT (Y/N)
set /p choice=DO YOU WANT TO RUN REBOOT (Y/N)
if '%choice%'=='Y' goto RUN_REBOOT

GOTO EXIT_END

:RUN_REBOOT
ECHO REBOOT WILL RUN

:EXIT_END

PAUSE
EXIT


@REM -------------
GOTO JUMP_NEXT
@REM -------------
@ECHO off
cls
:start
ECHO.
ECHO 1. Print Hello
ECHO 2. Print Bye
ECHO 3. Print Test
set /p choice=Type the number to print text.
rem if not '%choice%'=='' set choice=%choice:~0;1% ( don`t use this command, because it takes only first digit in the case you type more digits. After that for example choice 23455666 is choice 2 and you get "bye"
if '%choice%'=='' ECHO "%choice%" is not valid please try again
if '%choice%'=='1' goto hello
if '%choice%'=='2' goto bye
if '%choice%'=='3' goto test
ECHO.
goto start
:hello
ECHO HELLO
goto end
:bye
ECHO BYE
goto end
:test
ECHO TEST
goto end
:end
pause
exit



:JUMP_NEXT