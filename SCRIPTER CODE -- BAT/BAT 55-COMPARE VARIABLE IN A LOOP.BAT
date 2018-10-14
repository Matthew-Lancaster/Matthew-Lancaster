
REM setlocal EnableDelayedExpansion
@REM -- PROBLEM WITH EnableDelayedExpansion USE GOTO LOOP MAKE SURE VARIABLE HAS BEEN SET
@REM -- IS WHENEVER SET NEW ICACLS AND WANT TO RUN QUICCKER WANT TO PASTE CODE IN CMD PROMPT

SET PATH_TO_DO=

SET PATH_TO_DO_OLD=%PATH_TO_DO%
SET PATH_TO_DO="C:\Windows\System32\*.VBS"

ECHO %PATH_TO_DO%
ECHO %PATH_TO_DO_OLD%

set COUNTER=0

@echo off
if "%1"=="loop" (
  for /l %%f in (1,1,1000) do (
    echo %%f
	set /a COUNTER=COUNTER+1
    echo %COUNTER%
	REM IF NOT %PATH_TO_DO%==%PATH_TO_DO_OLD% GOTO:EXIT_LOOP
    if exist %%f exit
  )
  goto :eof
)
cmd /v:on /q /d /c "%0 loop"
echo done


ECHO DONE

PAUSE
EXIT /B


For %%B in (U V W X Y Z) DO (
   if %%B==X ( ECHO TEN )
   echo    Inner loop    Outer=%%A Inner=%%B
)


PAUSE
EXIT /B


@echo off
SetLocal EnableDelayedExpansion

set n=11
set m=12
set /a nme=3
set /a mdiff=nme-1
pause
if %n% NEQ %m% (
	for /l %%C in (1,1,3) do (
		if %%C EQU 1 (
			set mon=Apr
			set num=1!mon!
            )
    )
)
echo %num%
pause