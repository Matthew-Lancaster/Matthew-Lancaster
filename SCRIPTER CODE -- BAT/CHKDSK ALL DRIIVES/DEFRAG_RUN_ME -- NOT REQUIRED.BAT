@ECHO OFF
SET LABEL_CODER=DEFRAG
@ECHO -----------------------------------------------
@SET TIME_BEGIN=%DATE%--%TIME%%
@TITLE %1:\ -- %LABEL_CODER% %1: /F /X -- %TIME_BEGIN% -- IN PROCESS WORKING
@ECHO %1:\ -- BEGIN  -- %TIME_BEGIN%
@ECHO -----------------------------------------------
@VOL %1:
@ECHO.
@ECHO -----------------------------------------------
@ECHO CURRENT PATH ----------------------------------
@CD
@ECHO.
@ECHO -----------------------------------------------
@SET EXE_COM=START /B /HIGH %LABEL_CODER% %1: /F /V
@REM USE /WAIT AFTER /HIGH -- MAYBE
@REM 
@REM -- COMMAND DISPLAY
@ECHO %EXE_COM%
@ECHO.
@REM -- COMMAND EXECUTE
@%EXE_COM%
@ECHO.
@ECHO -----------------------------------------------
@ECHO CURRENT PATH ----------------------------------
@CD
@ECHO.
@ECHO -----------------------------------------------
@VOL %1:
@ECHO.
@SET TIME_FINISH=%DATE%--%TIME%%
@ECHO %1:\ -- BEGIN  -- %TIME_BEGIN%
@ECHO %1:\ -- FINISH -- %TIME_FINISH%
@ECHO -----------------------------------------------
@ECHO.
@TITLE %1:\ -- %LABEL_CODER% %1: /F /X - DONE FINISHED - %TIME_BEGIN%--%TIME_FINISH%

IF EXIST C:\WINDOWS\system32\sndrec32.exe (
	REM C:\WINDOWS\system32\sndrec32.exe /play /close "C:\Windows\Media\Notify.wav"
	REM ECHO -- PLAY AUDIO NOTIFY
)

@PAUSE
EXIT