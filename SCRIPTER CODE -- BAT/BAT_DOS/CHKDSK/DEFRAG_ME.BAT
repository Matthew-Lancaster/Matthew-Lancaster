@setlocal EnableDelayedExpansion
@ECHO OFF
SET LABEL_CODER=DEFRAG

@REM -- JOB TO DO ALL DRIVE IF EXIST
@REM -- MEANT TO BE RUN IN SAFE MODE WHERE ACCESS IS TOTAL 
@REM -- AND TO VIEW ON TASK MANAGER TO SEE TITLE CHANGE WHEN JOB ARE COMPLETE
@REM -- IF USE MULTI TASK MODE -- BUT DEFRAG DON'T WORK MULTI TASK -- SINGLE INSTANCE
@REM -- INSTANCE -- INSTANT -- IN-STANCE
@REM -- TIME CHECK -- SUN 28-AUG-2016
@REM                  BEGIN ABOUT 10 AM 
@REM --               FINISH AROUND 12:45 PM 

@ECHO -----------------------------------------------------------------------
@ECHO -- BEGIN %LABEL_CODER% -- %DATE% -- %TIME%%
@ECHO -----------------------------------------------------------------------
@ECHO.

@TITLE %LABEL_CODER% RUNNER ALL DRIVE AVAILABLE

@REM --------------------------------------------------
@REM -- REM BOOT OS SYSTEM DRIVE AT BOOT IS ONLY WAY TO UNLOCK IT FOR A CHKDSK 
@REM --------------------------------------------------
@REM -- SAFE MODE MAYBE PROBLEM IF NETWORK CABLE PLUGGED IN ON WIN XP
@REM --------------------------------------------------
@REM -- LOOK AT TASK MANAGER SEE RESULT OF THESE WORK DONE
@REM --------------------------------------------------
@REM -- BOOT SAFE MODE IN COMMAND MODE
@REM -- AND LAUNCH EXPLORER FROM IT FROM TASK MANAGER
@REM -- AND KILL EXPLORER PROCESS AGAIN -- IF WANT TO
@REM -- WHEN LOAD EXPLORER SAME AS SAFE MODE BUT NOT COMMAND MODE

@REM -- USE  /MIN /HIGH /WAIT IN THIS ORDER OR /MIN WON'T WORK

@SET TIME_BEGIN=%DATE%--%TIME%%

FOR %%G IN (C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (
	IF EXIST %%G:\. (
		
		@TITLE %LABEL_CODER% %%G:\ BEGIN %DATE%--%TIME% START %TIME_BEGIN% IN PROGRESS 
		@ECHO -----------------------------------------------------------------------
		@ECHO %%G:\ -- %DATE% -- %TIME%
		@ECHO -----------------------------------------------------------------------
		@ECHO VOLUME 
		VOL %%G:
		@ECHO -----------------------------------------------------------------------
		@ECHO.

		SET "START_COM="
		@SET START_COM=START /B /HIGH /WAIT %LABEL_CODER% %%G: /F /V

		@REM -- COMMAND DISPLAY
		ECHO !START_COM!
		@ECHO.

		@REM -- COMMAND EXECUTE
		!START_COM!
		@ECHO.
		@ECHO -----------------------------------------------------------------------
		@ECHO.
	)
)

@TITLE %LABEL_CODER% ALL DRIVE DONE -- START %TIME_BEGIN% -- END %DATE%--%TIME% 
@ECHO.
@ECHO -----------------------------------------------------------------------
@ECHO -- FINISH %LABEL_CODER% START %TIME_BEGIN% -- END %DATE%--%TIME%
@ECHO -----------------------------------------------------------------------
@ECHO.
@ECHO -----------------------------------------------------------------------
FOR %%G IN (C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (
	IF EXIST %%G:\. (
		@ECHO %LABEL_CODER% -- RAN ON DRIVE ---- %%G:\ --  %LABEL_CODER% 
		VOL %%G:

	)
)
@ECHO.
@ECHO -----------------------------------------------------------------------
@ECHO.
@ECHO PRESS KEY TO EXIT
@ECHO.
@PAUSE
@REM EXIT

@REM -- CODE REDUNDANT SCRAP 
@REM SET START_COM=START /B /HIGH  C:\BAT\CHKDSK\%LABEL_CODER%_RUN_ME %%G