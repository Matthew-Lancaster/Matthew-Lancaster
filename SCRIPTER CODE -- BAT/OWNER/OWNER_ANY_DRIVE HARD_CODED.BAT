@ECHO.
@ECHO ---------------------------------------
@ECHO ---------------------------------------

@REM PROBLEM WITH DRIVE FILING ACCESS __ RUN CHKDSK



@ECHO OFF
setlocal EnableDelayedExpansion

REM -------------------------------------------------------------------------------------------------
TIMEOUT /T 4
REM @FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (

REM D DRIVE
@FOR %%G IN (D) DO (
	IF EXIST %%G:\. (
		SET "target_="
		SET target_=%%G:
		CALL :IsDeviceReady !target_! isready_
		IF !isready_!==true (
			FOR /f "delims=" %%F in ('Dir !target_!\ /b /ad') do (
				SET "VAR_FLAG="
				SET VAR_FLAG=YES
				IF "%%F"=="System Volume Information" SET VAR_FLAG=NOT
				IF "%%F"=="$Recycle.Bin" SET VAR_FLAG=NOT
				IF "%%F"=="$RECYCLE.BIN" (
					SET VAR_FLAG=NOT
					REM PAUSE
				)
				IF "%%F"=="Recovery" SET VAR_FLAG=NOT
				IF "!VAR_FLAG!"=="YES" (
						ECHO.
						TITLE #01 OF #03 __ TAKEOWN ___________ "%%G:\%%F"
						ECHO "%%G:\%%F"
						ECHO #01 OF 03 __ TAKEOWN /F "%%G:\%%F" /R /A /D N 
						ECHO ---------------------------------------------------------------------------
						REM TAKEOWN /F "%%G:\%%F" /R /A /D Y >NUL
						REM TIMEOUT /T 2
						ECHO.
						TITLE #02 OF #03 __ ICACLS /SETOWNER __ "%%G:\%%F"
						ECHO "%%G:\%%F"
						ECHO #01 OF 02 __ ICACLS "%%G:\%%F" /setowner "Everyone"
						TIMEOUT /T 2
						ECHO ---------------------------------------------------------------------------
						ICACLS "%%G:\%%F" /T /C /L /Q /setowner "Everyone" >NUL

						ECHO.
						TITLE #03 OF #03 __ ICACLS /GRANT _____ "%%G:\%%F"
						ECHO "%%G:\%%F"
						ECHO #02 OF 02 __ ICACLS "%%G:\%%F" /grant Everyone:F
						ECHO ---------------------------------------------------------------------------
						ICACLS "%%G:\%%F" /T /C /L /Q /grant Everyone:F >NUL
						REM TIMEOUT /T 2
						ECHO.
				)
			)
		)
	)
)


GOTO VERY_ENDE

::
:IsDeviceReady
setlocal
set "ready_="
set ready_=true
dir "%~1" > nul 2>&1
if %errorlevel% NEQ 0 set ready_=false
endlocal & set "%2=%ready_%" & goto :EOF




:VERY_ENDE
PAUSE
