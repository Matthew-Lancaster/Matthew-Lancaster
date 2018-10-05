@CD\
@ECHO ------------------------------------------------------
@TITLE %~f0

%SystemRoot%\System32\tsdiscon.exe /v

@ECHO OFF
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
@FOR /F %%# IN ('COPY /Z "%~dpf0" NUL') DO SET "CR=%%#"
@FOR /L %%# IN (20,-1,1) DO (SET/P "=Script will end in %%# seconds. !CR!"<NUL:
	PING -n 2 127.0.0.1 >NUL:)
@ECHO.