@ECHO --------------------------------------------------------
@ECHO -- THIS IS A VB INSTALLED PROGRAM I THINK
@ECHO -- AND BELONG IN VB APP FOLDER
@ECHO -- IT IS ONE TYPE DRIVER THAT DONT RUN FROM SET APP.PATH 
@ECHO -- IT DON'T SHOW WHEN PRGRAM IN USE A A ATTACHED DRIVER
@ECHO --------------------------------------------------------


@ECHO -- TRY METHOD
@REM COPY ClipboardViewer.oca C:\WINDOWS\SYSTEM32\ClipboardViewer.oca

C:
CD \


@REM -- Regsvr32.exe C:\WINDOWS\SYSTEM32\ClipboardViewer.oca
@REM -- C:\Windows\System32\regsvr32.exe C:\Windows\System32\ClipboardViewer.oca
@REM -- C:\Windows\syswow64\regsvr32.exe C:\Windows\System32\ClipboardViewer.oca

@REM -- regsvr32.exe C:\Windows\System32\ClipboardViewer.oca


@ECHO -- -----------------------------------------------------
@ECHO -- DOUBLE CHECK -- CORRECT EXTIST HERE AFTER NEW INSTALL
@ECHO -- -----------------------------------------------------
REM regsvr32.exe "C:\Program Files\Microsoft Visual Studio\VB98\ClipboardViewer.oca"
@ECHO -- -----------------------------------------------------

@ECHO -- -----------------------------------------------------
@ECHO -- 64 BIT LOCATION 
@ECHO -- -----------------------------------------------------

@REM -- regsvr32.exe "C:\Program Files (x86)\Microsoft Visual Studio\VB98\ClipboardViewer.oca"
@ECHO -- -----------------------------------------------------
@ECHO -- SOLVED METHOD ----- 
@ECHO -- THE 32 BIT VERSION OF regsvr32 IS IN THIS ONE 
@ECHO -- C:\windows\syswow64\regsvr32 -- THE -- 64 BIT FOLDER 
@ECHO -- DEFAULT USE IS 64 BIT VERSION OF COURSE ON 64 BIT OS OPPO
@ECHO -- -----------------------------------------------------

CD \windows\syswow64

C:\windows\syswow64\regsvr32 "C:\Program Files (x86)\Microsoft Visual Studio\VB98\ClipboardViewer.oca"


@ECHO -- -----------------------------------------------------
@ECHO -- -----------------------------------------------------


@ECHO -- -----------------------------------------------------
@ECHO -- GUESS MAYBE REGISTER TWICE IF WANT RUN WITHOUT VB IDE INSTALLED
@ECHO -- AS PORTABLE APP
@ECHO -- -----------------------------------------------------


@ECHO -- -----------------------------------------------------
@ECHO.
@ECHO IS ALL REPORT OKAY
@ECHO.
@ECHO -- -----------------------------------------------------

@PAUSE
EXIT

@ECHO -- -----------------------------------------------------
REM --- REGISTER
REM --- regsvr32

REM --- UNREGISTER
REM --- regsvr32 /u
@ECHO -- -----------------------------------------------------
