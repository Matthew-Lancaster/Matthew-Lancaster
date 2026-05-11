@REM -----------------------------------------------------------
@REM --- HAD A PROBLEM NOT REGISTERED WITH THESE FILE AT INSTALL
@REM --- THEN HERE -- TRY THIS
@REM -----------------------------------------------------------
@REM -- RUN MSDN AND SERVICE PACK VER # SIX
@REM -- MAKE SURE
@REM -- THEY DO ONLY -- EXIST AFTER -- SERVICE PACK VER # SIX
@REM -- AND THEN CHECK YOUR PROGRAM CODE RUN IF OKAY 
@REM -- -- NEED -- HERE -- WON'T WANT
@REM -- --------------------------------------------------------
@REM -- RUN MSDN FROM THE MSDN FOLDER - NOT SETUP DISK ONE
@REM -- --------------------------------------------------------

@REM -----------------------------------------------------------
@REM -- YOU MIGHT FIND THESE FILE ALREADY THERE AND REGISTERED 
@REM -- BUT IF NOT THE LEGACY VERSION YOUR PROGRAM WONT LOAD IN 
@REM -- WITH GOT DRIVER ERROR BECUASE LAGACY VERSION
@REM -----------------------------------------------------------
@REM -- 

@REM -----------------------------------------------------------
@REM -----------------------------------------------------------
@REM -- I NEED THIS REGISTERED ALWAYS AT INSTALL - THINK A MOMENT -- USE 32 BIT REGSVR32
@REM -----------------------------------------------------------

@REM -----------------------------------------------------------
@REM -- Method 2: Use the 32-bit version of Regsvr32 to register a 32-bit DLL on a 64-bit version of Windows
@REM -- 
@REM -- If you receive an error when registering a 32-bit DLL on a 64-bit version of Windows, follow these steps:
@REM -- 
@REM -- Open an elevated command prompt.
@REM -- If the 32-bit DLL is in the %systemroot%\System32 folder, move it to the %systemroot%\SysWoW64 @REM -- folder.
@REM -- 
@REM -- Run the following command:
@REM -- %systemroot%\SysWoW64\regsvr32 <full path of the DLL>

@ECHO  -----------------------------------------------------------
@ECHO  -----------------------------------------------------------
@ECHO  -----------------------------------------------------------

@REM -----------------------------------------------------------
@REM -- I NEED THIS REGISTERED ALWAYS AT INSTALL - THINK A MOMENT -- USE 32 BIT REGSVR32
@REM -----------------------------------------------------------

@REM -----------------------------------------------------------
@REM -- C:\Program Files\Microsoft Visual Studio\VB98\ClipboardViewer.oca
@REM -----------------------------------------------------------
@REM -- AND OR
@REM -----------------------------------------------------------
@REM -- C:\Program Files (x86)\Microsoft Visual Studio\VB98\ClipboardViewer.oca
@REM -----------------------------------------------------------

@REM -----------------------------------------------------------
@REM -----------------------------------------------------------
@REM -- TRY REGISTER THIS FILE IS PROBLEM -- ClipboardViewer.oca
@REM -- MUST REGISTER THIS FILE -- ClipboardViewer.ocx
@REM -- PROBLEM BECAUSE .OCA IS IN EXPECTED FOLDER NOT .OCX - ClipboardViewer.ocx


@REM -----------------------------------------------------------
@REM DONT CHANGE PATH NEED LATER
@REM -----------------------------------------------------------
@REM C:
@REM CD \
@REM -----------------------------------------------------------

@REM -----------------------------------------------------------
@REM -- FIND 32 BIT OR 64 BIT VERSION
@REM -----------------------------------------------------------
@REM @If exist "C:\Program Files\Microsoft Visual Studio\VB98\ClipboardViewer.oca" SET FILE_PATH_ClipboardViewer="C:\Program Files\Microsoft Visual Studio\VB98\ClipboardViewer.oca"
@REM -----------------------------------------------------------
@REM @If exist "C:\Program Files (X86)\Microsoft Visual Studio\VB98\ClipboardViewer.oca" SET FILE_PATH_ClipboardViewer="C:\Program Files (X86)\Microsoft Visual Studio\VB98\ClipboardViewer.oca"
@REM -----------------------------------------------------------
@REM -----------------------------------------------------------
@REM %systemroot%\SysWoW64\regsvr32 %FILE_PATH_ClipboardViewer%
@REM @ECHO %systemroot%\SysWoW64\regsvr32 %FILE_PATH_ClipboardViewer%
@REM -----------------------------------------------------------
@REM @ECHO ----------------------------------------------------------
@REM @ECHO -- IS OKAY UPTO HERE -- AND THEN -- CONTINUE
@REM @ECHO ----------------------------------------------------------
@REM -----------------------------------------------------------

@REM ClipboardViewer.oca - NOT PROPER
@REM ClipboardViewer.ocx - PROPER - I KNOW FOUND IN ROOT FOLDER OF CLIPBOARD PROGRAM
@REM ClipboardViewer.ocx - THIS ALSO ONLY ONE THAT REGISTER PROPER

@REM -----------------------------------------------------------------------
@REM -- LAST NOTE -- FRI 27 END OF MAY
@REM -- SHOULDN'T BE A PROBLEM WITH REGISTER 
@REM -- ONE OF MY MY CODE -- TOOK A WHILE -- I MUST OF RUN CODE BEFORE READY
@REM -- WITH ALL INSTALLED SERVICE PACK SIX AND ALL
@REM -- AND THEN GOTO COMPONENTS AND RELOAD THEM CHECKED
@REM -----------------------------------------------------------------------
@REM Object={019FD1D9-2A35-4D5E-AB4E-BA793946941B}#1.0#0; ClipboardViewer.ocx
@REM Object={86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCT2.OCX
@REM Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX
@REM ------------------------------------------------------------------------
@REM -- IF YOUR .VBP CODE PROGRAM IS NOT DISPLAY THESE TYPE VERSION CODE
@REM -- LIKE HERE -- #2.0#0;
@REM -- THEN YOU ARE NOT USING THE LEGACY VERSION
@REM -- AND HAVE BEEN UPDATED WITH MORE UP-TO-DATE AND THAT IS A PROBLEM
@REM ------------------------------------------------------------------------


@REM @PAUSE

@REM -----------------------------------------------------------
@REM -----------------------------------------------------------

@REM -- IF MADE A MISTAKE -- THEN BEGIN WITH THIS
@REM -- @ECHO ----------------------------------------------------------
@REM -- SET FILE_PATH=C:\Windows\SysWOW64
@REM -- @ECHO ----------------------------------------------------------
@REM -- COPY MSCOMCT2.OCX.LEGACY %FILE_PATH%\MSCOMCT2.OCX
@REM -- COPY MSCOMCTL.OCX.LEGACY %FILE_PATH%\MSCOMCTL.OCX
@REM -- COPY MSCOMCT2.OCX.LEGACY %FILE_PATH%\MSCOMCT2.OCX.LEGACY
@REM -- COPY MSCOMCTL.OCX.LEGACY %FILE_PATH%\MSCOMCTL.OCX.LEGACY
@REM -- @ECHO ----------------------------------------------------------




@ECHO ----------------------------------------------------------
@ECHO ----------------------------------------------------------
@ECHO -- 01 OF 02 -- MSCOMCT2.OCX IN WINDOWS FOLDER
@ECHO ----------------------------------------------------------
@ECHO ----------------------------------------------------------
@If exist C:\Windows\System32\MSCOMCT2.OCX SET FILE_PATH=C:\Windows\System32
@If exist C:\Windows\SysWOW64\MSCOMCT2.OCX SET FILE_PATH=C:\Windows\SysWOW64
@REM -------------------------------------------------------------------------

@ECHO ----------------------------------------------------------
@ECHO -- MSCOMCT2.OCX -- FILE EXIST CHECK TRUE
@ECHO -- FILE EXIST CHECK TRUE @ -- %FILE_PATH%
@ECHO ----------------------------------------------------------

@REM -------------------------------------------------------------------------
@REM -- CHECK %FILE_PATH% -- VAR IS EMPTY -- USE THIS METHOD
@REM -------------------------------------------------------------------------
@if %FILE_PATH%!==! @echo -- -------------------------------------------------
@if %FILE_PATH%!==! @echo -- FILE DON'T EXIST C:\Windows\System32\MSCOMCT2.OCX
@if %FILE_PATH%!==! @echo -- AND OR --- FILE DON'T EXIST
@if %FILE_PATH%!==! @echo -- FILE DON'T EXIST C:\Windows\SysWOW64\MSCOMCT2.OCX
@if %FILE_PATH%!==! @echo -- -------------------------------------------------
@REM -------------------------------------------------------------------------


@ECHO ----------------------------------------------------------
@ECHO ----------------------------------------------------------
@ECHO -- 02 OF 02 -- MSCOMCTL.OCX IN WINDOWS FOLDER
@ECHO ----------------------------------------------------------
@ECHO ----------------------------------------------------------

@If exist C:\Windows\System32\MSCOMCTL.OCX SET FILE_PATH=C:\Windows\System32
@If exist C:\Windows\SysWOW64\MSCOMCTL.OCX SET FILE_PATH=C:\Windows\SysWOW64

@ECHO ----------------------------------------------------------
@ECHO -- MSCOMCTL.OCX -- FILE EXIST CHECK TRUE
@ECHO -- FILE EXIST CHECK TRUE @ -- %FILE_PATH%
@ECHO ----------------------------------------------------------

@ECHO ----------------------------------------------------------
@REM -------------------------------------------------------------------------
@if %FILE_PATH%!==! @echo -- -------------------------------------------------
@if %FILE_PATH%!==! @echo -- FILE DON'T EXIST C:\Windows\System32\MSCOMCTL.OCX
@if %FILE_PATH%!==! @echo -- AND OR --- FILE DON'T EXIST
@if %FILE_PATH%!==! @echo -- FILE DON'T EXIST C:\Windows\SysWOW64\MSCOMCTL.OCX
@if %FILE_PATH%!==! @echo -- -------------------------------------------------
@REM -------------------------------------------------------------------------
@if %FILE_PATH%!==! @echo -- GOTO EXIT 
@if %FILE_PATH%!==! @GOTO END_EXIT
@REM -------------------------------------------------------------------------
@REM -------------------------------------------------------------------------
@ECHO ----------------------------------------------------------

@ECHO -- NEXT STEP
@REM -------------------------------------------------------------------------


@ECHO ----------------------------------------------------------
@ECHO ----------------------------------------------------------
@IF NOT %FILE_PATH%!==! @ECHO -- MSCOMCT2.OCX AND OR MSCOMCTL.OCX -- FILE EXIST CHECK TRUE
@IF NOT %FILE_PATH%!==! @ECHO -- -- FILE EXIST CHECK TRUE @ -- %FILE_PATH%
@ECHO ----------------------------------------------------------
@ECHO ----------------------------------------------------------

@REM -- IF A PROBLEM AND FILE IS MISSING REM THIS LINE OUT TO CONTINUE AND PUT BACK
@if %FILE_PATH%!==! @GOTO END_EXIT


@REM --------------------------------------------------------------------------------------
@REM -- THESE TWO FILE NEED COPY TO WIN SYS FOLDER
@REM -- TO BE BETTER AND SAFE WHILE AT THIS JOB BETTER TO DO THE LEGACY VERSION WHILE AT IT
@REM --------------------------------------------------------------------------------------

@ECHO ----------------------------------------------------------
@ECHO -- CHECK EXIST -- MSCOMCT2.OCX.LEGACY
@ECHO ----------------------------------------------------------
@REM --------------------------------------------------------------------------------------
@If exist MSCOMCTL.OCX.LEGACY @echo -- MSCOMCTL.OCX.LEGACY -- FILE EXIST TRUE AT THIS LOCATION FOLDER
@ECHO -- CHECK EXIST MSCOMCT2.OCX.LEGACY
@If exist MSCOMCT2.OCX.LEGACY @echo -- MSCOMCT2.OCX.LEGACY -- FILE EXIST TRUE AT THIS LOCATION FOLDER

@SET EXIST_FLAG=TUFF

@ECHO ----------------------------------------------------------
@If exist MSCOMCTL.OCX.LEGACY SET EXIST_FLAG=YES
@If exist MSCOMCT2.OCX.LEGACY SET EXIST_FLAG=YES
@ECHO ----------------------------------------------------------

@ECHO ----------------------------------------------------------
@if %EXIST_FLAG%==TUFF @ECHO -- MSCOMCT2.OCX.LEGACY
@if %EXIST_FLAG%==TUFF @ECHO -- AND OR
@if %EXIST_FLAG%==TUFF @ECHO -- MSCOMCTL.OCX.LEGACY
@if %EXIST_FLAG%==TUFF @ECHO -- DON'T EXIST
@if %EXIST_FLAG%==TUFF @ECHO -- GOTO EXIT
@if %EXIST_FLAG%==TUFF @GOTO END_EXIT
@ECHO ----------------------------------------------------------

@ECHO ----------------------------------------------------------
COPY MSCOMCT2.OCX.LEGACY %FILE_PATH%\MSCOMCT2.OCX.LEGACY
COPY MSCOMCTL.OCX.LEGACY %FILE_PATH%\MSCOMCTL.OCX.LEGACY
@ECHO ----------------------------------------------------------

@PAUSE

@REM ------------------------------
@ECHO ----------- DO-- MSCOMCT2.OCX
@REM -------------
@ECHO ----------- UNREGISTER
Regsvr32.exe /u %FILE_PATH%\MSCOMCT2.OCX
@REM -------------------
@ECHO ----------- COPY WANT NEW IN
COPY .\MSCOMCT2.OCX.LEGACY %FILE_PATH%\MSCOMCT2.OCX
@REM -----------
@ECHO ----------- REGISTER
Regsvr32.exe %FILE_PATH%\MSCOMCT2.OCX
@REM ----------------------
@ECHO ----------- UNREGISTER -- AGAIN 
Regsvr32.exe /u %FILE_PATH%\MSCOMCT2.OCX
@REM ---------
@ECHO ----------- DELETE
DEL %FILE_PATH%\MSCOMCT2.OCX
@REM ----------------------------
@ECHO ----------- COPY WANT NEW IN -- AGAIN 
COPY %FILE_PATH%\MSCOMCT2.OCX.LEGACY  %FILE_PATH%\MSCOMCT2.OCX
@ECHO ----------- -----------
@ECHO ----------- REGISTER
Regsvr32.exe %FILE_PATH%\MSCOMCT2.OCX
@ECHO ----------- DONE-- MSCOMCT2.OCX


@REM ------------------------------
@ECHO ----------- DO-- MSCOMCTL.OCX
@REM -------------
@ECHO ----------- UNREGISTER
Regsvr32.exe /u %FILE_PATH%\MSCOMCTL.OCX
@REM -------------------
@ECHO ----------- COPY WANT NEW IN
COPY .\MSCOMCTL.OCX.LEGACY %FILE_PATH%\MSCOMCTL.OCX
@REM -----------
@ECHO ----------- REGISTER
Regsvr32.exe %FILE_PATH%\MSCOMCTL.OCX
@REM ----------------------
@ECHO ----------- UNREGISTER -- AGAIN 
Regsvr32.exe /u %FILE_PATH%\MSCOMCTL.OCX
@REM ---------
@ECHO ----------- DELETE
DEL %FILE_PATH%\MSCOMCTL.OCX
@REM ----------------------------
@ECHO ----------- COPY WANT NEW IN -- AGAIN 
COPY %FILE_PATH%\MSCOMCTL.OCX.LEGACY  %FILE_PATH%\MSCOMCTL.OCX
@ECHO ----------- -----------
@ECHO ----------- REGISTER
Regsvr32.exe %FILE_PATH%\MSCOMCTL.OCX
@ECHO ----------- DONE-- MSCOMCTL.OCX
@ECHO .
@ECHO ----------- CHECK ALL SHOULD SUCCEED -----------

:END_EXIT
PAUSE
EXIT