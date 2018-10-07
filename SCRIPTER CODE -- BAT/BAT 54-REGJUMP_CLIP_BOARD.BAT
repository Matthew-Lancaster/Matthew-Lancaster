@REM ----------------------------------------------------------------
@REM __ MATTHEW LANCASTER 
@REM __ MATT.LAN@BTINTERNET.COM
@REM ----------------------------------------------------------------
@REM __ ALL BATCH FILE COMMAND SETS 
@REM __ GOT AN IMPROVED ADMINISTRATOR LAUNCHER
@REM __ THAT DOESN'T GET STUCK IN A LOOP IF NOT SUCCESSFUL
@REM __ AT THE SACRIFICE OF COMMAND LINE ARGUMENTS BUT EASY TO 
@REM __ ADD CODE IF REQUIRE
@REM ----------------------------------------------------------------
@REM __ FROM ---- --- 20-May-2016 19:57:00 Creation Time
@REM __ TO   ---- Sun 07-Oct-2018 10:58:00 Last Update Header Info
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- 001 SESSION -- 
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- HERE BAT IS NOW SYNC IN 2 LOCATION YET IS SOME ABOUT
@REM -- CENTRALIZING
@REM ----------------------------------------------------------------
@REM -- C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\REGJUMP_CLIP_BOARD.BAT
@REM -- C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 54-REGJUMP_CLIP_BOARD.BAT
@REM ----------------------------------------------------------------
@REM -- IS A SYNCER _ SYNCED BY HERE :) 
@REM -- TO KEEP CENTRAL 
@REM -- MULTI JOB IN ONE
@REM ----------------------------------------------------------------
@REM -- C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 32-COPIER_SYNC.VBS
@REM ----------------------------------------------------------------



@ECHO OFF
@REM "regedit jump to"
@"C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\regjump.exe" -c | findstr /I "Complete." | findstr /I "Not a valid registry path"
if %errorlevel% == 0 (
    @ECHO OKAY REGISTRY PATH ----
) else (
    ECHO.
    ECHO ---------------------------------------
    ECHO ERROR REGISTRY PATH                   -
    ECHO ---------------------------------------
    ECHO.
    "C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\regjump.exe" -c
    ECHO.
    ECHO ---------------------------------------
    ECHO.
    PAUSE
)



@REM ECHO DONE ----
@REM PAUSE

EXIT

@REM ----------------------------------------------------------------------
@REM -- HELP ONLY ONE COMMAND LINE OPERATOR
@REM -- "C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\regjump.exe" /?
@REM ----------------------------------------------------------------------


@REM ----------------------------------------------------------------------
@REM IF ERRORLEVEL 0 ( @ECHO. && @PAUSE )
@REM dir c:\ /s /b | find "LOG" | more
@REM ----------------------------------------------------------------------


@REM ----
@REM windows - Search for a string in command line output - Stack Overflow
@REM http://stackoverflow.com/questions/12453733/search-for-a-string-in-command-line-output
@REM ----

@REM systeminfo | findstr /I /B /C:"system model" | findstr /I "virtual"
@REM if %errorlevel% == 0 (
@REM     echo virtual machine
@REM ) else (
@REM     echo real machine
@REM )


@REM ----------------------------------------------------------------------

@REM -- ----
@REM -- windows - Batch file: Find if substring is in string (not in a file) - Stack Overflow
@REM -- http://stackoverflow.com/questions/7005951/batch-file-find-if-substring-is-in-string-not-in-a-file
@REM -- ----
@REM -- Fri 26 August 2016 18:33:05

@REM -- For compatibility and ease of use it's always best to use FIND to do this, as you may be looking to match 
@REM -- output from a command, a variable in a loop, or a pointer variable in your Batch/CMD.
@REM -- ----
@REM -- ECHO.%Variable% | FIND /I "ABC">Nul && ( Echo Found "ABC" ) || ( Echo Did not find "ABC" )
@REM -- Fri 26 August 2016 18:45:07

@REM -- /B -- BARE NOT HEADER INFO
@REM -- /L -- Lower Case
@REM -- /S -- SUB FOLDERING
@REM -- /AD DIR ONLY AND HIDDEN SYS
