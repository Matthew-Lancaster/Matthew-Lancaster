@ECHO OFF


@REM -----------------------------------------------------
@REM TO DO LATER THE ROOT OF C 
@REM icacls c:\ /grant "Everyone":(OI)(CI)(F)


@REM ____ icacls c:\ /grant "MATT 01":(OI)(CI)(M)
@REM -- -------------------------
@REM -- UNABLE ON MICROSOFT SITE TO DIRECTLY LINK TO THE COMMENT THAT PROVIDE THE SOURCE FROM 2010
@REM -- PATHETIC REFERANCE SOURCE
@REM -- -------------------------
@REM -- NightMareForYou replied on  August 7, 2010See post history
@REM -- ReplyIn reply to Mac1973's post on August 6, 2010
@REM -- ----
@REM -- Error 0x80070522: A required privilege is not held by the client - Microsoft Community
@REM -- https://answers.microsoft.com/en-us/windows/forum/windows_vista-security/error-0x80070522-a-required-privilege-is-not-held/78cc11ec-c543-418d-9183-131318e9db8a
@REM -- --------
@REM -- error code 0x80070522
@REM -- https://social.technet.microsoft.com/Forums/windows/en-US/b5e69635-038b-4a80-8786-70e1b1d8be0a/error-code-0x80070522?forum=itprovistaannouncements
@REM -- ----


@REM --------------------------------------------------------
@REM --------------------------------------------------------

@REM ____ NOT SURE TO DO EXPANSION WHEN VARIABLE SYNCHRONOUS
@REM setlocal
Setlocal EnableDelayedExpansion

@REM __ VARIABLE __ V __ STAND FOR __ VARIABLE  __ AND VARIABLE __ VT __ STAND FOR __ VARIABLE TEXT __

ECHO ON

SET V=0

SET /A V=V+1 & SET VT%V%="C:\x86"
SET /A V=V+1 & SET VT%V%="C:\#_IRFAN_VIEW_USB"
SET /A V=V+1 & SET VT%V%="C:\0 00 LOGGERS TEXT"
SET /A V=V+1 & SET VT%V%="C:\Asus WebStorage"
SET /A V=V+1 & SET VT%V%="C:\autorun"
SET /A V=V+1 & SET VT%V%="C:\BACKUP"
SET /A V=V+1 & SET VT%V%="C:\BAT"
SET /A V=V+1 & SET VT%V%="C:\CAPTURE"
SET /A V=V+1 & SET VT%V%="C:\DOWNLOADS"
SET /A V=V+1 & SET VT%V%="C:\DOWNLOADS LINK FOLDER WHEN ASKER"
SET /A V=V+1 & SET VT%V%="C:\GOOD SYNC JOB & OPTION TIC MIRROR"
SET /A V=V+1 & SET VT%V%="C:\GoodSync"
SET /A V=V+1 & SET VT%V%="C:\HBCD"
SET /A V=V+1 & SET VT%V%="C:\images"
SET /A V=V+1 & SET VT%V%="C:\ISO"
SET /A V=V+1 & SET VT%V%="C:\Make_PE3"
SET /A V=V+1 & SET VT%V%="C:\MINIPE"
SET /A V=V+1 & SET VT%V%="C:\MSDOS6.22"
SET /A V=V+1 & SET VT%V%="C:\RF"
SET /A V=V+1 & SET VT%V%="C:\SCRIPTER"
SET /A V=V+1 & SET VT%V%="C:\TEMP"
SET /A V=V+1 & SET VT%V%="C:\UTILS"

SET /A V=V+1 & SET VT%V%="C:\WIN 10 PRO £189"
SET /A V=V+1 & SET VT%V%="C:\WIN_10_PRO_GRUB_4_DOS"
SET /A V=V+1 & SET VT%V%="C:\WINDOW98"
SET /A V=V+1 & SET VT%V%="C:\Windows10Upgrade"
SET /A V=V+1 & SET VT%V%="C:\BAT_WIN_XP"
SET /A V=V+1 & SET VT%V%="C:\BAT-DUPER"
SET /A V=V+1 & SET VT%V%="C:\PStart"
SET /A V=V+1 & SET VT%V%="C:\VB6"

SET /A V=V+1 & SET VT%V%="D:\# MY DOCS"
SET /A V=V+1 & SET VT%V%="D:\# MY DOCS COMMON"
SET /A V=V+1 & SET VT%V%="D:\#0 1 INSTALLATIONS"
SET /A V=V+1 & SET VT%V%="D:\#0 IRFAN SHOW SCRIPT FILE SET"
SET /A V=V+1 & SET VT%V%="D:\0 00 Art Loggers"
SET /A V=V+1 & SET VT%V%="D:\0 00 LOGGERS TEXT"
SET /A V=V+1 & SET VT%V%="D:\03_MICROSOFT"
SET /A V=V+1 & SET VT%V%="D:\BT Cloud Sync OLD"
SET /A V=V+1 & SET VT%V%="D:\DOWNLOADS"
SET /A V=V+1 & SET VT%V%="D:\DOWNLOADS LINK FOLDER WHEN ASKER"
SET /A V=V+1 & SET VT%V%="D:\DOWNLOADS MOVED"
SET /A V=V+1 & SET VT%V%="D:\DSC"
SET /A V=V+1 & SET VT%V%="D:\KAT MP3 RECORDER"
SET /A V=V+1 & SET VT%V%="D:\VB6"
SET /A V=V+1 & SET VT%V%="D:\VB6-EXE'S"
SET /A V=V+1 & SET VT%V%="D:\VICE_VERSA SCRIPT FILE"
SET /A V=V+1 & SET VT%V%="D:\VIDEO CCSS NVMS-1000"
SET /A V=V+1 & SET VT%V%="D:\Videos"


GOTO JUMPER
SET V=0
SET /A V=V+1 & SET VAR_00=VT!V!
SET V=0
SET /A V=V+1 & SET VAR_01=VT!V!

ECHO VAR_00
ECHO %VAR_00%
ECHO !%VAR_00!%
ECHO %VAR_00%
ECHO %VAR_01%
ECHO !VAR_00!
ECHO !VAR_01!
ECHO  VT%V%
ECHO  VT!V!
ECHO  %VT0%
ECHO  %VT1%

SET /A V=V+1 & SET VAR_00=VT!V!  & CALL :Set_Icacls !VAR_00! %%VAR_00%


:JUMPER
PAUSE 

SET V=0
SET /A V=V+1 & SET VAR_00=%VT%%V= & CALL :Set_Icacls !VAR_00! %%V_47
SET /A V=V+1 & SET VAR_00=%VT%%V= & CALL :Set_Icacls !VAR_00! %%V_47
SET /A V=V+1 & SET VAR_00=%VT%%V= & CALL :Set_Icacls !VAR_00! %%V_47
SET /A V=V+1 & SET VAR_00=%VT%%V= & CALL :Set_Icacls !VAR_00! %%V_47

PAUSE

SET VAR_00=%V_02% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_03% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_04% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_05% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_06% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_07% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_08% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_09% & CALL :Set_Icacls !VAR_00! %%V_47

SET VAR_00=%V_10% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_11% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_12% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_13% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_14% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_15% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_16% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_17% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_18% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_19% & CALL :Set_Icacls !VAR_00! %%V_47

SET VAR_00=%V_20% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_21% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_22% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_23% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_24% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_25% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_26% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_27% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_28% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_29% & CALL :Set_Icacls !VAR_00! %%V_47

SET VAR_00=%V_30% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_31% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_32% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_33% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_34% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_35% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_36% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_37% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_38% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_39% & CALL :Set_Icacls !VAR_00! %%V_47

SET VAR_00=%V_40% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_41% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_42% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_43% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_44% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_45% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_46% & CALL :Set_Icacls !VAR_00! %%V_47
SET VAR_00=%V_47% & CALL :Set_Icacls !VAR_00! %%V_47


ECHO.
ECHO DATE
ECHO TIME
ECHO ENDING --------------------------------------------------

PAUSE

GOTO VERY_ENDE

::
:Set_Icacls
setlocal
@REM ------------------------------------------------------------------------------------------------------------
@REM ---- EMPTY VALUE NOT GET LEARN __ STACKER THE CODE __ SLOT THEM IN
@REM ---- BETTER THAN A DIMENSION ARRAY __ USE THE KEYBOARD
@REM ------------------------------------------------------------------------------------------------------------

@REM if "%1%"=="" GOTO JUMPER_NOT_DO_SUBROUTINE
@REM if "!1%"=="" GOTO JUMPER_NOT_DO_SUBROUTINE

ECHO %2
TITLE %2

@ECHO ICACLS "%~1\*" /setowner Everyone /T /C /L /Q 
@ECHO ICACLS "%~1\*" /T /L /C /Q /Grant Everyone:F

:JUMPER_NOT_DO_SUBROUTINE

endlocal & goto :EOF
)


PAUSE
:VERY_ENDE



@REM ---- REFERENCE SET
@REM ---- 3 SEARCH BLOCK AND THEN 
@REM ---- 3 RESULT BLOCK

@REM -- ---------------------
@REM -- Count = 116 -- Tue 14-Feb-2017 21:21:38
@REM -- ---------------------
@REM -- ----
@REM -- STRIP THE QUOTE MARK OFF BATCH FILE - Google Search
@REM -- https://www.google.co.uk/search?num=50&q=STRIP+THE+QUOTE+MARK+OFF+BATCH+FILE&oq=STRIP+THE+QUOTE+MARK+OFF+BATCH+FILE&gs_l=serp.3...2218.4472.0.5143.2.2.0.0.0.0.87.163.2.2.0....0...1c.1.64.serp..0.1.86...30i10k1.uE1v2xu2TiI
@REM -- --------
@REM -- BATCH FILE TWO COMMAND ONE LINE - Google Search
@REM -- https://www.google.co.uk/search?q=BATCH+FILE+TWO+MAND+ONE+LINE&gws_rd=ssl
@REM -- --------
@REM -- BATCH FILE IF A STRING THAN NOTHING - Google Search
@REM -- https://www.google.co.uk/search?num=50&q=BATCH+FILE+IF+A+STRING+THAN+NOTHING&oq=BATCH+FILE+IF+A+STRING+THAN+NOTHING&gs_l=serp.3..30i10k1.5598.8706.0.8983.19.16.0.0.0.0.166.1287.10j4.14.0....0...1c.1.64.serp..9.8.703.bpKPNndiTdE
@REM -- ----


@REM -- ---------------------
@REM -- ---------------------
@REM -- Count = 117 -- Tue 14-Feb-2017 21:22:06
@REM -- ---------------------
@REM -- Form FindWindow ---
@REM -- string - Batch check for empty user input - Stack Overflow - Google Chrome
@REM -- ---------------------
@REM -- ----
@REM -- Removing double quotes from variables in batch file creates problems with CMD environment - Stack Overflow
@REM -- http://stackoverflow.com/questions/1964192/removing-double-quotes-from-variables-in-batch-file-creates-problems-with-cmd-en
@REM -- --------
@REM -- Multiple commands on a single line in a Windows batch file - Stack Overflow
@REM -- http://stackoverflow.com/questions/8922224/multiple-commands-on-a-single-line-in-a-windows-batch-file
@REM -- --------
@REM -- string - Batch check for empty user input - Stack Overflow
@REM -- http://stackoverflow.com/questions/32480655/batch-check-for-empty-user-input
@REM -- ----
