@ECHO OFF
@TITLE JOB BEGIN - SET OWNER EVERYONE AND GRANT PERMISSION EVERYONE - ALL DRIVE D:\ TO Z:\
@ECHO ---------------------------------------------------------------
@ECHO -- JOB BEGIN -- SET OWNER EVERYONE AND GRANT PERMISSION EVERYONE
@ECHO -- FOR ALL ABOVE D DRIVE
@ECHO -- AND SOME OPPO SYS OS C DRIVE
@ECHO ---------------------------------------------------------------

@ECHO -- LEARN METHOD JULY 13 MID-NIGHT ---
@ECHO -- CREDIT DUE TO IN REF IN PLACE
@ECHO -- CREATE THIS BATCH FILE 1:40 PM 13 JULY WEDNESDAY --- 
@ECHO -- CREATE THIS BATCH FILE 4:15 PM 13 JULY WEDNESDAY --- FINISH 
@REM ----------------------------------------------------------------
@ECHO -- RUN THIS BATCH FILE ADMINISTRATOOR PERMISSION
@ECHO -- TO DO IT ALWAYS CREATE A SHORTCUT TO THIS BATCH FILE AND PROPERTIES SET

@ECHO -- MY PERSONAL SETUP IS I SET ALL DRIVE WITH OWNER AND FULL PERMISSION
@ECHO -- BUT NOT ALL OF C DRIVE OR ERROR OCCUR

@ECHO -- IT IS THE SAME WITH NORTON NIS - I SET ALL DRIVE ABOVE C NOT TO INCLUDE CHECKER
@ECHO -- EXCEPT SOME FOLDER ON C ARE EXCLUDED
@ECHO -- AND THE DOWNLOAD FOLDER IS THERE AND NOT EXCLUDED

@ECHO -- RUN THE JOB FOR ALL DRIVE ABOVE DRIVE C - NOT NETWORK
@ECHO ------------------------------------------------------
@ECHO -- THE JOB IS RUN TO OUTPUT PIPE TO NULL OUTPUT AND THEN GRAPHIC DISPLAY SCROLLING
@ECHO -- WON'T HOLD UP THE SPEED OF OPERATION
@ECHO -- YOU
@ECHO -- THERE WILL BE DISPLAY WHEN ERROR OCCUR --
@ECHO -- THE PROGRAM Icacls OVERRIDE THE NULL FOR INFO FOR THAT
@ECHO -- IF WANTING YOU USE PROCESS TAMER TO SET THE PRIORITY REALTIME OF THE Icacls PROGRAM
@ECHO -- I GET 10% CPU USAGE OUT OF IT THEN TO SPEED UP
@ECHO -- OR I CAN DO THAT WITH START IN CMD COMMAND LANGAUGE BATCH FILE HERE
@ECHO -- LATER FOR THAT ONE
@ECHO -----------------------------------------------------------
@ECHO -- ALSO PUT A TITLE FOR WHEN COMPETE AND WHAT DRIVE IT UPTO
@ECHO -----------------------------------------------------------

@ECHO ------------------------------------------------------
@ECHO -- CREATE A TEMP FILE IN ROOT TO MARK IF A DRIVE EXIST  
@ECHO -- WHICH WILL REMOVE AFTER
@ECHO ------------------------------------------------------

@ECHO -- CREDIT DUE TO IN REF IN PLACE
@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ---- 1ST SOURCE -- LEARN FROM AND WORK OUT YOURSELF SOURCE
@REM ----------------------------------------------------------------
@REM ---- Echo | Windows CMD | SS64.com
@REM ---- http://ss64.com/nt/echo.html
@REM ----
@REM ---- MY SEARCH ----
@REM ---- BATCH FILE ECHO VARIABLE VALUE - Google Search
@REM ----  https://www.google.co.uk/search?num=30&q=batch+file+echo+variable+value&oq=CREATE+BATCH+FILE+WITH+ECHO+COMMAND+AND+PIPE&gs_l=serp.1.5.0i71l8.0.0.0.26979.0.0.0.0.0.0.0.0..0.0....0...1c..64.serp..0.0.0.ooVU2j8aWuc

@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ---- 2ND SOURCE -- LEARN ADD TWO COMMAND ONE LINE
@REM ----------------------------------------------------------------
@REM ---- batch file - How to run two commands in one line in Windows CMD? - Stack Overflow
@REM ---- http://stackoverflow.com/questions/8055371/how-to-run-two-commands-in-one-line-in-windows-cmd
@REM ----
@REM ---- MY SEARCH
@REM ---- TWO CMD COMMANDS ONE LINE - Google Search
@REM ---- https://www.google.co.uk/search?q=TWO+CMD+commands+one+line&oq=TWO+CMD+commands+one+line&aqs=chrome..69i57j0.57098j0j7&sourceid=chrome&ie=UTF-8

@REM ----------------------------------------------------------------
@REM ----------------------------------------------------------------
@REM ---- 3RD SOURCE -- LEARN ECHO WITH PIPE CARTE
@REM ----------------------------------------------------------------
@REM ---- 3RD SOURCE -- SEARCH 1
@REM ----------------------------------------------------------------
@REM ---- HOW USE >NUL WITH ECHO COMMAND - Google Search
@REM ---- https://www.google.co.uk/search?q=HOW+USE+%3ENUL+WITH+ECHO+COMMAND&oq=HOW+USE+%3ENUL+WITH+ECHO+COMMAND&aqs=chrome..69i57.14937j0j7&sourceid=chrome&ie=UTF-8
@REM ---- ----
@REM ---- 3RD SOURCE -- SEARCH 2 
@REM ----------------------------------------------------------------
@REM ---- CREATE BATCH FILE WITH ECHO COMMAND AND PIPE - Google Search
@REM ---- https://www.google.co.uk/search?q=CREATE+BATCH+FILE+WITH+ECHO+COMMAND+AND+PIPE&oq=CREATE+BATCH+FILE+WITH+ECHO+COMMAND+AND+PIPE&aqs=chrome..69i57.25849j0j7&sourceid=chrome&ie=UTF-8
@REM ---- ----
@REM ---- 3RD SOURCE -- SEARCH RESULT 
@REM ----------------------------------------------------------------
@REM ---- Batch files - How To ... Display and Redirect Output
@REM ---- http://www.robvanderwoude.com/battech_redirection.php
@REM ---- ----
@REM ----------------------------------------------------------------

@ECHO ---------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- UK TIME - CHECK - DON'T MATTER CMD PROMPT ALWAYS USA 
@REM -- WHEN AT RECOVERY BOOT
@REM -- @SET VARDATE_BEGIN=%DATE:~10,4%-%DATE:~4,2%-%DATE:~7,2%--%TIME%
@REM ----------------------------------------------------------------

@REM ----------------------------------------------------------------
@REM -- NOTICE NEW COMMAND LEARN IN SECOND COMMAND LINE
@REM -- SYMBOL CARET ^ IN FRONT OF PIPE ^> TELL THE PROCESS 
@REM -- TO IGNORE THE USE OF PIPE
@REM -- AS NOT ABLE USE TWO HERE
@REM -- PIPE THE COMMANDS WITH ECHO COMMAND TO A BATACH FILE
@REM -- TO LOAD FROM THE MORE FLASHY START COMMAND
@REM ----------------------------------------------------------------

@ECHO -- BEGIN TIME -- %DATE%--%TIME%
@REM ----------------------------------------------------------------
@REM -- MAKE TEMP FILE FOR RUN -- 
@REM -- DATE TIME VAR 1 -- 1 OF 2
@REM ----------------------------------------------------------------

@ECHO ---------------------------------------------------------------
@ECHO -- INFO SOME FOR THE MAIN BATCH FILE HERE

@ECHO ---------------------------------------------------------------

@ECHO -- KILL DELETE ALL BATCH FILE BEFORE BEGIN
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@IF EXIST "%%G:\OWNER_SET_TEMP_FILE.BAT" DEL "%%G:\OWNER_SET_TEMP_FILE.BAT")

PAUSE


@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO TESTING DOES EXIST -- %%G:\ -- %DATE%--%TIME% -- & @IF EXIST "%%G:\." @SET VARDATE_BEGIN_JOB1=%DATE%--%TIME% >%%G:\OWNER_SET_TEMP_FILE.BAT)

PAUSE 


@REM ----------------------------------------------------------------
@REM -- SET ICACLS COMMAND -- 1 OF 2
@REM ----------------------------------------------------------------

@REM @FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO ICALCS %%G:\* /setowner "Everyone" /T /C /L ^>NUL ^>>%%G:\OWNER_SET_TEMP_FILE.BAT)

@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO START "%%G:\* - %VARDATE_BEGIN_JOB1% - JOB 1 OF 2 - OWNERSHIP DO BEGIN" /D"%%G:\" /MIN /NORMAL "ICALCS %%G:\* /setowner "Everyone" ^>NUL"  >>%%G:\OWNER_SET_TEMP_FILE.BAT)

@REM ------------------------------------
@REM -- OWNERSHIP -- OBSESSION -- SESSION
@REM ------------------------------------

@REM ----------------------------------------------------------------
@REM -- DATE TIME VAR 2 --  2 OF 2 -- PIPE ADD APPEND
@REM ----------------------------------------------------------------
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@SET VARDATE_BEGIN_JOB2=%DATE%--%TIME% >>%%G:\OWNER_SET_TEMP_FILE.BAT)

@REM ----------------------------------------------------------------
@REM -- SET ICACLS COMMAND -- 2 OF 2 -- PIPE ADD APPEND
@REM ----------------------------------------------------------------

@REM @FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO ICALCS %%G:\* /setowner "Everyone" ^>NUL ^>>%%G:\OWNER_SET_TEMP_FILE.BAT)

@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO START "%%G:\* -  %VARDATE_BEGIN_JOB2% - JOB 2 OF 2 - GRANT PERMISSION DO BEGIN -- %VARDATE_BEGIN_JOB1% - JOB 1 OF 2 1ST BEGAN" /D"%%G:\" /MIN /NORMAL "ICALCS %%G:\* /t  /grant Everyone:F ^>NUL"  >>%%G:\OWNER_SET_TEMP_FILE.BAT)


@REM ----------------------------------------------------------------
@REM -- DATE TIME VAR 3 THE FINISH DATE TIME -- PIPE ADD APPEND
@REM -- INTERESTING LEARN NOW PAST 3 O'CLOCK IN DAY
@REM -- IF I PIPE THE VARIABLE TO A FILE FROM ECHO COMMAND
@REM -- AND THEN IN ORDER TO BE CORRECT THE VARIABLE MUST NOT BEEN USED BEFORE
@REM ----------------------------------------------------------------


@REM ------------------------------------------------------
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@SET VARDATE_JOB_END=%DATE%--%TIME% >>%%G:\OWNER_SET_TEMP_FILE.BAT)
@REM ------------------------------------------------------

@REM -- DO A TITLE ABOUT THE FINISH TIME ------------------
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@TITLE %%G:\* - %VARDATE_JOB_END% - FINISH  -- JOB BEGIN -  %VARDATE_BEGIN_JOB1% -- JOB PART 2 BEGIN - %VARDATE_BEGIN_JOB2% DONE OWNER SETTING FOR %%G:\* >>%%G:\OWNER_SET_TEMP_FILE.BAT)
@REM ------------------------------------------------------
@REM ------------------------------------------------------

@REM -- DONE A TITLE ABOUT THE FINISH TIME ----------------
@REM -- AND DO A ECHO TO BATCH FILE JOB ALSO END RESULT ---
@REM ------------------------------------------------------
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO %%G:\* - %VARDATE_JOB_END% - FINISH  -- DONE OWNER SETTING FOR %%G:\* >>%%G:\OWNER_SET_TEMP_FILE.BAT)
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO %%G:\* - %VARDATE_BEGIN_JOB1% - JOB BEGIN 1 OF 2 TAKE GIVE OWNER - DONE OWNER SETTING FOR %%G:\* >>%%G:\OWNER_SET_TEMP_FILE.BAT)
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO %%G:\* - %VARDATE_BEGIN_JOB2% - JOB BEGIN 2 OF 2 GRANT PERMISSION - DONE OWNER SETTING FOR %%G:\* >>%%G:\OWNER_SET_TEMP_FILE.BAT)
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO %%G:\* - JOB DONE OWNER SETTING FOR %%G:\* >>%%G:\OWNER_SET_TEMP_FILE.BAT)
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO %%G:\*-------------------------------------------------* >>%%G:\OWNER_SET_TEMP_FILE.BAT)



@TITLE JOB END - SET OWNER EVERYONE AND GRANT PERMISSION EVERYONE - ALL DRIVE D:\ TO Z:\
@ECHO ---------------------------------------------------------------
@ECHO -- JOB END -- SET OWNER EVERYONE AND GRANT PERMISSION EVERYONE
@ECHO -- FOR - ALL DRIVE D:\* TO Z:\*
@ECHO -- AND SOME WITH OPPO SYS-OS C DRIVE
@ECHO ---------------------------------------------------------------

@REM --------------------------------------------------------
@REM -- DO DELETE TEMP BATCH FILE --------- WITHIN RUN ITSELF
@REM --------------------------------------------------------
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO %%G:\OWN BATCH TEMP FILE DELETE >>%%G:\OWNER_SET_TEMP_FILE.BAT)
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@DEL %%G:\OWNER_SET_TEMP_FILE.BAT >>%%G:\OWNER_SET_TEMP_FILE.BAT)

@REM --------------------------------------------------------
@REM -- AND A PAUSE TO THE END -- RESULT TO SEE-ER
@REM --------------------------------------------------------
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@PAUSE >>%%G:\OWNER_SET_TEMP_FILE.BAT)
@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@EXIT >>%%G:\OWNER_SET_TEMP_FILE.BAT)

@ECHO ---------------------------------------------------------------
@ECHO -- JOB EXECUTE EVERY BATCH TEMP JOB FILE CREATED
@ECHO -- FOR - ALL DRIVE D:\* TO Z:\*
@ECHO ---------------------------------------------------------------

@REM @FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@SET VARDATE_BEGIN_JOB1=%DATE%--%TIME% & @TITLE %%G:\* -- %VARDATE_BEGIN_JOB1% -- 1 OF 2 -- SET OWNER EVERYONE FOR DRIVE %%G:\* & @IF EXIST "%%G:\OWNER_SET_TEMP_FILE.BAT" %%G:\OWNER_SET_TEMP_FILE.BAT


@PAUSE

@EXIT






