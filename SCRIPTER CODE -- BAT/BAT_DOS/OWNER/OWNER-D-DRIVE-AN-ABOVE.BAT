@ECHO OFF
setlocal EnableDelayedExpansion
@REM ENDLOCAL

@REM -----------------------------------------------------------------------
@REM -- FEW GEM OF EXPLAINED CODE IN HERE AND WORKING
@REM -----------------------------------------------------------------------
@REM -- EXPLAINED THE CODE WILL RUN TWO OPERATION OF THE ICACLS COMMAND ONE 
@REM -- AFTER THE OTHER IN A SUB WINDOW LAUNCHER FORM THE MAIN ROUTINE
@REM -- HAVEN'T FOUND HOW TO LAUNCH MINIMIZED YET AS SUPPOSED TO WITH START COMMAND
@REM -- OH YES FOUND ANSWER NEXT DAY LOOKER SUNDAY MORNING ORDER OF ARGUMENT 
@REM -- IS IMPORTANT WITH START COMMAND

@REM -- BUT START IS ALSO LAUNCHED WITH CMD
@REM -- IT TRICKS ARE MANY VOLUME NAME CAME LAST
@REM -- ALL DOUBLE CHECKING OVER AGAIN IN CASE LAST MINUTE ERROR SET PERMISSION EVERYONE
@REM -- ON WRONG C DRIVE
@REM -- 
@REM -- NOT ONLY VOLUME NAME ALSO DETECT IF DRIVE READY -- GUESS NOT MATTER 
@REM -- A LOT ABOUT IF IT IS A LOCAL -- AS ONLY ALPHABET SCANNER
@REM -- 
@REM -- IT SORT SENDING THE DATE PROPER TO REMOTE CREATED SUB CMD USE %%VAR%%

@REM -- ALL THE SEARCH REFERENCE WHERE I SEARCHED ARE COMPLETELY DOCUMENTED

@REM -- TOOK A TIME -- REFERNCE TO A WORKING PRODUCT

@REM -- THE -- setlocal EnableDelayedExpansion -- IS LEARN THAT 
@REM -- USE IT OFF 
@REM -- AND ON AND WITHIN A SUB ROUTINE WHILE STILL KEEP WITHIN THE MAIN ROUTINE

@REM -----------------------------------------------------------------------
@REM -----------------------------------------------------------------------

@REM -- 2 DAYS OF EFFORT -- GIVE YOU A JUMP HEAD START
@REM -- AND ME BE MY REFERENCE MATERIAL FOR THE BAT LIBRARY
@REM -----------------------------------------------------------------------
@REM ---------------------------------------------------------------------
@REM ---------------------------------------------------------------------
@REM -- NEAR END WORK BIG RUN TO LEAVE GOING
@REM ---------------------------------------------------------------------
@REM -- Fri 26 Aug 2016 15:22:56 -- FIRST RECORDED BEGIN TIME -- BECAUSE ADMIN RIGHT CLEAR UP
@REM -- PROJECT DAY BEFORE BEGIN UNTIL 3AM
@REM -- Sat 27 Aug 2016 10:57:30 -- FIRST RECORDED DATE IN FILE
@REM -- Sat 27 Aug 2016 18:03:49 -- 2 DAY WORK -- BEEN AT IT FULL  OUT
@REM -- Sat 27 Aug 2016 22:08:00 -- NEAR FINISHED TIME ONLY THESE NOTE -- THINK ABOUT PUBLISH IT 
@REM -- DINNER CALLING -- 31 HOUR -- 1 DAY 7 HOUR

@REM ---------------------------------------------------------------------
@REM -- PROBLEM GETTING TEMP FOLDER MY COMPUTER SO HERE NOT FINISHED YET
@REM ---------------------------------------------------------------------
@REM -- ADDED DRIVE NAME TO CREATED BATCH FILE
@REM -- TO RUN OR IF SAME BE OVER WRITE ONE AS WORKING ON ONE
@REM ---------------------------------------------------------------------

@REM -- I WILL ALSO LEAD AHEAD -- LEAP AND BOUND 
@REM -- NOTHING BETTER THAN QUICK CODE IN BATCH LANGUAGE
@REM -----------------------------------------------------------------------
@REM -- SEEM THE INTERNET PRODUCE TALK BUT NEVER REALLY WHAT I'M LOOKING FOR
@REM -- AFTER THEY DRIBBLE OUT SOMETHING THEY CLEARLY MOVE ON TO BETTER
@REM -- TALKING WWW.SUPER USER AND WWW.STACK EXCHANGE -- STUPID USER
@REM -- HELP FILE VS BOOK
@REM -- WHILE SOME ARE EXPERT SHAKING IT
@REM -- WELL KNOCK SPOT OFF THE CHKDSK I DONE BEFORE AND TO REDO
@REM -- NOT THE KAT MOUSE BIT UP PRO STANDARD
@REM -- SOME TALK IN THE LINK GIVEN ABOUT BEING A PROGRAMMER AND USING A IDE TO CODE
@REM -- AND BATCH LANGUAGE SUPPOSED NOT MADE FOR THAT
@REM -- BUT NOTEPAD++ PROGRAMMER MUST OF KNOWN A LOT ABOUT THE SYNTAX
@REM -- AND TALK A HTML PROGRAMMER MOSTLY USE A NOTEPAD++
@REM -- SO IT IS A NOTEPAD MUCH QUICK TO REACH TO AND EDIT THAN A IDE AND COMPILER

@REM -----------------------------------------------------------------------
@REM -- I DID SEE ON-LINE ON NORTON WEBSITE SOMEONE WROTE A SMALL BATCH UTIL
@REM -- AND NORTON DELETED IT FROM THE CONVERSATION TO AUTO RUN UPDATE METHOD
@REM -- AS PROTECTED FOLDER KEEP MOVE AROUND THE EXE EXECUTE
@REM -----------------------------------------------------------------------
@REM -- WELL THIS IS ALL I LEARN IN BATCH FILE LANGUAGE IN TWO DAY IMPROVER
@REM -- NEXT UP WILL BE THE REMOVE SPACE CHARACTER FROM A STRING TO GOT WITH
@REM -- MY SPECIAL FOLDERING -- I'M SURE SOME APPLICATION DON'T LIKE 
@REM -- SPECIAL CHARACTER IN FOLDER NAME -- I'M KNOW OF NIRSOFT WITH IT HELP FILE
@REM -- ANY # HASH IN THE PATH AND WON'T LIKE IT
@REM -- AND WHEN USE COMPUTER NAME AND USER NAME FOR IT -- WANT TO REMOVE THEM
@REM -- OR SOMETHING KEEP TESTING 1 2 3 4 MUSICKAL SCORE

@REM -- WORK TIME SPEND ON PROJECT


@ECHO OFF
@ECHO -----------------------------------------------------------------------------
@ECHO -----------------------------------------------------------------------------
@ECHO -- BEGIN TIME ICACLS -- SET ADMINISTRATOR PERMISSION EVERYONE
@ECHO -- DRIVE D TO Z -- %DATE%--%TIME%
@ECHO -----------------------------------------------------------------------------
@ECHO -----------------------------------------------------------------------------

SET "TEST_MODE_VAR="
@REM -- TEST -- COMPLETE DRY RUN TEST MODE SWITCH
@REM -- MAY NEED WORK ON TO PROVIDE INFO
@REM ----------------------------------------
@REM -- SET TEST_MODE_VAR=YES
@REM ----------------------------------------
@REM -- REM OUT SET LINE ABOVE TO TESTING USE
@REM ----------------------------------------

IF NOT EXIST %HOMEDRIVE%\TEMP (
	MKDIR %HOMEDRIVE%\TEMP
	ECHO #- ------------------------------------------------------------------------
	ECHO #- DIRECTORY MADE WITH MKDIR ON %%HOMEDRIVE%% -- %HOMEDRIVE%\TEMP
	ECHO #- MY TEMP SPECIAL SYSTEM FOLDER NOT WORKING ON THIS COMPUTER COMPLETELY
	ECHO #- DOING THE CREATING THIS CODE WITH
	ECHO #- BE BACK LATER TO BETTER FINISHED OFF CODE
	ECHO #- ------------------------------------------------------------------------
)

@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (

@REM @FOR %%G IN (E,F,G) DO (
@REM -- TEST ON LOW SCALE JOB TEST RUN


	@REM -- CHECK %%G:\ IS A VALID DRIVE LETTER AS NOTICE BEEN BUG DURING TESTING
	@REM -- AND IF IN C DRIVE WILL DO OPPO ICACLS ON THE C DRIVE WHICH NOT WANTER
	@REM -- WHEN %%G HAS A RETURN OF NOTHING I THINK

	IF EXIST %%G:\. (

		@REM ------------------------------------------------------------
		@REM -- GOT SOME BRANCHING SUB ROUTINE NOW FINAL MOMENT OF FINISH
		@REM ------------------------------------------------------------
		set "target_="
		set target_=%%G:

		call :IsDeviceReady !target_! isready_

		IF !isready_!==false (

			call :GetLabel !target_! label_
			echo %%G:\ -- The label of Volume !target_! is !label_!
			echo %%G:\ -- The Device !target_! ready: !isready_!

		)

		IF !isready_!==true (

			@REM echo Device !target_! ready: !isready_!

			@REM PAUSE
			@REM GOTO VERY_ENDE
			@REM -- TEST TO HERE --

			call :GetLabel !target_! label_
			@REM echo The label of Volume !target_! is !label_!
			SET "VOL="
			SET VOL=!label_!


			SET "VARDATE_BEGIN_11="
			SET VARDATE_BEGIN_11=%DATE% -- %TIME%
			SET "TEMP_FILE_VAR="
			SET TEMP_FILE_VAR=%HOMEDRIVE%\TEMP\OWNER_SET_TEMP_FILE--%%G.BAT


			@REM ---------------------------------------------------------------------
			@REM -- PART 01 OF 02 -- OWNERSHIP ---------------------------------------
			@REM ---------------------------------------------------------------------

			@ECHO #- DRIVE %%G:\ -- PROCESS BEGIN -- !VARDATE_BEGIN_JOB11! -- ICACLS -- JOB ALL DRIVE ABOVE
			@ECHO #- SET OWNERSHIP AND GRANT PERMISSION

			IF "!TEST_MODE_VAR!"=="" (

				ECHO @ECHO OFF >!TEMP_FILE_VAR!
				ECHO SET VARDATE_BEGIN_JOB11=%%DATE%% -- %%TIME%% >>!TEMP_FILE_VAR!
				ECHO TITLE "%%G:\---!vol! -- ICACLS -- OWNERSHIP -- SECTION 01 OF 02 -- IN PROGRESS %%G:\---!vol!">>!TEMP_FILE_VAR!
				ECHO ECHO # BEGIN ------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO # 1....................................................................... >>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- %%VARDATE_BEGIN_JOB11%% -- 01 OF 02 JOB BEGIN ICACLS -- OWNERSHIP>>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO # COMMAND LINE -- ICACLS %%G:\ /Setowner Everyone  /T /C /L /Q >>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!

				ECHO ECHO. >>!TEMP_FILE_VAR!
				@REM ----------------------- DISPLAY COMMAND TO BE EXECUTED
				ECHO ECHO START "" /D"%%G:\" /B /MIN /HIGH /WAIT ICACLS %%G:\ /setowner Everyone /T /C /L /Q >>!TEMP_FILE_VAR!

				@REM ECHO PAUSE>>!TEMP_FILE_VAR!
				@REM -- TEST --

				SET "TEST_FLAG_SECURE="
				@FOR %%D IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (
					IF %%G==%%D (
						@REM ---------------------------------------------------------------------------
						ECHO START "" /D"%%G:\" /B /MIN /HIGH /WAIT ICACLS %%G:\ /setowner Everyone /T /C /L /Q ^>NUL >>!TEMP_FILE_VAR!
						@REM ----------------------- EXECUTE COMMAND
						SET TEST_FLAG_SECURE=TRUE
						@REM ---------------------------------------------------------------------------
					)
				)

				IF NOT !TEST_FLAG_SECURE!==TRUE (
					@REM ---------------------------------------------------------------------------
					ECHO.
					ECHO WAS A FAULT WITH THE DRIVE LETTER SHOWN AS %%G:\ NOT RECOGNISED AS EXPECTED
					ECHO OPERATION ARE SKIPPED WITH THEM -- SOMETHING CLEARED THE VARIABLE
					@REM ---------------------------------------------------------------------------
					
					ECHO ECHO. >>!TEMP_FILE_VAR!
					ECHO ECHO WAS A FAULT WITH THE DRIVE LETTER SHOWN AS %%G:\ NOT RECOGNISED AS EXPECTED >>!TEMP_FILE_VAR!
					ECHO ECHO OPERATION ARE SKIPPED WITH THEM -- SOMETHING CLEARED THE VARIABLE >>!TEMP_FILE_VAR!
					@REM ---------------------------------------------------------------------------
				)

				@REM -----------------------
				ECHO ECHO. >>!TEMP_FILE_VAR!
				@REM -----------------------

				@REM ----------------------------------------------------------------
				@REM -- INFO BLOCK EXPLAIN PROGRESS
				@REM ----------------------------------------------------------------
				@REM -- I LOOK A IDIOT -- TO SET THE DATE IN A VAR WITH ECHO AND PIPE
				@REM -- THE DATE MUST NOT OF ALREADY BEEN SET BE THE VAR - HA HA HA
				@REM -- TRY AND REUSE IT -- RSI INJURY
				@REM -- REF: DETAILS ABOUT GOOGLE SEARCHING WITHOUT ANY ANSWER
				@REM -- UNTIL REMEMBER
				@REM -- WITH FRANTIC HOUR SPENT ON IT AT BOTTOM
				@REM -- STILL PROBLEM TIME DATE IS VARIABLE ALREADY SET
				@REM -- Sat 27-Aug-2016 11:54:47
				@REM -- Sat 27-Aug-2016 13:18:42
				@REM ----------------------------------------------------------------
				@REM -- FOUND ANSWER FROM WEB PAGE
				@REM -- ----
				@REM -- Batch files - Escape Characters
				@REM -- http://www.robvanderwoude.com/escapechars.php
				@REM -- ----
				@REM -- MUST USE A DOUBLE %% FOR VARIABLE PIPE OUTPUT
				@REM -- ANOTHER LONG TIME -- EVEN LONGER EXTRA 2 HOUR OF DAY FIND THAT BIT
				@REM -- Sat 27-Aug-2016 13:18:42
				@REM -- Sat 27-Aug-2016 15:36:50
				@REM ----------------------------------------------------------------

				ECHO SET VARDATE_END_JOB12=%%DATE%% -- %%TIME%% >>!TEMP_FILE_VAR!

				ECHO ECHO # OUTPUT RESULT AFTER EXECUTE 01 OF 02 >>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- %%VARDATE_BEGIN_JOB11%% -- JOB BEGAN -- 01 OF 02 -- OWNERSHIP>>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- %%VARDATE_END_JOB12%% -- JOB DONE  -- 01 OF 02 -- OWNERSHIP>>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- COMPLETED JOB OF ICACLS    -- SECTION   -- 01 OF 02 OF DRIVE %%G:\>>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!

				ECHO ECHO # COMMAND LINE -- ICACLS %%G:\ /Setowner Everyone  /T /C /L /Q >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO START "" /D"%%G:\" /B /MIN /HIGH /WAIT ICACLS %%G:\ /setowner Everyone  /T /C /L /Q >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!

			)

			SET "VARDATE_BEGIN_21="
			SET VARDATE_BEGIN_21=%DATE% -- %TIME%

			@REM ----------------------------------------------------------------------------
			@REM -- PART 02 OF 02 -- GRANT PERMISSION ---------------------------------------
			@ECHO #- %%G:\ -- BEGIN -- %DATE% -- %TIME% -- ICACLS -- JOB 02 OF 02 - GRANT PERMISSION

			IF "!TEST_MODE_VAR!"=="" (

				ECHO SET VARDATE_BEGIN_JOB21=%%DATE%% -- %%TIME%% >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO # 2....................................................................... >>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- %%VARDATE_BEGIN_JOB21%% -- 02 OF 02 JOB BEGIN ICACLS -- GRANT PERMISSION>>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO # COMMAND LINE -- ICACLS %%G:\ /T /Grant Everyone:F >>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!

				ECHO TITLE "%%G:\---!vol! -- ICACLS -- GRANT EVERYONE -- SECTION 02 OF 02 -- IN PROGRESS %%G:\---!vol!" >>!TEMP_FILE_VAR!

				ECHO ECHO. >>!TEMP_FILE_VAR!
				@REM ----------------------- DISPLAY COMMAND TO BE EXECUTED

				SET "TEST_FLAG_SECURE="
				@FOR %%D IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (
					IF %%G==%%D (
						@REM ---------------------------------------------------------------------------
						ECHO ECHO START "" /D"%%G:\" /B /MIN /HIGH /WAIT ICACLS %%G:\ /T /L /C /Q /Grant Everyone:F >>!TEMP_FILE_VAR!
						@REM ----------------------- EXECUTE COMMAND
						SET TEST_FLAG_SECURE=TRUE
						@REM ---------------------------------------------------------------------------
					)
				)

				IF NOT !TEST_FLAG_SECURE!==TRUE (
					@REM ---------------------------------------------------------------------------
					ECHO ECHO.
					ECHO WAS A FAULT WITH THE DRIVE LETTER SHOWN AS %%G:\ NOT RECOGNISED AS EXPECTED
					ECHO OPERATION ARE SKIPPED WITH THEM -- SOMETHING CLEARED THE VARIABLE
					@REM ---------------------------------------------------------------------------
					ECHO ECHO. >>!TEMP_FILE_VAR!
					ECHO ECHO WAS A FAULT WITH THE DRIVE LETTER SHOWN AS %%G:\ NOT RECOGNISED AS EXPECTED >>!TEMP_FILE_VAR!
					ECHO ECHO OPERATION ARE SKIPPED WITH THEM -- SOMETHING CLEARED THE VARIABLE >>!TEMP_FILE_VAR!
					@REM ---------------------------------------------------------------------------
				)



				@REM ECHO PAUSE>>!TEMP_FILE_VAR!
				@REM -- TEST --

				@REM -----------------------
				@REM -- --- WORK TO DO SUPPOSED TO LAUNCH AS A MINIMIZED WINDOW WITH /MIN
				@REM -- THE /Q ARGUMENT OPTION - DON'T REALLY DO MUCH TO SUPPRESS SUCCESS RESULT
				@REM -- USE A NULL WITH CARET ALSO ^>NULL
				@REM -- MORE WORK MAKE DO EACH FOLDER OF A DRIVE TO FILTER OUT DOING SOME 
				@REM -- FOLDER LIKE RECYCLE BIN OR IT CORRUPT IT
				@REM -- THAT IS WHAT I GOT FOR C DRIVE
				@REM --
				
				@REM -- THAT MAY MEAN LOT MORE SUB MENU -- POSSIBLE MAKE THEM EXIT WHEN DONE
				@REM -- REM OUT PAUSE AT END
				@REM -- BUT MISS TIME IT TOOK -- UNLESS APPEND RESULT TO A FILE TOTAL
				@REM --
				@REM -----------------------
				ECHO START "" /D"%%G:\" /B /MIN /HIGH /WAIT ICACLS %%G:\ /T /L /C /Q /Grant Everyone:F ^>NUL >>!TEMP_FILE_VAR!
				@REM -----------------------
				ECHO ECHO. >>!TEMP_FILE_VAR!
				@REM -----------------------

				@REM --------------------------------------------------------
				@REM -- SET VARIABLE AT CREATE BATCH NOT A ECHO VISIBLE THING
				@REM --------------------------------------------------------
				ECHO SET VARDATE_END_JOB22=%%DATE%% -- %%TIME%% >>!TEMP_FILE_VAR!
				@REM --------------------------------------------------------

				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!

				@REM ---------------------------
				@REM -- RESULT OUTPUT END RESULT
				@REM ---------------------------
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO   Note: >>!TEMP_FILE_VAR!
				ECHO ECHO   Sids may be in either numerical or friendly name form. If a numerical >>!TEMP_FILE_VAR!
				ECHO ECHO   form is given, affix a * to the start of the SID. >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO   /T indicates that this operation is performed on all matching >>!TEMP_FILE_VAR!
				ECHO ECHO   files/directories below the directories specified in the name. >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO   /C indicates that this operation will continue on all file errors. >>!TEMP_FILE_VAR!
				ECHO ECHO   Error messages will still be displayed. >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO   /L indicates that this operation is performed on a symbolic link >>!TEMP_FILE_VAR!
				ECHO ECHO   itself versus its target. >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO   /Q indicates that icacls should suppress success messages. >>!TEMP_FILE_VAR!
				ECHO ECHO   THE SOURCE THE DOS HELP HELP ICACLS /? >>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!

				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO # RESULT OUTPUT END RESULT >>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!

				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- %%VARDATE_BEGIN_JOB11%% -- JOB BEGAN -- 01 OF 02 -- OWNERSHIP>>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- %%VARDATE_END_JOB12%% -- JOB DONE  -- 01 OF 02 -- OWNERSHIP>>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- COMPLETED JOB OF ICACLS    -- SECTION   -- 01 OF 02 -- DRIVE %%G:\>>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO # AFTER RESULT ---- >>!TEMP_FILE_VAR!
				ECHO ECHO # COMMAND LINE -- ICACLS %%G:\ /Setowner Everyone /T /L /C /Q >>!TEMP_FILE_VAR!
				ECHO ECHO # COMMAND LINE -- >>!TEMP_FILE_VAR!
				ECHO ECHO # START "" /D"%%G:\" /B /MIN /HIGH /WAIT /NORMAL ICACLS %%G:\ /setowner Everyone /T /L /C /Q >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- %%VARDATE_BEGIN_JOB21%% -- JOB BEGAN -- 02 OF 02 -- GRANT PERMISSION>>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- %%VARDATE_END_JOB22%% -- JOB DONE  -- 02 OF 02 -- GRANT PERMISSION>>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- COMPLETED JOB OF ICACLS    -- SECTION   -- 02 OF 02 -- DRIVE %%G:\>>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO # AFTER RESULT ---- >>!TEMP_FILE_VAR!
				ECHO ECHO # COMMAND LINE -- ICACLS %%G:\ /T /L /C /Q /Grant Everyone:F >>!TEMP_FILE_VAR!
				ECHO ECHO # COMMAND LINE -->>!TEMP_FILE_VAR!
				ECHO ECHO # START "" /D"%%G:\" /B /MIN /HIGH /WAIT /NORMAL ICACLS %%G:\ /T /L /C /Q /Grant Everyone:F >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- TIME JOB PAIR BOTH DOUBLE TOTAL ---- OWNERSHIP AND GRANT PERMISSION>>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- TOTAL %%VARDATE_BEGIN_JOB11%% 01 OF 02 -- OWNERSHIP>>!TEMP_FILE_VAR!
				ECHO ECHO # %%G:\ -- TOTAL %%VARDATE_END_JOB22%% 02 OF 02 -- GRANT PERMISSION>>!TEMP_FILE_VAR!
				ECHO ECHO # ------------------------------------------------------------------------ >>!TEMP_FILE_VAR!

				ECHO ECHO # BOTH PAIR OF JOB EXECUTE FINISHED >>!TEMP_FILE_VAR!
				ECHO ECHO # OUTPUT RESULT FROM COMMAND ARE SUPPRESS SUCCESS RESULT ONLY ERROR RETURNED >>!TEMP_FILE_VAR!
				ECHO ECHO # WITHER -- ICACLS /Q OPTION >>!TEMP_FILE_VAR!

				ECHO ECHO. >>!TEMP_FILE_VAR!
				ECHO ECHO # PRESS KEY TO EXIT >>!TEMP_FILE_VAR!
				ECHO ECHO. >>!TEMP_FILE_VAR!

				ECHO TITLE "%%G:\---!vol! -- ICACLS OWNERSHIP -- COMPLETER -- %%G:\---!vol!" >>!TEMP_FILE_VAR!

				@REM --- END ASK WAIT -------------------------
				ECHO PAUSE >>!TEMP_FILE_VAR!
				ECHO EXIT >>!TEMP_FILE_VAR!
				@REM ------------------------------------------

				@REM ------------------------------------------
				@REM ------------------------------------------
				@REM ------------------------------------------
				@REM ------------------------------------------
				@REM ---- RUN EXECUTE --------------------------------------------
				IF EXIST "!TEMP_FILE_VAR!" START CMD /K !TEMP_FILE_VAR!
				@REM -- RUN IN SEQUENCE ONE PAIR AT A TIME
				@REM ------------------------------------------
				@REM ------------------------------------------
				@REM ------------------------------------------
				@REM ------------------------------------------

			)

			ECHO #- ------------------------------------------------------------------------
			ECHO #- SET OWNERSHIP AND GRANT PERMISSION -- JOB LAUNCHED IN OWN PROCESS
			ECHO #- ------------------------------------------------------------------------
			ECHO #- %%G:\!vol! -------- %DATE% -- %TIME%
			ECHO #- ------------------------------------------------------------------------
			ECHO START "" /D"%%G:\" /B /MIN /HIGH /WAIT /NORMAL ICACLS %%G:\ /setowner Everyone /T /L /C /Q
			ECHO START "" /D"%%G:\" /B /MIN /HIGH /WAIT /NORMAL ICACLS %%G:\ /T /L /C /Q /Grant Everyone:F
			ECHO.

			@REM -- TEST -- PAUSE
			@REM -- TEST -- GOTO VERY_ENDE
		)
		@Rem -- INFO -- IF STATMENT BLOCK WITH -- !isready_!
	)
	@Rem -- INFO -- IF STATMENT BLOCK WITH -- EXIST %%G:\.
)
@REM -- INFO -- IF STATMENT BLOCK WITH -- @FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (



@ECHO #- ------------------------------------------------------------------------
@ECHO #- ------------------------------------------------------------------------
@ECHO #- -- ALL JOB SET LAUNCHER D - Z
@ECHO #- -- DONE ICACLS D TO Z ------- %DATE%--%TIME%
@ECHO #- ------------------------------------------------------------------------
@ECHO #- -- IF THIS PROGRAM IS RUN AS OR IN CMD COMMAND SHELL WITH ADMINISTRATOR PRIV
@ECHO #- -- AND THEN ALL LAUNCH OF OTHER PROGRAM WILL BE IN  ADMINISTRATOR PRIV ELEVATED
@ECHO #- -- ACCORDING TO RESULT SHOWN IN TASK MANAGER
@ECHO #- -- AND ALSO
@ECHO #- -- ALSO SETTING THE TITLE OF THE COMMAND WINDOW WILL REVEAL
@ECHO #- -- IT IS IN ADMINISTRATOR
@ECHO #- ------------------------------------------------------------------------

GOTO VERY_ENDE

@REM ----------------------------------------
@REM -- COUPLE OF LITTLE SUBROUTINE TO FOLLOW
@REM -- IsDeviceReady AND GetLabel
@REM ----------------------------------------

@REM -- SORTED -- Sat 27 August 2016 19:10:49
@REM -- ----
@REM -- windows - How to read the label of a drive or volume in a batch file? - Stack Overflow
@REM -- http://stackoverflow.com/questions/8649934/how-to-read-the-label-of-a-drive-or-volume-in-a-batch-file
@REM -- ----
@REM -- ---- ORIGINAL SOURCE
@REM -- TSCMD101 FAQ: How to capture Volume Serial Number into an environment variable? Assorted NT/2000/XP/.. CMD.EXE script tricks written by Timo Salmi
@REM -- http://www.netikka.net/tsneti/info/tscmd101.htm#label
@REM -- ----

@REM -- set target_=D:
@REM -- ::
@REM -- call :IsDeviceReady %target_% isready_
@REM -- echo Device %target_% ready: %isready_%
@REM -- if /i "%isready_%"=="false" (endlocal & goto :EOF)
@REM -- ::
@REM -- call :GetLabel %target_% label_
@REM -- echo The label of Volume %target_% is %label_%
@REM endlocal & goto :EOF

@REM setlocal and endlocal -- MAKE IT WORK PROPER

::
:IsDeviceReady
setlocal
set "ready_="
set ready_=true
dir "%~1" > nul 2>&1
if %errorlevel% NEQ 0 set ready_=false
endlocal & set "%2=%ready_%" & goto :EOF
::
:GetLabel
setlocal
set "label_="
for /f "tokens=5*" %%a in (
  'vol "%~1"^|find "Volume in drive "') do (
    set label_=%%b)
endlocal & set "%2=%label_%" & goto :EOF

@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------


@REM -----------------------------------------------------
@REM -- EXAMPLE REDUNDANT CODE BEEN PLAYING WITH TIDY AWAY
@REM -----------------------------------------------------
@REM ECHO ECHO SET VARDATE_BEGIN_JOB2=%DATE% -- %TIME% >>!TEMP_FILE_VAR!
@REM ECHO ECHO SET VARDATE_BEGIN_JOB2=%DATE:~10,4%-%DATE:~4,2%-%DATE:~7,2%--%TIME% >>!TEMP_FILE_VAR!
@REM for /f "delims=" %%A in ('echo hello') do set "var=%%A"
@REM echo %var%
@REM ECHO ECHO for /f "tokens=1,2" %%u in ('date /t') set d=%%u >>!TEMP_FILE_VAR!
@REM ECHO ECHO for /f "tokens=1" %%u in ('time /t') set t=%%u >>!TEMP_FILE_VAR!
@REM ECHO ECHO %D% >>!TEMP_FILE_VAR!
@REM ECHO ECHO %T% >>!TEMP_FILE_VAR!
@REM if "%t:~1,1%"
@REM ECHO "For /f "tokens=1,2,3,4,5 delims=/. " %%a in ('date/T') do set CDate=%%a%%b%%c%%d" >>!TEMP_FILE_VAR!
@REM ECHO ECHO # JOB DONE -- %%G:\ - [%DATE% %TIME%] >>!TEMP_FILE_VAR!
@REM -------------------------------------------

@REM --
@REM -- ----
@REM -- BATCH FILE GET VOLUME LABEL - Google Search
@REM -- https://www.google.co.uk/search?num=30&q=BATCH+FILE+GET+VOLUME+LABEL&spell=1&sa=X&ved=0ahUKEwijld7zi-LOAhWBtBoKHWoZBvYQvwUIGygA&biw=1229&bih=558
@REM -- ----
@REM -- windows - How to read the label of a drive or volume in a batch file? - Stack Overflow
@REM -- http://stackoverflow.com/questions/8649934/how-to-read-the-label-of-a-drive-or-volume-in-a-batch-file
@REM -- ----
@REM SET "vol="
@REM for /f "tokens=1-5*" %%G in ('vol') do (
@REM 	set vol=%%6 & goto done
@REM )
@REM :done
@REM echo !vol!



@REM -- ----
@REM -- Append text to a text file ON THE SAME LINE using command line - Google Groups
@REM -- https://groups.google.com/forum/#!topic/alt.msdos.batch/083me7gc2ek
@REM -- ----
@REM -- cmd /c echo|set /p="# JOB DONE -- %DATE%">>!TEMP_FILE_VAR!
@REM -- USEFUL FOUND -- STILL NOT WORK WITH DATE VAR
@REM -- Sat 27 August 2016 14:40:31


@REM ----------------------------------------
@REM -- Sat 27 August 2016 10:57:30
@REM ----------------------------------------
@REM -- NEW CODE TO MAKE WORK BETTER
@REM -- TO EXECUTE TWO COMMAND ON ONE COMMAND LINE WITH START COMMAND MUST REQUIRE A BATCH FILE
@REM ----------------------------------------
@REM OLD CODE 01 -- START "%%G:\ - %VARDATE_BEGIN_JOB1% - JOB 1 OF 2 - OWNERSHIP DO BEGIN" /D"%%G:\" /HIGH "ICACLS" %%G:\ /setowner Everyone /T /C /L /Q >NUL
@REM OLD CODE 02 --


@REM -- RUN PROBLEM
@REM -- ----
@REM -- How to fix error “No mapping between account names and security IDs” in Active Directory - RebelAdmin
@REM -- http://www.rebeladmin.com/2016/01/how-to-fix-error-no-mapping-between-account-names-and-security-ids-in-active-directory/
@REM -- ----
@REM -- STAGE IN THIS FIX DEBUG LEVEL 2 SET BUT WAIT TOO LONG
@REM -- TEMP WORKAROUND USE CMD RATHER THAN START
@REM -- ----
@REM -- GOT OVER THE ERROR THE SYNTAX OF LINE NOT CORRECT WITH A EXTRA NOT WANTER " QUOTATION SYMBOL
@REM -- ----
@REM CMD /C "ICACLS" %%G:\ /setowner "Everyone >NUL
@REM CMD /C "ICACLS" "%%G:\ /t /grant Everyone:F" >NUL
@REM -- ----





PAUSE
PAUSE
PAUSE
PAUSE



















@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -- TOOL TO DETECT FILE SIZE RESULT IS NUMERIC 0 SIZE
@REM -- BUT NOT USE NOW WORK OUT MUCH BETTER METHOD DETECT DRIVE EXIST
@REM -----------------------------------------------------------------
@REM -- DEL C:\TEMP\TEST.TXT
@REM -- DIR E:\* /A-D-H-S-R /B >C:\TEMP\TEST.TXT
@REM -- IF EXIST C:\TEMP\TEST.TXT (
	@REM SET V1=findstr /R /N "^" C:\TEMP\TEST.TXT | find /C ":"

@REM -- 	IF findstr /R /N "^" C:\TEMP\TEST.TXT | find /C ":" EQU 0 ECHO HELLO

@REM -- 	ECHO.%%F | FIND /I "\Start Menu">Nul && Set "VAR_FLAG=" && SET VAR_FLAG=YES

@REM -- 	ECHO.0 | findstr /R /N "^" C:\TEMP\TEST.TXT | find /C ":" EQU 0 ECHO HELLO
@REM -- 	findstr /R /N "^" C:\TEMP\TEST.TXT | find /C ":" EQU 0 ECHO HELLO
@REM -- )

@REM -----------------------------------------------------------------
@REM -- TOOL TO DETECT FILE SIZE RESULT IS NUMERIC 0 SIZE
@REM -- BUT NOT USE NOW WORK OUT MUCH BETTER METHOD DETECT DRIVE EXIST
@REM -----------------------------------------------------------------
@REM -- ----
@REM -- How do I Detect (and Delete) a File if it is Empty using a Windows Batch File? - Stack Overflow
@REM -- http://stackoverflow.com/questions/9378972/how-do-i-detect-and-delete-a-file-if-it-is-empty-using-a-windows-batch-file
@REM -- --------
@REM -- windows - How to check if a file exists from inside a batch file - Stack Overflow
@REM -- http://stackoverflow.com/questions/4340350/how-to-check-if-a-file-exists-from-inside-a-batch-file
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------



	@REM IF "%%F"=="Windows" SET VAR_FLAG=NOT


	@REM ECHO.C:\TEMP\TEST.TXT | FIND /V 0>Nul && Set "VAR_FLAG=" && SET VAR_FLAG=YES

	@REM ECHO !VAR_FLAG!




@REM ECHO VALUE %V1%


REM if %~z1 eq 0 del %1

REM @IF EXIST "%%G:\OWNER_SET_TEMP_FILE.BAT" DEL "%%G:\OWNER_SET_TEMP_FILE.BAT")

REM ECHO.%%G | FIND /I "\Start Menu">Nul && Set "VAR_FLAG=" && SET VAR_FLAG=YES
REM if %~z1 eq 0 del %1




PAUSE
PAUSE
PAUSE
PAUSE



@REM -----------------------------------------------------------------------
@REM -----------------------------------------------------------------------
@REM ----------------------------------------------------------------------
@REM ----------------------------------------------------------------------
@REM -- NEW PROJECT ABOVE -- REFERENCE NOTE ALL GIVEN AT END DOCUMENT
@REM -- GOOGLE SEARCH -- OF LEARNING CURVE CREDIT DUE GIVEN
@REM -- AND HELP SYSTEM ON-LINE -- WITHOUT PAPER BOOK
@REM -- THINK BEFORE YOU PRINT -- AND AFTER SCAN RECYCLE PAPER
@REM ----------------------------------------------------------------------
@REM -- IMPROVER FROM MACH 2 VERSION -- FROM -- MACH 1 VERSION THAT FOR SURE
@REM -- MACH 2 VERSION OF SET OWNERSHIP WITH ICACLS.EXE
@REM -- AFTER LEARN ABOUT THE ONE COMMAND
@REM ----------------------------------------------------------------------
@REM ----------------------------------------------------------------------
@REM -- MACH 1 VERSION DATE REFERENCE FOUND
@REM ----------------------------------------------------------------------
@REM -- LEARN METHOD JULY 13 MID-NIGHT ---
@REM -- CREDIT DUE TO IN REF IN PLACE
@REM -- CREATE THIS BATCH FILE 1:40 PM 13 JULY WEDNESDAY ---
@REM -- CREATE THIS BATCH FILE 4:15 PM 13 JULY WEDNESDAY --- FINISH
@REM -----------------------------------------------------------------------
@REM -- QUICK LITTLE KNOCK UP -- LONGER THIS TIME AROUND MACH 2 VERSION
@REM -- ALL WEEKEND
@REM -----------------------------------------------------------------------
@REM -----------------------------------------------------------------------
@REM -----------------------------------------------------------------------



@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -- NEW MACH 2 VERSION OF SET OWNERSHIP -- ABOVE
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------



@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -- PREVIOUS EARLY 1ST -- MACH 1 VERSION OF SET OWNERSHIP
@REM -- AFTER LEARN ONE COMMAND
@REM -- BELOW -- MOST LOOK AS THOUGHT LEFT ORIGINAL WITHOUT
@REM -- REM LINE PUSHED ON SO SEE WORKING
@REM -- NEXT MACH VERSION MIGHT WRITE TO A NEW FILE
@REM -----------------------------------------------------------------
@REM -- AND REFERENCING FOLLOW AT END
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------
@REM -----------------------------------------------------------------




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
@ECHO -- RUN THIS BATCH FILE ADMINISTRATOR PERMISSION
@ECHO -- TO DO IT ALWAYS CREATE A SHORT-CUT TO THIS BATCH FILE AND PROPERTIES SET

@ECHO -- MY PERSONAL SET-UP IS I SET ALL DRIVE WITH OWNER AND FULL PERMISSION
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
@ECHO -- OR I CAN DO THAT WITH START IN CMD COMMAND LANGUAGE BATCH FILE HERE
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

@REM @ECHO -- KILL DELETE ALL BATCH FILE BEFORE BEGIN

@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (

	DEL C:\TEMP\TEST.TXT
	DIR %%G\* /A-D-H-S-R /B >C:\TEMP\TEST.TXT
	IF EXIST C:\TEMP\TEST.TXT (
	SET V1=findstr /R /N "^" C:\TEMP\TEST.TXT | find /C ":"




if %~z1 eq 0 del %1

@IF EXIST "%%G:\OWNER_SET_TEMP_FILE.BAT" DEL "%%G:\OWNER_SET_TEMP_FILE.BAT")

ECHO.%%G | FIND /I "\Start Menu">Nul && Set "VAR_FLAG=" && SET VAR_FLAG=YES
if %~z1 eq 0 del %1

@REM ----
@REM How do I Detect (and Delete) a File if it is Empty using a Windows Batch File? - Stack Overflow
@REM http://stackoverflow.com/questions/9378972/how-do-i-detect-and-delete-a-file-if-it-is-empty-using-a-windows-batch-file
@REM ----


PAUSE


@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO TESTING DOES EXIST -- %%G:\ -- %DATE%--%TIME% -- & @IF EXIST "%%G:\." @SET VARDATE_BEGIN_JOB1=%DATE%--%TIME% >%%G:\OWNER_SET_TEMP_FILE.BAT)

PAUSE


@REM ----------------------------------------------------------------
@REM -- SET ICACLS COMMAND -- 1 OF 2
@REM ----------------------------------------------------------------

@REM @FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO ICALCS %%G:\* /setowner "Everyone" /T /L /C /Q ^>NUL ^>>%%G:\OWNER_SET_TEMP_FILE.BAT)

@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO START "%%G:\* - %VARDATE_BEGIN_JOB1% - JOB 1 OF 2 - OWNERSHIP DO BEGIN" /D"%%G:\" /MIN /HIGH "ICALCS %%G:\* /setowner "Everyone" ^>NUL"  >>%%G:\OWNER_SET_TEMP_FILE.BAT)

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

@FOR %%G IN (D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z) DO (@ECHO START "%%G:\* -  %VARDATE_BEGIN_JOB2% - JOB 2 OF 2 - GRANT PERMISSION DO BEGIN -- %VARDATE_BEGIN_JOB1% - JOB 1 OF 2 1ST BEGAN" /D"%%G:\" /MIN /NORMAL "ICALCS %%G:\* /T /L /C /Q /grant Everyone:F ^>NUL"  >>%%G:\OWNER_SET_TEMP_FILE.BAT)


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



@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/--------------------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/--------------------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@REM ----------------------------------------------------------------------------
@REM -- THE END AND REFERENCE SECTION -------------------------------------------
@REM -- BYE BYE -----------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/--------------------\/\/\/\/\/\/\/\/\/\/\/\/\/\
@ECHO /\/\/\/\/\/\/\/\/\/\/\/\/\/--------------------\/\/\/\/\/\/\/\/\/\/\/\/\/\


@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM THE 1ST GREAT HALF OF GOOGLE SEARCHER
@REM BUT NOT THE GREATER HALF THAT IS IN THE 2ND SECTION
@REM 17 SEARCH TERM FOUND IN 1ST ---- 2ND SECTION 38 SEARCH TERM -- RUN AT IT
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------

@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ALL THE SEARCH AND PAGE WITHOUT ANSWER FROM SEARCH THEM PIPE CHARACTER
@REM ALL THE EFFORT SOME PAGES WHEN TO TO COVE THE WHOLE PICTURE
@REM AND FOUND A FEW INTERESTING
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM THE COUNT IS HOW MANY CLIPBOARD I DONE SINCE WEDNESDAY
@REM OR RELOAD THE PROGRAM CLIPBOARD LOGGER
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM
@REM Count = 661 -- Sat 27-Aug-2016 12:16:37
@REM Batch files - DATE and TIME - Google Chrome
@REM ----
@REM Batch files - DATE and TIME
@REM http://www.robvanderwoude.com/datetime.php
@REM ----
@REM Count = 662 -- Sat 27-Aug-2016 12:18:51
@REM How to Create FileName Variable with Date And Time From Command CMD Batch File For Use in Creating Unique Filename – ITProGuru Blog - Google Chrome
@REM ----
@REM How to Create FileName Variable with Date And Time From Command CMD Batch File For Use in Creating Unique Filename – ITProGuru Blog
@REM http://itproguru.com/expert/2013/02/how-to-create-filename-variable-with-date-and-time-from-command-cmd-batch-file-for-use-in-creating-unique-filename/
@REM ----
@REM Count = 663 -- Sat 27-Aug-2016 12:20:00
@REM MS-DOS Commands - Google Chrome
@REM ----
@REM MS-DOS Commands
@REM http://www.sprint.net.au/~terbut/usefulbox/msdoscmds.htm
@REM ----
@REM Count = 664 -- Sat 27-Aug-2016 12:34:31
@REM *C:\BAT\OWNER\OWNER-D-DRIVE-AN-ABOVE.BAT - Notepad++
@REM SET "VARDATE_BEGIN_JOB1="
@REM Count = 665 -- Sat 27-Aug-2016 12:35:04
@REM *C:\BAT\OWNER\OWNER-D-DRIVE-AN-ABOVE.BAT - Notepad++
@REM ECHO SET VARDATE_BEGIN_JOB2=%DATE% -- %TIME%
@REM Count = 666 -- Sat 27-Aug-2016 12:37:17
@REM Batch files - DATE and TIME - Google Chrome
@REM ----
@REM Batch files - DATE and TIME
@REM http://www.robvanderwoude.com/datetime.php
@REM ----
@REM Count = 667 -- Sat 27-Aug-2016 12:38:10
@REM How to get date / time in batch file - Google Chrome
@REM ----
@REM How to get date / time in batch file
@REM http://www.windows-commandline.com/get-date-time-batch-file/
@REM ----
@REM Count = 668 -- Sat 27-Aug-2016 12:39:50
@REM Batch Guide - Google Chrome
@REM ----
@REM Batch Guide
@REM http://www.infionline.net/~wtnewton/batch/batguide.html
@REM ----
@REM Count = 669 -- Sat 27-Aug-2016 12:45:51
@REM Windows batch assign output of a program to a variable - Stack Overflow - Google Chrome
@REM ----
@REM Windows batch assign output of a program to a variable - Stack Overflow
@REM http://stackoverflow.com/questions/2323292/windows-batch-assign-output-of-a-program-to-a-variable
@REM ----
@REM Count = 670 -- Sat 27-Aug-2016 12:47:12
@REM Windows batch file: set output of program to a variable? - Command Prompt - Windows XP - Google Chrome
@REM ----
@REM Windows batch file: set output of program to a variable? - Command Prompt - Windows XP
@REM http://www.tomshardware.co.uk/forum/230090-45-windows-batch-file-output-program-variable
@REM ----
@REM Count = 671 -- Sat 27-Aug-2016 12:49:20
@REM Redirect pipe to a variable in Windows batch file - Super User - Google Chrome
@REM ----
@REM Redirect pipe to a variable in Windows batch file - Super User
@REM http://superuser.com/questions/404737/redirect-pipe-to-a-variable-in-windows-batch-file
@REM ----
@REM Count = 672 -- Sat 27-Aug-2016 12:50:01
@REM How to redirect output of a batch file to another batch file? - Stack Overflow - Google Chrome
@REM ----
@REM How to redirect output of a batch file to another batch file? - Stack Overflow
@REM http://stackoverflow.com/questions/23474310/how-to-redirect-output-of-a-batch-file-to-another-batch-file
@REM ----
@REM Count = 673 -- Sat 27-Aug-2016 12:51:26
@REM Command Redirection, Pipes | Windows CMD | SS64.com - Google Chrome
@REM ----
@REM Command Redirection, Pipes | Windows CMD | SS64.com
@REM http://ss64.com/nt/syntax-redirection.html
@REM ----
@REM Count = 674 -- Sat 27-Aug-2016 12:54:05
@REM A Windows Batch File Programming Primer - Google Chrome
@REM ----
@REM A Windows Batch File Programming Primer
@REM http://www.vectorsite.net/tsbatch.html
@REM ----





@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ALL GOOGLE PAGE THAT ARRIVED AT THESE RESOURCE UTILITY
@REM INFO CREDIT TO OTHER-BODY PROGRAMMING
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM
@REM ----
@REM Windows batch: formatted date into variable - Stack Overflow
@REM http://stackoverflow.com/questions/10945572/windows-batch-formatted-date-into-variable
@REM --------
@REM MS-DOS Commands
@REM http://www.sprint.net.au/~terbut/usefulbox/msdoscmds.htm
@REM --------
@REM Echo | Windows CMD | SS64.com
@REM http://ss64.com/nt/echo.html
@REM --------
@REM Windows Batch Scripting - Wikibooks, open books for an open world
@REM https://en.wikibooks.org/wiki/Windows_Batch_Scripting
@REM --------
@REM Redirect pipe to a variable in Windows batch file - Super User
@REM http://superuser.com/questions/404737/redirect-pipe-to-a-variable-in-windows-batch-file
@REM --------
@REM Windows Programming/Programming CMD - Wikibooks, open books for an open world
@REM https://en.wikibooks.org/wiki/Windows_Programming/Programming_CMD
@REM --------
@REM Batch file - Wikipedia, the free encyclopedia
@REM https://en.wikipedia.org/wiki/Batch_file
@REM --------
@REM SEND A VARIABLE % TO ANOTHER BATCH FILE VIA ECHO AND PIPE - Google Search
@REM https://www.google.co.uk/search?q=SEND+A+VARIABLE+%25+TO+ANOTHER+BATCH+FILE+VIA+ECHO+AND+PIPE&oq=SEND+A+VARIABLE+%25+TO+ANOTHER+BATCH+FILE+VIA+ECHO+AND+PIPE&aqs=chrome..69i57.21897j0j7&sourceid=chrome&ie=UTF-8
@REM --------
@REM Cranky Bit » Redirecting Output to a File in Windows Batch Scripts
@REM http://blog.crankybit.com/redirecting-output-to-a-file-in-windows-batch-scripts/
@REM --------
@REM windows - How to append a date in batch files - Stack Overflow
@REM http://stackoverflow.com/questions/864718/how-to-append-a-date-in-batch-files
@REM ----




@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ALL GOOGLE SEARCH TERM AND INDEXED
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ANSWER WAS NOT FOUND ANYWHERE
@REM ----------------------------------------------------------------------------
@REM ABOUT PUT A VARIABLE ECHO PIPE TO A FILE TO USE AS A VARIABLE
@REM PROBLEM WAS I THOUGHT IT WAS -- UNABLE IF THE VARIOABLE HAD BEEN SET ONCE
@REM BUT FOUND THAT NOT WORKED ESPECIALLY FOR VARIABLE THAT ARE SET LIKE %DATE%
@REM SO ANOTHER ROUND FOLLOWED AND ANSWER WAS EASY SET WITH DOUBLE %%DATE%%
@REM ----------------------------------------------------------------------------
@REM BUT REMEMBER I DID FIND SOMEBODY TELL ONCE BEFORE
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------


@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM BRIEF SCRIPT OF MY 17 SEARCH TERM BELOW
@REM ----------------------------------------------------------------------------
@REM ALL THE FUN OF GOOGLE SEARCH PROGRAMMING AND HELP FILE
@REM THINK BEFORE YOU PRINT -- STORE IT ON MEDIA --
@REM AND THEN THROW AWAY THE PRINT FOR ECO RECYCLE
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM ---- START --- Sat 27-Aug-2016 11:54:47
@REM ----------------------------------------------------------------------------
@REM 01.. COMMAND LINE VARIABLE DATE ANOTHER BAT FILE - Google Search
@REM 02.. USE TIME %DATE% VARIABLE ECHO BATCH FILE - Google Search
@REM 03.. USE TIME %DATE% VARIABLE BATCH FILE - Google Search
@REM 04.. SEND TIME IN %DATE% VARIABLE TO BATCH FILE - Google Search
@REM 05.. SEND TIME IN VARIABLE TO BATCH FILE - Google Search
@REM 06.. SET THE %DATE% ECHO AND PIPE IT - Google Search
@REM 07.. SET THE %DATE% IN ECHO AND PIPE IT - Google Search
@REM 08.. ECHO SET VAR=%DATE% > PIPE NOT APPEAR CORRECT - Google Search
@REM 09.. ECHO SET VAR=%DATE% AND PIPE NOT APPEAR CORRECT - Google Search
@REM 10.. ECHO SET VAR=%DATE% -- %TIME% NOT APPEAR CORRECT - Google Search
@REM 11.. WINDOWS PIPE A VARIABLE WITH ECHO - Google Search
@REM 12.. PIPE A VARIABLE WITH ECHO - Google Search
@REM 13.. PIPE A VARIABLE AND READ IN FROM ANOTHER BATCH FILE - Google Search
@REM 14.. OUTPUT A VARIABLE LABLE %DATE% TO A BTACH FILE WITH ECHO TO READ IN AGAIN - Google Search
@REM 15.. SEND A VARIABLE % TO ANOTHER BATCH FILE VIA ECHO AND PIPE - Google Search
@REM 16.. WINDOWS NOT WORKING "ECHO ECHO TEXT - %DATE% >>" - Google Search
@REM 17.. NOT WORKING "ECHO ECHO TEXT - %DATE% >>" - Google Search
@REM ----------------------------------------------------------------------------
@REM ---- END ----- Sat 27-Aug-2016 13:18:42 -- LONG FRANTICLY SEARCH + 1 HOUR
@REM ----------------------------------------------------------------------------
@REM ----------------------------------------------------------------------------
@REM
@REM
@REM ---
@REM Count = 681 -- Sat 27-Aug-2016 13:16:34
@REM COMMAND LINE VARIABLE DATE ANOTHER BAT FILE - Google Search - Google Chrome
@REM ----
@REM COMMAND LINE VARIABLE DATE ANOTHER BAT FILE - Google Search
@REM https://www.google.co.uk/search?num=30&q=COMMAND+LINE+VARIABLE+DATE+ANOTHER+BAT+FILE&oq=COMMAND+LINE+VARIABLE+DATE+ANOTHER+BAT+FILE&gs_l=serp.3...31304.50640.0.51473.32.32.0.0.0.0.93.2251.32.32.0....0...1c.1.64.serp..0.17.1218...30i10k1j33i21k1j0i22i30k1j0i22i10i30k1.libgruvSXbc
@REM ----
@REM Count = 682 -- Sat 27-Aug-2016 13:16:43
@REM USE TIME %DATE% VARIABLE ECHO BATCH FILE - Google Search - Google Chrome
@REM ----
@REM USE TIME %DATE% VARIABLE ECHO BATCH FILE - Google Search
@REM https://www.google.co.uk/search?num=30&q=USE+TIME+%25DATE%25+VARIABLE+ECHO+BATCH+FILE&oq=USE+TIME+%25DATE%25+VARIABLE+ECHO+BATCH+FILE&gs_l=serp.3...9596.12529.0.13388.8.8.0.0.0.0.183.616.6j1.7.0....0...1c.1.64.serp..1.4.298...30i10k1.BRwOlP_z3AM
@REM ----
@REM Count = 683 -- Sat 27-Aug-2016 13:16:47
@REM USE TIME %DATE% VARIABLE BATCH FILE - Google Search - Google Chrome
@REM ----
@REM USE TIME %DATE% VARIABLE BATCH FILE - Google Search
@REM https://www.google.co.uk/search?num=30&q=USE+TIME+%25DATE%25+VARIABLE++BATCH+FILE&oq=USE+TIME+%25DATE%25+VARIABLE++BATCH+FILE&gs_l=serp.3...6254.15183.0.16101.14.13.0.0.0.0.157.1036.12j1.13.0....0...1c.1.64.serp..3.6.550...30i10k1j0i8i30k1.9fJ8fgeVuz4
@REM ----
@REM Count = 684 -- Sat 27-Aug-2016 13:16:51
@REM SEND TIME IN %DATE% VARIABLE TO BATCH FILE - Google Search - Google Chrome
@REM ----
@REM SEND TIME IN %DATE% VARIABLE TO BATCH FILE - Google Search
@REM https://www.google.co.uk/search?num=30&q=SEND+TIME+IN+%25DATE%25+VARIABLE+TO+BATCH+FILE&oq=SEND+TIME+IN+%25DATE%25+VARIABLE+TO+BATCH+FILE&gs_l=serp.3...45845.47316.0.47754.7.7.0.0.0.0.110.570.6j1.7.0....0...1c.1.64.serp..0.4.331...30i10k1.p1k_Ibnq67U
@REM ----
@REM Count = 685 -- Sat 27-Aug-2016 13:16:55
@REM SEND TIME IN VARIABLE TO BATCH FILE - Google Search - Google Chrome
@REM ----
@REM SEND TIME IN VARIABLE TO BATCH FILE - Google Search
@REM https://www.google.co.uk/search?q=SEND+TIME+IN+VARIABLE+TO+BATCH+FILE&oq=SEND+TIME+IN+VARIABLE+TO+BATCH+FILE&aqs=chrome..69i57.9080j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 686 -- Sat 27-Aug-2016 13:17:09
@REM SET THE %DATE% ECHO AND PIPE IT - Google Search - Google Chrome
@REM ----
@REM SET THE %DATE% ECHO AND PIPE IT - Google Search
@REM https://www.google.co.uk/search?num=30&q=SET+THE+%25DATE%25+ECHO+AND+PIPE+IT&oq=SET+THE+%25DATE%25+ECHO+AND+PIPE+IT&gs_l=serp.3...11024.11399.0.11957.3.3.0.0.0.0.85.208.3.3.0....0...1c.1.64.serp..0.0.0.vcXYTq5Uf5c
@REM ----
@REM Count = 687 -- Sat 27-Aug-2016 13:17:13
@REM SET THE %DATE% IN ECHO AND PIPE IT - Google Search - Google Chrome
@REM ----
@REM SET THE %DATE% IN ECHO AND PIPE IT - Google Search
@REM https://www.google.co.uk/search?q=SET+THE+%25DATE%25+IN+ECHO+AND+PIPE+IT&oq=SET+THE+%25DATE%25+IN+ECHO+AND+PIPE+IT&aqs=chrome..69i57.14675j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 688 -- Sat 27-Aug-2016 13:17:23
@REM ECHO SET VAR=%DATE% > PIPE NOT APPEAR CORRECT - Google Search - Google Chrome
@REM ----
@REM ECHO SET VAR=%DATE% > PIPE NOT APPEAR CORRECT - Google Search
@REM https://www.google.co.uk/search?num=30&q=ECHO+SET+VAR%3D%25DATE%25+%3E+PIPE+NOT+APPEAR+CORRECT&oq=ECHO+SET+VAR%3D%25DATE%25+%3E+PIPE+NOT+APPEAR+CORRECT&gs_l=serp.3...4031.6461.0.6841.6.6.0.0.0.0.867.1080.1j1j6-1.3.0....0...1c.1.64.serp..3.0.0.8FrxaXumOiE
@REM ----
@REM Count = 689 -- Sat 27-Aug-2016 13:17:26
@REM ECHO SET VAR=%DATE% AND PIPE NOT APPEAR CORRECT - Google Search - Google Chrome
@REM ----
@REM ECHO SET VAR=%DATE% AND PIPE NOT APPEAR CORRECT - Google Search
@REM https://www.google.co.uk/search?num=30&q=ECHO+SET+VAR%3D%25DATE%25+AND+PIPE+NOT+APPEAR+CORRECT&oq=ECHO+SET+VAR%3D%25DATE%25+AND+PIPE+NOT+APPEAR+CORRECT&gs_l=serp.3...9058.18395.0.19282.25.22.3.0.0.0.166.1710.20j2.22.0....0...1c.1.64.serp..2.13.873...30i10k1.zI1Kctw792k
@REM ----
@REM Count = 690 -- Sat 27-Aug-2016 13:17:31
@REM ECHO SET VAR=%DATE% -- %TIME% NOT APPEAR CORRECT - Google Search - Google Chrome
@REM ----
@REM ECHO SET VAR=%DATE% -- %TIME% NOT APPEAR CORRECT - Google Search
@REM https://www.google.co.uk/search?q=ECHO+SET+VAR%3D%25DATE%25+--+%25TIME%25+NOT+APPEAR+CORRECT&oq=ECHO+SET+VAR%3D%25DATE%25+--+%25TIME%25+NOT+APPEAR+CORRECT&aqs=chrome..69i57.15484j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 691 -- Sat 27-Aug-2016 13:17:46
@REM WINDOWS PIPE A VARIABLE WITH ECHO - Google Search - Google Chrome
@REM ----
@REM WINDOWS PIPE A VARIABLE WITH ECHO - Google Search
@REM https://www.google.co.uk/search?num=30&q=WINDOWS+PIPE+A+VARIABLE+WITH+ECHO&oq=WINDOWS+PIPE+A+VARIABLE+WITH+ECHO&gs_l=serp.3...10447.11746.0.11977.8.6.0.0.0.0.306.306.3-1.1.0....0...1c.1.64.serp..8.0.0.APZTdVsvtys
@REM ----
@REM Count = 692 -- Sat 27-Aug-2016 13:17:49
@REM PIPE A VARIABLE WITH ECHO - Google Search - Google Chrome
@REM ----
@REM PIPE A VARIABLE WITH ECHO - Google Search
@REM https://www.google.co.uk/search?q=PIPE+A+VARIABLE+WITH+ECHO&oq=PIPE+A+VARIABLE+WITH+ECHO&aqs=chrome..69i57.6218j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 693 -- Sat 27-Aug-2016 13:18:01
@REM PIPE A VARIABLE AND READ IN FROM ANOTHER BATCH FILE - Google Search - Google Chrome
@REM ----
@REM PIPE A VARIABLE AND READ IN FROM ANOTHER BATCH FILE - Google Search
@REM https://www.google.co.uk/search?q=PIPE+A+VARIABLE+AND+READ+IN+FROM+ANOTHER+BATCH+FILE&oq=PIPE+A+VARIABLE+AND+READ+IN+FROM+ANOTHER+BATCH+FILE&aqs=chrome..69i57.11749j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 694 -- Sat 27-Aug-2016 13:18:13
@REM OUTPUT A VARIABLE LABLE %DATE% TO A BTACH FILE WITH ECHO TO READ IN AGAIN - Google Search - Google Chrome
@REM ----
@REM OUTPUT A VARIABLE LABLE %DATE% TO A BTACH FILE WITH ECHO TO READ IN AGAIN - Google Search
@REM https://www.google.co.uk/search?q=OUTPUT+A+VARIABLE+LABLE+%25DATE%25+TO+A+BTACH+FILE+WITH+ECHO+TO+READ+IN+AGAIN&oq=OUTPUT+A+VARIABLE+LABLE+%25DATE%25+TO+A+BTACH+FILE+WITH+ECHO+TO+READ+IN+AGAIN&aqs=chrome..69i57.17382j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 695 -- Sat 27-Aug-2016 13:18:25
@REM SEND A VARIABLE % TO ANOTHER BATCH FILE VIA ECHO AND PIPE - Google Search - Google Chrome
@REM ----
@REM SEND A VARIABLE % TO ANOTHER BATCH FILE VIA ECHO AND PIPE - Google Search
@REM https://www.google.co.uk/search?q=SEND+A+VARIABLE+%25+TO+ANOTHER+BATCH+FILE+VIA+ECHO+AND+PIPE&oq=SEND+A+VARIABLE+%25+TO+ANOTHER+BATCH+FILE+VIA+ECHO+AND+PIPE&aqs=chrome..69i57.21897j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 696 -- Sat 27-Aug-2016 13:18:36
@REM WINDOWS NOT WORKING "ECHO ECHO TEXT - %DATE% >>" - Google Search - Google Chrome
@REM ----
@REM WINDOWS NOT WORKING "ECHO ECHO TEXT - %DATE% >>" - Google Search
@REM https://www.google.co.uk/search?num=30&q=WINDOWS+NOT+WORKING+%22ECHO+ECHO+TEXT+-+%25DATE%25+%3E%3E%22&oq=WINDOWS+NOT+WORKING+%22ECHO+ECHO+TEXT+-+%25DATE%25+%3E%3E%22&gs_l=serp.3...12315.13290.0.13530.8.7.0.0.0.0.305.898.0j2j1j1.4.0....0...1c.1.64.serp..7.0.0.cwEa1_4srq8
@REM ----
@REM Count = 697 -- Sat 27-Aug-2016 13:18:42
@REM NOT WORKING "ECHO ECHO TEXT - %DATE% >>" - Google Search - Google Chrome
@REM ----
@REM NOT WORKING "ECHO ECHO TEXT - %DATE% >>" - Google Search
@REM https://www.google.co.uk/search?q=NOT+WORKING+%22ECHO+ECHO+TEXT+-+%25DATE%25+%3E%3E%22&oq=NOT+WORKING+%22ECHO+ECHO+TEXT+-+%25DATE%25+%3E%3E%22&aqs=chrome..69i57.37223j0j7&sourceid=chrome&ie=UTF-8
@REM ----





@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM HERE THE 2ND GREATER HALF -- OF GOOGLE SEARCHING
@REM WITH 38 SEARCH TERM SHOWING
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------




@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ALL THE LAST PAGES LEFT OPEN IN BROWSER
@REM ---------------------------------------------------------------
@REM LOOK TO MIDDLE SECTION ALL THE SEARCH TERM USE 38 TOTAL THIS
@REM THE MORE SIZE GREATER HALF OF TWO TODAY THE 2ND BATCH TO FINISH
@REM 1ST -- RUN AT IT WAS 17 GOOGLE SEARCHER
@REM TO COMPLETER ON THE 2ND DAY -- RUN AT IT
@REM SATURDAY 27 AUGUST
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ----
@REM recursion - Windows Batch File Looping Through Directories to Process Files? - Stack Overflow
@REM http://stackoverflow.com/questions/8397674/windows-batch-file-looping-through-directories-to-process-files
@REM --------
@REM windows - Iterate all files in a directory using a for-loop - Stack Overflow
@REM http://stackoverflow.com/questions/138497/iterate-all-files-in-a-directory-using-a-for-loop
@REM --------
@REM for /f - Google Search
@REM https://www.google.co.uk/search?q=for+%2Ff&oq=for+%2Ff&aqs=chrome..69i57j69i60l2j69i65j69i60&sourceid=chrome&ie=UTF-8
@REM --------
@REM What does %~d0 mean in a Windows batch file? - Stack Overflow
@REM http://stackoverflow.com/questions/112055/what-does-d0-mean-in-a-windows-batch-file
@REM --------
@REM Parameters / Arguments | Windows CMD | SS64.com
@REM http://ss64.com/nt/syntax-args.html
@REM --------
@REM batch file - how to clear a variable which is set from command line by "set /p var1=" - Stack Overflow
@REM http://stackoverflow.com/questions/17106181/how-to-clear-a-variable-which-is-set-from-command-line-by-set-p-var1
@REM --------
@REM cmd - windows batch SET inside IF not working - Stack Overflow
@REM http://stackoverflow.com/questions/9102422/windows-batch-set-inside-if-not-working
@REM --------
@REM windows - Batch file: Find if substring is in string (not in a file) - Stack Overflow
@REM http://stackoverflow.com/questions/7005951/batch-file-find-if-substring-is-in-string-not-in-a-file
@REM --------
@REM windows - FORFILES: Only delete file if RELPATH contains substring - Stack Overflow
@REM http://stackoverflow.com/questions/33570132/forfiles-only-delete-file-if-relpath-contains-substring
@REM --------
@REM windows - Batch file: Find if substring is in string (not in a file)_KECHENGPUZI_课程铺子
@REM http://kechengpuzi.com/q/s7005951
@REM --------
@REM command line - Spaces problem in CMD ! - Stack Overflow
@REM http://stackoverflow.com/questions/6376113/spaces-problem-in-cmd
@REM --------
@REM DIR - Display information about files & subdirectories
@REM https://jpsoft.com/help/dir.htm
@REM --------
@REM New Tab
@REM chrome://newtab/
@REM --------
@REM batch file - How to run two commands in one line in Windows CMD? - Stack Overflow
@REM http://stackoverflow.com/questions/8055371/how-to-run-two-commands-in-one-line-in-windows-cmd
@REM --------
@REM How to grant permission to users for a directory using command line in Windows? - Stack Overflow
@REM http://stackoverflow.com/questions/2928738/how-to-grant-permission-to-users-for-a-directory-using-command-line-in-windows
@REM --------
@REM How do I Detect (and Delete) a File if it is Empty using a Windows Batch File? - Stack Overflow
@REM http://stackoverflow.com/questions/9378972/how-do-i-detect-and-delete-a-file-if-it-is-empty-using-a-windows-batch-file
@REM --------
@REM windows - How to check if a file exists from inside a batch file - Stack Overflow
@REM http://stackoverflow.com/questions/4340350/how-to-check-if-a-file-exists-from-inside-a-batch-file
@REM --------
@REM How to fix error “No mapping between account names and security IDs” in Active Directory - RebelAdmin
@REM http://www.rebeladmin.com/2016/01/how-to-fix-error-no-mapping-between-account-names-and-security-ids-in-active-directory/
@REM --------
@REM windows - How to prevent the command prompt from closing after execution? - Super User
@REM http://superuser.com/questions/306167/how-to-prevent-the-command-prompt-from-closing-after-execution
@REM --------
@REM windows - How to execute multiple commands in a single line - Stack Overflow
@REM http://stackoverflow.com/questions/13719174/how-to-execute-multiple-commands-in-a-single-line
@REM --------
@REM start cmd /k - Google Search
@REM https://www.google.co.uk/search?q=start+cmd+%2Fk&oq=START+%22CMD%22&aqs=chrome.3.69i57j0l5.13468j0j7&sourceid=chrome&ie=UTF-8
@REM --------
@REM Changing Windows process priority via command line - Super User
@REM http://superuser.com/questions/620724/changing-windows-process-priority-via-command-line
@REM --------
@REM COMMAND LINE OUTPUT A VARIABLE NAME WITH ENCLOSED %% TO A ECHO PIPE OUTPUT > - Google Search
@REM https://www.google.co.uk/search?num=30&biw=1229&bih=558&q=COMMAND+LINE+OUTPUT+A+VARIABLE+NAME+WITH+ENCLOSED+%25%25+TO+A+ECHO+PIPE+OUTPUT+%3E&oq=COMMAND+LINE+OUTPUT+A+VARIABLE+NAME+WITH+ENCLOSED+%25%25+TO+A+ECHO+PIPE+OUTPUT+%3E&gs_l=serp.3...10958.14102.0.14286.13.13.0.0.0.0.128.1144.12j1.13.0....0...1c.1.64.serp..1.11.991...30i10k1.0-AEy5Acsf0
@REM --------
@REM Batch File Commands -- Technical Notes
@REM http://academic.evergreen.edu/projects/biophysics/technotes/program/batch.htm
@REM --------
@REM ANSI.SYS Escape Sequences -- Technical Notes
@REM http://academic.evergreen.edu/projects/biophysics/technotes/program/ansi_esc.htm
@REM --------
@REM ANSI.SYS Key Codes -- Technical Notes
@REM http://academic.evergreen.edu/projects/biophysics/technotes/program/ansi_key.htm
@REM --------
@REM HTML Codes - Table of ascii characters and symbols
@REM http://www.ascii.cl/htmlcodes.htm
@REM --------
@REM Batch files - Escape Characters
@REM http://www.robvanderwoude.com/escapechars.php
@REM --------
@REM windows 7 - TEMP environment variable occasionally set incorrectly - Super User
@REM http://superuser.com/questions/414374/temp-environment-variable-occasionally-set-incorrectly
@REM --------
@REM windows - How to read the label of a drive or volume in a batch file? - Stack Overflow
@REM http://stackoverflow.com/questions/8649934/how-to-read-the-label-of-a-drive-or-volume-in-a-batch-file
@REM --------
@REM TSCMD101 FAQ: How to capture Volume Serial Number into an environment variable? Assorted NT/2000/XP/.. CMD.EXE script tricks written by Timo Salmi
@REM http://www.netikka.net/tsneti/info/tscmd101.htm#label
@REM --------
@REM Using multiple IF statements in a batch file - Stack Overflow
@REM http://stackoverflow.com/questions/6711615/using-multiple-if-statements-in-a-batch-file
@REM --------
@REM Testing If a Drive or Directory Exists from a Batch File
@REM https://support.microsoft.com/en-gb/kb/65994
@REM ----
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ALL THE MY SEARCH TERM LEFT IN BROWSER
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM SEARCH TERM INDEX
@REM ---------------------------------------------------------------
@REM BACK TO FRONT NEWEST FIRST
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM Count = 1059 -- Sat 27-Aug-2016 19:49:18 TIME AFTER NEAR END
@REM ---------------------------------------------------------------
@REM 01... DIR EXIST COMMAND LINE
@REM 02... DOUBLE IF COMMAND LINE
@REM 03... BATCH FILE GET VOLUME LABEL 01
@REM 04... BATCH FILE GET VOLUME LABLE 02
@REM 05... TEMP NOT SHOWING IN ENVIROMENT SET
@REM 06... WINDOWS COMMAND LINE OUTPUT PIPE SPECIAL CHARATOR
@REM 07... COMMAND LINE OUTPUT PIPE SPECIAL CHARATOR
@REM 08... CHARATOR CODE FOR %
@REM 09... WANT TO OUTPUT A STRING TO FILE FROM BATCH FILE %%
@REM 10... WANT TO OUTPUT A STRING TO FILE FROM BATCH FILE
@REM 11... COMMAND LINE OUTPUT A VARIABLE NAME WITH ENCLOSED %% TO A ECHO PIPE OUTPUT >
@REM 12... OUTPUT A VARIABLE NAME WITH ENCLOSED %% TO A ECHO PIPE OUTPUT >
@REM 13... OUTPUT A VAIRABLE NAME WITH ENCLOSED %% TO A ECHO PIPE OUTPUT >
@REM 14... CPU PRIORITY 13
@REM 15... start cmd /k
@REM 16... TWO COMMANDLINE ON ONE START CMD
@REM 17... COMMAND LINE START AND DON'T EXIT
@REM 18... No mapping between account names and security IDs was done.
@REM 19... COMMAND LINE FILE EXISTS
@REM 20... COMMAND LINE FILE BIGGER THAN EMPTY
@REM 21... COMMAND LINE FILE CONTAIN ANYTHING
@REM 22... Icacls .\* /t /Q /grant Everyone:F >NUL
@REM 23... Icacls "C:\USERS\All Users" /t /Q /grant Everyone:F >NUL
@REM 24... TWO COMMAND LINE IN ONE LINE
@REM 25... windows command line folder name in the last column
@REM 26... windows command line folder name in the 5th column
@REM 27... command line folder name in the 5th column
@REM 28... folder name in the 5th column
@REM 29... how use comandline and space in a line
@REM 30... if not x%str1:bcd=%==x%str1% echo It contains bcd
@REM 31... if not x%str1:bcd=%==x%str1% echo It contains
@REM 32... batch search string in variable
@REM 33... COMMAND LINE IF SET
@REM 34... COMMAND LINE CLEAR A VARIABLE
@REM 35... %~n1 - expands %1 to a file name only
@REM 36... for /f
@REM 37... batch loop through files in a directory
@REM 38... BATCH FILE TO RECURSE EVERY FOLDER
@REM ---------------------------------------------------------------
@REM Count = 1099 -- Sat 27-Aug-2016 13:18:42 -- TIME BEFORE
@REM ---------------------------------------------------------------
@REM CURRENT TIME -- Sat 27-Aug-2016 20:22:00
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM ---------------------------------------------------------------
@REM --------------------
@REM Count = 1059 -- Sat 27-Aug-2016 19:49:18
@REM DIR EXIST COMMAND LINE
@REM ----
@REM DIR EXIST COMMAND LINE - Google Search
@REM https://www.google.co.uk/search?q=DIR+EXIST+COMMAND+LINE&oq=DIR+EXIST+COMMAND+LINE&aqs=chrome..69i57j0l3j69i64.6070j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1060 -- Sat 27-Aug-2016 19:49:32
@REM DOUBLE IF COMMAND LINE
@REM ----
@REM DOUBLE IF COMMAND LINE - Google Search
@REM https://www.google.co.uk/search?q=DOUBLE+IF+COMMAND+LINE&oq=DOUBLE+IF+COMMAND+LINE&aqs=chrome..69i57.5594j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1061 -- Sat 27-Aug-2016 19:49:53
@REM BATCH FILE GET VOLUME LABEL
@REM ----
@REM BATCH FILE GET VOLUME LABEL - Google Search
@REM https://www.google.co.uk/search?num=30&q=BATCH+FILE+GET+VOLUME+LABEL&spell=1&sa=X&ved=0ahUKEwijld7zi-LOAhWBtBoKHWoZBvYQvwUIGygA&biw=1229&bih=558
@REM ----
@REM Count = 1062 -- Sat 27-Aug-2016 19:49:57
@REM BATCH FILE GET VOLUME LABLE
@REM ----
@REM BATCH FILE GET VOLUME LABLE - Google Search
@REM https://www.google.co.uk/search?q=BATCH+FILE+GET+VOLUME+LABLE&oq=BATCH+FILE+GET+VOLUME+LABLE&aqs=chrome..69i57j0.5695j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1063 -- Sat 27-Aug-2016 19:50:14
@REM TEMP NOT SHOWING IN ENVIROMENT SET
@REM ----
@REM TEMP NOT SHOWING IN ENVIROMENT SET - Google Search
@REM https://www.google.co.uk/search?q=TEMP+NOT+SHOWING+IN+ENVIROMENT+SET&oq=TEMP+NOT+SHOWING+IN+ENVIROMENT+SET&aqs=chrome..69i57.18883j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1064 -- Sat 27-Aug-2016 19:50:28
@REM WINDOWS COMMAND LINE OUTPUT PIPE SPECIAL CHARATOR
@REM ----
@REM WINDOWS COMMAND LINE OUTPUT PIPE SPECIAL CHARATOR - Google Search
@REM https://www.google.co.uk/search?num=30&q=WINDOWS+COMMAND+LINE+OUTPUT+PIPE+SPECIAL+CHARATOR&oq=WINDOWS+COMMAND+LINE+OUTPUT+PIPE+SPECIAL+CHARATOR&gs_l=serp.3..30i10k1.3030.4136.0.4578.8.8.0.0.0.0.84.549.7.7.0....0...1c.1.64.serp..1.7.546.or01cLk7524
@REM ----
@REM Count = 1065 -- Sat 27-Aug-2016 19:50:33
@REM COMMAND LINE OUTPUT PIPE SPECIAL CHARATOR
@REM COMMAND LINE OUTPUT PIPE SPECIAL CHARATOR
@REM ----
@REM COMMAND LINE OUTPUT PIPE SPECIAL CHARATOR - Google Search
@REM https://www.google.co.uk/search?q=COMMAND+LINE+OUTPUT+PIPE+SPECIAL+CHARATOR&oq=COMMAND+LINE+OUTPUT+PIPE+SPECIAL+CHARATOR&aqs=chrome..69i57.27822j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1066 -- Sat 27-Aug-2016 19:50:51
@REM CHARATOR CODE FOR %
@REM ----
@REM CHARATOR CODE FOR % - Google Search
@REM https://www.google.co.uk/search?q=CHARATOR+CODE+FOR+%25&oq=CHARATOR+CODE+FOR+%25&aqs=chrome..69i57j0l5.4959j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1067 -- Sat 27-Aug-2016 19:51:10
@REM WANT TO OUTPUT A STRING TO FILE FROM BATCH FILE %%
@REM ----
@REM WANT TO OUTPUT A STRING TO FILE FROM BATCH FILE %% - Google Search
@REM https://www.google.co.uk/search?num=30&biw=1229&bih=558&q=WANT+TO+OUTPUT+A+STRING+TO+FILE+FROM+BATCH+FILE+%25%25&oq=WANT+TO+OUTPUT+A+STRING+TO+FILE+FROM+BATCH+FILE+%25%25&gs_l=serp.3...1596.6853.0.8009.5.4.1.0.0.0.98.297.4.4.0....0...1c.1.64.serp..0.0.0.97ctxLs1Uds
@REM ----
@REM Count = 1068 -- Sat 27-Aug-2016 19:51:15
@REM WANT TO OUTPUT A STRING TO FILE FROM BATCH FILE
@REM ----
@REM WANT TO OUTPUT A STRING TO FILE FROM BATCH FILE - Google Search
@REM https://www.google.co.uk/search?num=30&q=WANT+TO+OUTPUT+A+STRING+TO+FILE+FROM+BATCH+FILE&spell=1&sa=X&ved=0ahUKEwjH_42b5OHOAhVBKsAKHQ5sB0AQvwUIGygA&biw=1229&bih=558
@REM ----
@REM Count = 1069 -- Sat 27-Aug-2016 19:51:36
@REM COMMAND LINE OUTPUT A VARIABLE NAME WITH ENCLOSED %% TO A ECHO PIPE OUTPUT >
@REM ----
@REM COMMAND LINE OUTPUT A VARIABLE NAME WITH ENCLOSED %% TO A ECHO PIPE OUTPUT > - Google Search
@REM https://www.google.co.uk/search?num=30&biw=1229&bih=558&q=COMMAND+LINE+OUTPUT+A+VARIABLE+NAME+WITH+ENCLOSED+%25%25+TO+A+ECHO+PIPE+OUTPUT+%3E&oq=COMMAND+LINE+OUTPUT+A+VARIABLE+NAME+WITH+ENCLOSED+%25%25+TO+A+ECHO+PIPE+OUTPUT+%3E&gs_l=serp.3...10958.14102.0.14286.13.13.0.0.0.0.128.1144.12j1.13.0....0...1c.1.64.serp..1.11.991...30i10k1.0-AEy5Acsf0
@REM ----
@REM Count = 1070 -- Sat 27-Aug-2016 19:51:42
@REM OUTPUT A VARIABLE NAME WITH ENCLOSED %% TO A ECHO PIPE OUTPUT >
@REM ----
@REM OUTPUT A VARIABLE NAME WITH ENCLOSED %% TO A ECHO PIPE OUTPUT > - Google Search
@REM https://www.google.co.uk/search?num=30&q=OUTPUT+A+VARIABLE+NAME+WITH+ENCLOSED+%25%25+TO+A+ECHO+PIPE+OUTPUT+%3E&spell=1&sa=X&ved=0ahUKEwjayrjQ4uHOAhWqB8AKHSeRBKQQvwUIGygA&biw=1229&bih=558
@REM ----
@REM Count = 1071 -- Sat 27-Aug-2016 19:51:48
@REM OUTPUT A VAIRABLE NAME WITH ENCLOSED %% TO A ECHO PIPE OUTPUT >
@REM ----
@REM OUTPUT A VAIRABLE NAME WITH ENCLOSED %% TO A ECHO PIPE OUTPUT > - Google Search
@REM https://www.google.co.uk/search?q=OUTPUT+A+VAIRABLE+NAME+WITH+ENCLOSED+%25%25+TO+A+ECHO+PIPE+OUTPUT+%3E&oq=OUTPUT+A+VAIRABLE+NAME+WITH+ENCLOSED+%25%25+TO+A+ECHO+PIPE+OUTPUT+%3E&aqs=chrome..69i57.21790j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1072 -- Sat 27-Aug-2016 19:52:02
@REM CPU PRIORITY 13
@REM CPU PRIORITY 13
@REM ----
@REM CPU PRIORITY 13 - Google Search
@REM https://www.google.co.uk/search?q=CPU+PRIORITY+13&oq=CPU+PRIORITY+13&aqs=chrome..69i57j0l3.3473j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1073 -- Sat 27-Aug-2016 19:52:09
@REM start cmd /k
@REM ----
@REM start cmd /k - Google Search
@REM https://www.google.co.uk/search?q=start+cmd+%2Fk&oq=START+%22CMD%22&aqs=chrome.3.69i57j0l5.13468j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1074 -- Sat 27-Aug-2016 19:52:21
@REM TWO COMMANDLINE ON ONE START CMD
@REM ----
@REM TWO COMMANDLINE ON ONE START CMD - Google Search
@REM https://www.google.co.uk/search?q=TWO+COMMANDLINE+ON+ONE+START+CMD&oq=TWO+COMMANDLINE+ON+ONE+START+CMD&aqs=chrome..69i57.7365j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1075 -- Sat 27-Aug-2016 19:52:31
@REM COMMAND LINE START AND DON'T EXIT
@REM ----
@REM COMMAND LINE START AND DON'T EXIT - Google Search
@REM https://www.google.co.uk/search?num=30&q=COMMAND+LINE+START+AND+DON%27T+EXIT&spell=1&sa=X&ved=0ahUKEwiXiprToOHOAhWMB8AKHVTKAY0QBQgbKAA&biw=1229&bih=558
@REM ----
@REM Count = 1076 -- Sat 27-Aug-2016 19:52:52
@REM No mapping between account names and security IDs was done.
@REM ----
@REM No mapping between account names and security IDs was done. - Google Search
@REM https://www.google.co.uk/search?q=No+mapping+between+account+names+and+security+IDs+was+done.&oq=No+mapping+between+account+names+and+security+IDs+was+done.&aqs=chrome..69i57.25114879j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1077 -- Sat 27-Aug-2016 19:53:09
@REM COMMAND LINE FILE EXISTS
@REM ----
@REM COMMAND LINE FILE EXISTS - Google Search
@REM https://www.google.co.uk/search?q=COMMAND+LINE+FILE+EXISTS&oq=COMMAND+LINE+FILE+EXISTS&aqs=chrome..69i57j0l5.5782j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1078 -- Sat 27-Aug-2016 19:53:22
@REM COMMAND LINE FILE BIGGER THAN EMPTY
@REM ----
@REM COMMAND LINE FILE BIGGER THAN EMPTY - Google Search
@REM https://www.google.co.uk/search?num=30&q=COMMAND+LINE+FILE+BIGGER+THAN+EMPTY&oq=COMMAND+LINE+FILE+BIGGER+THAN+EMPTY&gs_l=serp.3...2630.10214.0.10781.33.33.0.0.0.0.131.2639.26j7.33.0....0...1c.1.64.serp..0.24.1996...0j33i21k1j0i22i30k1j0i8i13i30k1j0i13k1.eAkvMTEIGb4
@REM ----
@REM Count = 1079 -- Sat 27-Aug-2016 19:53:31
@REM COMMAND LINE FILE CONTAIN ANYTHING
@REM ----
@REM COMMAND LINE FILE CONTAIN ANYTHING - Google Search
@REM https://www.google.co.uk/search?q=COMMAND+LINE+FILE+CONTAIN+ANYTHING&oq=COMMAND+LINE+FILE+CONTAIN+ANYTHING&aqs=chrome..69i57.19927j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1080 -- Sat 27-Aug-2016 19:53:43
@REM Icacls .\* /t /Q /grant Everyone:F >NUL
@REM ----
@REM Icacls .\* /t /Q /grant Everyone:F >NUL - Google Search
@REM https://www.google.co.uk/search?num=30&q=Icacls+.%5C*+%2Ft+%2FQ+%2Fgrant+Everyone%3AF+%3ENUL&oq=Icacls+.%5C*+%2Ft+%2FQ+%2Fgrant+Everyone%3AF+%3ENUL&gs_l=serp.3...16513.19899.0.20320.23.11.0.0.0.0.73.650.11.11.0....0...1c.1.64.serp..18.0.0.6nS5QibGbjk
@REM ----
@REM Count = 1081 -- Sat 27-Aug-2016 19:54:00
@REM Icacls "C:\USERS\All Users" /t /Q /grant Everyone:F >NUL
@REM ----
@REM Icacls "C:\USERS\All Users" /t /Q /grant Everyone:F >NUL - Google Search
@REM https://www.google.co.uk/search?q=Icacls+%22C%3A%5CUSERS%5CAll+Users%22+%2Ft+%2FQ+%2Fgrant+Everyone%3AF+%3ENUL&oq=Icacls+%22C%3A%5CUSERS%5CAll+Users%22+%2Ft+%2FQ+%2Fgrant+Everyone%3AF+%3ENUL&aqs=chrome..69i57.3425794j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1082 -- Sat 27-Aug-2016 19:54:12
@REM TWO COMMAND LINE IN ONE LINE
@REM ----
@REM TWO COMMAND LINE IN ONE LINE - Google Search
@REM https://www.google.co.uk/search?q=TWO+COMMAND+LINE+IN+ONE+LINE&oq=TWO+COMMAND+LINE+IN+ONE+LINE&aqs=chrome..69i57j0.5283j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1083 -- Sat 27-Aug-2016 19:54:37
@REM windows command line folder name in the last column
@REM ----
@REM windows command line folder name in the last column - Google Search
@REM https://www.google.co.uk/search?num=30&q=windows+command+line+folder+name+in+the+last++column&oq=windows+command+line+folder+name+in+the+last++column&gs_l=serp.3...5678.16006.0.16348.14.12.2.0.0.0.104.859.11j1.12.0....0...1c.1.64.serp..1.4.150...30i10k1.WE8QOmoGLPg
@REM ----
@REM Count = 1084 -- Sat 27-Aug-2016 19:54:43
@REM windows command line folder name in the 5th column
@REM ----
@REM windows command line folder name in the 5th column - Google Search
@REM https://www.google.co.uk/search?num=30&q=windows+command+line+folder+name+in+the+5th+column&oq=windows+command+line+folder+name+in+the+5th+column&gs_l=serp.3...7044.8398.0.8653.8.8.0.0.0.0.88.610.8.8.0....0...1c.1.64.serp..0.7.546...30i10k1.DrVHgIMwS5k
@REM ----
@REM Count = 1085 -- Sat 27-Aug-2016 19:54:47
@REM command line folder name in the 5th column
@REM ----
@REM command line folder name in the 5th column - Google Search
@REM https://www.google.co.uk/search?num=30&q=command+line+folder+name+in+the+5th+column&oq=command+line+folder+name+in+the+5th+column&gs_l=serp.3...7549.9664.0.10318.13.13.0.0.0.0.107.999.11j2.13.0....0...1c.1.64.serp..0.11.874...30i10k1.gWma8nS2N48
@REM ----
@REM Count = 1086 -- Sat 27-Aug-2016 19:54:52
@REM folder name in the 5th column
@REM ----
@REM folder name in the 5th column - Google Search
@REM https://www.google.co.uk/search?q=folder+name+in+the+5th+column&oq=folder+name+in+the+5th+column&aqs=chrome..69i57&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1087 -- Sat 27-Aug-2016 19:55:05
@REM how use comandline and space in a line
@REM ----
@REM how use comandline and space in a line - Google Search
@REM https://www.google.co.uk/search?q=how+use+comandline+and+space+in+a+line&oq=how+use+comandline+and+space+in+a+line&aqs=chrome..69i57.9009j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1089 -- Sat 27-Aug-2016 19:55:26
@REM if not x%str1:bcd=%==x%str1% echo It contains bcd
@REM ----
@REM if not x%str1:bcd=%==x%str1% echo It contains bcd - Google Search
@REM https://www.google.co.uk/search?q=if+not+x%25str1%3Abcd%3D%25%3D%3Dx%25str1%25+echo+It+contains+bcd&oq=if+not+x%25str1%3Abcd%3D%25%3D%3Dx%25str1%25+echo+It+contains+bcd&aqs=chrome..69i57&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1090 -- Sat 27-Aug-2016 19:55:37
@REM if not x%str1:bcd=%==x%str1% echo It contains
@REM ----
@REM if not x%str1:bcd=%==x%str1% echo It contains - Google Search
@REM https://www.google.co.uk/search?q=if+not+x%25str1%3Abcd%3D%25%3D%3Dx%25str1%25+echo+It+contains&oq=if+not+x%25str1%3Abcd%3D%25%3D%3Dx%25str1%25+echo+It+contains&aqs=chrome..69i57.4794357j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1091 -- Sat 27-Aug-2016 19:55:47
@REM batch search string in variable
@REM ----
@REM batch search string in variable - Google Search
@REM https://www.google.co.uk/search?q=batch+search+string+in+variable&oq=BATCH+TO+SEARCH+STRING+IN+VAR&aqs=chrome.1.69i57j0l2.21405j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1093 -- Sat 27-Aug-2016 19:56:01
@REM COMMAND LINE IF SET
@REM ----
@REM COMMAND LINE IF SET - Google Search
@REM https://www.google.co.uk/search?q=COMMAND+LINE+IF+SET&oq=COMMAND+LINE+IF+SET+&aqs=chrome..69i57.4387j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1094 -- Sat 27-Aug-2016 19:56:15
@REM COMMAND LINE CLEAR A VARIABLE
@REM ----
@REM COMMAND LINE CLEAR A VARIABLE - Google Search
@REM https://www.google.co.uk/search?q=COMMAND+LINE+CLEAR+A+VARIABLE&oq=COMMAND+LINE+CLEAR+A+VARIABLE&aqs=chrome..69i57j0l5.7332j0j7&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1096 -- Sat 27-Aug-2016 19:56:28
@REM %~n1 - expands %1 to a file name only
@REM ----
@REM %~n1 - expands %1 to a file name only - Google Search
@REM https://www.google.co.uk/search?q=%25~n1+-+expands+%251+to+a+file+name+only&oq=%25~n1++++++++-+expands+%251+to+a+file+name+only&aqs=chrome..69i57&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1097 -- Sat 27-Aug-2016 19:56:39
@REM for /f
@REM ----
@REM for /f - Google Search
@REM https://www.google.co.uk/search?q=for+%2Ff&oq=for+%2Ff&aqs=chrome..69i57j69i60l2j69i65j69i60&sourceid=chrome&ie=UTF-8
@REM ----
@REM Count = 1098 -- Sat 27-Aug-2016 19:56:51
@REM batch loop through files in a directory
@REM ----
@REM batch loop through files in a directory - Google Search
@REM https://www.google.co.uk/search?num=30&q=batch+loop+through+files+in+a+directory&oq=BATCH+FILE+TO+RECURSE+EVERY+FOLDER&gs_l=serp.1.5.0i71k1l8.0.0.0.21453.0.0.0.0.0.0.0.0..0.0....0...1c..64.serp..0.0.0.3wP3bt6FGD4
@REM ----
@REM Count = 1099 -- Sat 27-Aug-2016 19:56:56
@REM BATCH FILE TO RECURSE EVERY FOLDER
@REM ----
@REM BATCH FILE TO RECURSE EVERY FOLDER - Google Search
@REM https://www.google.co.uk/search?q=BATCH+FILE+TO+RECURSE+EVERY+FOLDER&oq=BATCH+FILE+TO+RECURSE+EVERY+FOLDER+&aqs=chrome..69i57.10901j0j7&sourceid=chrome&ie=UTF-8
----


:VERY_ENDE