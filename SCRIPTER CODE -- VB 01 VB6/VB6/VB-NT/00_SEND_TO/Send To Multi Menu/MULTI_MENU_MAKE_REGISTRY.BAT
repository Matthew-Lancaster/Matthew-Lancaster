@ECHO OFF
REM -- REM COMMENTER DIDN'T WORK

D:
CD\


REM -- COMMENT
REM -- ----------------------------------------------------------------
REM -- ----------------------------------------------------------------
REM -- FIRST READ NOTE AT END IF WANT LEARN ABOUT SET ADMIN FOR ALL
REM -- SEARCH HERE

REM -- ----------------------------------------------------------------
REM -- ----------------------------------------------------------------
REM -- NEW LINK SHOW HOW TO RUN ALL PROGRAM ALWAYS IN ADMIN
REM -- ----
REM -- How to grant or get Elevated Privileges in Windows 10/8/7
REM -- http://www.thewindowsclub.com/elevated-privileges-windows
REM -- ----
REM -- ----------------------------------------------------------------
REM -- ----------------------------------------------------------------
REM -- ENDCOMMENT


REM -- COMMENT
REM -- ----------------------------------------------------------------
REM -- ----------------------------------------------------------------
REM -- BEGIN CODE TO SET MULTI MENU ON THE PROGRAM
REM -- ALSO IN HERE
REM -- ----------------------------------------------------------------
REM -- AND SOURCE CODE ABOUT COMMENT LINE
REM -- ----------------------------------------------------------------
REM -- ----------------------------------------------------------------

REM -- ----------------------------------------------------------------
REM -- THIS SHELL MENU SETTING CAN BE VERIFIED HERE
REM -- FOR FILE AND FOLDER OPERATION IN EXPLORER FILES SCRIPT WINDOW

REM -- C:\PSTART\PROGS\0_NIRSOFT_PACKAGE\NIRSOFT\SHELLMENUNEW.EXE

REM -- THE 2 VERSION BOTH COMPATIBLE 64 BIT AND 32 
REM -- IF USE A PROGRAM TO LOOK AT THE PATH OF EXE BY SPY AND IT IS 32 BIT AND THEN OBSERVING IS 64 BIT THEN IT WONT

REM -- LIKE ELITE SPY VISUAL BASIC

REM -- THE FOLDER TREE ONE WILL BE ANOTHER
REM -- WHERE USUAL TREESIZE FREE WOULD BE
REM -- THE OTHER HERE AND IS ONLY ONE NIRSOFT 32 BIT VERSION
REM -- FAULT IN TALK
REM -- TOP VERSION THAT HAS TWO 32 AND 64 BIT
REM -- 32 BIT NOT WORK AS WELL 64 BIT FIND MORE

REM -- AND SHELLEXEVIEW -- IS 64 BIT ONLY --THINKER BECAUSE ELITE SPY NOT FIND
REM -- C:\PSTART\PROGS\0_NIRSOFT_PACKAGE\NIRSOFT\X64\SHEXVIEW.EXE
REM -- YES SO NIRSOFT SOMETIME SHOW BOTH OR ONE
REM -- YES TREESIZE IS IN THE FIRST ONE -- \SHELLMENUNEW.EXE
REM -- MOST ARE SET THE OPTION TO DISPLAY MORE

REM -- ----------------------------------------------------------------
REM -- ----------------------------------------------------------------
REM -- ENDCOMMENT

REM ----     CODE EXECUTE TO RUN THE REGISTRY SET MULTI_MENU FOR MY VISUAL BASIC CONTEXT MENU


ECHO RUN -- "D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe
@ECHO ON
REG.EXE ADD "HKEY_CLASSES_ROOT\*\shell\Multi --------\Command" /f /ve /t REG_EXPAND_SZ /d "D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe %%1"
@ECHO OFF
ECHO -------------------------------------------------------------------


ECHO.
ECHO RUN -- "D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe %%1"
@ECHO ON
REG.EXE ADD "HKEY_CLASSES_ROOT\Directory\shell\TreeSize Free\Command" /f /ve /t REG_EXPAND_SZ /d "D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe %%1"
@ECHO OFF
ECHO -------------------------------------------------------------------

ECHO.
ECHO RUN  MORE OR QUIT 
ECHO SET WILL BE DONE 
ECHO. 
ECHO 1.... TREESIZE AND TREESIZE PRO
ECHO ------------- 
ECHO AND
ECHO ------------- 
ECHO 2.... NOTEPAD++
ECHO. 
ECHO MIGHT MUCK UP NOTEPAD NEED ASSOCIATION BACK BY
ECHO RUN WITH THIS ALWAYS AND SET


REM -- COMMENT
REM -- ----------------------------------------------------------------
REM -- ----------------------------------------------------------------
REM -- OPTIONAL ---- SET TREESIZE FREE
REM -- DON'T FORGET TREESIZE FREE SOMETIME GET LOST IN THE CONTEXT MENU
REM -- ---- THAT IS HOW SET A DIRECTORY
REM -- ----------------------------------------------------------------
REM -- ----------------------------------------------------------------
REM -- ENDCOMMENT

REM -- ECHO.
REM -- ECHO RUN HERE THE TREESIZE FREE AND PROFESSIONAL LATER -- LATER WORK TO DO
REM -- ECHO.
REM -- ECHO "C:\Program Files (x86)\JAM Software\TreeSize Free\TreeSizeFree.exe %%1"
REM -- ECHO.

@ECHO OFF

REM -- @ECHO ON
REM -- HKEY_CLASSES_ROOT\Directory\shell\TreeSize Free\Command
REM -- REG.EXE ADD "HKEY_CLASSES_ROOT\Directory\shell\TreeSize Free\Command" /f /ve /t REG_EXPAND_SZ /d "C:\Program Files (x86)\JAM Software\TreeSize Free\TreeSizeFree.exe %%1"
REM -- @ECHO OFF

REM -- ECHO NOTEPADD++ IS HARDCODED REMED OUT
REM -- COMMENT
REM -- ----------------------------------------------------------------
REM -- ----------------------------------------------------------------
REM -- OPTIONAL ---- SET NOTEPAD++ TO ALL TXT FILE
REM -- ----------------------------------------------------------------
REM -- REG.EXE ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "C:\Program Files\Notepad++\notepad++.exe %%1"

REM -- ----------------------------------------------------------------
REM -- ----------------------------------------------------------------
REM -- ENDCOMMENT


ECHO.
ECHO -- ENDING HAPPEN NEXT ABOUT TO EXIT
ECHO.
PAUSE

EXIT

GOTO ENDE

REM -- LOOK LIKE THAT THE END OF IT 
MUST OF READ REMARK COMMENTER OR ANOTHER LANGUAGE

LEAVE THE REST TO GO THEN
EXIT ABOVE
REQUIRE GOTO END BEFORE EXIT
----------------------------------------------------------------
----------------------------------------------------------------


COMMENT
----------------------------------------------------------------
----------------------------------------------------------------
FROM NIRSOFT

Multi --------	*	Multi --------	*		No	No	No	No	10/08/2016 17:53:19	D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe	"D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe" "%1"		1.00.0053	#0 Send To Multi Menu	Magik -- The One -- Hammer HardCore Coder	HKEY_CLASSES_ROOT\*\shell\Multi --------\Command	


----------------------------------------------------------------
----------------------------------------------------------------
ENDCOMMENT





COMMENT
----------------------------------------------------------------
----------------------------------------------------------------
@ECHO ---------- SET NOTEPAD++ FOR EDIT FILES
-----------------------------------------------------------------------------
reg.exe ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "C:\Program Files\Notepad++\notepad++.exe %%1"
-----------------------------------------------------------------------------
-----------------------------------------------------------------------------
@ECHO ---------- RESTORE ORGINAL NOTEPAD ------------------------------------
-----------------------------------------------------------------------------
reg.exe ADD "HKEY_CLASSES_ROOT\batfile\shell\edit\command" /f /ve /t REG_EXPAND_SZ /d "%%SystemRoot%%\System32\NOTEPAD.EXE %%1"

----------------------------------------------------------------
ENDCOMMENT






COMMENT
---------------------------------------------------------

NEW LINK SHOW HOW TO RUN ALL PROGRAM ALWAYS IN ADMIN
----
How to grant or get Elevated Privileges in Windows 10/8/7
http://www.thewindowsclub.com/elevated-privileges-windows
----

--------------------------------------------------------------------
RUN THROUGH 1 STEP 2 3 
--------------------------------------------------------------------
OPEN COMMAND PROMPT WINDOW WITH ADMINISTRATOR 
--------------------------------------------------------------------
--------------------------------------------------------------------
WITH SHIFT OR CONTROL KEY WHILE RIGHT CLICK IN EXPLORER TO ENABLE THE CONTEXT OPTION TO SHOW
--------------------------------------------------------------------

--------------------------------------------------------------------
TYPE THIS LINE IN ---- PASTE IT IN ---- IF POSSIBLE
MIGHT TAKE FEW SECOND TO LOAD FIRST TIME
--------------------------------------------------------------------
SECPOL.MSC
--------------------------------------------------------------------

LOOK AT THE TREE AND NODE PLUS CLICKER
-------------------------------------------------------
1.. 2ND DOWN -- LOCAL POLICE'S
NEXT BRANCH 
-------------------------------------------------------
2.. SECURITY OPTIONS
NEXT BRANCH FEW PAGE DOWN AND THIRD FROM BOTTOM
-------------------------------------------------------
3.. User Account Control: RUN ALL ADMINISTRATORS IN ADMIN APPROVAL MODE

-------------------------------------------------------
HERE SET FROM ENABLED TO DISABLED
REQUIRE A REBOOT
-------------------------------------------------------

INTRUCTION FROM GUI FORM
-------------------------------------------------------
User Account Control: Turn on Admin Approval Mode

This policy setting controls the behavior of all User Account Control (UAC) policy settings for the computer. If you change this policy setting, you must restart your computer.

The options are:

• Enabled: (Default) Admin Approval Mode is enabled. This policy must be enabled and related UAC policy settings must also be set appropriately to allow the built-in Administrator account and all other users who are members of the Administrators group to run in Admin Approval Mode. 

• Disabled: Admin Approval Mode and all related UAC policy settings are disabled. Note: If this policy setting is disabled, the Security Center notifies you that the overall security of the operating system has been reduced.
-------------------------------------------------------

-------------------------------------------------------
WENT THOUGHT A LOAD OF WEB SITE 
ALL TALK ABOUT LOADING ONE UP
NOT GOOD IN MY EYE FOR A PROGRAMMER THAT CHANGE RECOMPILE THE EXE
AND THINGS CHANGE

AND FINALLY AT MS WINDOWS CLUB
LOOK LIKE THERE ARE ADVERTISEMENT ON SITE
AND IDIOT WORDIN
-------------------------------------------------------

1 STEP 2 STEP 3 DONE ALL I WANT WITH THE DRIBBLE DROOL

-------------------------------------------------------
-------------------------------------------------------
ENDCOMMENT 





COMMENT
----------------------------------------------------------------
----------------------------------------------------------------

----
COMMENT - Mark a block of comments in a batch file
https://jpsoft.com/help/comment.htm
----
NEW LEARN  Fri 04 November 2016 11:44:38----------
----------------------------------------------------------------

----------------------------------------------------------------
----------------------------------------------------------------
ENDCOMMENT 


EXIT
ENDE:
