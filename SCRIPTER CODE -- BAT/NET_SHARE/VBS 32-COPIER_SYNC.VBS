'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 32-COPIER_SYNC.VBS
'# __ 
'# __ VBS 32-COPIER_SYNC.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START     TIME [ Fri 05-Oct-2018 21:30:00 ]
'# END FINAL TIME [ Fri 05-Oct-2018 22:38:00 ]
'# __ 
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
' -------------------------------------------------------------------
' WHAT DOS'S XCOPY CAN'T DO
' OR ROBOCOPY IS ONLY DESIGN UP TO 32 BIT AND WIN XP
' XCOPY COPY ONE LATEST FILE OVER ANOTHER
' IF THINK GOING TO BE RECYCLING COPY OF PATH HIGHER THAN OTHER

' -------------------------------------------------------------------

'#-------------------------------------------------------------------
' -- 001 SESSION __
' -------------------------------------------------------------------
' WELL I MANAGED THE __  REDIM PRESERVE'S GOING IN VBS
' ONLY THING OF TWO JOB
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' -- 002 SESSION __
' -------------------------------------------------------------------
' ANOTHER ROUND OF ADDITION AND BETTER THE CODE ARRAY WORKING NICELY 
' WITH VARIABLE FLAG WITH COPY
' -------------------------------------------------------------------
' __ FROM  Sun 07-Oct-2018 04:26:00
' __ TO    Sun 07-Oct-2018 05:58:00 _ MIXED WITH OTHER CODE WRITER TIDY UP-PER
'#-------------------------------------------------------------------



Const ForReading = 1, ForWriting = 2, ForAppending = 8
const DontWaitUntilFinished = false, ShowWindow = 1, DontShowWindow = 0, WaitUntilFinished = true

DIM CUR_USER
DIM ComputerName

CUR_USER = CreateObject("WScript.Network").UserName
Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

DIM objSuperFolder

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO_2 = CreateObject("Scripting.FileSystemObject")

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")

DIM VAR_PATH_FILE_1()
DIM VAR_PATH_FILE_2()
DIM VAR_DO_HAVE_TO_EXIST()
DIM VAR_DO_SECONDARY_PATH_HAVE_TO_EXIST()
DIM VAR_DO_SECONDARY_FULL_PATH_TO_EXIST()

I_COUNTER = 30

REDIM VAR_PATH_FILE_1(I_COUNTER)
REDIM VAR_PATH_FILE_2(I_COUNTER)

' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' -- SET HERE FOR 1ST TIME USE AND FLAG IT TOGGLE OFF MODE AFTER
' -------------------------------------------------------------------
REDIM VAR_DO_HAVE_TO_EXIST(I_COUNTER)
' -------------------------------------------------------------------
' -- TRIM PATH TO END BACKSLASH FOLDER NAME 
' -------------------------------------------------------------------
REDIM VAR_DO_SECONDARY_PATH_HAVE_TO_EXIST(I_COUNTER)
' -------------------------------------------------------------------
' -- FULL PATH INCLUDE FILENAME
' -------------------------------------------------------------------
REDIM VAR_DO_SECONDARY_FULL_PATH_TO_EXIST(I_COUNTER)

I_COUNT=0

' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------

' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\BAT 29-NET_SHARE_ALL.BAT"
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\BAT 29-NET_SHARE_ALL.BAT"
I2="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 29-NET_SHARE_ALL.BAT"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------

' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\VBS 32-COPIER_SYNC.VBS"
I2="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 32-COPIER_SYNC.VBS"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------

' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- REG KEY SETTER\REG KEY\# Windows 7 REG SCRIPT\# TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT"
I2="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 00 # TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' -------------------------------------------------------------------
' 64 BIT NOTEPAD++ AUTOHOTKEYS LANGUAGE IN
' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\NOTEPAD++ AUTOHOTKEYS LANGUAGE PLUG-IN\userDefineLang_AHK.xml"
I2="C:\Program Files\Notepad++\userDefineLang_AHK.xml"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_DO_SECONDARY_PATH_HAVE_TO_EXIST(I_COUNT)="TRUE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2

' -------------------------------------------------------------------
' 32 BIT NOTEPAD++ AUTOHOTKEYS LANGUAGE IN
' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\NOTEPAD++ AUTOHOTKEYS LANGUAGE PLUG-IN\userDefineLang_AHK.xml"
I2="C:\Program Files (x86)\Notepad++\userDefineLang_AHK.xml"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_DO_SECONDARY_PATH_HAVE_TO_EXIST(I_COUNT)="TRUE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------
' -------------------------------------------------------------------

' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 45-SCRIPT RUN GITHUB.BAT"
I2="C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB.BAT"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------

' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 45-SCRIPT RUN GITHUB - BAT RUNNER.BAT"
I2="C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - BAT RUNNER.BAT"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------

' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 45-SCRIPT RUN GITHUB - AUTO UPDATE EXAMPLE.BAT"
I2="C:\SCRIPTER\SCRIPTER CODE -- GITHUB\BAT 45-SCRIPT RUN GITHUB - AUTO UPDATE EXAMPLE.BAT"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------

' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 17-NPP TO REPLACE NOTEPAD.BAT"
I2="C:\SCRIPTER\SCRIPTER CODE -- REG KEY SETTER\REG KEY\# Windows 7 REG SCRIPT\NOTEPAD++ WORK\notepad++ replace notepad\BAT 17-NPP TO REPLACE NOTEPAD.BAT"
' -------------------------------------------------------------------
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' AFTER SETTING PERMISSION TO EVERYONE IN PROPERTY GUI DIALOGUE
' IT TOOK A FEW SECOND TO UPDATE 
' FIRST ATTEMPT TO SET PERMISSION HERE NOT WORK FAIL
' AND THEN RUN THIS CODE AND OKAY
' AND THEN RUN SET THESE PERMISSION AND OKAY WITHOUT FAILING
' RUN IF YOU WANT TO ONCE SET NOT REQUIRE CODE
' GET AN MSGBOX ERROR ACCESS WHEN COPY IF FAIL 
' -------------------------------------------------------------------
IF MAYBE_REQUIRE_HERE=MAYBE_REQUIRE_HERE+2 THEN 
	' SET PATH_TO_DO="C:\Windows\System32\gatherNetworkInfo.vbs"
	' ECHO %PATH_TO_DO%
	' ICACLS %PATH_TO_DO% /setowner Everyone /T /C /L /Q 
	' ICACLS %PATH_TO_DO% /T /L /C /Q /Grant Everyone:F
	' ICACLS %PATH_TO_DO% /setintegritylevel (CI)H
END IF
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' FOUND IN THE WINDOWS FOLDER AS A SCHEDULER TASK RUNNER BY MICROSOFT
' WELL WRITTEN CODE FOR FINDER INFO ABOUT ALLSORT OF THING
' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 33-GatherNetworkInfo.vbs"
I2="C:\Windows\System32\gatherNetworkInfo.vbs"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_DO_SECONDARY_FULL_PATH_TO_EXIST(I_COUNT)="TRUE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------


' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 47-OWNER-HARDCODED ANYWHERE.BAT"
I2="C:\SCRIPTER\SCRIPTER CODE -- BAT\OWNER\#_OWNER-HARDCODED ANYWHERE.BAT"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------

' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 48-OWNER-C-DRIVE-ONLY & NOT-SUB-FOLDER.BAT"
I2="C:\SCRIPTER\SCRIPTER CODE -- BAT\OWNER\OWNER-C-DRIVE - ONLY & NOT-SUB-FOLDER.BAT"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------


' -------------------------------------------------------------------
' CAMERA 01 OF 02 
' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 52-XCOPY SYNC.BAT"
I2="D:\DSC\2018 Double Screen Cam\INFO _ DOUBLE SCREEN CAMERA\DOUBLE SCREEN CAMERA YELLOW WITHOUT HARDLY AUDIO.BAT"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' CAMERA 02 OF 02 
' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 52-XCOPY SYNC.BAT"
I2="D:\DSC\2018 GoPro 01\INFO _ HD DV CAM YELLOW\Sports HD DV Cam Ver 20170418 V1.0A.BAT"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------

' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 00 # TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT"
I2="E:\REG KEY SETTINGS\REG KEY\# Windows 7 REG SCRIPT\# TWEAK REG WIN 7 - NEW USER SETUP - SPECIAL FOLDERS.BAT"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------

' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 54-REGJUMP_CLIP_BOARD.BAT"
I2="C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\SysinternalsSuite\REGJUMP_CLIP_BOARD.BAT"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------


' -------------------------------------------------------------------
I1="C:\SCRIPTER\SCRIPTER CODE -- VB6\VB6 Programming Snippet-Set.TXT.BAS"
I2="D:\VB6\VB-NT\#1 Documents\Programming Snippets.txt"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------


' -------------------------------------------------------------------
I1="C:\Program Files\HandBrake\libdvdcss-2.dll"
I2="D:\#0 1 INSTALLATIONS\00 INSTALLER_SET\# 00 Install Progs\# Installed Now\#00 Paid For\HANDBRAKE DVD COPIER\VLC\LIBDVDCSS\1.2.12\WIN64\libdvdcss-2.dll"
' -------------------------------------------------------------------
I_COUNT=I_COUNT+1
VAR_DO_HAVE_TO_EXIST(I_COUNT)="FALSE"
VAR_PATH_FILE_1(I_COUNT)=I1
VAR_PATH_FILE_2(I_COUNT)=I2
' -------------------------------------------------------------------


' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------

REDIM PRESERVE VAR_PATH_FILE_1(I_COUNT)
REDIM PRESERVE VAR_PATH_FILE_2(I_COUNT)
REDIM PRESERVE VAR_DO_HAVE_TO_EXIST(I_COUNTER)
REDIM PRESERVE VAR_DO_SECONDARY_PATH_HAVE_TO_EXIST(I_COUNTER)
REDIM PRESERVE VAR_DO_SECONDARY_FULL_PATH_TO_EXIST(I_COUNTER)

FOR I_COUNT=1 TO UBOUND(VAR_PATH_FILE_1)

	FileModDate_1=0
	FileModDate_2=0

	IF VAR_PATH_FILE_1(I_COUNT)<>"" THEN 
		
		SET_GO=TRUE
		IF FSO.FileExists(VAR_PATH_FILE_1(I_COUNT))=FALSE THEN SET_GO=FALSE
		IF FSO.FileExists(VAR_PATH_FILE_2(I_COUNT))=FALSE THEN SET_GO=FALSE
		IF VAR_DO_HAVE_TO_EXIST(I_COUNT)<>"" THEN              SET_GO=TRUE
		
		IF VAR_DO_SECONDARY_PATH_HAVE_TO_EXIST(I_COUNT)="TRUE" THEN
			STRIP_FOLDER=MID(VAR_PATH_FILE_2(I_COUNT),1,INSTRREV(VAR_PATH_FILE_2(I_COUNT),"\"))
			If FSO.FolderExists(STRIP_FOLDER)=FALSE Then
				SET_GO=FALSE
			END IF
		END IF
		IF VAR_DO_SECONDARY_FULL_PATH_TO_EXIST(I_COUNT)="TRUE" THEN
			If FSO.FileExists(VAR_PATH_FILE_2(I_COUNT))=FALSE Then
				SET_GO=FALSE
			END IF
		END IF
		
		IF SET_GO=TRUE THEN 
		
			FileModDate_1=0
			FileModDate_2=0
			IF FSO.FileExists(VAR_PATH_FILE_1(I_COUNT)) THEN 
				Set F1 = fso.GetFile(VAR_PATH_FILE_1(I_COUNT))
				FileModDate_1 = F1.DateLastModified
			END IF
			IF FSO.FileExists(VAR_PATH_FILE_2(I_COUNT)) THEN 
				Set F2 = fso.GetFile(VAR_PATH_FILE_2(I_COUNT))
				FileModDate_2 = F2.DateLastModified
			END IF
			SET F1=NOTHING
			SET F2=NOTHING

			SET_GO=TRUE
			IF FileModDate_1=0 THEN SET_GO=FALSE
			IF FileModDate_2=0 THEN SET_GO=FALSE
			' IF SET ANSWER IS GIVE FALSE
			IF VAR_DO_HAVE_TO_EXIST(I_COUNT)<>"" THEN SET_GO=TRUE
			' -----------------------------------------------------------
			' NORMALLY ENTERED AS "FALSE" OR TRUE IS NOTHING EMPTY VOID
			' HOPEFULLY IT WILL PICK-UP THE ONE THE EXISTING BY DATE VALUE
			' -----------------------------------------------------------

			IF SET_GO=TRUE THEN

				IF FileModDate_1>FileModDate_2 THEN

				IF FSO.FileExists(VAR_PATH_FILE_1(I_COUNT))=FALSE THEN 
						MSGBOX "NOT EXIST "+VBCRLF+VAR_PATH_FILE_1(I_COUNT)
					END IF
					
					IF FSO.FileExists(VAR_PATH_FILE_1(I_COUNT)) THEN 
						Fso.CopyFile VAR_PATH_FILE_1(I_COUNT), VAR_PATH_FILE_2(I_COUNT)	
		
					END IF
				END IF

				IF FileModDate_1<FileModDate_2 THEN
					IF FSO.FileExists(VAR_PATH_FILE_2(I_COUNT))=FALSE THEN 
						MSGBOX "NOT EXIST "+VBCRLF+VAR_PATH_FILE_2(I_COUNT)
					END IF
					IF FSO.FileExists(VAR_PATH_FILE_2(I_COUNT)) THEN 
						Fso.CopyFile VAR_PATH_FILE_2(I_COUNT), VAR_PATH_FILE_1(I_COUNT)
					END IF
				END IF
			
			END IF
		END IF
	END IF
NEXT
' -------------------------------------------------------------------



' -------------------------------------------------------------------
' THIS PROCESS NOT TOO GOOD AS THESE SCRIPT IN BAT WAKE ALL THE DRIVE WHEN SCAN THEM FOR FOLDER
' AS PART OF SYNC DO
' INCORP IN TO CAMERA LOADER
' -------------------------------------------------------------------

REDIM VAR_PATH_FILE_1(I_COUNTER)
REDIM VAR_PATH_FILE_2(I_COUNTER)

I_COUNT=0
'I_COUNT=I_COUNT+1
'VAR_PATH_FILE_1(I_COUNT)="D:\DSC\2018 GoPro 01\INFO _ HD DV CAM YELLOW\Sports HD DV Cam Ver 20170418 V1.0A.BAT"
'VAR_PATH_FILE_2(I_COUNT)="D:\DSC\2018 Double Screen Cam\INFO _ DOUBLE SCREEN CAMERA\DOUBLE SCREEN CAMERA YELLOW WITHOUT HARDLY AUDIO.BAT"


REDIM PRESERVE VAR_PATH_FILE_1(I_COUNT)
REDIM PRESERVE VAR_PATH_FILE_2(I_COUNT)

FOR I_COUNT=1 TO UBOUND(VAR_PATH_FILE_1)
	IF VAR_PATH_FILE_1(I_COUNT)<>"" THEN 
		IF FSO.FileExists(VAR_PATH_FILE_1(I_COUNT)) THEN 
			WSHShell.Run """"+VAR_PATH_FILE_1(I_COUNT)+"""", DontShowWindow, DontWaitUntilFinished
		END IF
	END IF
NEXT

Wscript.QUIT

' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------

' ----
' VBScript Arrays: Using DIM, REDIM Statements and Ubound Array Functions
' https://www.softwaretestinghelp.com/vbscript-arrays-tutorial-7/
' ----

' ----
' VbsEdit - VBScript Editor with Debugger - VBS Editor
' https://www.vbsedit.com/
' ----

' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------



DIM JPG_BEEN_DONE

DIM STOP_PROGRESS
DIM IX_ERROR_MESSENGER

DIM DRIVE_SPACE_INFO

DIM objStartFolder
DIM objFolder
DIM I_COUNT_1
DIM I_COUNT_2
DIM VAR_PATH_SCANNER
DIM IMAGE_POINTER_NAME


Dim arr1
Dim oDict ' Key: Full filename, Item: Filename

Dim FileModDate
DIM FileModDate_3

'FOLDER_PATH_VAR="D:\DSC\2011"
'FOLDER_PATH_VAR="D:\DSC\z compare\NIKON TOTAL\DCIM"
'FOLDER_PATH_VAR="\"

' IF FSO.FolderExists(FOLDER_PATH_VAR)=FALSE THEN 
	' MSGBOX "FOLDER NOT EXIST"
	' Wscript.QUIT
' END IF


'Call the Search subroutine to start the recursive search.
'Search objFSO.GetFolder(FOLDER_PATH_VAR)
'---------------------------------------------------------
'FindWantedFiles objFSO.GetFolder(FOLDER_PATH_VAR)
'DisplayFindings
'---------------------------------------------------------


DIM FOLDER_PATH_INDEX

DIM FOLDER_PATH(7)
DIM FOLDER_PATH_DEST(7)
DIM FOLDER_PATH_VAR(7)
DIM FOLDER_PATH_DEST_VAR
DIM FOLDER_PATH_VAR_3

DIM EXPLORER_LOADED
DIM SHOW_EXPLORER_FLAG

FOLDER_PATH(1)="MC - HX60V"
FOLDER_PATH(2)="MC - HX60V0"
FOLDER_PATH(3)="FUJI XP90"
FOLDER_PATH(4)="GOPRO"
FOLDER_PATH(5)="DUAL SCRN" ' __ COPY FIRST TO OE PLACE
FOLDER_PATH(6)="DUAL SCRN" ' __ MOVE NEXT TO ANOTHER
FOLDER_PATH(7)="D DRIVE"

' NOKIA TO DO 1 2 3

FOLDER_PATH_VAR(1)="\"
FOLDER_PATH_VAR(2)="\"
FOLDER_PATH_VAR(3)="\"
FOLDER_PATH_VAR(4)="\"
FOLDER_PATH_VAR(5)="\"
FOLDER_PATH_VAR(6)="\"
FOLDER_PATH_VAR(7)="\Android\SDCard2_Sync\DCIM\"

FOLDER_PATH_DEST(1)="D:\DSC\2015-2018\2018 CyberShot HX60V"
FOLDER_PATH_DEST(2)="D:\VI_ DSC ME\2010-2018 CyberShot\2018 CyberShot HX60V_#"
FOLDER_PATH_DEST(3)="D:\DSC\2017-2018 FUJI XP90\2018 FinePix XP90"
IF Year(Now)>2018 THEN
	FOLDER_PATH_DEST(1)="D:\DSC\" + CSTR(Year(Now)) + " CyberShot HX60V"
	FOLDER_PATH_DEST(2)="D:\VI_ DSC ME\" + CSTR(Year(Now)) + " CyberShot HX60V_#"
	FOLDER_PATH_DEST(3)="D:\DSC\" + CSTR(Year(Now)) + " FUJI FinePix XP90"
END IF
FOLDER_PATH_DEST(4)="D:\DSC\" + CSTR(Year(Now)) + " GoPro 01"
FOLDER_PATH_DEST(5)="D:\DSC\" + CSTR(Year(Now)) + " Double Screen Cam"
FOLDER_PATH_DEST(6)="D:\VI_ DSC ME\" + CSTR(Year(Now)) + " Double Screen Cam"
FOLDER_PATH_DEST(7)="D:\DSC\" + CSTR(Year(Now)) + " Samsung Mini S5\DCIM"


FOR R_COUNT=1 TO UBOUND(FOLDER_PATH_DEST)
	If Not FSO.FolderExists(FOLDER_PATH_DEST(R_COUNT)) Then
		BuildFullPath FOLDER_PATH_DEST(R_COUNT)
	End If
NEXT

FOR R_COUNT=1 TO UBOUND(FOLDER_PATH)

	I_COUNT_1=0
	I_COUNT_2=0

	FOLDER_PATH_INDEX=CSTR(R_COUNT)+" "+FOLDER_PATH(R_COUNT)
	FOLDER_PATH_VAR_2=getDriveLetterFromVolumeName(FOLDER_PATH(R_COUNT))
	FOLDER_PATH_DEST_VAR=FOLDER_PATH_DEST(R_COUNT)
	
	FOLDER_PATH_VAR_3=FOLDER_PATH_VAR(R_COUNT)
	
	IF FOLDER_PATH_VAR_2<>"" THEN 
		
		WORK_GO_ER=TRUE
		
		
		FOLDER_PATH_VAR_2=FOLDER_PATH_VAR_2+FOLDER_PATH_VAR(R_COUNT)

		' MSGBOX FOLDER_PATH_VAR_2
		' Wscript.QUIT
		
		FOLDER_PATH_DEST_VAR_2=FOLDER_PATH_DEST(R_COUNT)
		
		'---------------------------------------------------------------------------------
		' DON'T SHOW EXPLORER ON DODGY COMPUTER AROUND BUILDING AND IF DODGY PATH IMAGE-ER
		' OR DODGY LESS QUICKER SPEC LOW END PC
		'---------------------------------------------------------------------------------
		' DON'T SHOW EXPLORER ON D DRIVE IF ONLY PICTURE IS NONE
		'---------------------------------------------------------------------------------
		SHOW_EXPLORER_FLAG=TRUE
		IF INSTR(FOLDER_PATH_DEST_VAR,"D:\VI_ DSC ME\") > 0 THEN 
			IF ComputerName="1-ASUS-X5DIJ" THEN SHOW_EXPLORER_FLAG=FALSE
			IF ComputerName="2-ASUS-EEE" THEN SHOW_EXPLORER_FLAG=FALSE
			IF ComputerName="3-LINDA-PC" THEN SHOW_EXPLORER_FLAG=FALSE
			IF ComputerName="4-ASUS-GL522VW" THEN SHOW_EXPLORER_FLAG=FALSE
		END IF
		
		IF FOLDER_PATH_INDEX="7 "+FOLDER_PATH(7) THEN  SHOW_EXPLORER_FLAG=FALSE
		
		EXPLORER_LOADED=FALSE
		IF SHOW_EXPLORER_FLAG=TRUE THEN 
			EXPLORER_LOADED=TRUE
			WSHShell.Run "EXPLORER """+FOLDER_PATH_DEST_VAR_2+"""" , ShowWindow, DontWaitUntilFinished
		END IF
		
		IMAGE_POINTER_FOLDER_1=MID(Wscript.ScriptFullName,1,INSTRREV(Wscript.ScriptFullName,"\"))+"VBS 29-COPY CAMERA\"
		IMAGE_POINTER_NAME=IMAGE_POINTER_FOLDER_1+"COPY_CAMERA_PHOTO_IMAGE_POINTER_"+FOLDER_PATH(R_COUNT)+".TXT"

		If Not FSO.FolderExists(IMAGE_POINTER_FOLDER_1) Then
			BuildFullPath IMAGE_POINTER_FOLDER_1
		End If

		IF FSO.FileExists(IMAGE_POINTER_NAME)=TRUE THEN 
			Set File_1 = fso.OpenTextFile(IMAGE_POINTER_NAME, ForReading)
			TT_DATE_POINTER = File_1.ReadAll
			set File_1=NOTHING
		END IF
			
		DIM FileModDate_2
		FileModDate_2=CDate(TT_DATE_POINTER)
		FileModDate_3=FileModDate_2
		
		'TEST ---------------------------------------------
		'FileModDate_2=CDate("14/06/2018")
		'FileModDate_2=NOW
		'Set objFile = FSO.CreateTextFile(IMAGE_POINTER_NAME)
		'objFile.Write FileModDate_2
		'SET objFile=NOTHING
		' --------------------------------------------------

		MAIN

		If FSO.FileExists(IMAGE_POINTER_NAME) Then 
			FSO.DeleteFile IMAGE_POINTER_NAME
		END IF

		Set objFile = FSO.CreateTextFile(IMAGE_POINTER_NAME)
		objFile.Write FileModDate_3
		SET objFile=NOTHING
		
		
	END IF

NEXT

IF WORK_GO_ER=FALSE THEN

	I=""
	I=I+"CAMERA COPY _ NOTHING DONE " + VBCRLF
	I=I+"CHECKED VOLUME NAME" + VBCRLF
	I=I+"--------------------------" + VBCRLF
	FOR R_COUNT=1 TO UBOUND(FOLDER_PATH)
		I=I+FOLDER_PATH(R_COUNT) + VBCRLF
	NEXT
	MSGBOX I

END IF


SET objFSO = NOTHING
Wscript.QUIT
'-------------------------------------------------------------------
'-------------------------------------------------------------------

Sub Main
    Set oDict = CreateObject("Scripting.Dictionary")
    ' Call Recursive Sub
    FindWantedFiles(objFSO.GetFolder(FOLDER_PATH_VAR_2))
    ' Display All Findings from Dictionary object
    DisplayFindings

	ShowDriveInfo(MID(FOLDER_PATH_VAR_2,1,2))

	' ---------------------------------------------------------------
	' CHECK EXPLORER BEEN LOADED BEFORE _ IF NOT
	' NOW IS TIME TO MAKE SURE _ AND DECIDE IF THE COUNTER HAS GONE UP FOR ANY COPY FILE
	' WHILE STILL MAINTAIN NOT FIR ANY DODGY COMPUTER OR FOLDER
	' INCLUDE D DRIVE OR BE SHOW EVERY TIME
	' ---------------------------------------------------------------
	IF SHOW_EXPLORER_FLAG=TRUE AND EXPLORER_LOADED=FALSE AND I_COUNT_1>0 THEN 
		WSHShell.Run "EXPLORER """+FOLDER_PATH_DEST_VAR_2+"""" , ShowWindow, DontWaitUntilFinished
	END IF
	
	I=""
	I=I+"DONE COPY COUNT: " 
	I=I+CSTR(I_COUNT_1)
	I=I+" of " + CSTR(I_COUNT_2) + VBCRLF
	I=I+DRIVE_SPACE_INFO
	
	IF SHOW_EXPLORER_FLAG=TRUE OR I_COUNT_1>0 THEN 
		MSGBOX I
	END IF
	
	IF I_COUNT_1>=0 AND JPG_BEEN_DONE=TRUE THEN 
		
		WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 33-#0 Send To Date Of File In Front FileName.ahk""", ShowWindow, DontWaitUntilFinished
	
	END IF
	
    Set oDict = Nothing
End Sub


function ShowDriveInfo(drvPath)
	strComputer = "." 

	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
	Set colItems = objWMIService.ExecQuery( _
		"SELECT * FROM Win32_LogicalDisk where DeviceID='"+drvPath+"'",,48) 

	For Each objItem in colItems 
		if len(objItem.VolumeName)>0 then
			S=S+"-----------------------------------------------------------" & vbCrLf
			S=S+"VolumeName:" & vbTab & objItem.VolumeName  & vbCrLf
			S=S+"-----------------------------------------------------------" & vbCrLf
			S=S+"Drive Letter 1st Found for Match Volume Name: " & drvPath & "\" & vbCrLf
			S=S+"-----------------------------------------------------------" & vbCrLf
			S=S+"FreeSpace:" & vbTab
			S=S+FormatNumber((CDbl(objItem.FreeSpace)/1024/1024/1024))  & " GB" & vbCrLf
			S=S+"Size:" & vbTab & vbTab
			S=S+FormatNumber((CDbl(objItem.Size)/1024/1024/1024)) & " GB" & vbCrLf
			S=S+"Occupied Space:" & vbTab
			S=S+FormatNumber((CDbl(objItem.Size - objItem.FreeSpace)/1024/1024/1024)) & " GB"
			IF IX_ERROR_MESSENGER<>"" THEN 
				S=S+vbCrLf
				S=IX_ERROR_MESSENGER
			END IF
		END IF
	Next
	DRIVE_SPACE_INFO=S
END FUNCTION

Sub FindWantedFiles(sFolder)
    On Error Resume Next
    Dim oFDR, oItem, file
    ' Check if wanted files are in this folder
    For Each file In sFolder.Files
		oDict.Add FILE.path, file
    Next
    ' Recurse into it's sub folders
    For Each oFDR In sFolder.SubFolders
        FindWantedFiles objFSO.GetFolder(oFDR.Path)
    Next
	
End Sub

Sub DisplayFindings()
    Dim oKeys, oKey
    oKeys = oDict.Keys
    For Each oKey In oKeys
		VAR_PATH_SCANNER=oKey
		IF FSO.FILEExists(VAR_PATH_SCANNER)=TRUE THEN 
			CALL PATH_SCANNER_COPY_CAMERA
		END IF
	Next
End Sub

SUB PATH_SCANNER_COPY_CAMERA
	SET_GO_FILE=FALSE
	IF mid(ucase(VAR_PATH_SCANNER),3,5)="\DCIM" then    SET_GO_FILE=TRUE
	IF mid(ucase(VAR_PATH_SCANNER),3,8)="\MP_ROOT" then SET_GO_FILE=TRUE
	
	DO_ALL_FILE=FALSE
	' THIS IS THE GOPRO DRIVE AND WILL WANT ALL COPY 
	IF FOLDER_PATH_INDEX="4 "+FOLDER_PATH(4) THEN SET_GO_FILE=TRUE : DO_ALL_FILE=TRUE

	MOVE_MODE_FILE=FALSE
	DO_ALL_FILE=FALSE
	IGNOR_DATE_INDEX=FALSE
	' THIS IS THE DUAL SCREEN DRIVE & DO HERE IS 2 STEP 1ST COPY TO ONE FOLDER 2ND MOVE ANOTHER FOLDER
	' AND IGNORE DATE INDEX REQUIRED WILL BE OKAY AS MOVE MODE IS AFTER THAT ALSO IGNORE DATE INDEX
	IF FOLDER_PATH_INDEX="5 "+FOLDER_PATH(5) THEN 
		SET_GO_FILE=TRUE
		DO_ALL_FILE=TRUE
		IGNOR_DATE_INDEX=TRUE
	END IF
	
	' THIS IS THE DUAL SCREEN DRIVE AND WILL WANT ALL COPY BUT MOVE_MODE VIDEO & IMAGE
	IF FOLDER_PATH_INDEX="6 "+FOLDER_PATH(6) THEN 
		SET_GO_FILE=TRUE
		DO_ALL_FILE=TRUE
		MOVE_MODE_FILE=TRUE
		IGNOR_DATE_INDEX=TRUE
	END IF
	
	' SAMSUNG MINI S 5 _ ALLOW OVER WRITER & ALLOW ALL FILE EXTENSION TYPE TYPEO TYPO P & E & O
	IF FOLDER_PATH_INDEX="7 "+FOLDER_PATH(7) THEN 
		SET_GO_FILE=TRUE
		DO_ALL_FILE=TRUE
		' MOVE_MODE_FILE=TRUE
		' IGNOR_DATE_INDEX=TRUE
	END IF

	' IF INSTR(Lcase(VAR_PATH_SCANNER),"_gsdata_")>0 then SET_GO_FILE=FALSE
	IF INSTR(VAR_PATH_SCANNER,"_gsdata_")>0 then SET_GO_FILE=FALSE
	IF INSTR(VAR_PATH_SCANNER,"System Volume Information")>0 then SET_GO_FILE=FALSE
	
	IF RIGHT(ucase(VAR_PATH_SCANNER),4)=".THM" then SET_GO_FILE=FALSE

	'if INSTR(Ucase(VAR_PATH_SCANNER),"AVF_INFO")>0 then SET_GO_FILE=FALSE

	IF SET_GO_FILE=FALSE THEN EXIT SUB
	
	I_COUNT_2=I_COUNT_2+1
	
	IF STOP_PROGRESS=TRUE THEN EXIT SUB
	
	Set f = fso.GetFile(VAR_PATH_SCANNER)
	FileModDate = f.DateLastModified
	SET f=NOTHING

	IF IGNOR_DATE_INDEX=FALSE THEN 
		IF FileModDate<=FileModDate_2 THEN EXIT SUB
	END IF 
	
	If RIGHT(ucase(VAR_PATH_SCANNER),4)=".JPG" then JPG_BEEN_DONE=TRUE
	
	MOVE_MODE_FLAG=FALSE
	If RIGHT(ucase(VAR_PATH_SCANNER),4)=".JPG" then MOVE_MODE_FLAG=TRUE
	If RIGHT(ucase(VAR_PATH_SCANNER),4)=".AVI" then MOVE_MODE_FLAG=TRUE

	IF FileModDate>FileModDate_3 THEN FileModDate_3=FileModDate
	
	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)
	
	' PATH_VAR="D:\DSC\2015-2018\2018 CyberShot HX60V"+MID(PATH_VAR,3)
	PATH_VAR=FOLDER_PATH_DEST_VAR+MID(PATH_VAR,3)
	PATH_VAR=REPLACE(PATH_VAR,MID(FOLDER_PATH_VAR_3,2),"")
	
	' TRICKY BUT WILL IT RUN
	
	' MSGBOX "PATH_VAR  "+PATH_VAR
	' Wscript.QUIT
	
	PATH_FILE_VAR=PATH_VAR+"\"+FILE_VAR

	If Not FSO.FolderExists(PATH_VAR) Then
		BuildFullPath PATH_VAR
	End If
	
	ON ERROR RESUME NEXT
	ERR.CLEAR
	
	If Not FSO.FileExists(PATH_FILE_VAR) OR DO_ALL_FILE=TRUE Then
		IF MOVE_MODE_FILE=FALSE THEN 
			Fso.CopyFile VAR_PATH_SCANNER, PATH_FILE_VAR
		END IF
		IF MOVE_MODE_FILE=TRUE THEN 
			IF MOVE_MODE_FLAG=TRUE THEN
				Fso.MoveFile VAR_PATH_SCANNER, PATH_FILE_VAR
			END IF
			IF MOVE_MODE_FLAG=FALSE THEN
				Fso.CopyFile VAR_PATH_SCANNER, PATH_FILE_VAR
			END IF
		END IF
			
		I_COUNT_1=I_COUNT_1+1
	END IF
	IF ERR.NUMBER>0 THEN
		IX=""+VBCRLF
		IX=IX+"ERROR COPY FILE _ TRY REDO ANOTHER"+VBCRLF+VBCRLF
		IX=IX+VAR_PATH_SCANNER +VBCRLF+VBCRLF
		IX=IX+PATH_FILE_VAR    +VBCRLF+VBCRLF
		IX=IX+ERR.DESCRIPTION  +VBCRLF+VBCRLF
		IX_ERROR_MESSENGER=IX
		STOP_PROGRESS=TRUE
		EXIT SUB
	END IF
	
END SUB

Sub BuildFullPath(ByVal FullPath)
	If Not fso.FolderExists(FullPath) Then
		BuildFullPath fso.GetParentFolderName(FullPath)
		fso.CreateFolder FullPath
	End If
End Sub

Function GetDriveLetterFromVolumeName( volumeName )
Dim volumes, volume

    ' Unless we found a matching volume, an empty string will be the returned value
    GetDriveLetterFromVolumeName=""

    ' Ask WMI for the list of volumes with the requested label
    Set volumes = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") _ 
                  .ExecQuery("SELECT DriveLetter FROM Win32_Volume WHERE Label='" & volumeName & "'")

    ' If exist an matching volume, get its drive letter
	If volumes.Count > 0 Then 
        For Each volume In volumes 
            GetDriveLetterFromVolumeName = volume.DriveLetter
            Exit For
        Next 
    End If

End Function

SUB PATH_SCANNER_RENAME_4

	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)

	NewFilename=FILE_VAR
	NewFilename=REPLACE(NewFilename,"__","_")
	NewFilename=REPLACE(NewFilename," _ ","_")
	NewFilename=REPLACE(NewFilename,"_ ","_")
	NewFilename=REPLACE(NewFilename," _","_")
	NewFilename=REPLACE(NewFilename,".jpg",".JPG")
	Extension=ucase(mid(NewFilename,len(NewFilename)-3))
	
	NewFilename=PATH_VAR+NewFilename
	
	SET_GO=FALSE
	if Extension=".JPG" then 
		SET_GO=TRUE
	end	if
	
	if VAR_PATH_SCANNER <> NewFilename AND SET_GO=TRUE then 
		'msgbox NewFilename
		I_COUNT_1=I_COUNT_1+1
		ON ERROR RESUME NEXT
		ERR.CLEAR
		Fso.MoveFile VAR_PATH_SCANNER, NewFilename
		IF ERR.NUMBER>0 THEN 
			MSGBOX ERR.DESCRIPTION & VBCRLF & VAR_PATH_SCANNER & VBCRLF & NewFilename
		END IF
		'Wscript.QUIT
	end if

END SUB


SUB PATH_SCANNER_RENAME_3

	NewFilename=REPLACE(VAR_PATH_SCANNER,".JPG",".jpg")
	
	if VAR_PATH_SCANNER <> NewFilename then 
		I_COUNT_1=I_COUNT_1+1
		Fso.MoveFile VAR_PATH_SCANNER, NewFilename
		'msgbox NewFilename
		'Wscript.QUIT
	end if

END SUB

SUB PATH_SCANNER_RENAME_2

	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)

	SET_GO=TRUE 
	IF INSTR(UCASE(VAR_PATH_SCANNER),".JPG")=0 THEN SET_GO=FALSE
	'IF MID(FILE_VAR,1,4)<>"2010" THEN  SET_GO=FALSE

	NewFilename=FILE_VAR

	IF LEN(FILE_VAR)>= 12 AND SET_GO=TRUE THEN SET_GO=FALSE
	
	ON ERROR RESUME NEXT
	if SET_GO=TRUE then
		NewFilename_COUNTER=CLNG(MID(NewFilename,4,INSTR(NewFilename,".")-4))
		IF ERR.NUMBER >0 THEN 
			msgbox "ERROR CLNG WITH "+VBCRLF+VBCRLF+NewFilename
		END IF
	END IF

	'IF NewFilename_COUNTER>= 761 THEN SET_GO=FALSE
	
	IF SET_GO=TRUE THEN
		'MSGBOX VAR_PATH_SCANNER
		NewFilename=FILE_VAR

		if FSO.FileExists(PATH_VAR+NewFilename)=TRUE THEN 
			'msgbox "FILE ALREADY EXISTS "+VBCRLF+VBCRLF+NewFilename
			'Wscript.QUIT
			X_D=0
			DO
				FOR R=1 TO 500
					X_D=X_D+1
					NewFilename="DSC"+CSTR(X_D)+".jpg"
					NewFilename_VAR_2=MID(NewFilename,4)
					 
					IF LEN(NewFilename)<12 THEN 
						X_C=1
						DO
							IF LEN(NewFilename)<12 THEN 
								NewFilename="DSC"+STRING(X_C,"0")+NewFilename_VAR_2
								X_C=X_C+1
							END IF
						LOOP UNTIL LEN(NewFilename)=12
					end if
					if FSO.FileExists(PATH_VAR+NewFilename)=FALSE THEN EXIT FOR
				NEXT
			LOOP UNTIL FSO.FileExists(PATH_VAR+NewFilename)=FALSE

		END IF

		NewFilename=PATH_VAR+NewFilename

		if VAR_PATH_SCANNER <> NewFilename then 
			msgbox NewFilename
			Wscript.QUIT
			Fso.MoveFile VAR_PATH_SCANNER, NewFilename

			I_COUNT_1=I_COUNT_1+1
		END IF
	END IF
	
	
END SUB

SUB PATH_SCANNER_RENAME

	SET_GO=FALSE
	IF INSTR(UCASE(VAR_PATH_SCANNER),"- SONY DSC-H7 -") THEN SET_GO=TRUE
	IF INSTR(UCASE(VAR_PATH_SCANNER),".JPG")=0 THEN SET_GO=FALSE
	
	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)
	
	'IF MID(UCASE(FILE_VAR),1,4)="2010" THEN SET_GO=TRUE
	
	SET_GO=TRUE 
	IF INSTR(UCASE(VAR_PATH_SCANNER),".JPG")=0 THEN SET_GO=FALSE
	
	'IF MID(FILE_VAR,1,4)="DSCN" THEN  SET_GO=FALSE
	'IF MID(FILE_VAR,1,4)<>"2010" THEN  SET_GO=FALSE

	IF LEN(FILE_VAR)>= 12 AND SET_GO=TRUE THEN SET_GO=FALSE
	
	NewFilename=FILE_VAR
	NewFilename="DSC"+MID(NewFilename,12)
	
	ON ERROR RESUME NEXT
	if SET_GO=TRUE then
		NewFilename_COUNTER=CLNG(MID(NewFilename,4,INSTR(NewFilename,".")-4))
		IF ERR.NUMBER >0 THEN 
			msgbox "ERROR CLNG WITH "+VBCRLF+VBCRLF+NewFilename
		END IF
	END IF

	'IF NewFilename_COUNTER>= 761 THEN SET_GO=FALSE
	
	IF SET_GO=TRUE THEN 
		
		NewFilename_COUNTER_STRING=MID(NewFilename,INSTR(NewFilename,".")-4)
		NewFilename_COUNTER_STRING=CLNG(MID(NewFilename,4,INSTR(NewFilename,".")-4))
		NewFilename_COUNTER_STRING=CLNG(MID(NewFilename,4,INSTR(NewFilename,".")-4))
		NewFilename
		IF LEN(NewFilename)<12 THEN 
			X_C=1
			DO
				IF LEN(NewFilename)<12 THEN 
					NewFilename="DSC"+STRING(X_C,"0")+MID(NewFilename,4)
					X_C=X_C+1
				END IF
			LOOP UNTIL LEN(NewFilename)=12
		end if


		
		' IF LEN(FILE_VAR)<12 THEN 
			' X_C=1
			' DO
				' IF LEN(FILE_VAR)<12 THEN 
					' NewFilename="DSC"+STRING(X_C,"0")+MID(FILE_VAR,4)
					' X_C=X_C+1
				' END IF
			' LOOP UNTIL LEN(NewFilename)=12
			' NewFilename=PATH_VAR+NewFilename
		' end if
			
		'msgbox NewFilename
		'Wscript.QUIT
		
		if FSO.FileExists(PATH_VAR+NewFilename)=TRUE THEN 
			'msgbox "FILE ALREADY EXISTS "+VBCRLF+VBCRLF+NewFilename
			'Wscript.QUIT
			FOR R=1 TO 400
				NewFilename="DSC"+CSTR(R)+".jpg"
				IF LEN(NewFilename)<12 THEN 
					X_C=1
					DO
						IF LEN(NewFilename)<12 THEN 
							NewFilename="DSC"+STRING(X_C,"0")+MID(NewFilename,4)
							X_C=X_C+1
						END IF
					LOOP UNTIL LEN(NewFilename)=12
				end if
				 
				msgbox NewFilename
				 
				 if FSO.FileExists(PATH_VAR+NewFilename)=FALSE THEN EXIT FOR
			 NEXT
		END IF
		
		NewFilename=PATH_VAR+NewFilename
		
		SET_GO=FALSE
		if instr(ucase(NewFilename),".JPG")>0 THEN SET_GO=TRUE
		PATH_VAR_4=MID(NewFilename,1,INSTRREV(NewFilename,"\"))
		IF INSTR(VAR_PATH_SCANNER,PATH_VAR_4)=0 THEN SET_GO=FALSE
		'MSGBOX PATH_VAR_4
		'MSGBOX VAR_PATH_SCANNER
		IF SET_GO=TRUE THEN 
			msgbox NewFilename
			'Wscript.QUIT
			Fso.MoveFile VAR_PATH_SCANNER, NewFilename
			I_COUNT_1=I_COUNT_1+1
		END IF
	END IF

END SUB





'------------------------------------------------
Sub Search(Folder)
'------------------------------------------------
	ON ERROR RESUME NEXT
	
	ERR.CLEAR
	for each file in Folder.Files
		IF ERR.NUMBER=0 THEN 
			VAR_PATH_SCANNER=file.PATH
			CALL PATH_SCANNER_RENAME
		END IF
	NEXT

	'Find EACH SUBFOLDER.
	For Each subFolder In Folder.SubFolders
		'Call the Search subroutine to start the recursive search on EACH SUBFOLDER.
		Search objFSO.GetFolder(subFolder.Path)
	Next

End Sub
		
		


'---------
' REFERANCE
'---------
' How Can I Get a List of All the Files in a Folder and Its Subfolders? – Hey, Scripting Guy! Blog
' https://blogs.technet.microsoft.com/heyscriptingguy/2004/10/20/how-can-i-get-a-list-of-all-the-files-in-a-folder-and-its-subfolders/
' ----

'------------------------------------------------------------------------
' THIS PROJECT USING THIS CODE
' BECAUSE A MOVE FILE HAPPEN MUST GATHER WHOLE FILE LIST BEFORE ACT ON IT
' AND IT WAS DEBUGED IN THE NEXT URL LINK LINE
'------------------------------------------------------------------------
'----
' windows - VBS to Search for Multiple Files Recursively in C:\Users - Stack Overflow
' https://stackoverflow.com/questions/25658004/vbs-to-search-for-multiple-files-recursively-in-c-users
'------------------------------------------------------------------------
'------------------------------------------------------------------------
' vbscript - Read files in subfolders - Stack Overflow
' https://stackoverflow.com/questions/16905278/read-files-in-subfolders
'----
'------------------------------------------------------------------------

' ----
' vbscript - Read files in subfolders - Stack Overflow
' https://stackoverflow.com/questions/16905278/read-files-in-subfolders
' ----