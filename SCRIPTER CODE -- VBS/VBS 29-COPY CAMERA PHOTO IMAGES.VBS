'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 29-COPY CAMERA PHOTO IMAGES 01.VBS
'# __ 
'# __ VBS 29-COPY CAMERA PHOTO IMAGES 01.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START     TIME [ Thu 14-Jun-2018 14:22:00 ]
'# END FINAL TIME [ Thu 14-Jun-2018 17:00:00 ] 2 hours, 38 minutes and 0 seconds
'# __ 
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
' -------------------------------------------------------------------
' OFF LOAD MY CAMERA REEL
' FROM TWO CAMERA DETECTED BY VOLUME NAME OF MEDIA CARD
' WORKS FOR FEW MEDIA CARD AT MOMENT AND EXPANDING
' RUN AND DISPLAY RESULT FOR EACH AT A TIME IN A FOR NEXT LOOP
'
' THIS THE MEDIA CARD DETECTOR & LAUNCHER OF THIS PROGRAM
' \SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 35-COPY CAMERA PHOTO IMAGES.AHK
'
' SYNCHRONIZER PROGRAMS NEVER HAVE WHAT I WANT
' GOODSYNC HAD A STEP IN RIGHT DIRECTION
' BUT ONLY WITH MOVE MODE
' WANTED MOVE LIKE COPY
' MOVE FILES OFF CAMERA REEL IS NOT WHAT I WANT
' I RATHER WAIT UNTIL MEDIA CARD IS FULL BEFORE DELETE AND START AGAIN
' IT TAKES KEEPING A RECORD TIME & DATE OF EACH FILE UPLOADED
' WITH SYNC PROGRAM I GOT TO SET THE TIME AND DATE OFTEN 
' AND ONLY WITHIN ONE WHOLE DAY OR TERRIBLE DIFFICULT
' -------------------------------------------------------------------
' IT IS BASED ON A COUPLE OF FOLDER FROM WHOLE MEDIA CARD
' IT WOULD BE HARDER TO DO MY MOBILE ONE AS MANY ROOT FOLDER
' SO EXTRA WORK I GOT TO DO SETTING THE BASE FOLDER NOT WHOLE CARD 
' SHOULD BE EASIER WITH ANOTHER ARRAY FOR MOBILE USE SOON
' -------------------------------------------------------------------
' CAREFUL ON FIRST TIME RUNS IF NEW TO THIS AND BEEN USE SYNCHRONIZER BEFORE
' AS WILL UPLOAD WHOLE CARD AGAIN
' WHY I LEFT A TEST DATE TO DEBUG PUT IN
' -------------------------------------------------------------------
' THE DATE MARKER FOR CAMERA FILE IS STORED IN A FILE PER MEDIA-CARD VOLUME-NAME 
' AND IS THE 
' ELDEST DATE FROM THE BATCH
' IT MIGHT BE BETTER TO DO IT FOR EACH FILE IN-CASE OF ERROR BUT IF GOT VIDEO FOLDER
' AND PHOTO IT THEN MAKES THE DATE OUT OF SYNC
' UNLESS YOU HAD A SORTING CODE AND IN VBS QUITE DIFFICULT
' NEEDS SOME THOUGHT ON THAT DISGUSTING ONE
' -------------------------------------------------------------------

'#-------------------------------------------------------------------
' 002 __ MORE WORK ADD __ FUJI CAMERA ON
' -------------------------------------------------------------------
' Sat 16-Jun-2018 10:35:23
' Sat 16-Jun-2018 11:31:44 __ 1 HOUR
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' 003 __ MORE WORK ADD __  PROCESS ONLY FILES SO EMPTY FOLDER DO GET ON
'                          DON'T UPLOAD THUMBS FOR VIDEO
'                          DETECT OF ANY JPG'S COPIED SO DON'T AUTO RUN RENAMER
' -------------------------------------------------------------------
' Mon 02-Jul-2018 09:18:00
' Mon 02-Jul-2018 10:23:00 __ 1 HOUR+ WITH BREAK TIME
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' 004 __ THE AUTOHOTKEYS PROGRAM SCRIPTOR THAT WORKS WITH THIS CODE
'        HAS NOW BEEN EXTRACTED FROM THE OTHER AND NOW RUNS AS AN INDIVIDUAL PROGRAM '        TO PARTNER WITH HERE CODE
' IS HERE
' C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 35-COPY CAMERA PHOTO IMAGES.AHK
' THIS PROGRAM SCRIPTOR VBS ALSO RUNS SOME OTHER VBA VB6 APP
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' 005 __ ADD NEW 
' -------------------------------------------------------------------
' 1.. CAMERA _ GO PRO _ & FOUND OUT MY GOPRO WAS ONLY A REPLICA CLONE & 
' NOT REAL GOPRO AND APP ON SMARTPHONE OUT-DATER NOT ANY LONGER DOES WIFI OVER
' 2.. NEW YELLOW CAMERA DOUBLE SCREEN ONE
'
' 3.. ADD MORE FEATURE LIKE GO WITH THE YEAR NUMBER AND CREATE FOLDER
' IF OUT OF TIMING A LITTLE BIT OVER NEW YEAR EASY MOVE FILE OVER
'
' 4.. ADD COPY EVERYTHING FROM A CERTAIN DRIVE LIKE GOPRO DRIVE 
' ONLY UNLESS CERTAIN OTHER EXCLUDED LIKE _GSDATA_
'
' 5.. ADDED MOVE MODE FOR DODGY PICTURE _ RISK IS FILE ERROR COPY AND DELETE AFTERWARD HOPE NOT SO
'
' 6.. ADD OTHER OPTION DEPEND ON COMPUTER NAME
'     DON'T SHOW EXPLORER ON DODGY COMPUTER AROUND BUILDING AND IF DODGY PATH IMAGE-ER
'     OR DODGY LESS QUICKER SPEC LOW END PC
' -------------------------------------------------------------------
' Mon 01-Oct-2018 22:07:00
' Mon 01-Oct-2018 23:51:09 __ TOOK A BREAK FOR A WHILE TEST OTHER THING NOT CODER HERE
' Mon 02-Oct-2018 00:40:00 __ 4..
' Mon 02-Oct-2018 01:38:00 __ 5..
' Mon 02-Oct-2018 01:50:00 __ 6..
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' 006 __ ADD NEW 
' -------------------------------------------------------------------
' 1.. CAMERA _ SAMSUNG MINI S5
' WHICH IS JUST A FOLDER ALREADY BEEN SYNC OVER AND PUT IN PHOTO AREA
' 
' 2.. ADD FOLDER TO SCAN IS ABLE TO NOW BE AT DEPTH _ QUITE A BIT OF WORKING
' HOPE NONE ERROR WITH OTHER FOLDER DRIVE WORKER KNOCK ON EFFECT
'
' 3.. ADD DIALOG AT END NOT DISPLAY IF COUNTER NOT UP ON NAUGHT AND D DRIVE AS SHOWING ALWAYS
' NORMAL THE TIMING OF FOLDER LOAD IS SOON AS BEGIN FOR A PRESENT EXIST DRIVE OR SAVE FOR LATER 
' WHEN COUNTER SHOW SOME DONE
' -------------------------------------------------------------------
' Tue 02-Oct-2018 06:04:58 __ 
' Mon 02-Oct-2018 07:35:00 __ WORK TOOK 80 CLIPBOARD EVENTS
' -------------------------------------------------------------------


'#-------------------------------------------------------------------
' 007 __
' -------------------------------------------------------------------
' 1.. ADD NEW CAMERA DASH_CAMERA
' 
' 2.. MAKE ALL VARIABLES THAT HARD_CODED INTO ARRAY FRONT END VARIABLE
'
' -------------------------------------------------------------------
' SESSION 01 ----------------
' Sun 25-Nov-2018 02:27:49 __ 
' Sun 25-Nov-2018 05:18:05 __ 
' BED TIME ------------------
' SESSION 02 ----------------
' Sun 25-Nov-2018 10:51:38 __ 
' Sun 25-Nov-2018 11:34:00 __ 
' -------------------------------------------------------------------

' CHECK IF ANY WSCRIPT ALREADY RUNNER AND QUIT IF IS WORKER
' -------------------------------------------------------------------
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")

i1 = 0 ' ANY PROGRAM WSCRIPT 
i2 = 0 ' MY  PROGRAM WSCRIPT 
FOR EACH objProcess IN colProcesses
    IF NOT (IsNull(objProcess.CommandLine )) THEN
		strScriptName = Trim(Right(objProcess.CommandLine,Len(objProcess.CommandLine) - InstrRev(objProcess.CommandLine,"\")))
		strScriptName = Left(strScriptName, Len(strScriptName) - 1)                             
		IF strScriptName = Wscript.ScriptName THEN
			i2=i2+1
		else
			i1=i1+1
		END IF
		IF i1 > 0 THEN 
				' Wscript.Quit
		END IF  
		IF i2 > 1 THEN 
				Wscript.Quit
		END IF  
    END IF  
NEXT

DIM EXPLORER_COUNTING

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'explorer.exe'")
' KNOCK UP THIS FOR USE IN MY OUTLOOK SCRIPT CODING
For Each objProcess In colProcesses
    If Not (IsNull(objProcess.ProcessID)) Then
        strScriptProcessID = objProcess.ProcessID
		EXPLORER_COUNTING=EXPLORER_COUNTING+1
        ' MsgBox "MY PROCESS NUMBER PID IS -- " + CStr(strScriptProcessID)
    End If
Next
                             
' Wscript.QUIT


Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

DIM CUR_USER
DIM ComputerName

CUR_USER = CreateObject("WScript.Network").UserName
Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

DIM JPG_BEEN_DONE
DIM MP4_BEEN_DONE
DIM AVI_BEEN_DONE

DIM STOP_PROGRESS
DIM IX_ERROR_MESSENGER

DIM DRIVE_SPACE_INFO

DIM objStartFolder
DIM objFolder
DIM I_COUNT_1
DIM I_COUNT_2
DIM VAR_PATH_SCANNER
DIM IMAGE_POINTER_NAME

DIM objSuperFolder

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO_2 = CreateObject("Scripting.FileSystemObject")

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")


Dim arr1
Dim oDict ' Key: Full filename, Item: Filename
Dim oDict_4

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

DIM R_COUNT

DIM FOLDER_PATH_INDEX

NUMBER_OF_ITEM = 100

DIM FOLDER_PATH()
DIM FOLDER_PATH_DEST()
DIM FOLDER_PATH_VAR()
DIM FOLDER_PATH_DO_ALL_FILE()
DIM FOLDER_PATH_MOVE_MODE_FILE()
DIM FOLDER_PATH_IGNOR_DATE_INDEX()
DIM FOLDER_PATH_DONT_SHOW_EXPLORER_UNLESS_COPY()


REDIM FOLDER_PATH(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_DEST(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_VAR(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_DO_ALL_FILE(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_MOVE_MODE_FILE(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_IGNOR_DATE_INDEX(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_DONT_SHOW_EXPLORER_UNLESS_COPY(NUMBER_OF_ITEM)

DIM FOLDER_PATH_AVOID_ITEM()
REDIM PRESERVE FOLDER_PATH_AVOID_ITEM(100)
NUMBER_COUNTER = 0
I=NUMBER_COUNTER
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="_gsdata_"
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="System Volume Information"
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="\F.A"                      ' DASH CAMERA
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=".THM"                      ' SONY
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=".AmazonDriveConfig"
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="CAPTURE.TXT"               ' 4K ULTRA ACTION CAM
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=""
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=""
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=""
' I=I+1:FOLDER_PATH_AVOID_ITEM(I)="AVF_INFO"
FOR R_COUNT_2=1 TO UBOUND(FOLDER_PATH_AVOID_ITEM)
	IF FOLDER_PATH_AVOID_ITEM(R_COUNT_2)<>"" THEN NUMBER_COUNTER=R_COUNT_2
NEXT
REDIM PRESERVE FOLDER_PATH_AVOID_ITEM(NUMBER_COUNTER)

DIM FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE()
REDIM PRESERVE FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(100)
NUMBER_COUNTER = 0
NUMBER_COUNTER = NUMBER_COUNTER + 1
I=NUMBER_COUNTER
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".JPG"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".AVI"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".MOV"                      ' DASH CAMERA
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=""
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=""
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=""
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=""
FOR R_COUNT_2=1 TO UBOUND(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE)
	IF FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_2)<>"" THEN NUMBER_COUNTER=R_COUNT_2
NEXT
REDIM PRESERVE FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(NUMBER_COUNTER)
	
DIM FOLDER_PATH_DEST_VAR
DIM FOLDER_PATH_VAR_3

DIM EXPLORER_LOADED
DIM SHOW_EXPLORER_FLAG

DIM FOLDER_PATH_MAIN_4

FOLDER_PATH(1)="MC - HX60V"
FOLDER_PATH(2)="MC - HX60V0"
FOLDER_PATH(3)="FUJI XP90"
FOLDER_PATH(4)="GOPRO"
FOLDER_PATH(5)="GOPRO"
FOLDER_PATH(6)="DUAL SCRN" ' __ COPY FIRST TO END PLACE
FOLDER_PATH(7)="DUAL SCRN" ' __ ON SECOND RUN MOVE NEXT TO ANOTHER
FOLDER_PATH(8)="D DRIVE"
FOLDER_PATH(9)="DASH_CAMERA" ' __ COPY FIRST TO END PLACE
FOLDER_PATH(10)="DASH_CAMERA" ' __ ON SECOND RUN MOVE NEXT TO ANOTHER
FOLDER_PATH(11)="4K-ULTRA-HD" ' __ COPY FIRST TO END PLACE
FOLDER_PATH(12)="4K-ULTRA-HD" ' __ ON SECOND RUN MOVE NEXT TO ANOTHER

' NOKIA TO DO 1 2 3

FOLDER_PATH_DONT_SHOW_EXPLORER_UNLESS_COPY(8)=TRUE

FOLDER_PATH_VAR(1)="\"
FOLDER_PATH_VAR(2)="\"
FOLDER_PATH_VAR(3)="\"
FOLDER_PATH_VAR(4)="\"
FOLDER_PATH_VAR(5)="\"
FOLDER_PATH_VAR(6)="\"
FOLDER_PATH_VAR(7)="\"
FOLDER_PATH_VAR(8)="\Android\SDCard2_Sync\DCIM\"
FOLDER_PATH_VAR(9)="\"
FOLDER_PATH_VAR(10)="\"
FOLDER_PATH_VAR(11)="\"
FOLDER_PATH_VAR(12)="\"

FOLDER_PATH_DEST(1)="D:\DSC\2015+Sony\" + CSTR(Year(Now)) + " CyberShot HX60V"
FOLDER_PATH_DEST(2)="D:\VI_ DSC ME\2010+SONY\" + CSTR(Year(Now)) + " CyberShot HX60V_#"
FOLDER_PATH_DEST(3)="D:\DSC\2017+FUJI XP90\" + CSTR(Year(Now)) + " FinePix XP90"
FOLDER_PATH_DEST(4)="D:\DSC\" + CSTR(Year(Now)) + " GoPro 01"
FOLDER_PATH_DEST(5)="D:\VI_ DSC ME\" + CSTR(Year(Now)) + " GoPro 01"
FOLDER_PATH_DEST(6)="D:\DSC\" + CSTR(Year(Now)) + " Double Screen Cam"
FOLDER_PATH_DEST(7)="D:\VI_ DSC ME\" + CSTR(Year(Now)) + " Double Screen Cam"
FOLDER_PATH_DEST(8)="D:\DSC\" + CSTR(Year(Now)) + " Samsung Mini S5\DCIM"
FOLDER_PATH_DEST(9)="D:\DSC\" + CSTR(Year(Now)) + " DASH CAMERA"
FOLDER_PATH_DEST(10)="D:\VI_ DSC ME\" + CSTR(Year(Now)) + " DASH CAMERA"
FOLDER_PATH_DEST(11)="D:\DSC\2019+4K ULTRA HD\" + CSTR(Year(Now)) + " 4K ULTRA"
FOLDER_PATH_DEST(12)="D:\VI_ DSC ME\2019+4K ULTRA HD\" + CSTR(Year(Now)) + " 4K ULTRA"

'THINKER WE WANT MOVE MODE FOR FUJI

' THIS IS THE GOPRO DRIVE AND WILL WANT ALL COPY WITH MOVE AND IGNOR_DATE_INDEX
FOLDER_PATH_MOVE_MODE_FILE(2)=TRUE
FOLDER_PATH_DO_ALL_FILE(4)=TRUE
FOLDER_PATH_DO_ALL_FILE(5)=TRUE
FOLDER_PATH_MOVE_MODE_FILE(5)=TRUE
'FOLDER_PATH_IGNOR_DATE_INDEX(4)=TRUE
'FOLDER_PATH_IGNOR_DATE_INDEX(5)=TRUE
' THIS IS THE DUAL SCREEN DRIVE AND WILL WANT ALL COPY BUT MOVE_MODE VIDEO & IMAGE
' THIS IS THE DUAL SCREEN DRIVE & DO HERE IS 2 STEP 1ST COPY TO ONE FOLDER 2ND MOVE ANOTHER FOLDER
' AND IGNORE DATE INDEX REQUIRED WILL BE OKAY AS MOVE MODE IS AFTER THAT ALSO IGNORE DATE INDEX
FOLDER_PATH_DO_ALL_FILE(6)=TRUE
FOLDER_PATH_DO_ALL_FILE(7)=TRUE
FOLDER_PATH_MOVE_MODE_FILE(7)=TRUE
' SAMSUNG MINI S 5 _ ALLOW OVER WRITER & ALLOW ALL FILE EXTENSION TYPE TYPEO TYPO P & E & O
FOLDER_PATH_DO_ALL_FILE(8)=TRUE
' DASH CAMERA
FOLDER_PATH_DO_ALL_FILE(9)=TRUE
FOLDER_PATH_DO_ALL_FILE(10)=TRUE
FOLDER_PATH_MOVE_MODE_FILE(10)=TRUE

' DO ALL THE FILE FROM SD CARD FIRST
' AND THEN DO AGAIN FROM SD CARD AND ERASE THEM AFTER
FOLDER_PATH_DO_ALL_FILE(11)=TRUE
FOLDER_PATH_DO_ALL_FILE(12)=TRUE
FOLDER_PATH_MOVE_MODE_FILE(12)=TRUE



' FOR R_COUNT=1 TO UBOUND(FOLDER_PATH_DEST)
	' If Not FSO.FolderExists(FOLDER_PATH_DEST(R_COUNT)) Then
		' BuildFullPath FOLDER_PATH_DEST(R_COUNT)
	' End If
' NEXT

FOR R_COUNT_2=1 TO UBOUND(FOLDER_PATH_DEST)
	IF FOLDER_PATH_DEST(R_COUNT_2)<>"" THEN NUMBER_COUNTER=R_COUNT_2
NEXT

REDIM PRESERVE FOLDER_PATH(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH_DEST(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH_VAR(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH_DO_ALL_FILE(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH_MOVE_MODE_FILE(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH_IGNOR_DATE_INDEX(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH_DONT_SHOW_EXPLORER_UNLESS_COPY(NUMBER_COUNTER)

DIM CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR
CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR=FALSE

DIM FOLDER_PATH_R_COUNT

FOR R_COUNT=1 TO UBOUND(FOLDER_PATH)

	I_COUNT_1=0
	I_COUNT_2=0

	FOLDER_PATH_INDEX=CSTR(R_COUNT)+" "+FOLDER_PATH(R_COUNT)
	FOLDER_PATH_VAR_2=getDriveLetterFromVolumeName(FOLDER_PATH(R_COUNT))
	FOLDER_PATH_DEST_VAR=FOLDER_PATH_DEST(R_COUNT)
	
	FOLDER_PATH_VAR_3=FOLDER_PATH_VAR(R_COUNT)
	FOLDER_PATH_R_COUNT=FOLDER_PATH(R_COUNT)
	
	IF FOLDER_PATH(1)="MC - HX60V" THEN 
		IF FOLDER_PATH(1)=FOLDER_PATH(R_COUNT) THEN 
			IF FOLDER_PATH_VAR_2<>"" THEN 
				CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR=TRUE
			END IF
		END IF
	END IF

	' ---------------------------------------------------------------
	' ONLY THESE CAMERA REQUIRE HAD DATE PUT ON FRONT FILENAME
	' ONLY CARRY EXIF
	' OTHER FOLDER ONLY DO FROM FILE _ DATE _ MODIFIED
	' ---------------------------------------------------------------
	IF INSTR("MC - HX60V -- FUJI XP90",FOLDER_PATH(R_COUNT)) THEN 
		IF FOLDER_PATH_VAR_2<>"" THEN 
			DO_SEND_TO_PUT_DATE_INFRONT_FILENAME=TRUE
		END IF
	END IF
	
	IF INSTR("MC - HX60V",FOLDER_PATH(R_COUNT)) THEN 
		IF FOLDER_PATH_VAR_2<>"" THEN 
			ONLY_SONY_CYBERSHOT_HAS_DO=TRUE
		END IF
	END IF
	
	
	

	

	
	
	
	IF FOLDER_PATH_VAR_2<>"" THEN 
		
		WORK_GO_ER=TRUE
		
		FOLDER_PATH_VAR_2=FOLDER_PATH_VAR_2+FOLDER_PATH_VAR(R_COUNT)

		' MSGBOX FOLDER_PATH_VAR_2
		' Wscript.QUIT
		
		FOLDER_PATH_DEST_VAR_2=FOLDER_PATH_DEST(R_COUNT)
		
		'------------------------------------------------------------
		' DON'T SHOW EXPLORER ON DODGY COMPUTER AROUND BUILDING AND IF DODGY PATH IMAGE-ER
		' OR DODGY LESS QUICKER SPEC LOW END PC
		'------------------------------------------------------------
		' DON'T SHOW EXPLORER ON D DRIVE IF ONLY PICTURE IS NONE
		'------------------------------------------------------------
		SHOW_EXPLORER_FLAG=TRUE
		IF INSTR(FOLDER_PATH_DEST_VAR,"D:\VI_ DSC ME\") > 0 THEN 
			IF ComputerName="1-ASUS-X5DIJ" THEN SHOW_EXPLORER_FLAG=FALSE
			IF ComputerName="2-ASUS-EEE" THEN SHOW_EXPLORER_FLAG=FALSE
			IF ComputerName="3-LINDA-PC" THEN SHOW_EXPLORER_FLAG=FALSE
			IF ComputerName="4-ASUS-GL522VW" THEN SHOW_EXPLORER_FLAG=FALSE
		END IF
		
		IF FOLDER_PATH_DONT_SHOW_EXPLORER_UNLESS_COPY(R_COUNT)=TRUE THEN
			SHOW_EXPLORER_FLAG=FALSE
		END IF
		
		IF EXPLORER_COUNTING=>2 THEN 
			SHOW_EXPLORER_FLAG=FALSE
		END IF
		
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


' -------------------------------------------------------------------
' THIS PROCESS NOT TOO GOOD AS THESE SCRIPT IN BAT WAKE ALL THE DRIVE WHEN SCAN THEM FOR FOLDER
' AS PART OF SYNC DO
' INCORP IN TO CAMERA LOADER
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' THIS SYNC NOT GOOD AS NOT CHECKER FOR CONFLICTS
' UNTIL DONE
' -------------------------------------------------------------------

' DIM RUN_COPY(1)

' RUN_COPY(1)="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 48-XCOPY SYNC.BAT"

' FOR R_COUNT=1 TO UBOUND(RUN_COPY)

	' IF FSO.FileExists(RUN_COPY(R_COUNT)) THEN 
		' WSHShell.Run """"+RUN_COPY(R_COUNT)+"""", DontShowWindow, DontWaitUntilFinished
	' END IF
	
' NEXT



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
	
	' ---------------------------------------------------------------
	' MP4_BEEN_DONE=TRUE
	' WOULD BE GOOD BUT NOT REQUIRE IF ONLY MP4
	' IT DOESN'T REQUIRE CONVERSION CODING STRICTLY
	' ---------------------------------------------------------------
	' Wed 20-Mar-2019 04:43:54
	' ---------------------------------------------------------------
	
	' IF THEN MP4_BEEN_DONE=TRUE
		' JPG_BEEN_DONE=TRUE
	' END IF
	' IF THEN AVI_BEEN_DONE=TRUE
		' JPG_BEEN_DONE=TRUE
	' END IF
	
	IF DO_SEND_TO_PUT_DATE_INFRONT_FILENAME=TRUE THEN 
		IF I_COUNT_1>=0 AND (JPG_BEEN_DONE=TRUE) THEN 
			WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 33-#0 Send To Date Of File In Front FileName.ahk""", ShowWindow, DontWaitUntilFinished
		END IF
	END IF

	IF ONLY_SONY_CYBERSHOT_HAS_DO=TRUE THEN
		IF I_COUNT_1>=0 AND (MP4_BEEN_DONE=TRUE) THEN 
			WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 54-CAMERA MOVE COPY FROM MP_ROOT TO DROP ONE PATH & IMAGER.VBS""", ShowWindow, DontWaitUntilFinished
		END IF
	END IF

	' ---------------------------------------------------------------	
	' ONLY THE SONY CYBERSHOT RUN HERE TO MAKE DUPE MC-CARD
	' ---------------------------------------------------------------	
	IF CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR=TRUE THEN 
		IF I_COUNT_1>=0 AND (JPG_BEEN_DONE=TRUE OR MP4_BEEN_DONE=TRUE) THEN 
			' ---------------------------------------
			FILE_NAME="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT"
			PARAM="CAMERA_BACKUP_MODE"
			' ---------------------------------------
			WSHShell.Run """"+FILE_NAME+"""" + " " + ""+PARAM+"" , ShowWindow, DontWaitUntilFinished
		END IF
	END IF


	' ---------------------------------------------------------------	
	' ---------------------------------------------------------------	
	' NOW LEFT TO FIND CODE
	' FOR CAMERA 
	' THAT DON'T USE USUAL RENAME METHOD FOR PUT DATE IN FRONT
	' MOST THE ONE THAT DON'T USE EXIF
	' SO TAKE FROM DATE FILE _ DATE _ MODIFIED
	' FOR VIDEO AND JPG
	' AND ALSO THE SONY CYBERSHOT VIDEO MP4 COULD DO WITH THAT ON
	' WHICH I HAVE CODE ABOUT FOR
	' BUT GUESS EASIER IF NOT BACKWARD COMPATIBLE
	' ---------------------------------------------------------------	
	' ---------------------------------------------------------------	
	' SOMETHING BEGIN LIKE HERE

	IF INSTR("4K-ULTRA-HD",FOLDER_PATH_R_COUNT) THEN 
		' DO SOME WORK TO RENAME FILE

		FOLDER_PATH_MAIN_4=FOLDER_PATH_DEST_VAR_2

		CALL MAIN_4
		
	END IF
	
	' ---------------------------------------------------------------	
	' ---------------------------------------------------------------	
	' ---------------------------------------------------------------	
	' ---------------------------------------------------------------	

    Set oDict = Nothing
End Sub

' -------------------------------------------------------------------




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
			S=S+"Drive Letter 1st Found for Match Volume Name: " & VBCRLF & drvPath & "\" & vbCrLf
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
		' MSGBOX VAR_PATH_SCANNER
		IF FSO.FILEExists(VAR_PATH_SCANNER)=TRUE THEN 
			CALL PATH_SCANNER_COPY_CAMERA
		END IF
	Next
End Sub

SUB PATH_SCANNER_COPY_CAMERA
	SET_GO_FILE=FALSE
	IF mid(ucase(VAR_PATH_SCANNER),3,5)="\DCIM"    then SET_GO_FILE=TRUE
	IF mid(ucase(VAR_PATH_SCANNER),3,8)="\MP_ROOT" then SET_GO_FILE=TRUE
	IF mid(ucase(VAR_PATH_SCANNER),3,6)="\DCIMA"   then SET_GO_FILE=TRUE
	
	DO_ALL_FILE=FALSE
	MOVE_MODE_FILE=FALSE
	IGNOR_DATE_INDEX=FALSE

	IF FOLDER_PATH_DO_ALL_FILE(R_COUNT)=TRUE THEN
		DO_ALL_FILE=TRUE
	END IF
	IF FOLDER_PATH_MOVE_MODE_FILE(R_COUNT)=TRUE THEN
		MOVE_MODE_FILE=TRUE
	END IF
	IF FOLDER_PATH_IGNOR_DATE_INDEX(R_COUNT)=TRUE THEN
		IGNOR_DATE_INDEX=TRUE
	END IF

	
	FOR R_COUNT_2=1 TO UBOUND(FOLDER_PATH_AVOID_ITEM)
		IF NOT MID(FOLDER_PATH_AVOID_ITEM(R_COUNT_2),1,1)="." THEN
			IF INSTR(UCASE(VAR_PATH_SCANNER),UCASE(FOLDER_PATH_AVOID_ITEM(R_COUNT_2)))>0 then 
				SET_GO_FILE=FALSE
			END IF 
		ELSE
			I_X__=LEN(FOLDER_PATH_AVOID_ITEM(R_COUNT_2))
			I_TXT=UCASE(FOLDER_PATH_AVOID_ITEM(R_COUNT_2))
			IF RIGHT(UCASE(VAR_PATH_SCANNER),I_X__)=I_TXT then SET_GO_FILE=FALSE
		END IF
	NEXT

	
	
	IF SET_GO_FILE=FALSE THEN EXIT SUB

	
	I_COUNT_2=I_COUNT_2+1
	
	IF STOP_PROGRESS=TRUE THEN EXIT SUB
	
	Set f = fso.GetFile(VAR_PATH_SCANNER)
	FileModDate = f.DateLastModified
	SET f=NOTHING

	IF IGNOR_DATE_INDEX=FALSE THEN 
		IF FileModDate<=FileModDate_2 THEN EXIT SUB
	END IF 

	' ONLY JPG IS REALLY CHECK FOR ANYTHING AT THE MOMENT
	' ---------------------------------------------------
	If RIGHT(ucase(VAR_PATH_SCANNER),4)=".JPG" then JPG_BEEN_DONE=TRUE
	If RIGHT(ucase(VAR_PATH_SCANNER),4)=".MP4" then MP4_BEEN_DONE=TRUE
	If RIGHT(ucase(VAR_PATH_SCANNER),4)=".AVI" then AVI_BEEN_DONE=TRUE
	
	' ".JPG"
	' ".AVI"
	' ".MOV"  ' DASH CAMERA
	MOVE_MODE_FLAG=FALSE
	FOR R_COUNT_2=1 TO UBOUND(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE)
		I_X__=LEN(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_2))
		I_TXT=UCASE(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_2))
		IF RIGHT(UCASE(VAR_PATH_SCANNER),I_X__)=I_TXT then 
			MOVE_MODE_FLAG=TRUE
		END IF
	NEXT

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
	PATH_FILE_VAR=REPLACE(PATH_FILE_VAR,"\\","\")

	If Not FSO.FolderExists(PATH_VAR) Then
		BuildFullPath PATH_VAR
	End If
	
	ON ERROR RESUME NEXT
	ERR.CLEAR
	

	If Not FSO.FileExists(PATH_FILE_VAR) OR DO_ALL_FILE=TRUE Then
		IF MOVE_MODE_FILE=FALSE THEN 
			' MSGBOX "1.. COPY "+VAR_PATH_SCANNER+" -- "+PATH_FILE_VAR
			Fso.CopyFile VAR_PATH_SCANNER, PATH_FILE_VAR
		END IF
		IF MOVE_MODE_FILE=TRUE THEN 
			IF MOVE_MODE_FLAG=TRUE THEN
				' MSGBOX "2.. MOVE "+VAR_PATH_SCANNER+" -- "+PATH_FILE_VAR
				Fso.MoveFile VAR_PATH_SCANNER, PATH_FILE_VAR
			END IF
			IF MOVE_MODE_FLAG=FALSE THEN
				' MSGBOX "3.. COPY "+VAR_PATH_SCANNER+" -- "+PATH_FILE_VAR
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



'-------------------------------------------------------------------
'-------------------------------------------------------------------

' -------------------------------------------------------------------
' THE BLOCK 4 ROUTINE BELOW IS USE FOR RENAME 
' OF TYPE PUT DATE IN FRONT FILE FOR CAMERA
' -------------------------------------------------------------------


Sub Main_4
	
	I_COUNT=0
	IF FSO.FolderExists(FOLDER_PATH_MAIN_4)=FALSE THEN 
		EXIT SUB
	END IF

    Set oDict_4 = CreateObject("Scripting.Dictionary")
    ' Call Recursive Sub
    FindWantedFiles_4(objFSO.GetFolder(FOLDER_PATH_MAIN_4))
    ' Display All Findings from Dictionary object
    DisplayFindings_4
	IF I_COUNT>0 THEN 
		MSGBOX "DONE RENAME COUNT AT" & VBCRLF & FOLDER_PATH_MAIN_4 & VBCRLF & CSTR(I_COUNT) & " of " & CSTR(oDict_4.count)
	END IF
	
    Set oDict_4 = Nothing
End Sub

Sub FindWantedFiles_4(sFolder)
    On Error Resume Next
    Dim oFDR, oItem, file
    ' Check if wanted files are in this folder
    For Each file In sFolder.Files
		IF INSTR(file,".JPG") THEN
			oDict_4.Add FILE.path, file
		END IF
		IF INSTR(file,".AVI") THEN
			oDict_4.Add FILE.path, file
		END IF
    Next
    ' Recurse into it's sub folders
    For Each oFDR In sFolder.SubFolders
        FindWantedFiles_4 objFSO.GetFolder(oFDR.Path)
    Next
	
End Sub

Sub DisplayFindings_4()
    Dim oKeys, oKey
    oKeys = oDict_4.Keys
    For Each oKey In oKeys
		VAR_PATH_SCANNER=oKey
		CALL PATH_SCANNER_RENAME_4
    Next
End Sub

SUB PATH_SCANNER_RENAME_4

	' ---------------------
	' 4K-ULTRA-HD -- FORMAT
	' ---------------------
	' JPG00000.JPG
	' REC00003.AVI
	' D:\DSC\2019+4K ULTRA HD\2019 4K ULTRA\DCIMA
	' ---------------------
	
	' --------------------------------
	' SONY CYBERSHOT AFTER DONE FORMAT
	' --------------------------------
	' 2019-01-05 00-13-12 - SONY DSC-HX60V - DSC03146.JPG
	' 2019_01_07 Jan_Mon 14_41_48__MAQ03178.MP4
	' --------------------------------
	
	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)

	FileModDate=0
	Set f = fso.GetFile(VAR_PATH_SCANNER)
		FileModDate = f.DateLastModified
	SET f=NOTHING	
	IF FileModDate=0 THEN 
		MSGBOX "ERROR FILE DATE = 0 " + VBCRLF + VAR_PATH_SCANNER
	END IF
	
	IF INSTR("4K-ULTRA-HD",FOLDER_PATH_R_COUNT) THEN 
		' IF LEN(FILE_VAR)<>12 THEN EXIT SUB
		i_TIME=FILE_VAR
		IF INSTR(FILE_VAR,".JPG") THEN 
			IF MID(FILE_VAR,1,3)="JPG" THEN 
				i_TIME=""
				i_TIME=i_TIME+LPad(Year(FileModDate), "0", 4)+"-"
				i_TIME=i_TIME+LPad(Month(FileModDate), "0", 2)+"-"
				i_TIME=i_TIME+LPad(Day(FileModDate), "0", 2)
				i_TIME=i_TIME+" "
				i_TIME=i_TIME+LPad(Hour(FileModDate), "0", 2)+"-"
				i_TIME=i_TIME+LPad(Minute(FileModDate), "0", 2)+"-"
				i_TIME=i_TIME+LPad(Second(FileModDate), "0", 2)
				i_TIME=i_TIME+" - 4K ULTRA A_CAM - "+FILE_VAR
			END IF
		END IF
		IF INSTR(FILE_VAR,".AVI") THEN 
			SET_GO=FALSE
			IF MID(FILE_VAR,1,3)="REC" THEN SET_GO=TRUE
			IF MID(FILE_VAR,1,3)="SOS" THEN SET_GO=TRUE
			IF SET_GO=TRUE THEN
				i_TIME=""
				i_TIME=i_TIME+LPad(Year(FileModDate), "0", 4)+"_"
				i_TIME=i_TIME+LPad(Month(FileModDate), "0", 2)+"_"
				i_TIME=i_TIME+LPad(Day(FileModDate), "0", 2)+" "
				i_TIME=i_TIME+MonthName(MONTH(FileModDate),TRUE)+"_"
				i_TIME=i_TIME+WeekdayName(Weekday(VAR_NOW),TRUE)

				i_TIME=i_TIME+" "
				i_TIME=i_TIME+LPad(Hour(FileModDate), "0", 2)+"_"
				i_TIME=i_TIME+LPad(Minute(FileModDate), "0", 2)+"_"
				i_TIME=i_TIME+LPad(Second(FileModDate), "0", 2)+"__"
				i_TIME=i_TIME+FILE_VAR
			END IF
		END IF

		NewFilename=i_TIME
		NewPATH=PATH_VAR
		IF INSTR(FILE_VAR,".AVI") THEN
		IF INSTR(PATH_VAR,"2 AVI\")=0 THEN 
				NewPATH=REPLACE(PATH_VAR,"4K ULTRA\DCIMA","4K ULTRA")+"2 AVI\"
			END IF
		END IF
		IF INSTR(FILE_VAR,".JPG") THEN 
			IF INSTR(PATH_VAR,"1 JPG\")=0 THEN 
				NewPATH=REPLACE(PATH_VAR,"4K ULTRA\DCIMA","4K ULTRA")+"1 JPG\"
			END IF
		END IF
		' msgbox PATH_VAR+VBCRLF+NewPATH

		NewFilename=NewPATH+NewFilename
		IF FSO.FolderExists(NewPATH)=FALSE THEN 
			CreateFolderTree(NewPATH)
		END IF
		
	END IF

	IF i_TIME="" THEN EXIT SUB
	
	' NewFilename=i_TIME
	
	' NewFilename=REPLACE(NewFilename,"__","_")
	' NewFilename=REPLACE(NewFilename," _ ","_")
	' NewFilename=REPLACE(NewFilename,"_ ","_")
	' NewFilename=REPLACE(NewFilename," _","_")
	' NewFilename=REPLACE(NewFilename,".jpg",".JPG")
	' Extension=ucase(mid(NewFilename,len(NewFilename)-3))
	
	' NewFilename=PATH_VAR+NewFilename
	
	if VAR_PATH_SCANNER <> NewFilename then 
		msgbox NewFilename
		I_COUNT=I_COUNT+1
		ON ERROR RESUME NEXT
		ERR.CLEAR
		Fso.MoveFile VAR_PATH_SCANNER, NewFilename
		IF ERR.NUMBER>0 THEN 
			MSGBOX ERR.DESCRIPTION & VBCRLF & VAR_PATH_SCANNER & VBCRLF & NewFilename
		END IF
		'Wscript.QUIT
		'Wscript.QUIT
	end if

END SUB

' -------------------------------------------------------------------
' THE BLOCK 4 ROUTINE ABOVE IS USE FOR RENAME 
' OF TYPE PUT DATE IN FRONT FILE FOR CAMERA
' -------------------------------------------------------------------

Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

Function CreateFolderTree(sPath)
    Dim nPos
    On Error RESUME NEXT

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        If nPos - 1 > 3 Then
            If Dir(Left(sPath, nPos - 1), vbDirectory) = "" Then
				newfolderpath=Left(sPath, nPos - 1)
                If Not FSO.FolderExists(newfolderpath) Then 
					Set newfolder = FSO.CreateFolder(newfolderpath) 
				END IF
            End If
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
	newfolderpath=sPath
	If Not FSO.FolderExists(newfolderpath) Then 
		Set newfolder = FSO.CreateFolder(newfolderpath) 
	END IF
    CreateFolderTree = True
	If Not FSO.FolderExists(newfolderpath) Then 
        CreateFolderTree = False
    End If

End Function




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