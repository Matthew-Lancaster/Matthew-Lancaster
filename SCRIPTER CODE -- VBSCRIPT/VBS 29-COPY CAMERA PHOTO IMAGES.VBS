'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 29-COPY CAMERA PHOTO IMAGES 01.VBS
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

' -------------------------------------------------------------------
' 008
' -------------------------------------------------------------------
' PUT LOAD OF CHANGE
' RENAME IS DO FOR THING THAT NOT EXIF -- BY MOD DATE
' CHECK EXPLORER IS RUN NOT RUN TOO MANY ON -- KNOW HOW THAT GET
' ROUTINE NEW MAIN_4 _ SEARCH FOR RENAME -- DO WITHIN HERE NOT EXTERNAL
' WHEN FOR SIMPLE RENAME -- BY MOD DATE
' ALSO MAKE 
' A GOODSYNC SCRIPT RUN FROM HERE
' THAT DO A BACKUP OF MY MEDIA CARD
' THAT IS SIMPLE AS LEAVE CARD CONNECT COMPUTER ALL THE TIME
' MY MC CARD HAS 2 YEAR OF STUFF ON
' AND IF LOST OR DAMAGER -- BE PAIN TO CORRECT
' MY PHOTOGRAPHY -- IS DANGER OF STOLEN
' PEOPLE DON'T LIKE IT HAVING CAMERA
' IF THERE WERE SABOTAGE TO CAMERA AND LOST MEDIA CARD
' HAVE NONE PROOF WHAT DO-ER
' LIKE TRAFFIC WARDEN
' BUT THEY HAVE IT EASIER
' MOST LIKELY THEY DELETE CAMERA ROLL BEFORE START NEW SHIFT
' AND IF IN CONFLICT WITH ANYONE THEY DON'T TALK 
' ONLY PHOTO-ER
' AND THEN YOU GOT PAPERERTZZE PHOTO ANYONE
' AND NOT MUCH CHANCE OF GET IT OFF THEM
' DAMAGE THEIR CAMERA OR THEM AND LOOK
' SOD ALL THAT
' BACK IT UP ANOTHER MEDIA CARD
' TAKE A PHOTO EVERYDAY BEFORE YOU GO OUT ANYWHERE
' AND -- A PROFESSIONAL SECURITY GUARD
' WILL ASK TAKE EVERY PHOTO OF MACHINE IF FOUND ON THEIR PROPERTY
' AND PROFESSIONAL SECURITY GUARD WILL HAVE READ
' -- DATA PROTECTION ACT 2018
' -- PREVENT PROFESSIONAL BURGLARY
' -- KEEP YOUR BANK HAPPY
' -------------------------------------------------------------------
' Sun 21-Jul-2019 17:54:45
' Mon 22-Jul-2019 00:58:00
' -------------------------------------------------------------------


' CHECK IF ANY WSCRIPT ALREADY RUNNER AND QUIT IF IS WORKER
' -------------------------------------------------------------------

DIM ORIGINAL_PATH
DIM EXPLORER_COUNTING
DIM LOAD_EXPLORER_HAPPEN

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")

' -------------------------------------------------------------------
' HERE WAS WORK FOR A BIT 
' AND BELOW A BETTER CODE
' -------------------------------------------------------------------
Dim iMyPID
' -------------------------------------------------------------------
' iMyPID = GetObject("winmgmts:root\cimv2").Get("Win32_Process.Handle='" & CreateObject("WScript.Shell").Exec("mshta.exe").ProcessID & "'").ParentProcessId
' -------------------------------------------------------------------
' MSGBOX iMyPID
' -------------------------------------------------------------------

' -------------------------------------------------------------------
FOR EACH objProcess IN colProcesses
	IF NOT (IsNull(objProcess.ProcessID )) THEN
		iMyPID = objProcess.ProcessID
	END IF
NEXT
' -------------------------------------------------------------------
' MSGBOX "MY PROCESS NUMBER PID IS -- "+ CSTR(iMyPID),vbMsgBoxSetForeground,Wscript.ScriptName 
' -------------------------------------------------------------------



Set args = Wscript.Arguments
For Each arg In args
	Wscript_Arguments = Wscript_Arguments + arg
Next







DIM MSGBOX_SHOW

MSGBOX_SHOW=TRUE

IF INSTR(Wscript_Arguments,"RUN_FROM_GOODSYNC")>0 THEN 
	MSGBOX_SHOW=FALSE
END IF

IF INSTR(Wscript_Arguments,"NOT_SHOW")>0 THEN ' ---- "RUN_FROM_GOODSYNC"
	MSGBOX_SHOW=FALSE
END IF




Dim strScriptName
Dim objProcess


i1 = 0 ' ANY PROGRAM WSCRIPT 
i2 = 0 ' MY  PROGRAM WSCRIPT 
FOR EACH objProcess IN colProcesses
    IF NOT (IsNull(objProcess.CommandLine )) THEN
		strScriptName = Trim(Right(objProcess.CommandLine,Len(objProcess.CommandLine) - InstrRev(objProcess.CommandLine,"\")))
        strScriptName = Left(strScriptName, Len(strScriptName) - 1)
		PID_Script = objProcess.ProcessID
		
		IF strScriptName = Wscript.ScriptName THEN
			i2=i2+1
		END IF
		IF strScriptName <> Wscript.ScriptName THEN
			i1=i1+1
		END IF
		
		IF i1 > 0 THEN ' ANY PROGRAM WSCRIPT 
			' Wscript.Quit
		END IF
		IF strScriptName = Wscript.ScriptName THEN ' MY PROGRAM WSCRIPT 
			' -------------------------------------------------------
			' KILL PROCESS NUMBER OF ALL OTHER INSTANCE 
			' HERE SCRIPT WHEN START
			' -------------------------------------------------------
			
			IF iMyPID<>PID_Script AND PID_Script>0 THEN 
			    ' MSGBOX CSTR(iMyPID) + VBCRLF + CSTR(PID_Script)
				objProcess.Terminate() 

			END IF
		END IF
    END IF
NEXT
                             
							 
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

DIM CUR_USER
DIM ComputerName

DIM FOLDER_PATH_DEST_VAR_2
DIM FOLDER_PATH_MP4__VAR_2
DIM SET_GO_MAIN_4

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




' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------

DIM DRIVE_FREE_SPACE_GB_BEGIN
DIM DRVPATH_VAR
DIM GET_SPACE_STRING
DIM FREE_SPACE_SET_TO_EQUAL
DIM	DRIVE_FREE_SPACE_GB
DIM	DRIVE_CAPACITY___GB
DIM I_COUNT_DEL
DIM I_COUNT_TOTAL
DIM DAY_BACK_TRIMMER
DIM LOW_DATE
DIM HIGH_DATE

LOW_DATE=NOW
HIGH_DATE=NOW-1000

DIM I_COUNT
DIM MSGBOX_I1 , MSGBOX_I2
DIM ODICT_TRIM_THE_DATE
FOLDER_PATH_TRIM_THE_DATE_OLDER_CAMERA="V:\"
CALL MAIN__TRIM_THE_DATE_OLDER_CAMERA


WSCRIPT.QUIT
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------





Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")


Dim arr1
Dim oDict     ' Key: Full filename, Item: Filename
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
' -------------------------------------------------------------------
'FindWantedFiles objFSO.GetFolder(FOLDER_PATH_VAR)
'DisplayFindings
' -------------------------------------------------------------------

DIM R_COUNT
DIM CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR_DONE_ONCE

DIM FOLDER_PATH_INDEX

NUMBER_OF_ITEM = 100

DIM FOLDER_PATH()
DIM FOLDER_PATH_DEST()
DIM FOLDER_PATH_VAR()
DIM FOLDER_PATH_DO_ALL_FILE()
DIM FOLDER_PATH_MOVE_MODE_FILE()
DIM FOLDER_PATH_IGNOR_DATE_INDEX()
DIM FOLDER_PATH_DONT_SHOW_EXPLORER_UNLESS_COPY()

DIM WORK_DOER
DIM FOLDER_PATH_HOLD
DIM DO_SEND_TO_PUT_DATE_INFRONT_FILENAME_DONE
DIM DO_SEND_TO_PUT_DATE_INFRONT_FILENAME
DIM DO_SEND_TO_PUT_DATE_INFRONT_FILENAME_SONY_CAMERA
DIM DO_SEND_TO_PUT_DATE_INFRONT_FILENAME_FUJI
DIM VBS_54_CAMERA_MOVE_COPY_FROM_MP_ROOT_TO_DROP_ONE_PATH_IMAGER

REDIM FOLDER_PATH(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_DEST(NUMBER_OF_ITEM)
REDIM FOLDER_PATH__MP4(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_VAR(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_DO_ALL_FILE(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_MOVE_MODE_FILE(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_IGNOR_DATE_INDEX(NUMBER_OF_ITEM)
REDIM FOLDER_PATH_DONT_SHOW_EXPLORER_UNLESS_COPY(NUMBER_OF_ITEM)

' -------------------------------------------------------------------
' SUB PATH_SCANNER_COPY_CAMERA
' CHECK TOP OF ROUTINE HERE 
' -------------------------------------------------------------------

DIM FOLDER_PATH_DEST_VAR
DIM FOLDER_PATH_VAR_3

DIM EXPLORER_LOADED
DIM SHOW_EXPLORER_FLAG
DIM SHOW_EXPLORER_QUESTION

DIM FOLDER_PATH_MAIN_4

WHERE_COPY_TO_FOR_EXPLORER_JPG=""
DIM PATH_VAR_FOR_EXPLORER


FOLDER_PATH(01)="MC - HX60V"
FOLDER_PATH(02)="MC - HX60V0"		' VIDEO
FOLDER_PATH(03)="FUJI XP90"
FOLDER_PATH(04)="GOPRO"
FOLDER_PATH(05)="GOPRO"
FOLDER_PATH(06)="DUAL SCRN"			' __ COPY FIRST TO END PLACE
FOLDER_PATH(07)="DUAL SCRN"			' __ ON SECOND RUN MOVE NEXT TO ANOTHER
FOLDER_PATH(08)="D DRIVE"
FOLDER_PATH(09)="DASH_CAMERA"		' __ COPY FIRST TO END PLACE
FOLDER_PATH(10)="DASH_CAMERA"		' __ ON SECOND RUN MOVE NEXT TO ANOTHER
FOLDER_PATH(11)="4K-ULTRA-HD"		' __ COPY FIRST TO END PLACE
FOLDER_PATH(12)="4K-ULTRA-HD"		' __ ON SECOND RUN MOVE NEXT TO ANOTHER
FOLDER_PATH(13)="4K-ULTRA-02"		' __ COPY FIRST TO END PLACE
FOLDER_PATH(14)="4K-ULTRA-02"		' __ ON SECOND RUN MOVE NEXT TO ANOTHER
FOLDER_PATH(15)="KEY_FOB"			' __ DATE ARRIVE 23-NOV-2019
FOLDER_PATH(16)="MEM-BANK-CAMERA"	' __ DATE ARRIVE 30-DEC-2019
FOLDER_PATH(17)="FRONT DOOR"		' __ AUDIO DEVICE FRONT DOOR
FOLDER_PATH(18)="HIKVISION#1"		' __ 128 SAMSUNG PEN DRIVE
FOLDER_PATH(19)="MAGICMINDR7"		' __ AUDIO DEVICE JACKET CAMERA
FOLDER_PATH(20)="HI-LOOK-04"		' __ HIK-VISION WHITE 4 PORT
FOLDER_PATH(21)="HI-LOOK-08"		' __ HIK-VISION BLACK 8 PORT
FOLDER_PATH(22)="HARD_REC_8G"		' __ £25.99 HTTPS://TINYURL.COM/Y3RWB8PF -- Tue 03-Mar-2020
									' ACTIONS HS USB FlashDisk USB ---- FAT32
									' Digital Voice Recorder, COOAU 8GB Rechargeable: Amazon.co.uk: Electronics 
									' https://www.amazon.co.uk/gp/product/B07G5BVWS8/ref=ppx_yo_dt_b_asin_title_o00_s00?ie=UTF8&psc=1
									' ---------------------------------------------------------------
FOLDER_PATH(23)="MAGICMINDR5"		' __ AUDIO DEVICE JACKET CAMERA NEW -- Fri 27-Mar-2020 15:47:25
FOLDER_PATH(24)="RETEKESS"
FOLDER_PATH(25)="128GB-V-OPA"		' __ £54.99 HTTPS://TINYURL.COM/Y9CN3UG3
									' ---------------------------------------------------------------
									' £54.99 HTTPS://TINYURL.COM/Y9CN3UG3
									' Spy Camera, OMOUP 10000mAh Power Bank with Camera,
									' 1080P HD Hidden Camera Nanny Cam with Motion Detection 
									' Night Vision Gravity Sensing Function
									' 17-JUN-2020 ORDER -- 18-JUN-2020 DELIVERY -- 20-JUN-2020 GET GO
									' NEW MODEL -- HAS EVERYTHING
									' AND OVER SPEC TAKE 128GB MC-CARD -- THAN SPECIFY MAXIMUM 64GB
									' ---------------------------------------------------------------
FOLDER_PATH(26)="HARD-RIPER2"   	' __ £19.88 HTTPS://TINYURL.COM/Y3RWB8PF -- Sat 20-Jun-2020
									' ACTIONS HS USB FlashDisk USB ---- FAT32
									' Digital Voice Recorder, COOAU 8GB Rechargeable: Amazon.co.uk: Electronics 
									' https://www.amazon.co.uk/gp/product/B07G5BVWS8/ref=ppx_yo_dt_b_asin_title_o00_s00?ie=UTF8&psc=1
									' ---------------------------------------------------------------



' FOLDER_PATH()=""
' FOLDER_PATH()=""
' FOLDER_PATH()=""
' FOLDER_PATH()=""
' FOLDER_PATH()=""
' FOLDER_PATH()=""
' FOLDER_PATH()=""
' FOLDER_PATH()=""
' FOLDER_PATH()=""
' FOLDER_PATH()=""

' -----------------
' NOKIA TO DO 1 2 3


DIM FOLDER_PATH_AVOID_ITEM()
REDIM PRESERVE FOLDER_PATH_AVOID_ITEM(100)
NUMBER_COUNTER = 0
I=NUMBER_COUNTER
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="_gsdata_"
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="System Volume Information"
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="\F.A"                          ' DASH CAMERA
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=".THM"                          ' SONY
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=".AmazonDriveConfig"
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="CAPTURE.TXT"                   ' 4K ULTRA ACTION CAM
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="TIME.TXT"                      ' KEY_FOB
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="VIDEO.LST"                     ' KEY_FOB
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="Settime"                       ' BLUE AUDIO FRONT DOOR
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="Baby One More Time.mp3"        ' BLUE AUDIO FRONT DOOR
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="\SetDigitRecorderTm.exe.lnk"   ' BLUE AUDIO FRONT DOOR
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="\SetTime.txt"                  ' BLUE AUDIO FRONT DOOR
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="#Send To ModifyDate.exe"
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="MAGICMINDR7--JACKET USB.txt"   ' BLUE AUDIO FRONT DOOR
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="\Settime - BACKER"             ' BLUE AUDIO FRONT DOOR
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="GUID_"                         ' GUID_C77607481_20190817085651.bin 		                
																' 2020_CCSE_HIKVISION_DS-7204H
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=".EXE"							' HIKVISION
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=".ZIP"							' HIKVISION
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="ipcCfg_"						' HIKVISION -- ipcCfg_20190519043855.xls
I=I+1:FOLDER_PATH_AVOID_ITEM(I)=".xls"							' HIKVISION -- ipcCfg_20190519043855.xls
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="\MUSIC"                        ' RETEKESS
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="index0.tmp"                    ' Spy Camera, OMOUP 10000mAh Power Bank with Camera, 1080P HD Hidden Camera Nanny Cam with Motion Detection Night Vision Gravity Sensing Function
I=I+1:FOLDER_PATH_AVOID_ITEM(I)="time.txt"                      ' Spy Camera, OMOUP 10000mAh Power Bank with Camera, 1080P HD Hidden Camera Nanny Cam with Motion Detection Night Vision Gravity Sensing Function
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
' NOT USE HERE

' ---------------------------------------------------------------
' DON'T PUT JPG OR MP4 HERE 
' ALL WILL GET INCLUDE WITH MOVE
' OR FILTER CAMERA DON'T WANT MOVE
' ONLY ODD TITLE LIKE AVI AND MOV
' ---------------------------------------------------------------
' HAVE TO USE HERE
' ---------------------------------------------------------------
' FOLDER_PATH_MOVE_MODE_FILE(2)=TRUE
' FOLDER_PATH_DO_ALL_FILE(2)=TRUE
' ---------------------------------------------------------------


'I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".JPG"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".AVI"
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".MOV"      ' DASH CAMERA
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".WAV"      ' KEY_FOB - AUDIO BLUE RECORDER
' I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".MP4"    ' HIKVISION WHITE 4 PORT
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=".MP3"      ' RETEKESS
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=""
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=""
I=I+1:FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(I)=""
FOR R_COUNT_2=1 TO UBOUND(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE)
	IF FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_2)<>"" THEN 
		NUMBER_COUNTER=R_COUNT_2
	END IF 
NEXT
REDIM PRESERVE FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(NUMBER_COUNTER)
	
' -------------------------------------------------------------------
	
FOLDER_PATH_DONT_SHOW_EXPLORER_UNLESS_COPY(8)=TRUE

FOR R=1 TO 40
	FOLDER_PATH_VAR(R)="\"
NEXT 
FOLDER_PATH_VAR(8)="\Android\SDCard2_Sync\DCIM\"




' ------------------------
FOLDER_PATH_DEST(1)="D:\DSC\2015+SONY\" + CSTR(Year(Now)) + " CyberShot HX60V"
' ------------------------

	

IF ComputerName="7-ASUS-GL522VW" THEN 
FOLDER_PATH__MP4(1)="\\7-asus-gl522vw\7_asus_gl522vw_80_3_samsung_4tb_d\DSC_4G_1TB\DSC\2015+SONY_MP4\" + CSTR(Year(Now)) + " CyberShot HX60V __ MP4"
END IF
	
	
IF ComputerName="4-ASUS-GL522VW" THEN 

	FOLDER_PATH__MP4(1)="F:\DSC\2015+SONY_MP4\" + CSTR(Year(Now)) + " 2020 CyberShot HX60V __ MP4"
	' ---------------------------------------------------------------
	' USER NETWORK PATH IF HASSLE DRIVE LETTER CHANGER
	' ---------------------------------------------------------------
	IF FSO.FolderExists(FOLDER_PATH__MP4(1))=FALSE THEN 
		FOLDER_PATH__MP4(1)="\\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY_MP4\" + CSTR(Year(Now)) + " CyberShot HX60V __ MP4"
		' \\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY_MP4\2020 CyberShot HX60V __ MP4
	END IF

END IF 

	

' ------------------------

FOLDER_PATH_DEST(2)="\\7-asus-gl522vw\7_asus_gl522vw_02_d_drive\VI_ DSC ME\2010+SONY\" + CSTR(Year(Now)) + " CyberShot HX60V_ IMAGE"
FOLDER_PATH__MP4(2)="\\7-asus-gl522vw\7_asus_gl522vw_80_3_samsung_4tb_d\VI_ DSC ME\2010+SONY\" + CSTR(Year(Now)) + " CyberShot HX60V_#"


' ------------------------

FOLDER_PATH_DEST(3)="D:\DSC\2017+FUJI XP90\" + CSTR(Year(Now)) + " FinePix XP90"
FOLDER_PATH_DEST(4)="D:\DSC\" + CSTR(Year(Now)) + " GoPro 01"
FOLDER_PATH_DEST(5)="D:\VI_ DSC ME\" + CSTR(Year(Now)) + " GoPro 01"
FOLDER_PATH_DEST(6)="D:\DSC\" + CSTR(Year(Now)) + " Double Screen Cam"
FOLDER_PATH_DEST(7)="D:\VI_ DSC ME\" + CSTR(Year(Now)) + " Double Screen Cam"

FOLDER_PATH_DEST(8)="D:\DSC\" + CSTR(Year(Now)) + " Samsung Mini S5\DCIM"

FOLDER_PATH_DEST(9)="D:\DSC\" + CSTR(Year(Now)) + " DASH CAMERA"
FOLDER_PATH_DEST(10)="D:\VI_ DSC ME\" + CSTR(Year(Now)) + " DASH CAMERA"

IF ComputerName="4-ASUS-GL522VW" THEN
	NET_PATH="\\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC--2019+4K ULTRA\"
	FOLDER_PATH_DEST(11)=NET_PATH + CSTR(Year(Now)) + " 4K ULTRA HD"
	FOLDER_PATH_DEST(13)=NET_PATH + CSTR(Year(Now)) + " 4K ULTRA 02"
END IF
IF ComputerName="7-ASUS-GL522VW" THEN 
	NET_PATH="\\7-asus-gl522vw\7_asus_gl522vw_80_3_samsung_4tb_d\DSC_4G_1TB\DSC--2019+4K ULTRA\"
	FOLDER_PATH_DEST(11)=NET_PATH + CSTR(Year(Now)) + " 4K ULTRA HD"
	FOLDER_PATH_DEST(13)=NET_PATH + CSTR(Year(Now)) + " 4K ULTRA 02"
END IF

FOLDER_PATH_DEST(15)="D:\DSC\2019+KEY_FOB\" + CSTR(Year(Now)) + " KEY_FOB"
FOLDER_PATH_DEST(16)="D:\DSC\2019+MEM-BANK-AR\" + CSTR(Year(Now)) + " MEMORY-BANK-CAMERA"
FOLDER_PATH_DEST(17)="D:\0 00 MOBILE-2\8GB RECORD AUDIO BLUE _ 1 _ FRONT DOOR\RECORD " + CSTR(Year(Now)) 
FOLDER_PATH_DEST(19)="D:\0 00 MOBILE-2\8GB RECORD AUDIO BLUE _ 7 _ OUTDOOR JACKET\RECORD " + CSTR(Year(Now))

IF ComputerName="4-ASUS-GL522VW" THEN 
	FOLDER_PATH_DEST(18)="F:\DSC--2018+CCSE_HIKVISION\"+CSTR(Year(Now))+"_CCSE_HIKVISION_DS-7204H"
	FOLDER_PATH_DEST(20)="F:\DSC--2018+CCSE_HIKVISION\"+CSTR(Year(Now))+"_CCSE_HIKVISION_DVR-104G-F1"
	FOLDER_PATH_DEST(21)="F:\DSC--2018+CCSE_HIKVISION\"+CSTR(Year(Now))+"_CCSE_HIKVISION_DVR-108G-F1"
END IF
IF ComputerName="7-ASUS-GL522VW" THEN 
	FOLDER_PATH_DEST(18)="T:\DSC_4G_1TB\DSC--2018+CCSE_HIKVISION\"+CSTR(Year(Now))+"_CCSE_HIKVISION_DS-7204H"
	FOLDER_PATH_DEST(20)="T:\DSC_4G_1TB\DSC--2018+CCSE_HIKVISION\"+CSTR(Year(Now))+"_CCSE_HIKVISION_DVR-104G-F1"
	FOLDER_PATH_DEST(21)="T:\DSC_4G_1TB\DSC--2018+CCSE_HIKVISION\"+CSTR(Year(Now))+"_CCSE_HIKVISION_DVR-108G-F1"
END IF

FOLDER_PATH_DEST(22)="D:\0 00 MOBILE-2\8GB RECORD HARD_REC_8G\" + CSTR(Year(Now))+"_HARD_REC_8G"
FOLDER_PATH_DEST(23)="D:\0 00 MOBILE-2\8GB RECORD AUDIO BLUE _ 5 _ OUTDOOR JACKET_02\RECORD " + CSTR(Year(Now))
FOLDER_PATH_DEST(24)="D:\0 00 MOBILE-2\RETEKESS\RETEKESS_RECORD " + CSTR(Year(Now))


IF ComputerName="4-ASUS-GL522VW" THEN 
	FOLDER_PATH_DEST(25)="F:\DSC--2020+128GB-V-OPA\"+CSTR(Year(Now))+"_128GB-V-OPA"
END IF
IF ComputerName="7-ASUS-GL522VW" THEN 
	FOLDER_PATH_DEST(25)="T:\DSC_4G_1TB\DSC--2020+128GB-V-OPA\"+CSTR(Year(Now))+"_128GB-V-OPA"
END IF
' ---------------------------------------------------------------
' £54.99 HTTPS://TINYURL.COM/Y9CN3UG3
' Spy Camera, OMOUP 10000mAh Power Bank with Camera,
' 1080P HD Hidden Camera Nanny Cam with Motion Detection 
' Night Vision Gravity Sensing Function
' 17-JUN-2020 ORDER -- 18-JUN-2020 DELIVERY -- 20-JUN-2020 GET GO
' NEW MODEL -- HAS EVERYTHING
' AND OVER SPEC TAKE 128GB MC-CARD -- THAN SPECIFY MAXIMUM 64GB
' ---------------------------------------------------------------

FOLDER_PATH_DEST(26)="D:\0 00 MOBILE-2\8GB RECORD HARD_RIPER_02_8G\" + CSTR(Year(Now))+"_HARD_RIPER_02_8G"

' WORK HERE TO DO -- NEW DELETE OLD MEDIA -- Fri 18-Sep-2020 10:17:37




FOLDER_PATH_MOVE_MODE_FILE(2)=TRUE
FOLDER_PATH_DO_ALL_FILE(2)=TRUE
FOLDER_PATH_DO_ALL_FILE(4)=TRUE
FOLDER_PATH_DO_ALL_FILE(5)=TRUE
FOLDER_PATH_MOVE_MODE_FILE(5)=TRUE
FOLDER_PATH_DO_ALL_FILE(6)=TRUE
FOLDER_PATH_DO_ALL_FILE(7)=TRUE
FOLDER_PATH_MOVE_MODE_FILE(7)=TRUE
' SAMSUNG MINI S 5 _ ALLOW OVER WRITER & ALLOW ALL FILE EXTENSION TYPE TYPEO TYPO P & E & O
FOLDER_PATH_DO_ALL_FILE(8)=TRUE
' FOLDER_PATH_IGNOR_DATE_INDEX(8)=TRUE
' DASH CAMERA
FOLDER_PATH_DO_ALL_FILE(9)=TRUE
FOLDER_PATH_DO_ALL_FILE(10)=TRUE
FOLDER_PATH_MOVE_MODE_FILE(10)=TRUE
' DO ALL THE FILE FROM SD CARD FIRST
' AND THEN DO AGAIN FROM SD CARD AND ERASE THEM AFTER
FOLDER_PATH_DO_ALL_FILE(11)=TRUE
FOLDER_PATH_DO_ALL_FILE(12)=TRUE
FOLDER_PATH_MOVE_MODE_FILE(12)=TRUE
' DO ALL THE FILE FROM SD CARD FIRST
' AND THEN DO AGAIN FROM SD CARD AND ERASE THEM AFTER
FOR R= 13 TO 50
	FOLDER_PATH_DO_ALL_FILE(R)=TRUE
	FOLDER_PATH_MOVE_MODE_FILE(R)=TRUE
NEXT

' -------------------------------------------------------------------
' PATH ARE BUILD LATER WHILE METHOD SEARCH LINE CODE
' If Not FSO.FolderExists(PATH_VAR) Then
' -------------------------------------------------------------------

FOR R_COUNT_2=1 TO UBOUND(FOLDER_PATH_DEST)
	IF FOLDER_PATH_DEST(R_COUNT_2)<>"" THEN NUMBER_COUNTER=R_COUNT_2
NEXT

REDIM PRESERVE FOLDER_PATH(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH_DEST(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH__MP4(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH_VAR(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH_DO_ALL_FILE(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH_MOVE_MODE_FILE(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH_IGNOR_DATE_INDEX(NUMBER_COUNTER)
REDIM PRESERVE FOLDER_PATH_DONT_SHOW_EXPLORER_UNLESS_COPY(NUMBER_COUNTER)

DIM CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR
CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR=FALSE

DIM FOLDER_PATH_R_COUNT
OLD_FOLDER_PATH_VAR_2=""

HDD_INFO=""
FOR R_COUNT=1 TO UBOUND(FOLDER_PATH)

	FOLDER_PATH_VAR_2=""
	FOLDER_PATH_INDEX=CSTR(R_COUNT)+" "+FOLDER_PATH(R_COUNT)
	' ---------------------------------------------------------------
	' TRY GET VOLUME LETTER LABEL FROM VOLUME_NAME OF PROVIDER
	' IF EXIST GOT A DRIVE LETTER OR EMPTY 
	' ---------------------------------------------------------------
	' AVOID D_DRIVE VOLUME NAME AS ALWAYS PRESENT
	' IF ANOTHER DRIVE PRESENT DROP D_DRIVE
	' ---------------------------------------------------------------
	IF R_COUNT<>8 THEN 
		FOLDER_PATH_VAR_2=getDriveLetterFromVolumeName(FOLDER_PATH(R_COUNT))
		HDD_INFO = HDD_INFO + FOLDER_PATH_VAR_2
	END IF
NEXT	
IF HDD_INFO<>"" THEN 
	FOLDER_PATH(8)=""
END IF


' -------------------------------------------------------------------
' GENERAL ONLY ONE USB AT A TIME BUTT MAYBE FEW END RESULT NOT SAME 
' ONLY WHAT IN FOR NEXT LOOP 
' -------------------------------------------------------------------
FOR R_COUNT=1 TO UBOUND(FOLDER_PATH)

	I_COUNT_1=0
	I_COUNT_2=0

	FOLDER_PATH_VAR_2=""
	FOLDER_PATH_INDEX=CSTR(R_COUNT)+" "+FOLDER_PATH(R_COUNT)
	' ---------------------------------------------------------------
	' TRY GET VOLUME LETTER LABEL FROM VOLUME_NAME OF PROVIDER
	' IF EXIST GOT A DRIVE LETTER OR EMPTY 
	' ---------------------------------------------------------------
	FOLDER_PATH_VAR_2=getDriveLetterFromVolumeName(FOLDER_PATH(R_COUNT))
	
	FOLDER_PATH_DEST_VAR=FOLDER_PATH_DEST(R_COUNT)
	FOLDER_PATH__MP4_VAR=FOLDER_PATH__MP4(R_COUNT)
	
	' ---------------------------------------------------------------
	' HERE GIVE THE VOLUME LABEL REFERENCE ALL PROCESS OF CODE
	' ---------------------------------------------------------------
	IF FOLDER_PATH_VAR_2<>"" THEN 
		FOLDER_PATH_HOLD_2=FOLDER_PATH(R_COUNT)
	END IF
	
	
	IF INSTR(OLD_FOLDER_PATH_VAR_2,FOLDER_PATH_VAR_2+"--")=0  THEN 
		EXPLORER_LOADED=FALSE
	END IF
	OLD_FOLDER_PATH_VAR_2=OLD_FOLDER_PATH_VAR_2+"--"+FOLDER_PATH_VAR_2
	
	FOLDER_PATH_VAR_3=FOLDER_PATH_VAR(R_COUNT)
	FOLDER_PATH_R_COUNT=FOLDER_PATH(R_COUNT)
	
	FOLDER_PATH_HOLD=FOLDER_PATH(R_COUNT)

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
	IF INSTR("FUJI XP90 --",FOLDER_PATH(R_COUNT)) THEN 
		IF FOLDER_PATH_VAR_2<>"" THEN 
			DO_SEND_TO_PUT_DATE_INFRONT_FILENAME_FUJI=TRUE         ' FUTURE
			DO_SEND_TO_PUT_DATE_INFRONT_FILENAME=TRUE 
		END IF
	END IF

	IF INSTR("MC - HX60V -- MC - HX60V0 --",FOLDER_PATH(R_COUNT)) THEN 
		IF FOLDER_PATH_VAR_2<>"" THEN 
			DO_SEND_TO_PUT_DATE_INFRONT_FILENAME_SONY_CAMERA=TRUE  ' FUTURE
			DO_SEND_TO_PUT_DATE_INFRONT_FILENAME=TRUE
		END IF
	END IF
	
	IF INSTR("MC - HX60V",FOLDER_PATH(R_COUNT)) THEN 
		IF FOLDER_PATH_VAR_2<>"" THEN 
			ONLY_SONY_CYBERSHOT_HAS_DO=TRUE
			RUN_FILE="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 00-DELETE_EMPTY_FOLDER_CAMERA.VBS"
			IF FSO.FileExists(RUN_FILE) THEN 
				WSHShell.Run """"+RUN_FILE+"""", DontShowWindow, WaitUntilFinished
			END IF

		END IF
	END IF

	IF FOLDER_PATH_VAR_2<>"" THEN 
		IF R_COUNT<>8 THEN 
			' -------------------------------------------------------
			' SHOW FIND-PATH AND DESTINATION-PATH
			' -------------------------------------------------------
			' MSGBOX "1.."+FOLDER_PATH_VAR_2 + VBCRLF +"2.."+ FOLDER_PATH_DEST_VAR
			' -------------------------------------------------------
		END IF

		' CHECK HERE AND ANOTHER IN FILE OPERATION
		IF FOLDER_PATH_VAR_2<>"D:" THEN 
			HDD_WAS_HERE=FOLDER_PATH_VAR_2
			HDD_WAS_HERE_2=FOLDER_PATH(R_COUNT)
			' WORK_GO_ER=TRUE
			' MSGBOX FOLDER_PATH_VAR_2
		END IF

		' -----------------------------------------------------------
		' NOW GOT THE DRIVE LETTER ADD PATH GOING TO USER
		' -----------------------------------------------------------
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
			' IF ComputerName="1-ASUS-X5DIJ" THEN SHOW_EXPLORER_FLAG=FALSE
			' IF ComputerName="2-ASUS-EEE" THEN SHOW_EXPLORER_FLAG=FALSE
			' IF ComputerName="3-LINDA-PC" THEN SHOW_EXPLORER_FLAG=FALSE
			' IF ComputerName="4-ASUS-GL522VW" THEN SHOW_EXPLORER_FLAG=FALSE
		END IF
		
		IF FOLDER_PATH_DONT_SHOW_EXPLORER_UNLESS_COPY(R_COUNT)=TRUE THEN
			SHOW_EXPLORER_FLAG=FALSE
		END IF
		
		IF EXPLORER_COUNTING=>2 THEN 
			' SHOW_EXPLORER_FLAG=FALSE
		END IF

		IF LOAD_EXPLORER_HAPPEN=FALSE THEN 
			LOAD_EXPLORER_HAPPEN=TRUE
			' IF SHOW_EXPLORER_FLAG=TRUE AND EXPLORER_LOADED=FALSE THEN 
			' EXPLORER_LOADED=TRUE
			
			' WSHShell.Run "EXPLORER """+FOLDER_PATH_DEST_VAR_2+"""" , ShowWindow, DontWaitUntilFinished
			' REMMER ---- Tue 10-Mar-2020 17:38:35
			' WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 29 VBS 29-COPY CAMERA LOAD EXPLORER.VBS.AHK"""+ " """+FOLDER_PATH_DEST_VAR_2+"""" , ShowWindow, DontWaitUntilFinished

			FOLDER_PATH_DEST_VAR_2=PATH_FILE_VAR
			FOLDER_PATH_DEST_VAR_2=FOLDER_PATH_DEST_VAR
			FOLDER_PATH_DEST_VAR_PARAM=REPLACE(FOLDER_PATH_DEST_VAR_2," ","*")
			WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 29 VBS 29-COPY CAMERA LOAD EXPLORER.VBS.AHK"""+ " """+FOLDER_PATH_DEST_VAR_PARAM+"""" , ShowWindow, DontWaitUntilFinished
			
		END IF
		
		IMAGE_POINTER_FOLDER_1=MID(Wscript.ScriptFullName,1,INSTRREV(Wscript.ScriptFullName,"\"))+"VBS 29-COPY CAMERA\"
		IMAGE_POINTER_NAME=IMAGE_POINTER_FOLDER_1+"COPY_CAMERA_PHOTO_IMAGE_POINTER_"+FOLDER_PATH(R_COUNT)+".TXT"
		IMAGE_POINTER_NAME=REPLACE(IMAGE_POINTER_NAME,"\SCRIPTER\","\SCRIPTOR DATA\")
		
		IMAGE_POINTER_FOLDER_DATA=MID(Wscript.ScriptFullName,1,INSTRREV(Wscript.ScriptFullName,"\"))+"VBS 29-COPY CAMERA\"
		IMAGE_POINTER_FOLDER_DATA=REPLACE(IMAGE_POINTER_FOLDER_DATA,"\SCRIPTER\","\SCRIPTOR DATA\")
		
		' C:\SCRIPTOR DATA\SCRIPTER CODE -- VBSCRIPT\VBS 29-COPY CAMERA		
		
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

		MAIN

		If FSO.FileExists(IMAGE_POINTER_NAME) Then 
			FSO.DeleteFile IMAGE_POINTER_NAME
		END IF

		' -----------------------------------------------------------
		' BOTH VIDEO AND DSC FOLDER MC - HX60V0 & MC - HX60V 
		' -----------------------------------------------------------
		' -----------------------------------------------------------
		' -----------------------------------------------------------
		' -----------------------------------------------------------
		' -----------------------------------------------------------
		' -----------------------------------------------------------
		' -----------------------------------------------------------
		' -----------------------------------------------------------
		' -----------------------------------------------------------
		' -----------------------------------------------------------
		' -----------------------------------------------------------
		X_FIND=MID(FOLDER_PATH(R_COUNT),1,10)
		IF X_FIND="MC - HX60V" THEN 
			DeleteFolder__=FOLDER_PATH_VAR_2
			Set objFolder = FSO.GetFolder(DeleteFolder__)
			Set oDict = CreateObject("Scripting.Dictionary")
			CALL FindWantedFiles_THUMB_OFF_CAMERA_HX60(objFolder)
			CALL DisplayFinder_THUMB_OFF_CAMERA_HX60

			DeleteFolder__=FOLDER_PATH_VAR_2+"DCIM"
			Set objFolder = FSO.GetFolder(DeleteFolder__)
			ORIGINAL_PATH=UCASE(objFolder.PATH)
			DeleteEmptyFolders objFolder

			DeleteFolder__=FOLDER_PATH_VAR_2+"MP_ROOT"
			Set objFolder = FSO.GetFolder(DeleteFolder__)
			ORIGINAL_PATH=UCASE(objFolder.PATH)
			DeleteEmptyFolders objFolder
		END IF
		
		' --------------------------------------------------
		If Not FSO.FolderExists(IMAGE_POINTER_FOLDER_DATA) Then
			BuildFullPath IMAGE_POINTER_FOLDER_DATA
		End If
		' --------------------------------------------------
		Set objFile = FSO.CreateTextFile(IMAGE_POINTER_NAME)
			objFile.Write FileModDate_3
		SET objFile=NOTHING
		' --------------------------------------------------

		IF I_COUNT_1>0 THEN 
			WORK_GO_ER=TRUE
		END IF
		
		
		IF FOLDER_PATH_HOLD_2<>"D DRIVE" THEN 
			WORK_DOER_2=FOLDER_PATH_HOLD_2
		END IF

		IF WORK_GO_ER=TRUE THEN
			WORK_DOER=WORK_DOER+FOLDER_PATH(R_COUNT)+VBCRLF
			EXIT FOR
			' -------------------------------------------------------
			' IF DONE ONE DO NOT ANY MORE
			' EXIT FOR -- EXIT THE FOR NEXT LOOP
			' -------------------------------------------------------
		END IF
	END IF
NEXT

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

IF INSTR(WORK_DOER,"MC - HX60V")>0 THEN 
	RENAME_VAR="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 27-RENAME FILE+SUBFOLDER_MC_HX60V.VBS"
	IF FSO.FileExists(RENAME_VAR) THEN 
		' WSHShell.Run """"+RENAME_VAR+"""", DontShowWindow, DontWaitUntilFinished
	END IF
END IF

IF INSTR(WORK_DOER,"MC - HX60V")>0 THEN 
	RENAME_VAR="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 27-RENAME FILE+SUBFOLDER_FRONT_DOOR.VBS"
	PARAM="MC_HX60V"  
	IF FSO.FileExists(RENAME_VAR) THEN
		' WSHShell.Run """"+RENAME_VAR+"""" + " " + PARAM, DontShowWindow, DontWaitUntilFinished
	END IF
END IF
IF INSTR(WORK_DOER,"MC - HX60V1")>0 THEN 
	RENAME_VAR="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 27-RENAME FILE+SUBFOLDER_FRONT_DOOR.VBS"
	PARAM="MC_HX60V1"  
	IF FSO.FileExists(RENAME_VAR) THEN
		' WSHShell.Run """"+RENAME_VAR+"""" + " " + PARAM, DontShowWindow, DontWaitUntilFinished
	END IF
END IF

IF INSTR(WORK_DOER,"KEY_FOB")>0 THEN 
	RENAME_VAR="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 27-RENAME FILE+SUBFOLDER_KEY_FOB.VBS"
	IF FSO.FileExists(RENAME_VAR) THEN 
		WSHShell.Run """"+RENAME_VAR+"""", DontShowWindow, DontWaitUntilFinished
	END IF
END IF

IF INSTR(WORK_DOER,"MEM-BANK-CAMERA")>0 THEN 
	RENAME_VAR="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 27-RENAME FILE+SUBFOLDER_MEMORY-BANK-CAMERA.VBS"
	IF FSO.FileExists(RENAME_VAR) THEN 
		WSHShell.Run """"+RENAME_VAR+"""", DontShowWindow, DontWaitUntilFinished
	END IF
END IF

IF INSTR(WORK_DOER,"FRONT DOOR")>0 THEN
	' ALWAYS SAME
	RENAME_VAR="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 27-RENAME FILE+SUBFOLDER_FRONT_DOOR.VBS"
	PARAM="BLUE_1_FRONTDOOR"  
	IF FSO.FileExists(RENAME_VAR) THEN
		WSHShell.Run """"+RENAME_VAR+"""" + " " + PARAM, DontShowWindow, DontWaitUntilFinished
	END IF
END IF

IF INSTR(WORK_DOER,"OUTDOOR JACKET")>0 THEN
	' ALWAYS SAME
	RENAME_VAR="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 27-RENAME FILE+SUBFOLDER_FRONT_DOOR.VBS"
	PARAM="OUTDOOR_JACKET"
	IF FSO.FileExists(RENAME_VAR) THEN
		WSHShell.Run """"+RENAME_VAR+"""", DontShowWindow, DontWaitUntilFinished
	END IF
END IF

' -----------------------------------------------------------
' GOT THE DRIVE WHERE VOLUME SEARCH ARE -- FOLDER_PATH_HOLD_2
' -----------------------------------------------------------
DO_THE_CLOCK=FALSE
IF INSTR(FOLDER_PATH_HOLD_2,"FRONT DOOR")   THEN DO_THE_CLOCK=TRUE
IF INSTR(FOLDER_PATH_HOLD_2,"MAGICMINDR")  THEN DO_THE_CLOCK=TRUE  ' "OUTDOOR JACKET"

IF DO_THE_CLOCK=TRUE THEN 
	EXE_VAR="D:\0 00 MOBILE-2\8GB RECORD AUDIO BLUE _ 7 _ OUTDOOR JACKET\Settime\SetDigitRecorderTm.exe"
	IF FSO.FileExists(EXE_VAR) THEN
		WSHShell.Run """"+EXE_VAR+"""", ShowWindow, DontWaitUntilFinished
	END IF
END IF
' -------------------------------------------------------------------


' -------------------------------------------------------------------
' ---- I=I+"DONE COPY COUNT: "  ---- I=I+CSTR(I_COUNT_1)
' ---- VAR_FORCE_GOODSYNC_GO=TRUE
' -------------------------------------------------------------------
IF INSTR(WORK_DOER,"MC - HX60V")=0 OR I_COUNT_1=0 THEN 
	VAR_FORCE_GOODSYNC_GO=FALSE
END IF
' -------------------------------------------------------------------
' -------------------------------------------------------------------
IF VAR_FORCE_GOODSYNC_GO=TRUE THEN
	CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR_DONE_ONCE=TRUE
END IF

IF CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR_DONE_ONCE=FALSE THEN
	CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR_DONE_ONCE=TRUE
	' ---------------------------------------
	FILE_NAME="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT"
	PARAM="CAMERA_BACKUP_MODE"
	' ---------------------------------------
	' WSHShell.Run """"+FILE_NAME+"""" + " " + ""+PARAM+"" , ShowWindow, DontWaitUntilFinished
	
	
	' ---------------------------------------------------------------
	' TAKE FROM HERE INSERT NEW
	' GO TO GET SD-CARD FROM CAMERA AND HOLDER COMPUTER AND DRIVE
	' AND TRIM THEM FROM OLDER
	' HELP THE SYNC DO BETTER
	' SCAN THE DIR
	' ONLY WHEN COPY SOME
	' ---------------------------------------------------------------
	' ---------------------------------------------------------------
	
	
		
	
	
	
	
	
END IF


INF=""
FOR R_COUNT=1 TO UBOUND(FOLDER_PATH)
	IF INSTR(INF,FOLDER_PATH(R_COUNT)+"  ")=0 THEN 
		INF=INF+FOLDER_PATH(R_COUNT)+"  " + VBCRLF
		IX=IX+1
	END IF
NEXT


IF WORK_GO_ER=TRUE OR 1=1 THEN
	I=""
	I=I+ Wscript.ScriptName + VBCRLF
	I=I+ VBCRLF
	I=I+ "DONE COPY COUNT: " + CSTR(I_COUNT_1)
	I=I+ " of " + CSTR(I_COUNT_2) + VBCRLF + VBCRLF
	I=I+ "From Date -- Index Pointer -- " + TT_DATE_POINTER + VBCRLF + VBCRLF
	I=I+ "HDD IS HERE -- " + HDD_WAS_HERE + " -- "+ WORK_DOER_2 + VBCRLF
	' I=I+ "WORK DRIVE VOLUME NAME -- " + HDD_WAS_HERE_2 + VBCRLF 
	
	I=I+ VBCRLF
	I=I+ DRIVE_SPACE_INFO
	I=I+ VBCRLF
	I=I+ VBCRLF
	
	I=I+"CHECKER VOLUME NAME OF -- "+CSTR(IX)+" TOTAL DUPE FILTER" + VBCRLF
	I=I+"--------------------------------------------------------" + VBCRLF
	I=I+INF
	'IF SHOW_EXPLORER_FLAG=TRUE OR I_COUNT_1>0 THEN 

	IF MSGBOX_SHOW=TRUE THEN 
		MSGBOX I,vbMsgBoxSetForeground,Wscript.ScriptName 
		IF I_COUNT_1=0 THEN 
			I=""
			I=I+ Wscript.ScriptName + VBCRLF
			I=I+ VBCRLF
			I=I+ "NOTHING DONE COPY COUNT: " + CSTR(I_COUNT_1) + VBCRLF + VBCRLF
			I=I+ "FROM DATE -- INDEX POINTER -- " + TT_DATE_POINTER + VBCRLF + VBCRLF
			I=I+ "WORK DRIVE VOLUME NAME ------ " + HDD_WAS_HERE_2 + VBCRLF + VBCRLF
			
			I=I+ "IF WANT REWIND BACK -- INDEX POINTER -- A DAY FOR OVERSHOOT" + VBCRLF
			I=I+ "OPERATION OR SKIPPER "
			lRet = MsgBox(I, vbYesNo + vbMsgBoxSetForeground,Wscript.ScriptName)
			IF lRet=VBYES THEN
				IMAGE_POINTER_FOLDER_1=MID(Wscript.ScriptFullName,1,INSTRREV(Wscript.ScriptFullName,"\"))+"VBS 29-COPY CAMERA\"
				IMAGE_POINTER_NAME=IMAGE_POINTER_FOLDER_1+"COPY_CAMERA_PHOTO_IMAGE_POINTER_"+HDD_WAS_HERE_2+".TXT"
				IMAGE_POINTER_NAME=REPLACE(IMAGE_POINTER_NAME,"\SCRIPTER\","\SCRIPTOR DATA\")
					
				' DIM FileModDate_2
				FileModDate_2=CDate(TT_DATE_POINTER)
				FileModDate_3=FileModDate_2-1

				If FSO.FileExists(IMAGE_POINTER_NAME) Then 
					FSO.DeleteFile IMAGE_POINTER_NAME
				END IF
				
				Set objFile = FSO.CreateTextFile(IMAGE_POINTER_NAME)
				objFile.Write FileModDate_3
				SET objFile=NOTHING
			END IF
		END IF
	END IF
END IF


SET objFSO = NOTHING
Wscript.QUIT
'-------------------------------------------------------------------
'-------------------------------------------------------------------

SUB SET_VAR_DESTINATION_FOLDER
	SET_GO_MAIN_4=FALSE
	IF INSTR("4K-ULTRA-HD",FOLDER_PATH_R_COUNT) THEN SET_GO_MAIN_4=TRUE
	IF INSTR("MC - HX60V",FOLDER_PATH_R_COUNT) THEN SET_GO_MAIN_4=TRUE
	IF INSTR("MC - HX60V0",FOLDER_PATH_R_COUNT) THEN SET_GO_MAIN_4=TRUE
	IF SET_GO_MAIN_4=TRUE THEN
		' IF INSTR(FOLDER_PATH_DEST(1),"\MP4")=0 THEN 
			' FOLDER_PATH_DEST(1)=FOLDER_PATH_DEST(1)+"\MP4"
			' IF FOLDER_PATH__MP4(1)<>"" THEN 
				' FOLDER_PATH__MP4(1)=FOLDER_PATH__MP4(1)+"\MP4"
			' END IF
		' END IF
		' IF INSTR(FOLDER_PATH_DEST(2),"\MP4")=0 THEN 
			' FOLDER_PATH_DEST(2)=FOLDER_PATH_DEST(2)+"\MP4"
			' IF FOLDER_PATH__MP4(2)<>"" THEN 
				' FOLDER_PATH__MP4(2)=FOLDER_PATH__MP4(2)+"\MP4"
			' END IF
		' END IF
		FOLDER_PATH_DEST_VAR_2=FOLDER_PATH_DEST(R_COUNT)
		FOLDER_PATH_MP4__VAR_2=FOLDER_PATH__MP4(R_COUNT)
		' MSGBOX FOLDER_PATH_MP4__VAR_2

	END IF
END SUB

Sub Main
    Set oDict = CreateObject("Scripting.Dictionary")
    ' Call Recursive Sub
    FindWantedFiles(objFSO.GetFolder(FOLDER_PATH_VAR_2))
    ' Display All Findings from Dictionary object
	
	FOLDER_PATH_DEST_VAR_2=FOLDER_PATH_DEST(R_COUNT)
	FOLDER_PATH_MP4__VAR_2=FOLDER_PATH__MP4(R_COUNT)
	' CALL SET_VAR_DESTINATION_FOLDER
	
    DisplayFindings

	IF MID(FOLDER_PATH_VAR_2,1,2)<>"D:" THEN 
		ShowDriveInfo(MID(FOLDER_PATH_VAR_2,1,2))
	END IF
	
	' ---------------------------------------------------------------
	' CHECK EXPLORER BEEN LOADED BEFORE _ IF NOT
	' NOW IS TIME TO MAKE SURE _ AND DECIDE IF THE COUNTER HAS GONE UP FOR ANY COPY FILE
	' WHILE STILL MAINTAIN NOT FIR ANY DODGY COMPUTER OR FOLDER
	' INCLUDE D DRIVE OR BE SHOW EVERY TIME
	' ---------------------------------------------------------------
	
	IF SHOW_EXPLORER_FLAG=TRUE AND EXPLORER_LOADED=FALSE AND I_COUNT_1>0 THEN 
		' WSHShell.Run "EXPLORER """+FOLDER_PATH_DEST_VAR_2+"""" , ShowWindow, DontWaitUntilFinished
		' REMMER ---- Tue 10-Mar-2020 17:38:35
		' WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 29 VBS 29-COPY CAMERA LOAD EXPLORER.VBS.AHK"""+ " """+FOLDER_PATH_DEST_VAR_2+"""" , ShowWindow, DontWaitUntilFinished
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
	
	IF AVI_BEEN_DONE=TRUE THEN
		MP4_BEEN_DONE=TRUE
	END IF
	
	IF DO_SEND_TO_PUT_DATE_INFRONT_FILENAME_DONE=FALSE THEN 
		IF DO_SEND_TO_PUT_DATE_INFRONT_FILENAME=TRUE THEN 
		' IF DO_SEND_TO_PUT_DATE_INFRONT_FILENAME_SONY_CAMERA=TRUE THEN 
			IF I_COUNT_1>=0 AND (JPG_BEEN_DONE=TRUE) THEN 
			' IF I_COUNT_1>=0 THEN 
				' VB6
				DO_SEND_TO_PUT_DATE_INFRONT_FILENAME_DONE=TRUE
				
				' AHK FOR RUN VB6 CODE
				' ---------------------------------------------------
				' WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 33-#0 Send To Date Of File In Front FileName.ahk""", ShowWindow, DontWaitUntilFinished
				
				' NEW WAY HERE
				CALL DATE_OF_FILE_IN_FRONT_FILENAME_SONY_CAMERA
				
				
			END IF
		END IF
	END IF

	IF VBS_54_CAMERA_MOVE_COPY_FROM_MP_ROOT_TO_DROP_ONE_PATH_IMAGER=FALSE THEN
		IF I_COUNT_1>=0 AND (MP4_BEEN_DONE=TRUE) THEN
			VBS_54_CAMERA_MOVE_COPY_FROM_MP_ROOT_TO_DROP_ONE_PATH_IMAGER=TRUE
			WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 54-CAMERA MOVE COPY FROM MP_ROOT TO DROP ONE PATH & IMAGER.VBS""", ShowWindow, DontWaitUntilFinished
		END IF
	END IF

	' ---------------------------------------------------------------	
	' ONLY THE SONY CYBERSHOT RUN HERE TO MAKE DUPE MC-CARD
	' ---------------------------------------------------------------	

	IF CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR=TRUE THEN 
		IF CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR_DONE_ONCE=FALSE THEN 
			IF I_COUNT_1>=0 AND (JPG_BEEN_DONE=TRUE OR MP4_BEEN_DONE=TRUE) THEN 

				CAMERA_BACKUP_MC_CARD_CODE_TO_ACTOR_DONE_ONCE=TRUE
				' ---------------------------------------
				FILE_NAME="C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 59-RUN GOODSYNC SET SCRIPTOR.BAT"
				PARAM="CAMERA_BACKUP_MODE"
				' ---------------------------------------
				' ---------------------------------------
				' ---------------------------------------
				' ---------------------------------------
				' TEMP REMOVE -- SOMETHING IS DELETE THING 
				' GOODSYNC
				' DOS CMD START
				' ---------------------------------------
				' ---------------------------------------
				' ---------------------------------------
				' ---------------------------------------
				'WSHShell.Run """"+FILE_NAME+"""" + " " + ""+PARAM+"" , ShowWindow, DontWaitUntilFinished
				' ---------------------------------------
				' ---------------------------------------
				' ---------------------------------------
				' ---------------------------------------
			END IF
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

	SET_GO_MAIN_4=FALSE
	IF INSTR("4K-ULTRA-HD",FOLDER_PATH_R_COUNT) THEN SET_GO_MAIN_4=TRUE
	IF INSTR("MC - HX60V",FOLDER_PATH_R_COUNT) THEN SET_GO_MAIN_4=TRUE
	IF INSTR("MC - HX60V0",FOLDER_PATH_R_COUNT) THEN SET_GO_MAIN_4=TRUE
	FOLDER_PATH_DEST_VAR_2=FOLDER_PATH_DEST(R_COUNT)
	FOLDER_PATH_MP4__VAR_2=FOLDER_PATH__MP4(R_COUNT)

	IF SET_GO_MAIN_4=TRUE THEN
		CALL MAIN_4
	END IF
	
	' ---------------------------------------------------------------	
	' ---------------------------------------------------------------	
	' ---------------------------------------------------------------	
	' ---------------------------------------------------------------	

    Set oDict = Nothing
End Sub
' -------------------------------------------------------------------










SUB DATE_OF_FILE_IN_FRONT_FILENAME_SONY_CAMERA

EXIT SUB

FOLDER_PATH_DEST(1)="D:\DSC\2015+SONY\" + CSTR(Year(Now)) + " CyberShot HX60V"
FOLDER_PATH__MP4(1)="\\7-asus-gl522vw\7_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY_MP4\" + CSTR(Year(Now)) + " CyberShot HX60V"
FOLDER_PATH_DEST(2)="\\7-asus-gl522vw\7_asus_gl522vw_02_d_drive\VI_ DSC ME\2010+SONY_MP4\" + CSTR(Year(Now)) + " CyberShot HX60V_#"

FOR R_CNT = 1 TO 2

	IF FSO.FolderExists(FOLDER_PATH_DEST(R_CNT))=TRUE THEN  
		PAD
	END IF

	IF FSO.FolderExists(FOLDER_PATH__MP4(R_CNT))=TRUE THEN  
	
	END IF


NEXT




End Sub








SUB GET_DRIVE_SPACE(drvPath)
	STRCOMPUTER = "." 
	SET objWMIService = GetObject("winmgmts:\\" & STRCOMPUTER & "\root\CIMV2") 
	SET colItems = objWMIService.ExecQuery( _
		"SELECT * FROM Win32_LogicalDisk where DeviceID='"+drvPath+"'",,48) 
	FOR EACH objItem in colItems 
		IF LEN(objItem.VolumeName)>0 then
			DRIVE_FREE_SPACE_GB=CDbl(objItem.FreeSpace)/1024/1024/1024
			DRIVE_CAPACITY___GB=CDbl(objItem.Size)/1024/1024/1024
		END IF
	NEXT
END SUB



FUNCTION SHOWDRIVEINFO(drvPath)
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
			S=S+FORMATNUMBER((CDbl(objItem.FreeSpace)/1024/1024/1024))  & " GB" & vbCrLf
			S=S+"Size:" & vbTab & vbTab
			S=S+FORMATNUMBER((CDbl(objItem.Size)/1024/1024/1024)) & " GB" & vbCrLf
			S=S+"Occupied Space:" & vbTab
			S=S+FORMATNUMBER((CDbl(objItem.Size - objItem.FreeSpace)/1024/1024/1024)) & " GB"
			IF IX_ERROR_MESSENGER<>"" THEN 
				S=S+vbCrLf
				S=IX_ERROR_MESSENGER
			END IF
			DRIVE_FREE_SPACE_GB=CDbl(objItem.FreeSpace)/1024/1024/1024
			DRIVE_CAPACITY___GB=CDbl(objItem.Size)/1024/1024/1024
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
	IF mid(ucase(VAR_PATH_SCANNER),3,5)="\DCIM"    then SET_GO_FILE=TRUE
	IF mid(ucase(VAR_PATH_SCANNER),3,8)="\MP_ROOT" then SET_GO_FILE=TRUE
	IF mid(ucase(VAR_PATH_SCANNER),3,6)="\DCIMA"   then SET_GO_FILE=TRUE
	IF mid(ucase(VAR_PATH_SCANNER),3,6)="\AUDIO"   then SET_GO_FILE=TRUE   ' KEY_FOB
	IF mid(ucase(VAR_PATH_SCANNER),3,6)="\PHOTO"   then SET_GO_FILE=TRUE   ' KEY_FOB
	IF mid(ucase(VAR_PATH_SCANNER),3,6)="\VIDEO"   then SET_GO_FILE=TRUE   ' KEY_FOB
	IF mid(ucase(VAR_PATH_SCANNER),3,7)="\RECORD"  then SET_GO_FILE=TRUE   ' BLUE AUDIO
	IF mid(ucase(VAR_PATH_SCANNER),3,8)="\MRECORD"  then SET_GO_FILE=TRUE  ' RETEKESS
	

	IF INSTR(FOLDER_PATH_HOLD,"HIKVISION#1")>0 THEN
		SET_GO_FILE=TRUE  ' ----------------------- WANT ALL FROM ROOT
	END IF
	IF INSTR(FOLDER_PATH_HOLD,"HI-LOOK-04")>0 THEN
		SET_GO_FILE=TRUE  ' ----------------------- WANT ALL FROM ROOT
	END IF
	IF INSTR(FOLDER_PATH_HOLD,"HI-LOOK-08")>0 THEN
		SET_GO_FILE=TRUE  ' ----------------------- WANT ALL FROM ROOT
	END IF
	
	DO_ALL_FILE=FALSE
	MOVE_MODE_FILE=FALSE
	IGNOR_DATE_INDEX=FALSE

	IF FOLDER_PATH_DO_ALL_FILE(R_COUNT)=TRUE THEN
		DO_ALL_FILE=TRUE
		IGNOR_DATE_INDEX=TRUE
	END IF
	IF FOLDER_PATH_MOVE_MODE_FILE(R_COUNT)=TRUE THEN
		MOVE_MODE_FILE=TRUE
	END IF
		' MSGBOX MOVE_MODE_FILE
		
	IF FOLDER_PATH_IGNOR_DATE_INDEX(R_COUNT)=TRUE THEN
		IGNOR_DATE_INDEX=TRUE
	END IF
		' MSGBOX IGNOR_DATE_INDEX
	
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
	
	FileModDate=0
	IF FSO.FileExists(VAR_PATH_SCANNER) THEN 
		Set f = fso.GetFile(VAR_PATH_SCANNER)
		ON ERROR RESUME NEXT
		FileModDate = f.DateLastModified
		SET f=NOTHING
		IF ERR.NUMBER>0 THEN 
			MSGBOX "FILE NOT GET DATE "+VBCR+VAR_PATH_SCANNER+VBCR+VBCR+"ERROR NUMBER"+VBCR+CSTR(ERR.NUMBER)+VBCR+"ERROR DESCRIPTION"+VBCR+ERR.DESCRIPTION+VBCR+VBCR+"PUT CURRENT DATE",vbMsgBoxSetForeground,Wscript.ScriptName 
			FileModDate=NOW
		END IF
		ON ERROR GOTO 0
	END IF
	
	' MSGBOX ucase(VAR_PATH_SCANNER)
	
	IF IGNOR_DATE_INDEX=FALSE THEN 
		IF FileModDate<=FileModDate_2 THEN EXIT SUB
	END IF 
	
	IF FileModDate>FileModDate_3 THEN FileModDate_3=FileModDate
	
	' ---------------------------------------------------
	If RIGHT(ucase(VAR_PATH_SCANNER),4)=".JPG" then JPG_BEEN_DONE=TRUE
	If RIGHT(ucase(VAR_PATH_SCANNER),4)=".MP4" then MP4_BEEN_DONE=TRUE
	If RIGHT(ucase(VAR_PATH_SCANNER),4)=".AVI" then AVI_BEEN_DONE=TRUE
	
	' ---------------------------------------------------------------
	' ".JPG"
	' ".AVI"
	' ".MOV"                  ' DASH CAMERA
	'   MOVE_MODE_FILE=FALSE  ' 
	' ---------------------------------------------------------------
	' DON'T PUT JPG HERE
	' ALL WILL GET INCLUDE WITH MOVE
	' OR FILTER CAMERA DON'T WANT MOVE
	' ONLY ODD TITLE LIKE AVI AND MOV AND WAV
	' ---------------------------------------------------------------
	' A BUGG HERE GOT TO CHECK IF STRING EMPTY
	' Sat 08-Feb-2020 06:40:20
	' ---------------------------------------------------------------
	' ALL MY FILE OUTDOOR CAMERA ARE WIPE NOW
	' ONLY OLDER BACKUP WITH A GAP
	' STOP ERROR HAPPEN ANOTHER WITH CODE FIX
	' ---------------------------------------------------------------
	' POSSIBLE CONVERSION PUT SOME BACK ON
	' ---------------------------------------------------------------
	FOR R_COUNT_2=1 TO UBOUND(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE)
		I_TXT=TRIM(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_2))
		IF I_TXT<>"" THEN 
			I_X__=LEN(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_2))
			I_TXT=UCASE(FOLDER_PATH_EXT_ONLY_TO_USE_WITH_MOVE(R_COUNT_2))
			IF RIGHT(UCASE(VAR_PATH_SCANNER),I_X__)=I_TXT then 
				MOVE_MODE_FILE=TRUE
			END IF
		END IF
	NEXT

	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)
	
	' PATH_VAR="D:\DSC\2015-2018\2018 CyberShot HX60V"+MID(PATH_VAR,3)
	
	PATH_VAR=FOLDER_PATH_DEST_VAR+MID(PATH_VAR,3)
	PATH_VAR=REPLACE(PATH_VAR,MID(FOLDER_PATH_VAR_3,2),"")
	PATH_VAR_8=FOLDER_PATH_DEST_VAR
	
	IF INSTR(UCASE(PATH_VAR),"CYBERSHOT HX60V")>0 THEN 
	IF INSTR(UCASE(FOLDER_PATH_DEST_VAR),"\DSC\2015+SONY")>0 THEN 
	IF INSTR(UCASE(FILE_VAR),".MP4")>0 THEN 
		PATH_VAR=FOLDER_PATH_MP4__VAR_2
		PATH_VAR_8=PATH_VAR
	END IF
	END IF
	END IF
	IF INSTR(UCASE(PATH_VAR),"CYBERSHOT HX60V")>0 THEN 
	IF INSTR(UCASE(FOLDER_PATH_DEST_VAR),"\VI_ DSC ME\")>0 THEN   ' ALERT HERE IF EVER CHANGE PATH DRASTIC-AR 
																  ' THE STRING MUST FIND
																  ' AND THEN ENABLE JUMP EXPLORER BETWEEN JPG AND MP4
																  ' THE FOLDER OPEN AS JPG AND MP4 DIFFERENT LOCATION
																  ' AND OPEN FOR BOTH THEM WHEN FIRST FILE WRITE COPIER
																  '	Thu 26-Mar-2020 01:38:00
																  ' ---------------------------------------------------
	IF INSTR(UCASE(FILE_VAR),".MP4")>0 THEN 
		PATH_VAR=FOLDER_PATH_MP4__VAR_2
		PATH_VAR_8=PATH_VAR
	END IF
	END IF
	END IF

	IF PATH_VAR_FOR_EXPLORER <> PATH_VAR_8 THEN
		WHERE_COPY_TO_FOR_EXPLORER_JPG=""
		' -----------------------------------------------------------
		' RESET SHALL DISPLAY EXPLORER ANOTHER 
		' BETWEEN JPG AND MP4 FOLDER
		' -----------------------------------------------------------
	END IF
	PATH_VAR_FOR_EXPLORER=PATH_VAR_8
	
	PATH_FILE_VAR=PATH_VAR+"\"+FILE_VAR
	
	FileModDate_1=FileModDate
	' IF FSO.FileExists(VAR_PATH_SCANNER) THEN 
		' Set F1 = fso.GetFile(VAR_PATH_SCANNER)
		' ON ERROR RESUME NEXT
		' FileModDate_1 = F1.DateLastModified
		' ON ERROR GOTO 0
	' END IF

	' ---------------------------------------------------------------
	' CYBERSHOT HX60V -- MP4
	' ---------------------------------------------------------------
	IF INSTR(UCASE(PATH_VAR),"CYBERSHOT HX60V")>0 THEN 
	IF INSTR(UCASE(FILE_VAR),"MP4")>0 THEN 
	IF LEN(FILE_VAR)<41 THEN 
		' 41
		' 2020_01_03 Jan_Fri 22_47_16__MAH00824.MP4
		IF LEN(FILE_VAR)< 41 THEN
			i_TIME=""
			i_TIME=i_TIME+LPad(Year(FileModDate_1), "0", 4)+"_"
			i_TIME=i_TIME+LPad(Month(FileModDate_1), "0", 2)+"_"
			i_TIME=i_TIME+LPad(Day(FileModDate_1), "0", 2)+" "
			i_TIME=i_TIME+MonthName(MONTH(FileModDate_1),TRUE)+"_"
			i_TIME=i_TIME+WeekdayName(Weekday(FileModDate_1),TRUE)+" "
			i_TIME=i_TIME+""
			i_TIME=i_TIME+LPad(Hour(FileModDate_1), "0", 2)+"_"
			i_TIME=i_TIME+LPad(Minute(FileModDate_1), "0", 2)+"_"
			i_TIME=i_TIME+LPad(Second(FileModDate_1), "0", 2)+"__"
			NewFilename=i_TIME + UCASE(FILE_VAR)
		END IF
		STRIP_PART="\MP_ROOT\100ANV01"
		STRIP_PART="\MP_ROOT\"
		X_SP_1=INSTRREV(PATH_VAR,STRIP_PART)
		IF X_SP_1>0 THEN 
			X_SP_2=INSTRREV(PATH_VAR,"\")
			X_SP_4=MID(PATH_VAR,X_SP_1,X_SP_2-X_SP_1)
			PATH_VAR=REPLACE(PATH_VAR,X_SP_4,"")
		END IF
		PATH_FILE_VAR=PATH_VAR+"\"+NewFilename
	END IF
	END IF
	END IF


	' ---------------------------------------------------------------
	' CYBERSHOT HX60V -- JPG
	' 2020-03-10 09-48-46 - SONY DSC-HX60V - DSC01697.JPG
	' ---------------------------------------------------------------
	IF INSTR(UCASE(PATH_VAR),"CYBERSHOT HX60V")>0 THEN 
	IF INSTR(UCASE(FILE_VAR),"JPG")>0 THEN 
	IF LEN(FILE_VAR)<41 THEN 
		' 41
		' 2020_01_03 Jan_Fri 22_47_16__MAH00824.MP4
		IF LEN(FILE_VAR)< 41 THEN
			i_TIME=""
			i_TIME=i_TIME+LPad(Year(FileModDate_1), "0", 4)+"-"
			i_TIME=i_TIME+LPad(Month(FileModDate_1), "0", 2)+"-"
			i_TIME=i_TIME+LPad(Day(FileModDate_1), "0", 2)+" "
			FOLDER_DATE=""
			FOLDER_DATE=FOLDER_DATE+LPad(Year(FileModDate_1), "0", 4)+" "
			FOLDER_DATE=FOLDER_DATE+LPad(Month(FileModDate_1), "0", 2)+" "
			FOLDER_DATE=FOLDER_DATE+LPad(Day(FileModDate_1), "0", 2)
			
			i_TIME=i_TIME+""
			i_TIME=i_TIME+LPad(Hour(FileModDate_1), "0", 2)+"-"
			i_TIME=i_TIME+LPad(Minute(FileModDate_1), "0", 2)+"-"
			i_TIME=i_TIME+LPad(Second(FileModDate_1), "0", 2)+" "
			i_TIME=i_TIME+"- SONY DSC-HX60V - "
			NewFilename=i_TIME + UCASE(FILE_VAR)
		END IF
		STRIP_PART="\MP_ROOT\100ANV01"
		STRIP_PART="\DCIM\"
		X_SP_1=INSTRREV(PATH_VAR,STRIP_PART)
		' -----------------------------------------------------------
		' STRIP THE END PATH LEVEL AWAY AND REPLACE THE BIT VARIABLE ABOVE -- "\DCIM\"
		' THE END PATH IS FOLDER FROM THE CAMERA AND LAST 4 DIGIT AND DATE OF MONTH AND DAY
		' FROM HERE FOLDER DATE WILL MAKE FROM FILENAME DATE 
		' RATHER THAN THAT FOLDER THING STRIPPER
		' -----------------------------------------------------------
		IF X_SP_1>0 THEN 
			X_SP_2=INSTRREV(PATH_VAR,"\")
			X_SP_4=MID(PATH_VAR,X_SP_1,X_SP_2-X_SP_1)
			PATH_VAR=REPLACE(PATH_VAR,X_SP_4,"")
			PATH_VAR=PATH_VAR+"JPG\"+FOLDER_DATE+"\"
		END IF
		PATH_FILE_VAR=PATH_VAR+NewFilename
	END IF
	END IF
	END IF

	' ---------------------------------------------------------------
	' HIKVISION#1
	' ---------------------------------------------------------------
	IF INSTR(FOLDER_PATH_HOLD,"HIKVISION#1")>0 THEN
		IF INSTR(UCASE(FILE_VAR),".MP4")>0 THEN 
			' -----------------------------------------------------------
			' EXAMPLE ---- ch01_20190808094856.mp4
			' EXAMPLE ---- CH01_20190808--094856.MP4
			' -----------------------------------------------------------
			PATH_FILE_VAR=UCASE(PATH_VAR+MID(FILE_VAR,1,13)+"--"+MID(FILE_VAR,14))
		END IF
		IF INSTR(UCASE(FILE_VAR),".TXT")>0 THEN 
			EXIT SUB
		END IF
	END IF

	' ---------------------------------------------------------------
	' HI-LOOK-04
	' HI-LOOK-08
	' ---------------------------------------------------------------
	SET_GO=0
	IF INSTR(FOLDER_PATH_HOLD,"HI-LOOK-04")>0 THEN SET_GO=1
	IF INSTR(FOLDER_PATH_HOLD,"HI-LOOK-08")>0 THEN SET_GO=1
	IF	SET_GO=1 THEN 
		IF INSTR(UCASE(FILE_VAR),".MP4")>0 THEN 
			' -----------------------------------------------------------
			' EXAMPLE ---- ch01_20190808094856.mp4
			' EXAMPLE ---- CH01_20190808--094856.MP4
			' -----------------------------------------------------------
			PATH_FILE_VAR=UCASE(PATH_VAR+MID(FILE_VAR,1,13)+"--"+MID(FILE_VAR,14))
		END IF
		IF INSTR(UCASE(FILE_VAR),".TXT")>0 THEN 
			EXIT SUB
		END IF
	END IF

	

	If Not FSO.FolderExists(PATH_VAR) Then
		BuildFullPath PATH_VAR
	End If
	
	ON ERROR RESUME NEXT
	ERR.CLEAR
	
	' DOUBLE DASH IS BUILD SOMEWHERE WORTH A DEBUG LATER
	PATH_FILE_VAR_TMP=MID(PATH_FILE_VAR,3)
	IF INSTRREV(PATH_FILE_VAR_TMP,"\\")>0 THEN 
		PATH_FILE_VAR_TMP=REPLACE(PATH_FILE_VAR_TMP,"\\","\")
		PATH_FILE_VAR=MID(PATH_FILE_VAR,1,2)+PATH_FILE_VAR_TMP
	END IF

	PATH_VAR_TMP=MID(PATH_FILE_VAR,1,INSTRREV(PATH_FILE_VAR,"\"))
	If Not FSO.FolderExists(PATH_VAR_TMP) Then
		MSGBOX "DESTINATION FOLDER NOT EXIST -- POSSIBLE FORMAT WRONG"+VBCRLF++VBCRLF+PATH_VAR_TMP
	END IF 
	
	
	
	' DO_ALL_FILE=TRUE ---- OUGHT TO BE 
	' DO ALL FILE WITH OVERWRITE IF LOOK AR
	' ---------------------------------------------------------------
	SET_GO_COPY=FALSE
	' IF DO_ALL_FILE=TRUE                  THEN SET_GO_COPY=TRUE
	IF NOT FSO.FileExists(PATH_FILE_VAR) THEN SET_GO_COPY=TRUE
	If SET_GO_COPY=TRUE  THEN
		IF MOVE_MODE_FILE=FALSE  THEN 
			' WSCRIPT.QUIT

			Fso.CopyFile VAR_PATH_SCANNER, PATH_FILE_VAR
			I_COUNT_1=I_COUNT_1+1
			' LOAD EXPLORER FIRST FOLDER DESTINATION
			' --------------------------------------
			IF WHERE_COPY_TO_FOR_EXPLORER_JPG="" THEN 
				WHERE_COPY_TO_FOR_EXPLORER_JPG=PATH_FILE_VAR
				FOLDER_PATH_DEST_VAR_2=PATH_FILE_VAR
				' MSGBOX FOLDER_PATH_DEST_VAR_2
				FOLDER_PATH_DEST_VAR_PARAM=REPLACE(FOLDER_PATH_DEST_VAR_2," ","*")
				WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 29 VBS 29-COPY CAMERA LOAD EXPLORER.VBS.AHK"""+ " """+FOLDER_PATH_DEST_VAR_PARAM+"""" , ShowWindow, DontWaitUntilFinished
				LOAD_EXPLORER_HAPPEN=TRUE
			END IF
		END IF
		
		IF MOVE_MODE_FILE=TRUE THEN
			Fso.CopyFile VAR_PATH_SCANNER, PATH_FILE_VAR
			I_COUNT_1=I_COUNT_1+1
			IF WHERE_COPY_TO_FOR_EXPLORER_JPG="" THEN 
				WHERE_COPY_TO_FOR_EXPLORER_JPG=PATH_FILE_VAR
				FOLDER_PATH_DEST_VAR_2=PATH_FILE_VAR
				' MSGBOX FOLDER_PATH_DEST_VAR_2
				FOLDER_PATH_DEST_VAR_PARAM=REPLACE(FOLDER_PATH_DEST_VAR_2," ","*")
				WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 29 VBS 29-COPY CAMERA LOAD EXPLORER.VBS.AHK"""+ " """+FOLDER_PATH_DEST_VAR_PARAM+"""" , ShowWindow, DontWaitUntilFinished
				LOAD_EXPLORER_HAPPEN=TRUE
			END IF
			If FSO.FileExists(PATH_FILE_VAR) Then
				FSO.DeleteFile(VAR_PATH_SCANNER) 
			END IF
		END IF
	END IF
	
	' COPY ROUTINE ABOVE --------------------------------------------
	
	IF ERR.NUMBER>0 THEN
		IX=""+VBCRLF
		IX=IX+"ERROR COPY FILE _ TRY REDO ANOTHER"+VBCRLF+VBCRLF
		IX=IX+VAR_PATH_SCANNER +VBCRLF+VBCRLF
		IX=IX+PATH_FILE_VAR    +VBCRLF+VBCRLF
		IX=IX+ERR.DESCRIPTION  +VBCRLF+VBCRLF
		IX_ERROR_MESSENGER=IX
		IF IGNOR_DATE_INDEX=FALSE THEN 
			STOP_PROGRESS=TRUE
			EXIT SUB
		END IF
	END IF
END SUB

Sub BuildFullPath(ByVal FullPath)
	If Not fso.FolderExists(FullPath) Then
		ON ERROR RESUME NEXT
		BuildFullPath fso.GetParentFolderName(FullPath)
		fso.CreateFolder FullPath
		ON ERROR GOTO 0
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




' -------------------------------------------------------------------
' -------------------------------------------------------------------
Sub MAIN__TRIM_THE_DATE_OLDER_CAMERA
	I_COUNT=0
	IF FSO.FolderExists(FOLDER_PATH_TRIM_THE_DATE_OLDER_CAMERA)=FALSE THEN 
		EXIT SUB
	END IF
    Set ODICT_TRIM_THE_DATE = CreateObject("Scripting.Dictionary")
    ' CALL RECURSIVE SUB
    FINDWANTEDFILE_TRIM_THE_DATE(objFSO.GetFolder(FOLDER_PATH_TRIM_THE_DATE_OLDER_CAMERA))
    ' DISPLAY ALL FIND FROM DICTIONARY OBJECT


	
	CALL DISPLAYFIND_TRIM_THE_DATE
	
	IF I_COUNT_DEL>0 THEN 
		I = ""
		I = I + "DONE DELETE COUNT LESS THAN DAY BEFORE ____ MP4 AND JPG ____" & VBCRLF 
		I = I + "____ " & TRIM(CSTR(DAY_BACK_TRIMMER)) & " DAY" & VBCRLF 
		I = I + "____ " & FOLDER_PATH_TRIM_THE_DATE_OLDER_CAMERA & VBCRLF & VBCRLF 
		I = I + CSTR(I_COUNT_DEL) & " ____ DELETE ON ____ " & CSTR(I_COUNT_TOTAL) & " ____ TOTAL" & VBCRLF & VBCRLF
		I = I + "--------------------------------" & VBCRLF 
		I = I + "--------------------------------" & VBCRLF 
		I = I + "OLDER DATE OF ______ " & VBCRLF 
		I = I + "--------------------------------" & VBCRLF 
		I = I + MSGBOX_I1 & VBCRLF
		I = I + "NEWER DATE BEFORE ____ " & VBCRLF 
		I = I + "--------------------------------" & VBCRLF 
		I = I + MSGBOX_I2 & VBCRLF 
		I = I + "--------------------------------" & VBCRLF 
		I = I + "--------------------------------" & VBCRLF 
		I = I + "TOTAL FREE_SPACE_SET_TO_EQUAL____ " & FORMATNUMBER(FREE_SPACE_SET_TO_EQUAL,0) & " __ GB" & VBCRLF & VBCRLF 
		I = I + "TOTAL DRIVE_FREE_SPACE_BEGIN_____GB ____ " & VBCRLF & FORMATNUMBER(DRIVE_FREE_SPACE_GB_BEGIN,2) & VBCRLF & VBCRLF 
		I = I + "TOTAL DRIVE_FREE_SPACE_____________GB ____ " & VBCRLF & FORMATNUMBER(DRIVE_FREE_SPACE_GB,2) & VBCRLF & VBCRLF 
		I = I + "TOTAL DRIVE_CAPACITY_________________GB ____ " & VBCRLF & FORMATNUMBER(DRIVE_CAPACITY___GB,2) & VBCRLF 
		I = I + "--------------------------------" & VBCRLF 

		MSGBOX I,vbMsgBoxSetForeground,Wscript.ScriptName 
	END IF
	' DEL EMPTY
	
    Set ODICT_TRIM_THE_DATE = Nothing
	WSCRIPT.QUIT
End Sub
' -------------------------------------------------------------------
' -------------------------------------------------------------------
SUB FINDWANTEDFILE_TRIM_THE_DATE(sFolder)
    On Error Resume Next
    Dim oFDR, oItem, file
    ' CHECK IF WANTED FILES ARE IN THIS FOLDER
    For Each file In sFolder.Files
		STR_F="-- .JPG .MP4" 
		PATH_FIND_01="\DCIM"
		PATH_FIND_02="\MP_ROOT"
		EXTENSION_PART=MID(FILE.path,INSTRREV(FILE.path,"."))
		PATH_FIND_RESULT=FALSE
		IF INSTR(STR_F , EXTENSION_PART) >0 THEN PATH_FIND_RESULT=TRUE
		IF INSTR(UCASE(FILE.path) , PATH_FIND_01) >0 THEN PATH_FIND_RESULT=TRUE
		IF INSTR(UCASE(FILE.path) , PATH_FIND_02) >0 THEN PATH_FIND_RESULT=TRUE
		FILEMODDATE=0
		ERR.CLEAR
		SET F = FSO.GETFILE(FILE.path)
			FILEMODDATE = F.DATELASTMODIFIED
		SET F=NOTHING
		IF ERR.NUMBER>0 THEN FILEMODDATE=0
		IF FILEMODDATE=0 THEN 
			MSGBOX "ERROR RESULT GET DATE"
			PATH_FIND_RESULT=FALSE
		END IF
		IF PATH_FIND_RESULT=TRUE THEN 
			I_COUNT_TOTAL=I_COUNT_TOTAL+1
		END IF
		
		
		' DO HERE METHOD DAY GONE
		' -----------------------------------------------------------
		DAY_BACK_TRIMMER=NOW-120 ' DAY -- LESS NUMBER MORE IN FUTURE
		IF PATH_FIND_RESULT=TRUE THEN 
			IF FILEMODDATE>DAY_BACK_TRIMMER THEN 
				PATH_FIND_RESULT=FALSE
			END IF
		END IF
		IF PATH_FIND_RESULT=TRUE THEN 
			ODICT_TRIM_THE_DATE.Add FILE.path, file
			IF FILEMODDATE>HIGH_DATE THEN 
				HIGH_DATE=FILEMODDATE
				MSGBOX_I2=FILE.path & VBCR & FILEMODDATE & " ---- HIGH" 
			END IF
			IF FILEMODDATE<LOW_DATE THEN 
				LOW_DATE=FILEMODDATE
				MSGBOX_I1=FILE.path & VBCR & FILEMODDATE & " ---- LOW" 
			END IF
		END IF
		' -----------------------------------------------------------
		' DO AR METHOD FREESPACE
		' IN NEXT ROUTINE
		' -----------------------------------------------------------
		

    Next
    ' RECURSE INTO IT'S SUB FOLDERS
    For Each oFDR In sFolder.SubFolders
        FINDWANTEDFILE_TRIM_THE_DATE objFSO.GetFolder(oFDR.Path)
    Next
End Sub
' -------------------------------------------------------------------
' -------------------------------------------------------------------
Sub DISPLAYFIND_TRIM_THE_DATE()
    Dim oKeys, oKey
    oKeys = ODICT_TRIM_THE_DATE.Keys
    For Each oKey In oKeys
		ON ERROR RESUME NEXT
			GOGO=0
			ERR.CLEAR
			IF FSO.FILEEXISTS(oKey)=TRUE THEN GOGO=1
			IF ERR.NUMBER>0 THEN              GOGO=0
			' -----------------------------------------------------------
			' DO AR METHOD FREESPACE
			' IN NEXT ROUTINE
			' -----------------------------------------------------------
			IF GET_SPACE_STRING="" THEN
				GET_SPACE_STRING="IN IT"
				' -------------------------------------------------------
				DRVPATH_VAR=MID(FOLDER_PATH_TRIM_THE_DATE_OLDER_CAMERA,1,2)
				DRIVE_FUNCTION=SHOWDRIVEINFO(DRVPATH_VAR)
				IF DRIVE_CAPACITY___GB>40 THEN ' EXAMPLE 64GB CARD
					FREE_SPACE_SET_TO_EQUAL=19 ' GB
				END IF
				DRIVE_FREE_SPACE_GB_BEGIN=DRIVE_FREE_SPACE_GB
				' -------------------------------------------------------
				' CHECK DATE HAPPEN AND ALSO WANT 20 GB FREE
				' BUTTER NOT OVER DO HER
				' SET DATE MORE TRIM
				' -------------------------------------------------------
			END IF

			IF GET_SPACE_STRING<>"" THEN 
				CALL GET_DRIVE_SPACE(DRVPATH_VAR)
			END IF

			IF DRIVE_FREE_SPACE_GB<FREE_SPACE_SET_TO_EQUAL THEN 
				GOGO=1
				' ---------------------------------------------------
				' MSGBOX DRIVE_FREE_SPACE_GB & FREE_SPACE_SET_TO_EQUAL & "GB"
				' ---------------------------------------------------
			END IF
			' -------------------------------------------------------
			' EXAMPLE ON -- FORMATNUMBER
			' FORMATNUMBER(DRIVE_FREE_SPACE_GB,2) GIVE 2 DECIMAL PLACE 00.00
			' -------------------------------------------------------
			' MSGBOX FORMATNUMBER(DRIVE_FREE_SPACE_GB,2) & " ---- " &  FORMATNUMBER(FREE_SPACE_SET_TO_EQUAL,2)  & " -- GB"
			' -------------------------------------------------------

			IF GOGO=1 THEN 
				' MSGBOX OKEY
				' FSO.DeleteFile(oKey) 
				I_COUNT_DEL=I_COUNT_DEL+1
			END IF
		ON ERROR GOTO 0
		' -----------------------------------------------------------
    Next
End Sub
' -------------------------------------------------------------------
' -------------------------------------------------------------------




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
		MSGBOX "DONE RENAME COUNT AT" & VBCRLF & FOLDER_PATH_MAIN_4 & VBCRLF & CSTR(I_COUNT) & " of " & CSTR(oDict_4.count),vbMsgBoxSetForeground,Wscript.ScriptName 
	END IF
	
    Set oDict_4 = Nothing
End Sub



' -------------------------------------------------------------------
' MAH00004.THM
' &
' THUMBS.DB -- Thumbs.db
' -------------------------------------------------------------------
Sub FindWantedFiles_THUMB_OFF_CAMERA_HX60(sFolder)
    On Error Resume Next
    Dim oFDR, oItem, file
    ' CHECK IF WANTED FILES ARE IN THIS FOLDER
    For Each file In sFolder.Files
		IF INSTR(UCASE(FILE.path),".THM") THEN
			oDict.Add FILE.path, file
		END IF
		IF UCASE(FILE.path)="THUMBS.DB" THEN
			oDict.Add FILE.path, file
		END IF
    Next
    ' RECURSE INTO IT'S SUB FOLDERS
    For Each oFDR In sFolder.SubFolders
        FindWantedFiles_THUMB_OFF_CAMERA_HX60 objFSO.GetFolder(oFDR.Path)
    Next
End Sub
Sub DisplayFinder_THUMB_OFF_CAMERA_HX60()
    Dim oKeys, oKey
    oKeys = oDict.Keys
    For Each oKey In oKeys
		GOGO=0
		XX=MID(oKey,1,INSTRREV(oKey,"."))+"MP4"
		ON ERROR RESUME NEXT
		ERR.CLEAR
		IF FSO.FILEEXISTS(XX)=FALSE THEN GOGO=1
		IF ERR.NUMBER>0 THEN GOGO=0
		ON ERROR RESUME NEXT
		IF GOGO=1 THEN 
			' THE THUMB _.THM BEEN DELETE THAT NOT WITHER _.MP4
			FSO.DeleteFile(oKey) 
		END IF
		' -----------------------------------------------------------
		IF INSTR(UCASE(oKey),"THUMBS.DB") THEN XX=oKey
		IF FSO.FILEEXISTS(oKey)=TRUE THEN
			' THE -- THUMBS.DB -- Thumbs.db -- BEEN DELETE
			FSO.DeleteFile(oKey) 
		END IF
    Next
End Sub
' -------------------------------------------------------------------



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
		IF INSTR(file,".MP4") THEN
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
	' SOS00003.AVI
	' D:\DSC\2019+4K ULTRA\2019 4K ULTRA\DCIMA
	' ---------------------
	
	' --------------------------------
	' SONY CYBERSHOT AFTER DONE FORMAT
	' "MC - HX60V"
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
		MSGBOX "ERROR FILE DATE = 0 " + VBCRLF + VAR_PATH_SCANNER,vbMsgBoxSetForeground,Wscript.ScriptName 
	END IF
	
	' 4K-ULTRA-HD
	' ---------------------------------------------------------------
	' ---------------------------------------------------------------
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

		NewFilename=NewPATH+NewFilename
		IF FSO.FolderExists(NewPATH)=FALSE THEN 
			CreateFolderTree(NewPATH)
			' BuildFullPath PATH_VAR
		END IF
	END IF
	' ---------------------------------------------------------------
	' ---------------------------------------------------------------
	' 4K-ULTRA-HD  ---- ENDER
	' ---------------------------------------------------------------
	' ---------------------------------------------------------------
	
	
	' MC - HX60V
	' ---------------------------------------------------------------
	' ---------------------------------------------------------------
	SET_GO=FALSE
	IF INSTR("MC - HX60V",FOLDER_PATH_R_COUNT) THEN SET_GO=TRUE
	IF INSTR("MC - HX60V0",FOLDER_PATH_R_COUNT) THEN SET_GO=TRUE
	IF SET_GO=TRUE THEN
	
		' IF LEN(FILE_VAR)<>12 THEN EXIT SUB
		i_TIME=FILE_VAR
		IF INSTR(UCASE(FILE_VAR),".MP4") THEN 
			SET_GO=FALSE
			IF MID(FILE_VAR,1,2)="MA" THEN SET_GO=TRUE
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

		NewFilename=NewPATH+NewFilename
		IF FSO.FolderExists(NewPATH)=FALSE THEN 
			MSGBOX "FOLDER NOT EXIST ERROR"+VBCRLF+VBCLF+NewPATH,vbMsgBoxSetForeground,Wscript.ScriptName 
			' CreateFolderTree(NewPATH)
			' BuildFullPath PATH_VAR
		END IF
	END IF
	' ---------------------------------------------------------------
	' ---------------------------------------------------------------
	' MC - HX60V  ---- ENDER
	' ---------------------------------------------------------------
	' ---------------------------------------------------------------
	

	if VAR_PATH_SCANNER <> NewFilename then 
		' MSGBOX NewFilename
		I_COUNT=I_COUNT+1
		ON ERROR RESUME NEXT
		ERR.CLEAR
		Fso.MoveFile VAR_PATH_SCANNER, NewFilename
		IF ERR.NUMBER>0 THEN 
			MSGBOX ERR.DESCRIPTION & VBCRLF & VAR_PATH_SCANNER & VBCRLF & NewFilename,vbMsgBoxSetForeground,Wscript.ScriptName 
		END IF
		'WSCRIPT.QUIT
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
			msgbox "ERROR CLNG WITH "+VBCRLF+VBCRLF+NewFilename,vbMsgBoxSetForeground,Wscript.ScriptName 
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
				 
				msgbox NewFilename,vbMsgBoxSetForeground,Wscript.ScriptName 
				 
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
			msgbox NewFilename,vbMsgBoxSetForeground,Wscript.ScriptName 
			'Wscript.QUIT
			Fso.MoveFile VAR_PATH_SCANNER, NewFilename
			I_COUNT_1=I_COUNT_1+1
			' MSGBOX CSTR(I_COUNT_1)

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


Sub DeleteEmptyFolders(folder)
    Dim subfolder
	ON ERROR RESUME NEXT
	ERR.CLEAR
    For Each subfolder in folder.SubFolders
			DeleteEmptyFolders subfolder
    Next
	
	' ---------------------------------------------------------------
	' MORE ERROR SAFE METHOD
	' REMEMBER DECLARE ON AS OPTION EXPLICIT 
	' AND ERROR RESULT FROM DECLARE ERROR MAKE NOT PROPER TRAPP IT
	' ---------------------------------------------------------------
	READY_01=-1
	READY_02=-1
	READY_03=0
	READY_04=-1
	ERR.CLEAR
	READY_01=folder.SubFolders.Count
	READY_02=folder.Files.Count
	READY_04=ERR.NUMBER
	IF UCASE(ORIGINAL_PATH)=UCASE(folder.Path) THEN READY_03=1
	READY_04=ERR.NUMBER
	' ---------------------------------------------------------------
    If READY_01 = 0 And READY_02 = 0 And READY_03 = 0 AND READY_04=0 Then
		ON ERROR GOTO 0
		TEST_COUNTER=TEST_COUNTER+1
		ON ERROR RESUME NEXT
		ERR.CLEAR
		' -----------------------------------------------------------
		' SUCCESS ACHIEVE ON TEST RUN
		' ALLOW NOT MESSENGER LEAVER
		' Sun 08-Mar-2020 17:20:40
		' -----------------------------------------------------------
		'MSGBOX "DeleteFolder " + folder.Path
		FSO.DeleteFolder folder.Path, TRUE
		IF ERR.NUMBER>0 THEN 
			ERR_COUNTER = ERR_COUNTER +1
			ERROR_WITH_SOME = folder.Path & VBCRLF & ERR.DESCRIPTION
			ERR.CLEAR
			' objFolder.DELETE, TRUE
			IF ERR.NUMBER=0 THEN ERR_COUNTER = ERR_COUNTER - 1
		END IF
	End If
End Sub

		


'---------
' REFERENCE
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