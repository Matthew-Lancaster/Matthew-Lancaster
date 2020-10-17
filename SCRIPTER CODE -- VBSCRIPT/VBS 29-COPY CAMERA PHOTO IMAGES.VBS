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

' -------------------------------------------------------------------
' SESSION __ 009
' -------------------------------------------------------------------
' A LOAD OF WORK COMPLETER
' 1 2 3
' 1 MAKE ALL OUT OF DATE IMAGE A MP4 TRUNCATE OFF THE END
' SO BY DATE AND ALSO ALWAYS HAVE ON ABOUT 20 GB SPARE 
' GO OUT TOOLED WITH CARD
' I DO ENJOY THIS PART
' I BEEN WAIT FOR THIS CODER LONG TIME
' AND ALMOST COMPLETE EVERYTHING TO DO
' 2 -- I RECURSIVE DELETE ALL THE \FOUND.???
' 3 -- I MAKE DELETE EMPTY FOLDER -- EVERY GO
' -------------------------------------------------------------------
' AS TALK MANY
' I KEEP WHAT MY SODDING SD-CARD GOT THERE FOR OUTDOOR CAMERA
' AND ABLE SHOW ANYONE
' EVERYBODY HAPPY A CUSTOMER
' ONE THING WAS STUCK WITH SYNCHRONIZE PROFESSIONAL GOODSYNC
' ALL CORRECT EXCEPT ONE THING
' IT NOT LIKE HUGE NEAR RUN OUT OF SPACE
' AND ALSO SOME SD-CARD GET TO CERTAIN LEVEL AND NOT WORK WELL
' OF HOLD CONTAIN
' SO TRIM -- IS EASIER NOW
' ONLY SET HOW MANY GB WANT FREE 
' WHEN SYNCER
' AND NOW GOODSYNC GONNA BE FABBAR
' GOODSYNC AND SD-CARD HAVE TO OPERATE THAT NONE EVER DELETE
' MAKE MORE SUPERIOR WAY
' AND NONE SYNC PROGRAM CODE COMMERCIAL WILL DO THIS LOT
' WE LIKE IT KEEP ON 
' -------------------------------------------------------------------
' TIME SPEND FORM THIS MORNING WITHOUT ANY ASLEEP 
' -------------------------------------------------------------------
' Count = 062 -- Mon 05-Oct-2020 04:50:20
' -------------------------------------------------------------------
' HOW LONG
' -------------------------------------------------------------------
' GOT NEW CODE ROUTINE THAT NOT FIND ON-LINE HARDLY
' HOW JUMP INTO RECURSIVE ROUTINE AND NOT BIGG CODE LEAD IN
' -------------------------------------------------------------------
' THREE ROUTINE WORKER 
' -------------------------------------------------------------------
' RUN ROUTINE END NUMBER 03 AND RUN IT FOUR THAT POWER
' -------------------------------------------------------------------
' 0001 OF 0004 -------------------------------------
' CALL DELETE_FOUND_000_ROOT_PATH(ARRAY_SONY_FOLDER)
' 0002 OF 0004 -------------------------------------
' CALL DELETE_EMPTY_FOLDER(ARRAY_SONY_FOLDER)
' 0003 OF 0004 -------------------------------------
' CALL STRESSOR__TRIM_THE_DATE_OLDER_CAMERA
' 0004 OF 0004 -------------------------------------
' LINE 2800 AT END
' CALL RUNN_IT_HARDCORE_THREE_ROUTINE
' --------------------------------------------------
' -------------------------------------------------------------------
' AND SAVE ME FROM HAVE 3 BACK UP COPY WALK AROUND WITH
' AS IT STAND ABLE SLIP MIRROR COPY SIDE OF SD-CARD SLOT CAMERA 
' AND TWO OF THEM
' MAKE NICE PRESS NOT SLIP ANYWHERE
' BUTTER DON'T NEED THEM
' THE THREE ARE ONE BLANK AND ONE BEFORE
' WHEN CROSS OVER
' I HAVE ABOUT 8 SUPPOSE
' TWO BACK UP BESIDE COMPUTER
' 3 COMPUTER SD-CARD BACKUP ONE EACH AND NETWORK PROVIDE
' AND HARD DRIVE SOME THEM
' THAT AMOUNT ABOUT 8 64-GB FOR ONE OUTDOOR
' AND NOW SLIM DOWN TO TWO WHERE ONLY REQUIRE ONE BLANK
' AND BIFF AR TO MAXIMUM CARD SIZE
' HA HA
' AFRAID A FEW 128-GB
' BUTTY THE RELY GOOD WHEN SPACE STORE
' AND CAMERA BOOT TIME ACCORD TO HOW MUCH ON THERE
' GUESS 64-GB DO END OF DAY
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' Mon 05-Oct-2020 04:50:20
' Mon 05-Oct-2020 16:50:00 -- 12 HOUR 00 MINUTE
' Mon 05-Oct-2020 17:52:00 -- 13 HOUR 00 MINUTE
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' SESSION __ 10
' -------------------------------------------------------------------
' SUB FIND_HASH_TRASH_32_CHAR(FOLDER)
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' ALL-IN-ONE -- EASIER -- RECURSIVE NEST LOOP
' FILE AND SUB-FOLDER --- NICELY PRESENTABLE
' EASY ENTRY EITHER STRING FOLDER NAME OR OBJECT
' NOT EASY FINDER INTERNET PLUS SWEET AND SIMPLE
' -------------------------------------------------------------------
' WED 07-OCT-2020 01:22:36
' WED 07-OCT-2020 04:18:30 -- 2 HOUR 55 MINUTE
' EXTEND WORK
' WED 07-OCT-2020 06:24:00 -- 5 HOUR 00 MINUTE
' -------------------------------------------------------------------
' ANOTHER ROUTINE ADD FOR CAMERA AS PART NEW SYNC ROUTINE STUFF
' SOMETIME -- GOODSYNC WHEN NOT SAFEMODE OR SUPPOSE IS SAFEMODE
' FILE TRASH HASH LEFT BEHIND FROM ABORT COPY
' THE HASH FILE NAME ARE 32 CHAR LONG
' EXAMPLE HERE
' \1d6c1d40497153e04ffd0ea3f9dbb124.JPG
' WE KILL THEM WHEN ALL LOWERCASE AND NUMBER LEN(32)
' -------------------------------------------------------------------
' AND BETTER SUBROUTINE FILE SUB-FOLDER ONE SHOT AND! 
' ONE ROUTINE CODE -- SIMPLICITY TO RUNNER
' NONE FOLDER COMPLICATE GET FROM API -- USE STRING PATH
' EASY LAY DOWN
' -------------------------------------------------------------------
' UNLESS WANT FILE SCAN DOUBLE NEST IN THE SUB FOLDER ROUTINE
' -------------------------------------------------------------------
' SOME PEOPLE USER 3 ROUTINE FOR THE PART HERE
' ONE TO CALL 
' ONE TO COLLECT STRING ARRAY
' ANOTHER TO DEAL WITH THEM
' -------------------------------------------------------------------
' I FIND DELETE WORK WELL IT NOT MESS THE FIND OVER TO DELETE 
' WHILE GO
' -------------------------------------------------------------------
' -------------------------------------------------------------------


' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' WORK TO DO 
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' HERE CODE SET REQUIRE MOVE IN HERE PROJECT
' -------------------------------------------------------------------
' VBS 54-CAMERA MOVE COPY FROM MP_ROOT TO DROP ONE PATH & IMAGER.VBS
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------


DIM MESSENGER_DISPLAY_MODE
MESSENGER_DISPLAY_MODE=FALSE

DIM ORIGINAL_PATH
DIM EXPLORER_COUNTING
DIM LOAD_EXPLORER_HAPPEN

CALL ROUTINE_DELETE_WSCRIPT_FINDER_RUN

SET args = Wscript.Arguments
FOR Each arg In args
	Wscript_Arguments = Wscript_Arguments + arg
NEXT


' -------------------------------------------------------------------
' With CreateObject("Scripting.FileSystemObject")
' MsgBox .GetDrive("\\8-msi-gp62m-7rd\8_msi_gp62m_7rd_82_1_mc_card_camera\").AvailableSpace & " bytes"
' End With
' ----
' check free space on share - VBScript - Tek-Tips 
' https://www.tek-tips.com/viewthread.cfm?qid=1692870
' ----
' WSCRIPT.QUIT
' -------------------------------------------------------------------


DIM MSGBOX_SHOW

' -------------------------------------------------------------------
MSGBOX_SHOW=TRUE
' -------------------------------------------------------------------

IF INSTR(Wscript_Arguments,"RUN_FROM_GOODSYNC")>0 THEN 
	MSGBOX_SHOW=FALSE
END IF

IF INSTR(Wscript_Arguments,"NOT_SHOW")>0 THEN ' ---- "RUN_FROM_GOODSYNC"
	MSGBOX_SHOW=FALSE
END IF

DIM FSO
SET FSO = CreateObject("Scripting.FileSystemObject")




Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

DIM FOLDER_PATH_DEST_VAR_2
DIM FOLDER_PATH_MP4__VAR_2
DIM SET_GO_MAIN_4

DIM GETUSERNAME
DIM GETCOMPUTERNAME
DIM COMPUTERNAME

DIM WSHSHELL
SET WSHSHELL = CreateObject("WScript.Shell")
GETUSERNAME = CreateObject("WScript.Network").USERNAME
GETCOMPUTERNAME = CreateObject("WScript.Network").COMPUTERNAME
COMPUTERNAME = CreateObject("WScript.Network").COMPUTERNAME
' ------------------------------------------------------------------
' ANOTHER EXAMPLE
' ------------------------------------------------------------------
' COMPUTERNAME = WSHSHELL.ExpandEnvironmentStrings("%computername%")
' ------------------------------------------------------------------

DIM JPG_BEEN_DONE
DIM MP4_BEEN_DONE
DIM AVI_BEEN_DONE

DIM STOP_PROGRESS
DIM IX_ERROR_MESSENGER

DIM DRIVE_SPACE_INFO

DIM objStartFolder
DIM OBJFOLDER
DIM I_COUNT_1
DIM I_COUNT_2
DIM VAR_PATH_SCANNER
DIM IMAGE_POINTER_NAME

' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
DIM ARRAY_SONY_FOLDER
DIM DRIVE_LETTER__
DIM RESULT
DIM ARRAY_FOR_NEXT
DIM ARRAY_PATH_01()
DIM ARRAY_PATH_02()
' -------------------------------------------------------------------
REDIM ARRAY_PATH_01(10)
REDIM ARRAY_PATH_02(10)
I=0
' -------------------------------------------------------------------
' LOCAL COMPUTER MC-CARD
' -------------------------------------------------------------------
' I=I+1:ARRAY_PATH_01(I) = "MC - HX60V"
' I=I+1:ARRAY_PATH_01(I) = "MC - HX60V1"
' NETWORK COMPUTER MC-CARD
' -------------------------------------------------------------------
' ALL NETWORK IMPROVER NOT LOCAL
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' MC-CARD NETWORK -- 4G 7G 8M
' -------------------------------------------------------------------
I=I+1:ARRAY_PATH_01(I) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_82_1_MC_CARD_CAMERA" 
I=I+1:ARRAY_PATH_01(I) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_82_2_MC_CARD_CAMERA"
' -------------------------------------------------------------------
I=I+1:ARRAY_PATH_01(I) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW_82_1_MC_CARD_CAMERA"
I=I+1:ARRAY_PATH_01(I) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW_82_2_MC_CARD_CAMERA"
' -------------------------------------------------------------------
I=I+1:ARRAY_PATH_01(I) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_82_1_MC_CARD_CAMERA"
I=I+1:ARRAY_PATH_01(I) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_82_2_MC_CARD_CAMERA"
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' HDD NETWORK     -- 4G 7G 8M
' -------------------------------------------------------------------
I=I+1:ARRAY_PATH_01(I) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_05_1_SAMSUNG_4TB_D\MC - HX60V1_MEDIA_CARD_CAMERA"
I=I+1:ARRAY_PATH_01(I) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW_80_3_SAMSUNG_4TB_D\MC - HX60V1_MEDIA_CARD_CAMERA"
I=I+1:ARRAY_PATH_01(I) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_05_2_SAMSUNG_4TB_D\MC - HX60V1_MEDIA_CARD_CAMERA"
' -------------------------------------------------------------------
' AS HERE -- ABLE INCLUDE HDD BACK WHEN IGNORE OCCUPY SPACE 
' ALL DATE NOW BE EXACT -- FOR DELETE OFF END
' -------------------------------------------------------------------
' CROSS HERE OFF
' HERE IS HARD DRIVE STORE AND NOT REQUIRE SPACE CHECKER
' MORE IN-LINE TALK COMPARE TO SD CARD
' TEMP PUT ON HOLD
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' WHEN MAIN MC-CARD FOR CAMERA IN
' NOT OPERATION COMPUTER TAKEN AWAY OF FOUND.???
' -------------------------------------------------------------------
' CODE IS THERE STOP THAT DO
' -------------------------------------------------------------------
' I=I+1:ARRAY_PATH_01(I) = "MC - HX60V"
' -------------------------------------------------------------------
' NOT GET - HDD SIZE & FREESPACE AT MOMENT SOMETHING CODE WENT LOON
' -------------------------------------------------------------------
' -------------------------------------------------------------------
'  "MC - HX60V"   BACKUP INDOOR MEDIA-CARD 64GB TRIPLE THREE NETWORK COMPUTER ON SHARE DRIVE
'  "MC - HX60V1"  BACKUP INDOOR MEDIA-CARD 64GB TRIPLE THREE NETWORK COMPUTER ON SHARE DRIVE  
'  "MC - HX60V"   ANY USUAL DRIVE LETTER
'  "MC - HX60V1"  ARE __ V:\
' -------------------------------------------------------------------
REDIM PRESERVE ARRAY_PATH_01(I)
REDIM PRESERVE ARRAY_PATH_02(I)
' -------------------------------------------------------------------
' -------------------------------------------------------------------
FOR ARRAY_FOR_NEXT=1 TO UBOUND(ARRAY_PATH_01)
	ARRAY_PATH_02(I) =ARRAY_PATH_01(I)
NEXT
' -------------------------------------------------------------------

DIM DELETE_LIMITITOR
DIM DUMMY_RUN_GET_VALUE_HOW_MANY_DELETE
DIM COUNTER_RR
DIM DEL_SCRIPT_DATE_GONE
DIM I_DELETE_FOUND_000
DIM I_DELETE_EMPTY
DIM I_COUNT_HASH_TRASH_32_CHAR
DIM VOLUME_NAME_STRING_
DIM NOT_DO_OCCUPY_SPACE
DIM DAY_BACK_TRIMMER_COUNTER
DIM DRIVE_FREE_SPACE_GB_BEGIN
DIM DRVPATH_VAR
DIM GET_SPACE_STRING
DIM FREE_SPACE_SET_TO_EQUAL
DIM	DRIVE_FREE_SPACE_GB
DIM	DRIVE_CAPACITY___GB
DIM I_COUNT_DEL
DIM I_COUNT_TOTAL
DIM DAY_BACK_TRIMMER
DIM I_COUNT
DIM MSGBOX_I1 , MSGBOX_I2
DIM ODICT_TRIM_THE_DATE
' -------------------------------------------------------------------
DIM LOW_DATE
DIM HIGH_DATE

''''''''''''''''''''''''''''''''''''''''''''
' I = ARRAY_SONY_FOLDER & VBCR & VBCR & "HASH TRASH FOLDER GONE BY DELETER" & VBCR & VBCR & "NONE CONTENT" & VBCR & VBCR & "TAKE THERE EXPLORER.EXE" & VBCR & VBCR & "MESSENGER POP FEW SECOND"

' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' CALL RUNN_IT_HARDCORE_THREE_ROUTINE_0002
' HERE DO LATER LAST THING
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------




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
		FOLDER_PATH_VAR_2=GET_VOLUME_LETTER__(FOLDER_PATH(R_COUNT))
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
	FOLDER_PATH_VAR_2=GET_VOLUME_LETTER__(FOLDER_PATH(R_COUNT))
	
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
		' IF FOLDER_PATH_VAR_2<>"" THEN 
			' ONLY_SONY_CYBERSHOT_HAS_DO=TRUE
			' RUN_FILE="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 00-DELETE_EMPTY_FOLDER_CAMERA.VBS"
			' IF FSO.FileExists(RUN_FILE) THEN 
				' SET WSHShell = CreateObject("WScript.Shell")
				' RESULT = WSHShell.Run (""""+RUN_FILE+"""", DontShowWindow, WaitUntilFinished)
				' SET WSHShell = Nothing
			' END IF
		' END IF
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
			SET WSHShell = CreateObject("WScript.Shell")
			WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 29 VBS 29-COPY CAMERA LOAD EXPLORER.VBS.AHK"""+ " """+FOLDER_PATH_DEST_VAR_PARAM+"""" , ShowWindow, DontWaitUntilFinished
			SET WSHShell = Nothing
			
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
			DELETE_EMPTY_FOLDER objFolder

			DeleteFolder__=FOLDER_PATH_VAR_2+"MP_ROOT"
			Set objFolder = FSO.GetFolder(DeleteFolder__)
			ORIGINAL_PATH=UCASE(objFolder.PATH)
			DELETE_EMPTY_FOLDER objFolder
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
		' SET WSHShell = CreateObject("WScript.Shell")
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
		SET WSHShell = CreateObject("WScript.Shell")
		WSHShell.Run """"+RENAME_VAR+"""", DontShowWindow, DontWaitUntilFinished
	END IF
END IF

IF INSTR(WORK_DOER,"MEM-BANK-CAMERA")>0 THEN 
	RENAME_VAR="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 27-RENAME FILE+SUBFOLDER_MEMORY-BANK-CAMERA.VBS"
	IF FSO.FileExists(RENAME_VAR) THEN 
		SET WSHShell = CreateObject("WScript.Shell")
		WSHShell.Run """"+RENAME_VAR+"""", DontShowWindow, DontWaitUntilFinished
	END IF
END IF

IF INSTR(WORK_DOER,"FRONT DOOR")>0 THEN
	' ALWAYS SAME
	RENAME_VAR="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 27-RENAME FILE+SUBFOLDER_FRONT_DOOR.VBS"
	PARAM="BLUE_1_FRONTDOOR"  
	IF FSO.FileExists(RENAME_VAR) THEN
		SET WSHShell = CreateObject("WScript.Shell")
		WSHShell.Run """"+RENAME_VAR+"""" + " " + PARAM, DontShowWindow, DontWaitUntilFinished
	END IF
END IF

IF INSTR(WORK_DOER,"OUTDOOR JACKET")>0 THEN
	' ALWAYS SAME
	RENAME_VAR="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 27-RENAME FILE+SUBFOLDER_FRONT_DOOR.VBS"
	PARAM="OUTDOOR_JACKET"
	IF FSO.FileExists(RENAME_VAR) THEN
		SET WSHShell = CreateObject("WScript.Shell")
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
		SET WSHShell = CreateObject("WScript.Shell")
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
		' SET WSHShell = CreateObject("WScript.Shell")
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

	' ---------------------------------------------------------------
	' HERE A BIT LATER
	' ---------------------------------------------------------------
	' RUNN_IT_HARDCORE_THREE_ROUTINE:
	' ---------------------------------------------------------------
	
END IF


INF=""
FOR R_COUNT=1 TO UBOUND(FOLDER_PATH)
	IF INSTR(INF,FOLDER_PATH(R_COUNT)+"  ")=0 THEN 
		INF=INF+FOLDER_PATH(R_COUNT)+"  " + VBCRLF
		IX=IX+1
	END IF
NEXT

' IF WORK_GO_ER=TRUE OR 1=1 THEN
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
	
	' CALL GET_DRIVE_SPACE(DRVPATH)
	
	I=I+ DRIVE_SPACE_INFO
	I=I+ VBCRLF
	I=I+ VBCRLF
	
	I=I+"CHECKER VOLUME NAME OF -- "+CSTR(IX)+" TOTAL DUPE FILTER" + VBCRLF
	I=I+"--------------------------------------------------------" + VBCRLF
	I=I+INF + VBCRLF
	I=I+"--------------------------------------------------------" + VBCRLF

	IF MSGBOX_SHOW=TRUE THEN 
		' -----------------------------------------------------------
		' -----------------------------------------------------------
		IF I_COUNT_1>0 THEN 
			I=I+ VBCRLF
			I=I+ "SOME COPY GET COUNT: " + CSTR(I_COUNT_1) + VBCRLF + VBCRLF
			I=I+ "FROM DATE -- INDEX POINTER -- " + TT_DATE_POINTER + VBCRLF + VBCRLF
			I=I+ "WORK DRIVE VOLUME NAME ------ " + HDD_WAS_HERE_2 + VBCRLF + VBCRLF
			I = I + "-------------------------------" & VBCRLF
			I = I + "MESSENGER BOX GONE OF 10 SECOND"

			IF TRUE=TRUE OR MESSENGER_DISPLAY_MODE=TRUE THEN 
			SET WSHShell = CreateObject("WScript.Shell")
			RESULT = WSHSHELL.POPUP(I, 10, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
			SET WSHShell = Nothing
			END IF
		END IF
		' -----------------------------------------------------------
		' -----------------------------------------------------------
		IF I_COUNT_1=0 THEN 
			I=I+ VBCRLF
			I=I+ "NOTHING DO COPY COUNT: " + CSTR(I_COUNT_1) + VBCRLF + VBCRLF
			I=I+ "FROM DATE -- INDEX POINTER -- " + TT_DATE_POINTER + VBCRLF + VBCRLF
			I=I+ "WORK DRIVE VOLUME NAME ------ " + HDD_WAS_HERE_2 + VBCRLF + VBCRLF
			
			I=I+ "IF WANT REWIND BACK -- INDEX POINTER -- A DAY FOR OVERSHOOT" + VBCRLF
			I=I+ "OPERATION OR SKIPPER " & VBCRLF
			' lRet = MsgBox(I, vbYesNo + vbMsgBoxSetForeground,Wscript.ScriptName)
			
			I = I + "-------------------------------" & VBCRLF
			I = I + "MESSENGER BOX GONE OF 10 SECOND"
			IF TRUE=TRUE OR MESSENGER_DISPLAY_MODE=TRUE THEN 
			SET WSHShell = CreateObject("WScript.Shell")
			RESULT = WSHSHELL.POPUP(I, 10, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
			SET WSHShell = Nothing
			END IF
			' -------------------------------------------------------
			' REQUIRE NOT HERE VBOKONLY
			' -------------------------------------------------------
			lRet=RESULT
			IF lRet=VBYES THEN
				IMAGE_POINTER_FOLDER_1=MID(Wscript.ScriptFullName,1,INSTRREV(Wscript.ScriptFullName,"\"))+"VBS 29-COPY CAMERA\"
				IMAGE_POINTER_NAME=IMAGE_POINTER_FOLDER_1+"COPY_CAMERA_PHOTO_IMAGE_POINTER_"+HDD_WAS_HERE_2+".TXT"
				IMAGE_POINTER_NAME=REPLACE(IMAGE_POINTER_NAME,"\SCRIPTER\","\SCRIPTOR DATA\")
					
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


' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
CALL RUNN_IT_HARDCORE_THREE_ROUTINE_0002
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------


SET objFSO = NOTHING

I_1=""
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"THE ENDER" & VBCR
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"----------------------------------------" & VBCR
I_1=I_1+"----------------------------------------" & VBCR

IF TRUE=TRUE OR MESSENGER_DISPLAY_MODE=TRUE THEN 
	SET WSHShell = CreateObject("WScript.Shell")
	RESULT = WSHSHELL.POPUP(I_1, 8, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
	SET WSHShell = Nothing
END IF



'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------

'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------

'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------

'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
Wscript.QUIT
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------

'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------

'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------

'-------------------------------------------------------------------
'-------------------------------------------------------------------
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
    FindWantedFiles(FSO.GetFolder(FOLDER_PATH_VAR_2))
    ' Display All Findings from Dictionary object
	
	FOLDER_PATH_DEST_VAR_2=FOLDER_PATH_DEST(R_COUNT)
	FOLDER_PATH_MP4__VAR_2=FOLDER_PATH__MP4(R_COUNT)
	' CALL SET_VAR_DESTINATION_FOLDER
	
    DisplayFindings

	' DRVPATH_VAR=MID(ARRAY_PATH_01,1,2)
	' ' WHEN NETWORK PATH GIVE FULL PATH
	' IF MID(ARRAY_PATH_01,1,2)="\\" THEN 
		' DRVPATH_VAR=ARRAY_PATH_01
	' END IF


	IF MID(FOLDER_PATH_VAR_2,1,2)<>"D:" THEN 
		CALL GET_DRIVE_SPACE(MID(FOLDER_PATH_VAR_2,1,2))
	END IF
	
	' ---------------------------------------------------------------
	' CHECK EXPLORER BEEN LOADED BEFORE _ IF NOT
	' NOW IS TIME TO MAKE SURE _ AND DECIDE IF THE COUNTER HAS GONE UP FOR ANY COPY FILE
	' WHILE STILL MAINTAIN NOT FIR ANY DODGY COMPUTER OR FOLDER
	' INCLUDE D DRIVE OR BE SHOW EVERY TIME
	' ---------------------------------------------------------------
	
	IF SHOW_EXPLORER_FLAG=TRUE AND EXPLORER_LOADED=FALSE AND I_COUNT_1>0 THEN 
		' SET WSHShell = CreateObject("WScript.Shell")
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
				' SET WSHShell = CreateObject("WScript.Shell")
				' WSHShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 33-#0 Send To Date Of File In Front FileName.ahk""", ShowWindow, DontWaitUntilFinished
				
				' NEW WAY HERE
				CALL DATE_OF_FILE_IN_FRONT_FILENAME_SONY_CAMERA
				
				
			END IF
		END IF
	END IF

	IF VBS_54_CAMERA_MOVE_COPY_FROM_MP_ROOT_TO_DROP_ONE_PATH_IMAGER=FALSE THEN
		IF I_COUNT_1>=0 AND (MP4_BEEN_DONE=TRUE) THEN
			VBS_54_CAMERA_MOVE_COPY_FROM_MP_ROOT_TO_DROP_ONE_PATH_IMAGER=TRUE
			SET WSHShell = CreateObject("WScript.Shell")
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
				'SET WSHShell = CreateObject("WScript.Shell")
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
			MSGBOX "LINE 1714 __ LINER __ PAD __ SOMETHING SUPPOSE TO BE HERE"& VBCR & VBCR & "SUB DATE_OF_FILE_IN_FRONT_FILENAME_SONY_CAMERA"
		END IF
		IF FSO.FolderExists(FOLDER_PATH__MP4(R_CNT))=TRUE THEN  
			MSGBOX "LINE 1717 __ LINER __ PAD __ SOMETHING SUPPOSE TO BE HERE"& VBCR & VBCR & "SUB DATE_OF_FILE_IN_FRONT_FILENAME_SONY_CAMERA"
		END IF
	NEXT
End Sub


Sub FindWantedFiles(sFolder)
    On Error Resume NEXT
    Dim oFDR, oItem, file
    ' Check if wanted files are in this folder
    For Each file In sFolder.Files
		oDict.Add FILE.path, file
    NEXT
    ' Recurse into it's sub folders
    For Each oFDR In sFolder.SubFolders
        FindWantedFiles FSO.GetFolder(oFDR.Path)
    NEXT
End Sub

Sub DisplayFindings()
    Dim oKeys, oKey
    oKeys = oDict.Keys
    For Each oKey In oKeys
		VAR_PATH_SCANNER=oKey
		IF FSO.FILEExists(VAR_PATH_SCANNER)=TRUE THEN 
			CALL PATH_SCANNER_COPY_CAMERA
		END IF
	NEXT
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
			i_TIME=i_TIME+LPAD(Year(FileModDate_1), "0", 4)+"_"
			i_TIME=i_TIME+LPAD(Month(FileModDate_1), "0", 2)+"_"
			i_TIME=i_TIME+LPAD(Day(FileModDate_1), "0", 2)+" "
			i_TIME=i_TIME+MonthName(MONTH(FileModDate_1),TRUE)+"_"
			i_TIME=i_TIME+WeekdayName(Weekday(FileModDate_1),TRUE)+" "
			i_TIME=i_TIME+""
			i_TIME=i_TIME+LPAD(Hour(FileModDate_1), "0", 2)+"_"
			i_TIME=i_TIME+LPAD(Minute(FileModDate_1), "0", 2)+"_"
			i_TIME=i_TIME+LPAD(Second(FileModDate_1), "0", 2)+"__"
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
			i_TIME=i_TIME+LPAD(Year(FileModDate_1), "0", 4)+"-"
			i_TIME=i_TIME+LPAD(Month(FileModDate_1), "0", 2)+"-"
			i_TIME=i_TIME+LPAD(Day(FileModDate_1), "0", 2)+" "
			FOLDER_DATE=""
			FOLDER_DATE=FOLDER_DATE+LPAD(Year(FileModDate_1), "0", 4)+" "
			FOLDER_DATE=FOLDER_DATE+LPAD(Month(FileModDate_1), "0", 2)+" "
			FOLDER_DATE=FOLDER_DATE+LPAD(Day(FileModDate_1), "0", 2)
			
			i_TIME=i_TIME+""
			i_TIME=i_TIME+LPAD(Hour(FileModDate_1), "0", 2)+"-"
			i_TIME=i_TIME+LPAD(Minute(FileModDate_1), "0", 2)+"-"
			i_TIME=i_TIME+LPAD(Second(FileModDate_1), "0", 2)+" "
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
				SET WSHShell = CreateObject("WScript.Shell")
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
				SET WSHShell = CreateObject("WScript.Shell")
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


SUB GET_DRIVE_SPACE(DRVPATH)

	' ---------------------------------------------------------------
	IF DRVPATH="" THEN 
		MSGBOX "DRVPATH = NOTHING " & VBCR & UCASE(TYPENAME(DRVPATH))
		EXIT SUB
	END IF
	' ---------------------------------------------------------------
	IF UCASE(TYPENAME(DRVPATH))="FOLDER" THEN 
		' PERFECT -- FOLDER OBJECT 
	END IF
	' ---------------------------------------------------------------
	IF UCASE(TYPENAME(DRVPATH))="STRING" THEN
		IF FSO.FOLDEREXISTS(DRVPATH)=TRUE THEN 
			Set DRVPATH = FSO.GETFOLDER(DRVPATH)
			GET_DRIVE_SPACE(DRVPATH)
		ELSE
			VALUE_WAIT=4
			IF TRUE=TRUE OR MESSENGER_DISPLAY_MODE=TRUE THEN
				I = "ERROR DRIVE NOT READY " & VBCR & DRVPATH & VBCR
				I=I+ "MESSENGER BOX DISAPPEAR ON " & CSTR(VALUE_WAIT) & " SECOND"
				DIM WSHShell
				SET WSHShell = CreateObject("WScript.Shell")
				RESULT = WSHSHELL.POPUP(I, VALUE_WAIT, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
				SET WSHShell = Nothing
			END IF
		END IF
	END IF
	' ---------------------------------------------------------------
	IF FSO.FOLDEREXISTS(DRVPATH)=FALSE THEN 
	IF UCASE(TYPENAME(DRVPATH))="STRING" THEN
		' -----------------------------------------------------------
		' LAST TIME HERE 
		' 13-OCT-2020 20:34:48 TUE
		' -----------------------------------------------------------
		' MSGBOX "8888 IF TRIM(DRVPATH)= THEN -- EXIT SUB" 
		EXIT SUB
	END IF 
	END IF 
	IF TRIM(DRVPATH)="" THEN 
		MSGBOX "5555 IF TRIM(DRVPATH)= THEN -- EXIT SUB" 
		EXIT SUB
	END IF
	
	IF UCASE(TYPENAME(DRVPATH))="FOLDER" THEN 
		IF FSO.FOLDEREXISTS(DRVPATH.PATH)=FALSE THEN 
			' THE SD-CARD WORK OKAY SEEM TO -- SOMETIME ASSESS
			' MSGBOX "4444 IF TRIM(DRVPATH)= THEN -- EXIT SUB" 
			' HERE K DRIVE SD-CARD OKAY BUTT NOT POPUP
			EXIT SUB
		END IF
	END IF

	ON ERROR RESUME NEXT
	ERR.CLEAR
	SET D = FSO.GETDRIVE(DRVPATH)
	IF ERR.NUMBER>0 THEN 
	IF FSO.FOLDEREXISTS(DRVPATH)=FALSE THEN 
		' ERROR AT LINE HERE -- SET D = FSO.GETDRIVE(DRVPATH)
		' Line:	2194
		' Char:	2
		' Error:	Invalid procedure call or argument
		' Code:	800A0005
		' Source: 	Microsoft VBScript runtime error
		EXIT SUB
	END IF 
	END IF 
	IF FSO.FOLDEREXISTS(DRVPATH)=FALSE THEN 
		EXIT SUB
	END IF 

	' ---------------------------------------------------------------
	DRIVE_FREE_SPACE_GB=CDbl(D.FREESPACE)/1024/1024/1024
	DRIVE_CAPACITY___GB=CDbl(D.TOTALSIZE)/1024/1024/1024
	' ---------------------------------------------------------------
	IF FREE_SPACE_SET_TO_EQUAL=0 THEN DRIVE_FREE_SPACE_GB_BEGIN = DRIVE_FREE_SPACE_GB
	FREE_SPACE_SET_TO_EQUAL=20 ' GB
	
	IF D.ISREADY = TRUE THEN
		DRIVE_IS_READY_____="DRIVE IS READY __ TRUE"
		VOLUME_NAME_STRING_=D.VOLUMENAME
	ELSE
		DRIVE_IS_READY_____="DRIVE IS READY __ FALSE"
		VOLUME_NAME_STRING_=D.VOLUMENAME
	END IF

	IF D.ISREADY = TRUE THEN
		S=S+"-----------------------------------------------------------" & VBCRLF
		S=S+"VOLUMENAME:" & VBTAB & VOLUME_NAME_STRING  & VBCRLF
		S=S+"-----------------------------------------------------------" & VBCRLF
		S=S+"DRIVE LETTER 1ST FOUND FOR MATCH VOLUME NAME: " & VBCRLF & DRVPATH & "\" & VBCRLF
		S=S+"-----------------------------------------------------------" & VBCRLF
		S=S+"FREESPACE:" & VBTAB
		S=S+FORMATNUMBER((CDBL(D.FREESPACE)/1024/1024/1024))  & " GB" & VBCRLF
		S=S+"SIZE:" & VBTAB & VBTAB
		S=S+FORMATNUMBER((CDBL(D.TOTALSIZE)/1024/1024/1024)) & " GB" & VBCRLF
		S=S+"OCCUPIED SPACE:" & VBTAB
		S=S+FORMATNUMBER((CDBL(D.TOTALSIZE - D.FREESPACE)/1024/1024/1024)) & " GB"
		IF IX_ERROR_MESSENGER<>"" THEN 
			S=S+VBCRLF
			S=IX_ERROR_MESSENGER
		END IF
	END IF
	' ---------------------------------------------------------------
	DRIVE_SPACE_INFO=S
	' ---------------------------------------------------------------
	
	' ---------------------------------------------------------------
	' EXAMPLE ON -- FORMATNUMBER
	' FORMATNUMBER(DRIVE_FREE_SPACE_GB,2) GIVE 2 DECIMAL PLACE 00.00
	' ---------------------------------------------------------------
	' MSGBOX FORMATNUMBER(DRIVE_FREE_SPACE_GB,2) & " ---- " &  
	' MSGBOX FORMATNUMBER(FREE_SPACE_SET_TO_EQUAL,2)  & " -- GB"
	' ---------------------------------------------------------------
END SUB


Sub BuildFullPath(ByVal FullPath)
	If Not fso.FolderExists(FullPath) Then
		ON ERROR RESUME NEXT
		BuildFullPath fso.GetParentFolderName(FullPath)
		fso.CreateFolder FullPath
		ON ERROR GOTO 0
	End If
End Sub



Function GET_VOLUME_LETTER__(VOLUMENAME)
	
	' ---------------------------------------------------------------
	' RETURN -- NONE
	' RETURN -- LETTER OF DRIVE
	' RETURN -- NETWORK DRIVE IS GOOD
	' ---------------------------------------------------------------
	
	GET_VOLUME_LETTER__=""
	IF MID(VOLUMENAME,1,2)="\\" THEN 
		' NETWORK DRIVE CHECK RESULT GOOD
		' -------------------------------
		IF FSO.DRIVEEXISTS(VOLUMENAME) = TRUE THEN 
			' DRIVE_LETTER__
			GET_VOLUME_LETTER__="DRIVE EXIST GOOD NETWORK TYPE"
			EXIT FUNCTION
		END IF
	ELSE
		FOR R1 = ASC("C") TO ASC("Z")
			Set FS0 = CreateObject("Scripting.FileSystemObject")
			DRIVE_LETTER__=CHR(R1)+":"
			Set D = FS0.GetDrive(FSO.GetDriveName(FSO.GetAbsolutePathName(DRIVE_LETTER__)))
			ON ERROR RESUME NEXT
			ERR.CLEAR
			XX=D.ISREADY
			IF XX=TRUE AND ERR.NUMBER=0 THEN 
				IF FSO.DRIVEEXISTS(DRIVE_LETTER__) = TRUE THEN 
					IF UCASE(D.VOLUMENAME)=UCASE(VOLUMENAME) THEN
						' -------------------------------------------------------
						GET_VOLUME_LETTER__=UCASE(DRIVE_LETTER__+"\")
						EXIT FUNCTION
						' -------------------------------------------------------
					END IF
				END IF
			END IF
		NEXT
	END IF
END FUNCTION




Function GET_VOLUME_LETTER___OLD_MESSY(volumeName)
Dim volumes, volume

    ' UNLESS WE FOUND A MATCHING VOLUME, AN EMPTY STRING WILL BE THE RETURNED VALUE
    GET_VOLUME_LETTER__=""

    ' ASK WMI FOR THE LIST OF VOLUMES WITH THE REQUESTED LABEL
    Set volumes = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") _ 
                  .ExecQuery("SELECT DriveLetter FROM Win32_Volume WHERE Label='" & volumeName & "'")

	ON ERROR RESUME NEXT
	ERR.CLEAR
	X = VOLUMES.COUNT
	IF ERR.NUMBER>0 THEN 
	EXIT FUNCTION
	END IF
	ON ERROR GOTO 0

    ' IF EXIST AN MATCHING VOLUME, GET ITS DRIVE LETTER
	IF VOLUMES.COUNT > 0 THEN 
        For Each volume In volumes 
            GET_VOLUME_LETTER__ = volume.DriveLetter
            Exit For
        NEXT 
    End If

End Function




SUB ROUTINE_DELETE_WSCRIPT_FINDER_RUN
' -------------------------------------------------------------------
' 01 OF 02 
' SOME CONFLICT DISAGREE-ER IN CODE STYLE 
' SOME PEOPLE TALK THIS HAS ERROR 
' THIS ONE IS CORRECT 
' THIS WAS CORRECT 
' THING ARE EXPLAIN HERE
' MSHTA.EXE DELETE ITSELF TOO QUICK FOR RESULT FIND
' NOW WITH BEST EXAMPLE ALSO FIND HIDDEN PROCESS TOP SPEED
' -------------------------------------------------------------------
' PID - FIND MY OWN PROCESS ID IN VBSCRIPT - STACK OVERFLOW 
' https://stackoverflow.com/questions/8296037/find-my-own-process-id-in-vbscript
' -------------------------------------------------------------------
On Error Resume Next
Dim iMyPID
' iMyPID = GetObject("winmgmts:root\cimv2").Get("Win32_Process.Handle='" & CreateObject("WScript.Shell").Exec("mshta.exe").ProcessID & "'").ParentProcessId
iMyPID=ThisProcessId
On Error Goto 0
' -------------------------------------------------------------------
' GOOD CODE TO PASTE ON VB6
' -------------------------------------------------------------------
DIM WSHSHELL
SET WSHSHELL = CreateObject("WScript.Shell")

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")
' -------------------------------------------------------------------
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
		
		IF iMyPID<>PID_Script THEN 
			' MSGBOX iMyPID & VBCR & PID_Script
			objProcess.Terminate() 
			' WSHSHELL.Run "TASKKILL /F /IM """+CSTR(PID_Script)+"""", DONTShowWindow, WaitUntilFinished
		END IF
    END IF
NEXT
' -------------------------------------------------------------------
END SUB
' -------------------------------------------------------------------
' THING ARE EXPLAIN HERE
' MSHTA.EXE DELETE ITSELF TOO QUICK FOR RESULT FIND
' NOW WITH BEST EXAMPLE ALSO FIND HIDDEN PROCESS TOP SPEED
' -------------------------------------------------------------------
' PID - FIND MY OWN PROCESS ID IN VBSCRIPT - STACK OVERFLOW 
' https://stackoverflow.com/questions/8296037/find-my-own-process-id-in-vbscript
' -------------------------------------------------------------------
' I like Kul-Tigin's idea (+1), and Asok Smith's idea (based on .Exec) deserve respect (+1), 
' and it w'd been even better if .Exec run hidden process. So, to feed my curiosity, 
' I also toyed with this and this's what I did.
' -------------------------------------------------------------------
Function ThisProcessId
    ThisProcessId = 0
    Dim sTFile, oPrc
    With CreateObject("Scripting.FileSystemObject")
        sTFile = .BuildPath(.GetSpecialFolder(2), "sleep.vbs")
        With .OpenTextFile(sTFile, 2, True)
            .Write "WScript.Sleep 1000"
        End With
    End With
    With CreateObject("WScript.Shell").Exec("WScript " & sTFile)
        For Each oPrc In GetObject("winmgmts:\\.\root\cimv2").ExecQuery(_
        "Select * From Win32_Process Where ProcessId=" & .ProcessID)
        Exit For : Next
        ThisProcessId = oPrc.ParentProcessId
    End With
End Function







SUB RUNN_IT_HARDCORE_THREE_ROUTINE_0002
	' -------------------------------------------------------------------
	' DIM DAY_BACK_TRIMMER_COUNTER ' =102 ---- HERE LATER ON -- DAY COUNT BACK
	' -------------------------------------------------------------------
	' SET ABOVE FOR CODE ROUTINE 
	' -------------------------------------------------------------------
	' -------------------------------------------------------------------
	' RUNN_IT_HARDCORE_THREE ROUTINE GIVER GIVE BACK SPACE _ REMOVE \_FOUND_000 KILL EMPTY
	' CALL RUNN_IT_HARDCORE_THREE_ROUTINE:
	' -------------------------------------------------------------------
	' 0001 
	' START BEFORE LINE  380
	' 0002
	' AHEAD LINE 2800 AT END
	' -------------------------------------------------------------------
	' -------------------------------------------------------------------
	' RUN ROUTINE END NUMBER 03 AND RUN IT FOUR THAT POWER
	' -------------------------------------------------------------------
	' 0001 OF 0004 -------------------------------------
	' CALL DELETE_FOUND_000_ROOT_PATH(ARRAY_SONY_FOLDER)
	' 0002 OF 0004 -------------------------------------
	' CALL DELETE_EMPTY_FOLDER(ARRAY_SONY_FOLDER)
	' 0003 OF 0004 -------------------------------------
	' CALL STRESSOR__TRIM_THE_DATE_OLDER_CAMERA
	' 0004 OF 0004 -------------------------------------
	' LINE 2800 AT END
	' CALL RUNN_IT_HARDCORE_THREE_ROUTINE
	' -------------------------------------------------------------------

	FOR ARRAY_FOR_NEXT=1 TO UBOUND(ARRAY_PATH_01)

	DELETE_LIMITITOR=0
	DUMMY_RUN_GET_VALUE_HOW_MANY_DELETE=0
	COUNTER_RR=0
	DEL_SCRIPT_DATE_GONE=NULL
	I_DELETE_FOUND_000=0
	I_DELETE_EMPTY=0
	I_COUNT_HASH_TRASH_32_CHAR=NULL
	VOLUME_NAME_STRING_=NULL
	NOT_DO_OCCUPY_SPACE=0
	DAY_BACK_TRIMMER_COUNTER=0
	DRIVE_FREE_SPACE_GB_BEGIN=0
	DRVPATH_VAR=NULL
	GET_SPACE_STRING=NULL
	FREE_SPACE_SET_TO_EQUAL=0
	DRIVE_FREE_SPACE_GB=0
	DRIVE_CAPACITY___GB=0
	I_COUNT_DEL=0
	I_COUNT_TOTAL=0
	DAY_BACK_TRIMMER=0
	I_COUNT=0
	MSGBOX_I1=NULL
	MSGBOX_I2=NULL
	ODICT_TRIM_THE_DATE=0
	' -------------------------------------------------------------------
	' REDIM ODICT_TRIM_THE_DATE
	' -------------------------------------------------------------------
	LOW_DATE=NULL
	HIGH_DATE=NULL
	' -------------------------------------------------------------------

	' -------------------------------------------------------------------
	' REDIM ODICT_TRIM_THE_DATE
	' -------------------------------------------------------------------
	LOW_DATE=NOW
	HIGH_DATE=NOW-1000
	' -------------------------------------------------------------------
	I=GET_VOLUME_LETTER__(ARRAY_PATH_01(ARRAY_FOR_NEXT))
	IF MID(I,2,1)=":" THEN 
		ARRAY_PATH_01(ARRAY_FOR_NEXT)=I
	END IF
	ARRAY_SONY_FOLDER=ARRAY_PATH_01(ARRAY_FOR_NEXT)

	' ---------------------------------------------------
	' DO LATER AT RUN 
	' TEST RUN
	' ---------------------------------------------------
	CALL STRESSOR__TRIM_THE_DATE_OLDER_CAMERA
	' ---------------------------------------------------

	' ---------------------------------------------------
	' WSCRIPT.QUIT
	' ---------------------------------------------------
 	
	I_COUNT_HASH_TRASH_32_CHAR=0
	CALL FIND_HASH_TRASH_32_CHAR(ARRAY_SONY_FOLDER)
	
	IF I_COUNT_HASH_TRASH_32_CHAR>0 THEN 
		FOLDER_JUMP_TO_HASH_TRASH_EXPLORER = ARRAY_SONY_FOLDER & "\TEMP_HASH_TRASH"
		' COUNTER EXPLAIN GOT HASH TRASH BUTTY FOLDER NOT THERE TRY ROOT\
		' ---------------------------------------------------------------
		IF FSO.FOLDEREXISTS(ARRAY_SONY_FOLDER & "\TEMP_HASH_TRASH")=FALSE  THEN
			I=""
			I=I+"SOME HASH TRASH COUNTER AND FOLDER NOT THERE - ERROR" & VBCR 
			I=I+ VBCR 
			I=I+"ARRAY_SONY_FOLDER" & "\TEMP_HASH_TRASH"
			I=I+ VBCR 
			I=I+ "CHANGE PATH TO \ AND TAKE THERE EXPLORER"
			I=I+ VBCR 
			I=I+ ARRAY_SONY_FOLDER
			I=I+ VBCR 

			FOLDER_JUMP_TO_HASH_TRASH_EXPLORER = ARRAY_SONY_FOLDER
			
			IF MESSENGER_DISPLAY_MODE=TRUE THEN 
			SET WSHShell = CreateObject("WScript.Shell")
				RESULT = WSHSHELL.POPUP(I, 15, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
			SET WSHShell = Nothing
			END IF
		END IF

		' FOLDER HASH TRASH STILL NOT THERE AGAIN LOCATE FOLDER OR ROOT\
		' TUFF ONE
		' --------------------------------------------------------------
		IF FSO.FOLDEREXISTS(FOLDER_JUMP_TO_HASH_TRASH_EXPLORER) = FALSE THEN 
			I="THE FOLDER PATH NOT THERE OR ACCESS " & VBCR 
			I=I+ VBCR 
			I=I+ FOLDER_JUMP_TO_HASH_TRASH_EXPLORER
			I=I+ VBCR 
			I=I+ VBCR 
			I=I+ "GET TO HASH TRASH THE COUNT WAS SHOW"
			I=I+ VBCR 

			IF MESSENGER_DISPLAY_MODE=TRUE THEN 
			SET WSHShell = CreateObject("WScript.Shell")
				RESULT = WSHSHELL.POPUP(I, 15, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
			SET WSHShell = Nothing
			END IF
		END IF

		' MAIN FOLDER HAS TRASH IS THERE AND EMPTY -- DELETER
		' -----------------------------------------------------------
		IF FSO.FOLDEREXISTS(ARRAY_SONY_FOLDER & "\TEMP_HASH_TRASH")=TRUE  THEN
			SET F = FSO.GETFOLDER(ARRAY_SONY_FOLDER & "\TEMP_HASH_TRASH")
			IF F.SIZE=0 THEN
				' FOLDER NOT EXIST DELETER
				FSO.DELETEFOLDER ARRAY_SONY_FOLDER & "\TEMP_HASH_TRASH", TRUE
					
				I = ARRAY_SONY_FOLDER & VBCR & VBCR & "HASH TRASH FOLDER GONE BY DELETER" & VBCR & VBCR & "NONE CONTENT" & VBCR & VBCR & "TAKE THERE EXPLORER.EXE" & VBCR & VBCR & "MESSENGER POP FEW SECOND"
				IF MESSENGER_DISPLAY_MODE=TRUE THEN 
				SET WSHShell = CreateObject("WScript.Shell")
					RESULT = WSHSHELL.POPUP(I, 8, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
				SET WSHShell = Nothing
				END IF
			END IF
		END IF



		IF FSO.FOLDEREXISTS(FOLDER_JUMP_TO_HASH_TRASH_EXPLORER) = TRUE THEN 
			' -----------------------------------------------------------
			' RUN EXPLORER
			' -----------------------------------------------------------
			SET WSHShell = CreateObject("WScript.Shell")
			' MSGBOX "EXPLORER """+ARRAY_SONY_FOLDER & "\TEMP_HASH_TRASH\"+""""
			RESULT = WSHShell.Run("EXPLORER """+ARRAY_SONY_FOLDER & "\TEMP_HASH_TRASH\"+"""" , ShowWindow, DontWaitUntilFinished)
			SET WSHShell = Nothing
			' -----------------------------------------------------------
			' INFO MSGBOX
			' -----------------------------------------------------------
			I = ARRAY_SONY_FOLDER & VBCR & VBCR & "FIND SOME LOSS HASH TRASH" & VBCR & VBCR & "TAKE THERE EXPLORER.EXE" & VBCR & VBCR & "MESSENGER POP FEW SECOND"

			IF MESSENGER_DISPLAY_MODE=TRUE THEN 
			SET WSHShell = CreateObject("WScript.Shell")
				RESULT = WSHSHELL.POPUP(I, 8, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
			SET WSHShell = Nothing
			END IF
		END IF
	END IF

	I_DELETE_FOUND_000=0
	IF ARRAY_PATH_02(ARRAY_FOR_NEXT) <> "MC - HX60V" THEN 
	' ---------------------------------------------------------------
		CALL DELETE_FOUND_000_ROOT_PATH(ARRAY_SONY_FOLDER)
	' ---------------------------------------------------------------
	END IF

	IF I_DELETE_FOUND_000>0 THEN 
		I = ARRAY_SONY_FOLDER & VBCR & VBCR & "FIND TO DELETE SOME FOUND_000"
		IF MESSENGER_DISPLAY_MODE=TRUE THEN 
		SET WSHShell = CreateObject("WScript.Shell")
			RESULT = WSHSHELL.POPUP(I, 4, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
		SET WSHShell = Nothing
		END IF 
	END IF 
	
	' ---------------------------------------------------------------
	I_DELETE_EMPTY_COUNTER=0
	CALL DELETE_EMPTY_FOLDER(ARRAY_SONY_FOLDER)
	' ---------------------------------------------------------------
	IF I_DELETE_EMPTY_COUNTER>0 THEN 
		' -----------------------------------------------------------
		' SOME DELETE EMPTY
		' -----------------------------------------------------------
		I = ARRAY_SONY_FOLDER & VBCR & VBCR & "SOME DELETE __ EMPTY __ WERE DONE"
		IF MESSENGER_DISPLAY_MODE=TRUE THEN 
		SET WSHShell = CreateObject("WScript.Shell")
			RESULT = WSHSHELL.POPUP(I, 4, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
		SET WSHShell = Nothing
		END IF
	END IF
	' ---------------------------------------------------------------
	' ---------------------------------------------------
	' DO LATER AT RUN 
	' ---------------------------------------------------
	' CALL STRESSOR__TRIM_THE_DATE_OLDER_CAMERA
	' ---------------------------------------------------
NEXT
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------


' -----------------------------------------------------------
' DO THE CAMERA SYNCER NOW
' -----------------------------------------------------------
IF MESSENGER_DISPLAY_MODE=TRUE THEN 
I = ARRAY_SONY_FOLDER & VBCR & VBCR & "DO THE CAMERA SYNCER NOW"
SET WSHShell = CreateObject("WScript.Shell")
	RESULT = WSHSHELL.POPUP(I, 4, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
SET WSHShell = Nothing
END IF


' ---------------------------------
' NOT ABLE RELOCATE THIS PART 
' RESULT COME THROUGH BLANK
' RUNNER AT FRONT -- WHILE
' ---------------------------------
' CALL RUNN_IT_HARDCORE_THREE_ROUTINE
' CALL RUNN_IT_HARDCORE_THREE_ROUTINE_0002

END SUB


' -------------------------------------------------------------------
' -------------------------------------------------------------------
Sub STRESSOR__TRIM_THE_DATE_OLDER_CAMERA
	' ---------------------------------------------------------------
	' VERIFY WHAT DELETE BEFORE -- HOW MANY
	' ---------------------------------------------------------------
	' 0002 OF 0002
	' ---------------------------------------------------------------
	I_COUNT=0
	I_COUNT_DEL=0
	I_COUNT_TOTAL=0
	MSGBOX_I1=""
	MSGBOX_I2=""
	FREE_SPACE_SET_TO_EQUAL=20
	LOW_DATE=NOW
	HIGH_DATE=NOW-1000
	CALL GET_DRIVE_SPACE(ARRAY_SONY_FOLDER) ' BEFORE NUMBER
	' ---------------------------------------------------------------
	DUMMY_RUN_GET_VALUE_HOW_MANY_DELETE="GET_FIND_INFO"
	CALL SCAN__TRIM_THE_DATE_OLDER_CAMERA(ARRAY_SONY_FOLDER)
	' ---------------------------------------------------------------
	CALL GET_DRIVE_SPACE(ARRAY_SONY_FOLDER) ' AFTER
	' ---------------------------------------------------------------
	IF I_COUNT_DEL>0 THEN 
		I=""
		I=I+ "FOLDER -- " & VBCR 
		I=I+ ARRAY_SONY_FOLDER & VBCR 
		I=I+ "FILE TO DELETE" & VBCR 
		I=I+ CSTR(I_COUNT_DEL) & VBCR 
		I=I+ "----" & VBCR 
		I=I+ "HIGH_DATE  " & HIGH_DATE & VBCR 
		I=I+ "LOW_DATE   " & LOW_DATE & VBCR 
		I=I+ "----" & VBCR 
		I=I+ "NEWER DATE BEFORE _____ " & VBCR
		I=I+ MSGBOX_I2 & VBCR 
		I=I+ "----" & VBCR 
		I=I+ "OLDER DATE OF _________ " & VBCR 
		I=I+ MSGBOX_I1 & VBCR 
		I=I+ "----" & VBCR 
		I=I+ "DELETE IS LIMIT TO __ " & CSTR(DELETE_LIMITITOR) 
		IF I_COUNT_DEL<DELETE_LIMITITOR THEN 
			I=I+ VBCR & "CLEAR FOR LIMITITOR" & VBCR
			I=I+ "THE DELETE LIMIT HAVE NOT BEEN REACH __ " 
			I=I+ CSTR(I_COUNT_DEL) & "_<_" & CSTR(DELETE_LIMITITOR) & VBCR
		END IF
		IF I_COUNT_DEL=>DELETE_LIMITITOR THEN 
			I=I+ " __ DO ANOTHER" & VBCR 
			I=I+ "THE DELETE LIMIT GIVE OVER  __ " 
			I=I+ CSTR(I_COUNT_DEL) & "_>_" & CSTR(DELETE_LIMITITOR) & VBCR
		END IF
		I=I+ "----" & VBCR 
		I=I+ "FRONT HALF SCRIPT -- DELETE TO DO" & VBCR 
		I=I+ "----" & VBCR 
		I=I+ DEL_SCRIPT_DATE_GONE & VBCR 
		I=I+ "----" & VBCR 
		VALUE_WAIT=20
		I=I+ "MESSENGER BOX DISAPPEAR ON " & CSTR(VALUE_WAIT) & " SECOND"
		I = TRIM(I)
		IF TRUE=TRUE OR MESSENGER_DISPLAY_MODE=TRUE THEN
			DIM WSHShell
			SET WSHShell = CreateObject("WScript.Shell")
			RESULT = WSHSHELL.POPUP(I, VALUE_WAIT, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
			SET WSHShell = Nothing
		END IF
	END IF
	' ---------------------------------------------------------------
	
	' ---------------------------------------------------------------
	' CONFIRM AND GO
	' ---------------------------------------------------------------
	' 0002 OF 0002
	' ---------------------------------------------------------------
	IF I_COUNT_DEL>0 THEN 
		I_COUNT=0
		I_COUNT_DEL=0
		I_COUNT_TOTAL=0
		MSGBOX_I1=""
		MSGBOX_I2=""
		LOW_DATE=NOW
		HIGH_DATE=NOW-1000
		DUMMY_RUN_GET_VALUE_HOW_MANY_DELETE=""
		CALL SCAN__TRIM_THE_DATE_OLDER_CAMERA(ARRAY_SONY_FOLDER)
	END IF

	' -------------------------------------------------------
	' 0002 OF 0002 -- GET FREE SPACE
	' -------------------------------------------------------
	' -------------------------------------------------------
	CALL GET_DRIVE_SPACE(ARRAY_SONY_FOLDER)
	' -------------------------------------------------------
	DRIVE_FREE_SPACE_GB_BEGIN=DRIVE_FREE_SPACE_GB
	' -----------------------------------------------------------

	IF I_COUNT_DEL>0 OR 1 = 1 THEN 
		
		DISPLAY_MESSENGER_FOLDER_CAMERA=ARRAY_SONY_FOLDER
		
		I = ""
		IF I_COUNT_DEL=0 THEN 
			I = I + "RESULT NONE DELETER" & VBCR
		ELSE
			I = I + "RESULT SOME DELETER" & VBCR
		END IF

		ON ERROR RESUME NEXT
		ERR.CLEAR
		I = I + LPAD(I_COUNT_DEL, "0", 4) & " ____ DELETE ON ____ " & LPAD(I_COUNT_TOTAL, "0", 4) & " ____ TOTAL" & VBCRLF 
		IF ERR.NUMBER>0 THEN 
			MSGBOX "2553 ERROR LINE " & VBCR & LPAD(I_COUNT_DEL, "0", 4) & VBCR & LPAD(I_COUNT_TOTAL, "0", 4) 
			MSGBOX "2553 ERROR LINE " & VBCR & CSTR(LPAD(I_COUNT_DEL, "0", 4)) & VBCR & CSTR(LPAD(I_COUNT_TOTAL, "0", 4) )
		END IF
		I = I + LPAD(ARRAY_FOR_NEXT, "0", 4) & " ____ OF ____ " & LPAD(UBOUND(ARRAY_PATH_01), "0", 4) & VBCRLF
		I = I + DISPLAY_MESSENGER_FOLDER_CAMERA & VBCRLF
		I = I + VOLUME_NAME_STRING_ & VBCRLF
		I = I + "-----------------------------------------------------------" & VBCRLF 
		I = I + "DONE DELETE COUNT LESS DAY SPAN COUNTER" & VBCRLF 
		I = I + "MP4 __AND__ JPG" & VBCRLF 
		I = I + "-----------------------------------------------------------" & VBCRLF 
		I = I + LPAD(DAY_BACK_TRIMMER_COUNTER, "0", 4) & "                           DAY __ LIMIT COUNTER" & VBCRLF 
		I = I + CSTR(DAY_BACK_TRIMMER) & " DAY __ TRIMMER" & VBCRLF 
		I = I + "-----------------------------------------------------------" & VBCRLF 
		I = I + "---------------------------------------------------" & VBCRLF 
		I = I + "NEWER DATE BEFORE ____ " & VBCRLF 
		I = I + "------------------------------------------" & VBCRLF 
		I = I + MSGBOX_I2 & VBCRLF 
		I = I + "------------------------------------------" & VBCRLF 
		I = I + "OLDER DATE OF _________ " & VBCRLF 
		I = I + "------------------------------------------" & VBCRLF 
		I = I + MSGBOX_I1 & VBCRLF 
		I = I + "--------------------------------------------------------------" & VBCRLF 
		I = I + "TOTAL FREE_SPACE_SET_TO_EQUAL__GB _ " & FORMATNUMBER(FREE_SPACE_SET_TO_EQUAL,2) & VBCRLF 
		I = I + "TOTAL DRIVE_FREE_SPACE___________GB _ " & FORMATNUMBER(DRIVE_FREE_SPACE_GB,2) & VBCRLF 
		I = I + "TOTAL DRIVE_FREE_SPACE_BEGIN____GB _ " & FORMATNUMBER(DRIVE_FREE_SPACE_GB_BEGIN,2) & VBCRLF 
		I = I + "TOTAL DRIVE_CAPACITY_______________GB _ " & FORMATNUMBER(DRIVE_CAPACITY___GB,2) & VBCRLF 
		I = I + "--------------------------------------------------------------" & VBCRLF 
		IF I_COUNT_DEL=0 THEN 
			VALUE_WAIT=8
		ELSE
			VALUE_WAIT=40
		END IF
		I = I + "MESSENGER BOX DISAPPEAR ON " & CSTR(VALUE_WAIT) & " SECOND"
		IF TRUE=TRUE OR MESSENGER_DISPLAY_MODE=TRUE THEN
		IF I_COUNT_DEL>0 THEN 
			SET WSHShell = CreateObject("WScript.Shell")
			RESULT = WSHSHELL.POPUP(I, VALUE_WAIT, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
			SET WSHShell = Nothing
		END IF
		END IF
	END IF
End Sub


' ---------------------------------------------------------------
' ALL-IN-ONE -- EASIER -- RECURSIVE NEST LOOP
' FILE AND SUB-FOLDER --- NICELY PRESENTABLE
' EASY ENTRY EITHER STRING FOLDER NAME OR OBJECT
' NOT EASY FINDER INTERNET PLUS SWEET AND SIMPLE
' ---------------------------------------------------------------
' WED 07-OCT-2020 01:22:36
' WED 07-OCT-2020 04:18:30 -- 2 HOUR 55 MINUTE
' EXTEND WORK
' WED 07-OCT-2020 06:24:00 -- 5 HOUR 00 MINUTE
' ---------------------------------------------------------------
' ANOTHER ROUTINE ADD FOR CAMERA AS PART NEW SYNC ROUTINE STUFF
' SOMETIME -- GOODSYNC WHEN NOT SAFEMODE OR SUPPOSE IS SAFEMODE
' FILE TRASH HASH LEFT BEHIND FROM ABORT COPY
' THE HASH FILE NAME ARE 32 CHAR LONG
' EXAMPLE HERE
' \1d6c1d40497153e04ffd0ea3f9dbb124.JPG
' WE KILL THEM WHEN ALL LOWERCASE AND NUMBER LEN(32)
' ---------------------------------------------------------------
' AND BETTER SUBROUTINE FILE SUB-FOLDER ONE SHOT AND! 
' ONE ROUTINE CODE -- SIMPLICITY TO RUNNER
' NONE FOLDER COMPLICATE GET FROM API -- USE STRING PATH
' EASY LAY DOWN
' ---------------------------------------------------------------
' UNLESS WANT FILE SCAN DOUBLE NEST IN THE SUB FOLDER ROUTINE
' ---------------------------------------------------------------
' SOME PEOPLE USER 3 ROUTINE FOR THE PART HERE
' ONE TO CALL 
' ONE TO COLLECT STRING ARRAY
' ANOTHER TO DEAL WITH THEM
' ---------------------------------------------------------------
' I FIND DELETE WORK WELL IT NOT MESS THE FIND OVER TO DELETE 
' WHILE GO
' ---------------------------------------------------------------
' ---------------------------------------------------------------
SUB SCAN__TRIM_THE_DATE_OLDER_CAMERA(FOLDER)
	DIM SUBFOLDER
	DIM FILE
	
	ON ERROR RESUME NEXT
	' ---------------------------------------------------------------
	IF UCASE(TYPENAME(FOLDER))="FOLDER" THEN 
		' PERFECT -- FOLDER OBJECT 
		' ABLE ARGUMENT PARAMETER AS STRING ANYTIME AND EXTRA CONVERT TAKE PLACE
		' USUAL ONCE AS STRING AND FOREVER AFTER FOLDER OBJECT VARIABLE TYPE
	END IF
	' ---------------------------------------------------------------
	IF UCASE(TYPENAME(FOLDER))="STRING" THEN
		Set FOLDER = FSO.GETFOLDER(FOLDER)
		SCAN__TRIM_THE_DATE_OLDER_CAMERA(FOLDER)
	END IF
	' ---------------------------------------------------------------
	

	FOR EACH FILE IN FOLDER.FILES
		FILE_PATH_STRING=FILE.path
		STR_F="-- .JPG .MP4" 
		PATH_FIND_01="\DCIM"
		PATH_FIND_02="\MP_ROOT"
		EXTENSION_PART=MID(FILE.path,INSTRREV(FILE.path,"."))
		PATH_FIND_RESULT=FALSE
		' -----------------------------------------------------------
		' NOT GO IN DIRECTORY THAT DON;T HAVE ACCESS
		' LIKE \SYSTEM INFORMATION FOLDER FROM ROOT
		' OR BOMB OUT OF SCAN
		' SAFE NOT PROCESS ANY INTRICATE INFO ABOUT THEM
		' -----------------------------------------------------------
		IF INSTR(STR_F , EXTENSION_PART) >0 THEN PATH_FIND_RESULT=TRUE
		IF INSTR(UCASE(FILE.path) , PATH_FIND_01) >0 THEN PATH_FIND_RESULT=TRUE
		IF INSTR(UCASE(FILE.path) , PATH_FIND_02) >0 THEN PATH_FIND_RESULT=TRUE
		
		IF PATH_FIND_RESULT=TRUE THEN 
			ON ERROR RESUME NEXT
			ERR.CLEAR
			FILEMODDATE=0
			SET F = FSO.GETFILE(FILE.path)
				FILEMODDATE = F.DATELASTMODIFIED
			SET F=NOTHING
			IF ERR.NUMBER>0 THEN FILEMODDATE=0
			IF FILEMODDATE=0 THEN 
				MSGBOX "ERROR RESULT GET DATE" & VBCR & FILE.path,vbMsgBoxSetForeground,Wscript.ScriptName 
				PATH_FIND_RESULT=FALSE
			END IF

			I_COUNT_TOTAL=I_COUNT_TOTAL+1
			' -------------------------------------------------------
			' -------------------------------------------------------
			' NOT USE NOW -- HIGH NUMBER SOUND
			' -------------------------------------------------------
			DELETE_LIMITITOR=1000 
			' -------------------------------------------------------
			DAY_BACK_TRIMMER_COUNTER=150
			DAY_BACK_TRIMMER = NOW - DAY_BACK_TRIMMER_COUNTER 
			' -------------------------------------------------------
			' DAY -- HIGH NUMBER MORE BEFORE
			' -------------------------------------------------------
			' EXAMPLE 400 DAY - SUPPLY - 2020-10-08 GIVE - 04/09/2019
			' EXAMPLE 200 DAY - SUPPLY - 2020-10-08 GIVE - 22/03/2020 
			' EXAMPLE 150 DAY - SUPPLY - 2020-10-08 GIVE - 11/05/2020
			' -------------------------------------------------------
			' DEBUG REM VERIFY -- WHAT DATE GIVE HOW FAR BACK
			' MSGBOX DAY_BACK_TRIMMER
			' -------------------------------------------------------
			' -------------------------------------------------------

			
			IF PATH_FIND_RESULT=TRUE THEN 
				IF FILEMODDATE>DAY_BACK_TRIMMER THEN 
					PATH_FIND_RESULT=FALSE
				END IF
			END IF
			' -------------------------------------------------------
			' OBTAIN HIGH LOW DATE OF SELECTION WORKER
			' THAT WILL BE CLOSE TO LAST ONE LEFT OF DATE HIGH VALUE
			' AND LOW IS WHERE WIPE ALL THEM FROM
			' -------------------------------------------------------
			IF PATH_FIND_RESULT=TRUE THEN 
				IF FILEMODDATE>HIGH_DATE THEN 
					HIGH_DATE=FILEMODDATE
					' STRIP THE FRONT PART PATH TOO LONG 
					' ---------------------------------------------------
					FILE_PATH_TWO_LINE=REPLACE(FILE.PATH,ARRAY_SONY_FOLDER,"")
					MSGBOX_I2=FILE_PATH_TWO_LINE & VBCR & FILEMODDATE & " ---- HIGH"
					MSGBOX_I2=UCASE(MSGBOX_I2)
				END IF
				IF FILEMODDATE<LOW_DATE THEN 
					LOW_DATE=FILEMODDATE
					' STRIP THE FRONT PART PATH TOO LONG 
					' ---------------------------------------------------
					FILE_PATH_TWO_LINE=REPLACE(FILE.PATH,ARRAY_SONY_FOLDER,"")
					MSGBOX_I1=FILE_PATH_TWO_LINE & VBCR & FILEMODDATE & " ---- LOW" 
					MSGBOX_I1=UCASE(MSGBOX_I1)
				END IF
				' MAXIMUM HIGH DATE WHEN WANT
				' AND LOW
				' RATHER THAN DATE SPLITTER
				' -------------------------------------------------------
			END IF
		
			ON ERROR RESUME NEXT
			ERR.CLEAR

			IF PATH_FIND_RESULT=TRUE THEN 

				IF I_COUNT_DEL<=DELETE_LIMITITOR THEN 
					' MSGBOX FILE.PATH & VBCR & FILEMODDATE & VBCR & HIGH_DATE & VBCR & LOW_DATE & VBCR & "HIGH_DATE  " & HIGH_DATE & VBCR & "LOW_DATE  " & LOW_DATE 					' ---------------------------------------------------
					IF DUMMY_RUN_GET_VALUE_HOW_MANY_DELETE<>"GET_FIND_INFO" THEN 
						IF I_COUNT_DEL=0 THEN 
							I=""
							I=I+"1ST FILE DELETE" & VBCR 
							I=I+"NONE INFO PROVIDE AFTER" & VBCR 
							I = I + "-------------------------------" & VBCR
							I=I+FILE.PATH & VBCR
							I=I+FILEMODDATE & VBCR 
							I = I + "-------------------------------" & VBCR
							I=I+HIGH_DATE & VBCR
							I=I+LOW_DATE & VBCR
							I = I + "-------------------------------" & VBCR
							I=I+"HIGH_DATE  " & HIGH_DATE & VBCR
							I=I+"LOW_DATE   " & LOW_DATE & VBCR
							VALUE_WAIT=20
							I = I + "-------------------------------" & VBCR
							I = I + "MESSENGER BOX GONE IN " & CSTR(VALUE_WAIT) & " SECOND"
							DIM WSHShell
							SET WSHShell = CreateObject("WScript.Shell")
							RESULT = WSHSHELL.POPUP(I, VALUE_WAIT, Wscript.ScriptName , vbOKOnly + vbMsgBoxSetForeground)
							SET WSHShell = Nothing
							
						END IF
						FSO.DeleteFile(FILE.path) 
					END IF
					' ---------------------------------------------------
				END IF
				I_COUNT_DEL=I_COUNT_DEL+1
				DEL_SCRIPT_DATE_GONE=DEL_SCRIPT_DATE_GONE & FILE.path & VBCRLF
			END IF
			' -------------------------------------------------------
		END IF
	NEXT
	' ON ERROR GOTO 0
	' NOT CODE DO -- ON ERROR GOTO 0 -- UNREM 

	' ---------------------------------------------------------------
    ' RECURSE INTO IT'S SUB FOLDERS
	' ---------------------------------------------------------------
    FOR EACH SUBFOLDER IN FOLDER.SUBFOLDERS
		SCAN__TRIM_THE_DATE_OLDER_CAMERA(SUBFOLDER)
    NEXT
	
	' ---------------------------------------------------------------
End Sub
' -------------------------------------------------------------------


' -------------------------------------------------------------------
' -------------------------------------------------------------------
' TWO ROUTINE THIS ONE -- THUMB WRITTEN
' -------------------------------------------------------------------
' MAH00004.THM
' &
' THUMBS.DB -- Thumbs.db
' -------------------------------------------------------------------
' -------------------------------------------------------------------
Sub FIND_THUMB_OFF_CAMERA_HX60(sFolder)
    On Error Resume NEXT
    Dim oFDR, oItem, file
    ' CHECK IF WANTED FILES ARE IN THIS FOLDER
    For Each file In sFolder.Files
		IF INSTR(UCASE(FILE.path),".THM") THEN
			oDict.Add FILE.path, file
		END IF
		IF UCASE(FILE.path)="THUMBS.DB" THEN
			oDict.Add FILE.path, file
		END IF
    NEXT
    ' RECURSE INTO IT'S SUB FOLDERS
    For Each oFDR In sFolder.SubFolders
        FindWantedFiles_THUMB_OFF_CAMERA_HX60 FSO.GetFolder(oFDR.Path)
    NEXT
End Sub
' -------------------------------------------------------------------
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
    NEXT
End Sub
' -------------------------------------------------------------------





' -------------------------------------------------------------------
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
    FindWantedFiles_4(FSO.GetFolder(FOLDER_PATH_MAIN_4))
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
' THUMBS.DB -- THUMBS.DB
' -------------------------------------------------------------------
Sub FindWantedFiles_THUMB_OFF_CAMERA_HX60(sFolder)
    On Error Resume NEXT
    Dim oFDR, oItem, file
    ' CHECK IF WANTED FILES ARE IN THIS FOLDER
    For Each file In sFolder.Files
		IF INSTR(UCASE(FILE.path),".THM") THEN
			oDict.Add FILE.path, file
		END IF
		IF UCASE(FILE.path)="THUMBS.DB" THEN
			oDict.Add FILE.path, file
		END IF
    NEXT
    ' RECURSE INTO IT'S SUB FOLDERS
    For Each oFDR In sFolder.SubFolders
        FindWantedFiles_THUMB_OFF_CAMERA_HX60 FSO.GetFolder(oFDR.Path)
    NEXT
End Sub
' -------------------------------------------------------------------
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
    NEXT
End Sub
' -------------------------------------------------------------------



Sub FindWantedFiles_4(sFolder)
    On Error Resume NEXT
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
    NEXT
    ' Recurse into it's sub folders
    For Each oFDR In sFolder.SubFolders
        FindWantedFiles_4 FSO.GetFolder(oFDR.Path)
    NEXT
	
End Sub

Sub DisplayFindings_4()
    Dim oKeys, oKey
    oKeys = oDict_4.Keys
    For Each oKey In oKeys
		VAR_PATH_SCANNER=oKey
		CALL PATH_SCANNER_RENAME_4
    NEXT
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
				i_TIME=i_TIME+LPAD(Year(FileModDate), "0", 4)+"-"
				i_TIME=i_TIME+LPAD(Month(FileModDate), "0", 2)+"-"
				i_TIME=i_TIME+LPAD(Day(FileModDate), "0", 2)
				i_TIME=i_TIME+" "
				i_TIME=i_TIME+LPAD(Hour(FileModDate), "0", 2)+"-"
				i_TIME=i_TIME+LPAD(Minute(FileModDate), "0", 2)+"-"
				i_TIME=i_TIME+LPAD(Second(FileModDate), "0", 2)
				i_TIME=i_TIME+" - 4K ULTRA A_CAM - "+FILE_VAR
			END IF
		END IF
		IF INSTR(FILE_VAR,".AVI") THEN 
			SET_GO=FALSE
			IF MID(FILE_VAR,1,3)="REC" THEN SET_GO=TRUE
			IF MID(FILE_VAR,1,3)="SOS" THEN SET_GO=TRUE
			IF SET_GO=TRUE THEN
				i_TIME=""
				i_TIME=i_TIME+LPAD(Year(FileModDate), "0", 4)+"_"
				i_TIME=i_TIME+LPAD(Month(FileModDate), "0", 2)+"_"
				i_TIME=i_TIME+LPAD(Day(FileModDate), "0", 2)+" "
				i_TIME=i_TIME+MonthName(MONTH(FileModDate),TRUE)+"_"
				i_TIME=i_TIME+WeekdayName(Weekday(VAR_NOW),TRUE)

				i_TIME=i_TIME+" "
				i_TIME=i_TIME+LPAD(Hour(FileModDate), "0", 2)+"_"
				i_TIME=i_TIME+LPAD(Minute(FileModDate), "0", 2)+"_"
				i_TIME=i_TIME+LPAD(Second(FileModDate), "0", 2)+"__"
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
				i_TIME=i_TIME+LPAD(Year(FileModDate), "0", 4)+"_"
				i_TIME=i_TIME+LPAD(Month(FileModDate), "0", 2)+"_"
				i_TIME=i_TIME+LPAD(Day(FileModDate), "0", 2)+" "
				i_TIME=i_TIME+MonthName(MONTH(FileModDate),TRUE)+"_"
				i_TIME=i_TIME+WeekdayName(Weekday(VAR_NOW),TRUE)

				i_TIME=i_TIME+" "
				i_TIME=i_TIME+LPAD(Hour(FileModDate), "0", 2)+"_"
				i_TIME=i_TIME+LPAD(Minute(FileModDate), "0", 2)+"_"
				i_TIME=i_TIME+LPAD(Second(FileModDate), "0", 2)+"__"
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
		FSO.MOVEFILE VAR_PATH_SCANNER, NEWFILENAME
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

Function LPAD (str, pad, length)
	LPAD=""
	ON ERROR RESUME NEXT
	LPAD = String(length - Len(str), pad) & str
	IF ERR.NUMBER>0 THEN
		LPAD=STR
		' BEFORE ERROR WERE PAD = 0 WHICH NOT RUN
		MSGBOX "LINE 3318 ERROR " & VBCR & "LPAD = String(length - Len(str), pad) & str" & VBCR & CSTR(str) & VBCR & CSTR(pad) & VBCR & CSTR(length)
	END IF
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
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' -------------------------------------------------------------------
SUB SEARCH(FOLDER)
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
		Search FSO.GetFolder(subFolder.Path)
	NEXT
End Sub
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' -------------------------------------------------------------------
SUB DELETE_EMPTY_FOLDER(FOLDER)
	DIM SUBFOLDER
	ON ERROR RESUME NEXT
	ERR.CLEAR
	' ----------------------------------------
	' HERE FOLDER OBJECT
	' ----------------------------------------
	IF UCASE(TYPENAME(FOLDER))="FOLDER" THEN 
		' PERFECT 
	END IF
	' ----------------------------------------
	IF UCASE(TYPENAME(FOLDER))="STRING" THEN
		SET FOLDER = FSO.GETFOLDER(FOLDER)
		CALL DELETE_EMPTY_FOLDER(FOLDER)
	END IF
	
	' MSGBOX "DELETE_EMPTY_FOLDER     " & VBCR & ARRAY_PATH_01(ARRAY_FOR_NEXT) & VBCR & ARRAY_PATH_02(ARRAY_FOR_NEXT) & UCASE(FOLDER.PATH)

	' -----------------------------------------------------------
	' -----------------------------------------------------------
	' REM HERE -- STOP LOOP SUBFOLDER
	' HERE NOT A FILE NESTER ONLY FOLDER
	' WOULD BE BETTER NOT SUBFOLDER -- ROOT DOWN
	' -----------------------------------------------------------
    FOR EACH SUBFOLDER IN FOLDER.SUBFOLDERS
		DELETE_EMPTY_FOLDER(subfolder)
    NEXT
	
	' ---------------------------------------------------------------
	' MORE ERROR SAFE METHOD
	' REMEMBER DECLARE ON AS OPTION EXPLICIT 
	' AND ERROR RESULT FROM DECLARE ERROR MAKE NOT PROPER TRAP IT
	' ---------------------------------------------------------------
	READY_01=-1
	READY_02=-1
	READY_03= 0
	READY_04=-1
	ERR.CLEAR
	READY_01=FOLDER.SUBFOLDERS.COUNT
	READY_02=FOLDER.FILES.COUNT
	READY_04=ERR.NUMBER
	' IF UCASE(ORIGINAL_PATH)=UCASE(FOLDER.PATH) THEN 
		' ' NOT ORIGINAL PATH
		' READY_03 = 0
	' END IF
	' ---------------------------------------------------------------
    IF READY_01 = 0 AND READY_02 = 0 AND READY_03 = 0 AND READY_04=0 THEN
		ON ERROR RESUME NEXT
		ERR.CLEAR
		' MSGBOX "CHECKER __ DEL EMPTY " & FOLDER.PATH
		FSO.DELETEFOLDER FOLDER.PATH, TRUE
		I_DELETE_EMPTY_COUNTER=I_DELETE_EMPTY_COUNTER+1
		TEST_COUNTER=TEST_COUNTER+1
		IF ERR.NUMBER>0 THEN 
			ERR_COUNTER = ERR_COUNTER +1
			ERROR_WITH_SOME = FOLDER.PATH & VBCRLF & ERR.DESCRIPTION
			ERR.CLEAR
			' OBJFOLDER.DELETE, TRUE
			IF ERR.NUMBER=0 THEN ERR_COUNTER = ERR_COUNTER - 1
		END IF
	End If
End Sub
' -------------------------------------------------------------------	

' -------------------------------------------------------------------	
' -------------------------------------------------------------------
SUB DELETE_FOUND_000_ROOT_PATH(FOLDER)
	DIM SUBFOLDER
	ON ERROR RESUME NEXT
	ERR.CLEAR
	' ----------------------------------------
	' HERE FOLDER OBJECT
	' ----------------------------------------
	IF UCASE(TYPENAME(FOLDER))="FOLDER" THEN 
		' PERFECT 
	END IF
	' ----------------------------------------
	IF UCASE(TYPENAME(FOLDER))="STRING" THEN
		Set FOLDER = FSO.GETFOLDER(FOLDER)
		CALL DELETE_FOUND_000_ROOT_PATH(FOLDER)
	END IF
	
	' ----------------------------------------
	' -----------------------------------------------------------
	' REM HERE -- STOP LOOP SUBFOLDER
	' HERE NOT A FILE NESTER ONLY FOLDER
	' WOULD BE BETTER NOT SUBFOLDER -- ROOT DOWN
	' -----------------------------------------------------------
    FOR EACH SUBFOLDER IN FOLDER.SUBFOLDERS
		DELETE_FOUND_000_ROOT_PATH subfolder
    NEXT
	' ---------------------------------------------------------------
	' MORE ERROR SAFE METHOD
	' REMEMBER DECLARE ON AS OPTION EXPLICIT 
	' AND ERROR RESULT FROM DECLARE ERROR MAKE NOT PROPER TRAP IT
	' ---------------------------------------------------------------
	READY_01=-1
	READY_02=-1
	READY_04=-1
	ERR.CLEAR
	READY_01=0
	READY_02=0
	' ---------------------------------------------------------------
	' SEARCH FOR EMPTY
	' ---------------------------------------------------------------
	' READY_01=FOLDER.SUBFOLDERS.COUNT
	' READY_02=FOLDER.FILES.COUNT
	' ---------------------------------------------------------------
	READY_04=ERR.NUMBER
	' ---------------------------------------------------------------
	' MSGBOX "DELETE_FOUND_000_ROOT_PATH     " & VBCR & ARRAY_PATH_01(ARRAY_FOR_NEXT) & VBCR & ARRAY_PATH_02(ARRAY_FOR_NEXT) & UCASE(FOLDER.PATH)
    If READY_01 = 0 And READY_02 = 0 AND READY_04=0 Then
		ON ERROR GOTO 0
		TEST_COUNTER=TEST_COUNTER+1
		ON ERROR RESUME NEXT
		ERR.CLEAR
		' -----------------------------------------------------------
		' SUCCESS ACHIEVE ON TEST RUN
		' ALLOW NOT MESSENGER LEAVER
		' Sun 08-Mar-2020 17:20:40
		' -----------------------------------------------------------

		IF INSTR(FOLDER.PATH,"\FOUND.")>0  THEN 
			MSGBOX "CHECKER __ FOUND 000 " & FOLDER.PATH
			FSO.DELETEFOLDER FOLDER.PATH, TRUE
			I_DELETE_FOUND_000=I_DELETE_FOUND_000+1
			' FOLDER.DELETE, TRUE
		END IF
		
		IF ERR.NUMBER>0 THEN 
			ERR_COUNTER = ERR_COUNTER +1
			ERROR_WITH_SOME = ERROR_WITH_SOME + folder.Path & VBCRLF & ERR.DESCRIPTION + VBCR
			ERR.CLEAR
			IF ERR.NUMBER=0 THEN ERR_COUNTER = ERR_COUNTER - 1
		END IF
	End If
End Sub
' -------------------------------------------------------------------	

' -------------------------------------------------------------------
' -------------------------------------------------------------------
SUB FIND_HASH_TRASH_32_CHAR(FOLDER)
	' ---------------------------------------------------------------
	' ALL-IN-ONE -- EASIER -- RECURSIVE NEST LOOP
	' FILE AND SUB-FOLDER --- NICELY PRESENTABLE
	' EASY ENTRY EITHER STRING FOLDER NAME OR OBJECT
	' NOT EASY FINDER INTERNET PLUS SWEET AND SIMPLE
	' ---------------------------------------------------------------
	' WED 07-OCT-2020 01:22:36
	' WED 07-OCT-2020 04:18:30 -- 2 HOUR 55 MINUTE
	' EXTEND WORK
	' WED 07-OCT-2020 06:24:00 -- 5 HOUR 00 MINUTE
	' ---------------------------------------------------------------
	' ANOTHER ROUTINE ADD FOR CAMERA AS PART NEW SYNC ROUTINE STUFF
	' SOMETIME -- GOODSYNC WHEN NOT SAFEMODE OR SUPPOSE IS SAFEMODE
	' FILE TRASH HASH LEFT BEHIND FROM ABORT COPY
	' THE HASH FILE NAME ARE 32 CHAR LONG
	' EXAMPLE HERE
	' \1d6c1d40497153e04ffd0ea3f9dbb124.JPG
	' WE KILL THEM WHEN ALL LOWERCASE AND NUMBER LEN(32)
	' ---------------------------------------------------------------
	' AND BETTER SUBROUTINE FILE SUB-FOLDER ONE SHOT AND! 
	' ONE ROUTINE CODE -- SIMPLICITY TO RUNNER
	' NONE FOLDER COMPLICATE GET FROM API -- USE STRING PATH
	' EASY LAY DOWN
	' ---------------------------------------------------------------
	' UNLESS WANT FILE SCAN DOUBLE NEST IN THE SUB FOLDER ROUTINE
	' ---------------------------------------------------------------
	' SOME PEOPLE USER 3 ROUTINE FOR THE PART HERE
	' ONE TO CALL 
	' ONE TO COLLECT STRING ARRAY
	' ANOTHER TO DEAL WITH THEM
	' ---------------------------------------------------------------
	' I FIND DELETE WORK WELL IT NOT MESS THE FIND OVER TO DELETE 
	' WHILE GO
	' ---------------------------------------------------------------
	' ---------------------------------------------------------------
	DIM SUBFOLDER
	DIM FILE
	ON ERROR RESUME NEXT
	' ---------------------------------------------------------------
	IF UCASE(TYPENAME(FOLDER))="FOLDER" THEN 
		' PERFECT -- FOLDER OBJECT 
		' ABLE ARGUMENT PARAMETER AS STRING ANYTIME AND EXTRA CONVERT TAKE PLACE
		' USUAL ONCE AS STRING AND FOREVER AFTER FOLDER OBJECT VARIABLE TYPE
	END IF
	' ---------------------------------------------------------------
	IF UCASE(TYPENAME(FOLDER))="STRING" THEN
		Set FOLDER = FSO.GETFOLDER(FOLDER)
		CALL FIND_HASH_TRASH_32_CHAR(FOLDER)
	END IF
	' ---------------------------------------------------------------
    
	' ---------------------------------------------------------------
	' FIND THE ROOT PATH AND NETWORK OR BE A BIT HARD
	' WORK RECYCLE BIN STORE FILE RISK PRECIOUS
	' ONLY DETECT IS LEN() 32 WITH LOWER CASE ALPHA
	' ---------------------------------------------------------------
	DIM F
	SET F = FSO.GETFOLDER(FOLDER.PATH)
	FOR R1=1 TO 10
		IF F.ISROOTFOLDER=FALSE THEN 
			SET F = F.PARENTFOLDER
		ELSE
			EXIT FOR
		END IF
	NEXT
	ROOT_FOLDER = F.PATH
	
	SET F = NOTHING
	
	FOR EACH FILE IN FOLDER.FILES
		FILE_01_PART=MID(FILE.path     ,INSTRREV(FILE.path,"\")+1)
		FILE_02_PART=MID(FILE_01_PART,1,INSTRREV(FILE_01_PART,".")-1)
		FILE_02_PART_NOT_EXTENSION=FILE_02_PART
		FILE_02_NOT_EXTENSION_____=FILE_02_PART
		' ALL LOWERCASE VERIFY
		IF LCASE(FILE_02_NOT_EXTENSION_____)=FILE_02_NOT_EXTENSION_____ THEN
		' LENGTH 32 CHARACTER
		IF LEN(FILE_02_NOT_EXTENSION_____)=32 THEN
			' -------------------------------------------------------
			' FIND ITEM MAKE TEMP FOLDER 
			' HAVE TO VERIFY NOT ANOTHER FILE OF LEN()=32 
			' RECYCLE NOT WORK MC-CARD
			' -------------------------------------------------------
			' GET THE ROOT GOOD IDEA BUTTY REALLY WANT ORIGINAL PATH WITH DEPTH
			' MAKE SURE END SLASH NOT THERE AS ENTER
			' DOUBLE CHECK NONE TO DO ALREADY API FSO OBJECT
			' -------------------------------------------------------
			' NOW DO BUILD DIRECTORY TREE WHEN REQUEST MADE MOVE FILE
			' -------------------------------------------------------
			
			' -------------------------------------------------------
			' CALL BUILDFULLPATH(ROOT_FOLDER & "\TEMP_HASH_TRASH")
			' -------------------------------------------------------

			' -------------------------------------------------------
			' MOVE THE FILE -- NOT \TEMP_HASH_TRASH FOLDER 
			' DONE BEFORE
			' -------------------------------------------------------

			' -------------------------------------------------------
			' VERIFY FILE SIZE IS NAUGHT BYTE 
			' NOT REQUIRE MOVE THEN ONLY DELETEE 
			' MOST COMMON IS THAT ONE
			' -------------------------------------------------------
			ASIZE1=-10
			SET F = FS.GETFILE(FILE.path)
			ASIZE1 = F.SIZE
			SET F = NOTHING
			
			If ASIZE1=0 AND FSO.FileExists(FILE.path) = TRUE Then 
				FSO.DeleteFile FILE.path
			END IF
			If ASIZE1>0 AND FSO.FileExists(FILE.path) = TRUE Then 
				IF INSTR(FILE.path,"\TEMP")=0 THEN 
					' ---------------------------------------------------
					IF FSO.FOLDEREXISTS(ARRAY_SONY_FOLDER & "\TEMP_HASH_TRASH") = FALSE THEN 
						CALL BUILDFULLPATH(ARRAY_SONY_FOLDER & "\TEMP_HASH_TRASH")
					END IF
					' ---------------------------------------------------
					FSO.MOVEFILE FILE.path, ARRAY_SONY_FOLDER & "\TEMP_HASH_TRASH\" & FILE_01_PART
					I_COUNT_HASH_TRASH_32_CHAR = I_COUNT_HASH_TRASH_32_CHAR + 1
					' ---------------------------------------------------
				END IF
			END IF
		END IF
		END IF
    NEXT
	' ---------------------------------------------------------------
    FOR EACH SUBFOLDER IN FOLDER.SUBFOLDERS
		FIND_HASH_TRASH_32_CHAR(SUBFOLDER)
    NEXT
	' ---------------------------------------------------------------
    If INSTR(FOLDER.PATH,"\FOUND.")  Then
	End If
End Sub
' -------------------------------------------------------------------	

' -------------------------------------------------------------------	
' -------------------------------------------------------------------	
Sub BuildFullPath(ByVal FullPath)
    If Not fso.FolderExists(FullPath) Then
        BuildFullPath fso.GetParentFolderName(FullPath)
        fso.CreateFolder FullPath
    End If
End Sub
' -------------------------------------------------------------------	
' scripting - Help To create Folder1/Folder2 in Windows using VBScript ( Both the folders not exists before, i mean to create multilevel folders @ a strech.) - Stack Overflow 
' https://stackoverflow.com/questions/4407386/help-to-create-folder1-folder2-in-windows-using-vbscript-both-the-folders-not
' -------------------------------------------------------------------	


' -------------------------------------------------------------------	
' -------------------------------------------------------------------	
' HERE MOVE LOWER AND NOT RUN SMALL VERSION ABOVE GOT THERE
' RUNN_IT_HARDCORE_THREE_ROUTINE_0002
' -------------------------------------------------------------------	
' -------------------------------------------------------------------	
SUB RUNN_IT_HARDCORE_THREE_ROUTINE:
	FOR ARRAY_FOR_NEXT=1 TO UBOUND(ARRAY_PATH_01)
		VOLUME_NAME_STRING_=""
		DAY_BACK_TRIMMER_COUNTER=0
		DRIVE_FREE_SPACE_GB_BEGIN=0
		DRVPATH_VAR=""
		GET_SPACE_STRING=""
		FREE_SPACE_SET_TO_EQUAL=0
		DRIVE_FREE_SPACE_GB=0
		DRIVE_CAPACITY___GB=0
		I_COUNT_DEL=0
		I_COUNT_TOTAL=0
		DAY_BACK_TRIMMER=0
		I_COUNT=0
		MSGBOX_I1=""
		MSGBOX_I2=""
		' REDIM ODICT_TRIM_THE_DATE
		' MSGBOX "0002 DIM ODICT_TRIM_THE_DATE"

		' -------------------------------------------------------------------
		LOW_DATE=NOW
		HIGH_DATE=NOW-1000
		' -------------------------------------------------------------------

		I=GET_VOLUME_LETTER__(ARRAY_PATH_01(ARRAY_FOR_NEXT))
		IF MID(I,2,1)=":" THEN 
			ARRAY_PATH_01(ARRAY_FOR_NEXT)=I
		END IF
		ARRAY_SONY_FOLDER=ARRAY_PATH_01(ARRAY_FOR_NEXT)

		' ---------------------------------------------------
		' 0003 OF 0003
		CALL STRESSOR__TRIM_THE_DATE_OLDER_CAMERA
		' ---------------------------------------------------
		' 0002 OF 0003
		CALL DELETE_EMPTY_FOLDER(ARRAY_SONY_FOLDER)
		' ---------------------------------------------------
	NEXT
END SUB
' -------------------------------------------------------------------
' RUNN_IT_HARDCORE_THREE ROUTINE GIVER GIVE BACK SPACE _ REMOVE \_FOUND_000 KILL EMPTY
' -------------------------------------------------------------------
' RUN ROUTINE END NUMBER 03 AND RUN IT FOUR THAT POWER
' -------------------------------------------------------------------
' 0001 OF 0004 -------------------------------------
' CALL DELETE_FOUND_000_ROOT_PATH(ARRAY_SONY_FOLDER)
' 0002 OF 0004 -------------------------------------
' CALL DELETE_EMPTY_FOLDER(ARRAY_SONY_FOLDER)
' 0003 OF 0004 -------------------------------------
' CALL STRESSOR__TRIM_THE_DATE_OLDER_CAMERA
' 0004 OF 0004 -------------------------------------
' LINE 2800 AT END
' CALL RUNN_IT_HARDCORE_THREE_ROUTINE
' -------------------------------------------------------------------





' -------------------------------------------------------------------
' 0001 OF 0002
' -------------------------------------------------------------------
' REFERENCE
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' How Can I Get a List of All the Files in a Folder and Its Subfolders? – Hey, Scripting Guy! Blog
' https://blogs.technet.microsoft.com/heyscriptingguy/2004/10/20/how-can-i-get-a-list-of-all-the-files-in-a-folder-and-its-subfolders/
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' THIS PROJECT USING THIS CODE
' BECAUSE A MOVE FILE HAPPEN MUST GATHER WHOLE FILE LIST BEFORE ACT ON IT
' AND IT WAS DEBUGED IN THE NEXT URL LINK LINE
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' windows - VBS to Search for Multiple Files Recursively in C:\Users - Stack Overflow
' https://stackoverflow.com/questions/25658004/vbs-to-search-for-multiple-files-recursively-in-c-users
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' vbscript - Read files in subfolders - Stack Overflow
' https://stackoverflow.com/questions/16905278/read-files-in-subfolders
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' vbscript - Read files in subfolders - Stack Overflow
' https://stackoverflow.com/questions/16905278/read-files-in-subfolders
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------


' -------------------------------------------------------------------
' -------------------------------------------------------------------
' 0002 OF 0002
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' EXTRA CODE ROUTINE SET
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
SUB VBS_82_RECURSE_SUBFOLDER___FILE_ONE_SHOT_ROUTINE_EXAMPLE_TOP_PLAYER_VBS(FOLDER)
	' -------------------------------------------------------------------
	' HERE ARE PERFECT CODE
	' NOT FOUND ON-LINE EVER EXCEPT
	' YOU ABLE CALL ROUTINE FROM OUT ROUTINE
	' OR GIVE FOLDER WITHIN
	' THERE NONE NEST BESIDE IT LOOP ITSELF
	' NONE MULTIPLE ROUTINE EFFORT
	' -------------------------------------------------------------------
	' EASY PASS 1ST PARAMETER -- WHAT FOLDER
	' -------------------------------------------------------------------
	' IF YOU AIN'T KICKING TOP ROUTINE SCANNER 
	' LOOK AT MY SHIT
	' -------------------------------------------------------------------
	' WITH AR SOME EXAMPLE WHAT TO DO WITH INNER
	' -------------------------------------------------------------------
	' [ WED 12:44:10 PM_07 OCTOBER 2020 ]
	' -------------------------------------------------------------------
	' -------------------------------------------------------------------
	DIM SUBFOLDER
	DIM FILE
	ON ERROR RESUME NEXT
	' ---------------------------------------------------------------
	IF UCASE(TYPENAME(FOLDER))="FOLDER" THEN 
		' PERFECT -- FOLDER OBJECT 
		' ABLE ARGUMENT PARAMETER AS STRING ANYTIME AND EXTRA CONVERT TAKE PLACE
		' USUAL ONCE AS STRING AND FOREVER AFTER FOLDER OBJECT VARIABLE TYPE
	END IF
	' ---------------------------------------------------------------
	IF UCASE(TYPENAME(FOLDER))="STRING" THEN
		FOLDER="E:\HARDWARE"
		Set FOLDER = FSO.GETFOLDER(FOLDER)
		VBS_82_RECURSE_SUBFOLDER___FILE_ONE_SHOT_ROUTINE_EXAMPLE_TOP_PLAYER_VBS(FOLDER)
	END IF
	' ---------------------------------------------------------------
	FOR EACH FILE IN FOLDER.FILES
		' DO STUFF
		' ----------------------------------------------
		' NOT GO IN DIRECTORY THAT DON'T HAVE ACCESS
		' LIKE \SYSTEM INFORMATION FOLDER FROM ROOT
		' OR BOMB OUT OF SCAN
		' SAFE NOT PROCESS ANY INTRICATE INFO ABOUT THEM
		' ----------------------------------------------
		FILE_PATH_STRING=FILE.path
		STR_F="-- .JPG .MP4" 
		PATH_FIND_01="\DCIM"
		PATH_FIND_02="\MP_ROOT"
		EXTENSION_PART=MID(FILE.path,INSTRREV(FILE.path,"."))
		PATH_FIND_RESULT=FALSE
		IF INSTR(STR_F , EXTENSION_PART) >0 THEN PATH_FIND_RESULT=TRUE
		IF INSTR(UCASE(FILE.path) , PATH_FIND_01) >0 THEN PATH_FIND_RESULT=TRUE
		IF INSTR(UCASE(FILE.path) , PATH_FIND_02) >0 THEN PATH_FIND_RESULT=TRUE
		PATH_FIND_RESULT=TRUE
		IF PATH_FIND_RESULT=TRUE THEN 
			ERR.CLEAR
			FILEMODDATE=0
			SET F = FSO.GETFILE(FILE.path)
				FILEMODDATE = F.DATELASTMODIFIED
			SET F=NOTHING
			COUNTER_RR=COUNTER_RR+1
		END IF
	NEXT
	' ---------------------------------------------------------------
	' CALL GET_DRIVE_SPACE(FOLDER.path)
	' ---------------------------------------------------------------
	COUNTER_RR=COUNTER_RR+1
	' ---------------------------------------------------------------
	' DO STUFF ON FOLDER HERE
	' ---------------------------------------------------------------
	MYFOLDER = FOLDER.path
	' ---------------------------------------------------------------
	IF MYFOLDER = "EXCELLENT" THEN 
		' DO STUFF
	END IF
	' ---------------------------------------------------------------
	
	' ---------------------------------------------------------------
    ' RECURSE INTO AR SUB FOLDER COLLECTION
	' ---------------------------------------------------------------
    FOR EACH SUBFOLDER IN FOLDER.SUBFOLDERS
		VBS_82_RECURSE_SUBFOLDER___FILE_ONE_SHOT_ROUTINE_EXAMPLE_TOP_PLAYER_VBS(SUBFOLDER)
    NEXT
END SUB
' -------------------------------------------------------------------
' -------------------------------------------------------------------



