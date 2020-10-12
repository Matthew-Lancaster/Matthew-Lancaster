' -------------------------------------------------------------------
' HERE ARE PERFECT CODE
' NOT FOUND ON-LINE EVER EXCEPT THIS ONE
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
DIM FSO
SET FSO = CreateObject("Scripting.FileSystemObject")

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

' ------------------------------------------------------------------
' EXAMPLE 0001 CALL ROUTINE WITH PATH
' ------------------------------------------------------------------
FOLDER_NAME="C:\Users\"+GETUSERNAME+"\AppData\Local\Packages\Microsoft.Windows.Cortana_cw5n1h2txyewy\LocalState\DeviceSearchCache"

' ------------------------------------------------------------------
' EXAMPLE 0002 CALL ROUTINE NOTHING PATH AND THEN MINIMAL ONE SHOT SET PATH WITHIN 
' ------------------------------------------------------------------
FOLDER_NAME=""

COUNTER_RR=0
COUNTER_TOTAL=0

CALL SCAN__TRIM_THE_DATE_OLDER_CAMERA(FOLDER_NAME)

MSGBOX "DELETE DONE "& VBCR & COUNTER_RR & VBCR & "OF" & VBCR & COUNTER_TOTAL

WSCRIPT.QUIT

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
		' FOLDER= "LET ME FIND A PATH"
		IF FOLDER="" THEN 
			FOLDER="E:\HARDWARE"
		END IF
		Set FOLDER = FSO.GETFOLDER(FOLDER)
		SCAN__TRIM_THE_DATE_OLDER_CAMERA(FOLDER)
		EXIT SUB
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
		STR_FIND="-- .JPG .MP4" 
		STR_FIND="-- .TXT" 
		PATH_FIND_01="\DCIM"
		PATH_FIND_02="\MP_ROOT"
		EXTENSION_PART=MID(FILE.path,INSTRREV(FILE.path,"."))
		PATH_FIND_RESULT=FALSE
		
		IF INSTR(STR_FIND , UCASE(EXTENSION_PART)) >0 THEN PATH_FIND_RESULT=TRUE
		' IF INSTR(UCASE(FILE.path) , PATH_FIND_01) >0 THEN PATH_FIND_RESULT=TRUE
		' IF INSTR(UCASE(FILE.path) , PATH_FIND_02) >0 THEN PATH_FIND_RESULT=TRUE
		
		IF PATH_FIND_RESULT=TRUE THEN
			ERR.CLEAR
			FILEMODDATE=0
			SET F = FSO.GETFILE(FILE.path)
				FILEMODDATE = F.DATELASTMODIFIED
				FILECREATEDATE = F.DATECREATED
			SET F=NOTHING
			' --------------------------------------------
			' EXAMPLE DELETE FILE OVER 48 HOUR IN THE PAST
			' --------------------------------------------
			' IF FILEMODDATE+TIMESERIAL(48,0,0)<NOW THEN 
			' --------------------------------------------
			' EXAMPLE DELETE FILE OVER 20 DAY IN THE PAST
			' --------------------------------------------
			IF FILEMODDATE+20<NOW THEN 
				ON ERROR RESUME NEXT
				ERR.CLEAR
				MSGBOX FILE.path
				'''' FSO.DeleteFile FILE.path
				COUNTER_RR=COUNTER_RR+1
			ELSE
			END IF
		END IF
		COUNTER_TOTAL=COUNTER_TOTAL+1
	NEXT

	' ---------------------------------------------------------------
	' CALL GET_DRIVE_SPACE(FOLDER.path)
	' ---------------------------------------------------------------
	
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
		SCAN__TRIM_THE_DATE_OLDER_CAMERA(SUBFOLDER)
    NEXT

END SUB
