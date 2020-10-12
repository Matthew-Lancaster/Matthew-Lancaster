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
		FOLDER="E:\HARDWARE"
		Set FOLDER = FSO.GETFOLDER(FOLDER)
		SCAN__TRIM_THE_DATE_OLDER_CAMERA(FOLDER)
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
		SCAN__TRIM_THE_DATE_OLDER_CAMERA(SUBFOLDER)
    NEXT

END SUB

