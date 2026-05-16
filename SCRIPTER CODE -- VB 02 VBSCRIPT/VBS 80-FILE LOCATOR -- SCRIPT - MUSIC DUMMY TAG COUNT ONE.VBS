' -------------------------------------------------------------------
' EXAMPLE PUT THIS FILE AS SCRIPT IN \FILELOCATORPRO.EXE
' VBS 80-FILE LOCATOR -- SCRIPT -  MUSIC DUMMY TAG COUNT ONE.VBS
' -------------------------------------------------------------------
' THE CODE WILL FIND FROM PATH 
' 0#00--*.MP3;0#01--*.MP3;0#02--*.MP3
' ANY OF HERE WILD-CARD FORMAT FOR FILE LOCATER
' ANY THESE WILD-CARD REPEAT MORE THAN ONE 
' EXCESS WILL BE GONE
' USE TO HAPPEN BY PROGRAM 
' WITH VB6
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' Sun 20-Sep-2020 21:11:11
' Sun 20-Sep-2020 22:02:00 -- 50 MINUTE
' -------------------------------------------------------------------


Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true
Const ForReading = 1, ForWriting = 2, ForAppending = 8

DIM FSO
SET FSO = CreateObject("Scripting.FileSystemObject")

DIM XX_INDEX
DIM XX
DIM PATH_STRING_PUT_OLD

DIM HOLD_PATH
DIM HOLD_NAME

DIM FILE_TEST_1(50)
DIM FILE_TEST_2(50)
DIM FILE_DRIVE_(50)
DIM PAD_LEVEL(50)
DIM FILE_TEST_ARRAY_INDEX
DIM WSHSHELL
DIM ARRAY_UBOUND


Set Sound = CreateObject("WMPlayer.OCX.7")
Sound.URL = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV"
Sound.Controls.play



XX_INDEX=0

FILE_TEST_ARRAY_INDEX=1
' -------------------------------------------------------------------
' RETEKESSMTB -- MUSIC AND TALKER BOOK
' -------------------------------------------------------------------
' I=0
' ' 0001
' I=I+1
' FILE_DRIVE_(I)="RETEKESSMTB"
' TEXT_INFO=FILE_DRIVE_(I) + "_MTB_MUSIC_&_TALK_BOOK"
' PAD_LEVEL(I)=4
' PAD_INDEX=LPad(I, "0", 2) 
' FILE_TEST_1(I)="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBSCRIPT 68__" & PAD_INDEX & "__" & GET_DATE_FORMATOR + "__" + TEXT_INFO + "__PLAYLIST_FULL____.TXT"
' FILE_TEST_2(I)="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBSCRIPT 68__" & PAD_INDEX & "__" & GET_DATE_FORMATOR + "__" + TEXT_INFO + "__PLAYLIST_HEADER__.TXT"

DIM GIVER_ONE_FILE_DONE_NOTEPAD

ARRAY_UBOUND=I

DIM CURRENT_DIRECTORY
CURRENT_DIRECTORY = fso.GetAbsolutePathName(".")
' -------------------------------------------------------------------
' AFTER RUN TIDY
' GET CURRENT_DIRECTORY TO FIND IF NOT RUN BY FILE_LOCATOR
' TAKE ACTION WHEN NOTEPAD++
' COPY SOME FILE
' -------------------------------------------------------------------
' EXAMPLE
' -------------------------------------------------------------------
' NOTEPAD++ --------- CURRENT_WORKING_DIRECTORY -- C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT
' FILE_LOCATOR.EXE -- CURRENT_WORKING_DIRECTORY -- DRIVE ASK TO SCAN FIRST OR MAYBE MULTIPLE
' -------------------------------------------------------------------
IF 1=2 THEN 
IF CURRENT_DIRECTORY="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT" THEN 
	FOR R1 = ASC("C") TO ASC("Z")
		Set FS = CreateObject("Scripting.FileSystemObject")
		DRIVE_LETTER=CHR(R1)+":"
		Set D = FS.GetDrive(FS.GetDriveName(FS.GetAbsolutePathName(DRIVE_LETTER)))
		ON ERROR RESUME NEXT
		FOR R2 = 1 TO ARRAY_UBOUND
			IF D.ISREADY THEN 
			IF UCASE(D.VOLUMENAME)=UCASE(FILE_DRIVE_(R2)) THEN
			' -------------------------------------------------------
				FILE_TEST_01=FILE_TEST_1(R2)
				FILE_TEST_02=FILE_TEST_2(R2)
				FILE_TEST_INDEX=R

				IF FSO.FILEEXISTS(FILE_TEST_01) THEN 
					PATH_1ST_MID_PART=MID(DRIVE_LETTER,1,2)
					X1=INSTRREV(FILE_TEST_01,"\")
					PATH_2ND_MID_PART=MID(FILE_TEST_01,X1+1)
					FULL_PATH_OUT=PATH_1ST_MID_PART & "\00_" & PATH_2ND_MID_PART
					FSO.CopyFile FILE_TEST_01, FULL_PATH_OUT

					' HERE NOT GET RUN 
					' FILE MAYBE GONE NOT BEGIN
					' -------------------------------------------------------------------
					' NOTEPAD++
					' -----------------------------------------------------------
					Set WSHShell = CreateObject("WScript.Shell")
					WSHShell.Run """C:\Program Files (x86)\Notepad++\notepad++.exe""" + " " + """"+FILE_TEST_01+""""
					Set WSHShell = Nothing
				END IF

				IF FSO.FILEEXISTS(FILE_TEST_02) THEN 
					PATH_1ST_MID_PART=MID(DRIVE_LETTER,1,2)
					X1=INSTRREV(FILE_TEST_02,"\")
					PATH_2ND_MID_PART=MID(FILE_TEST_02,X1+1)
					FULL_PATH_OUT=PATH_1ST_MID_PART & "\00_" & PATH_2ND_MID_PART
					FSO.CopyFile FILE_TEST_02, FULL_PATH_OUT

					' -------------------------------------------------------------------
					' NOTEPAD++
					' -----------------------------------------------------------
					Set WSHShell = CreateObject("WScript.Shell")
					WSHShell.Run """C:\Program Files (x86)\Notepad++\notepad++.exe""" + " " + """"+FILE_TEST_02+""""
					Set WSHShell = Nothing
				END IF

				IF FSO.FILEEXISTS(FILE_TEST_01) THEN 
					FILE_IN="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 68-FILE LOCATOR -- SCRIPT -  RETEKESS.VBS"
					PATH_1ST_MID_PART=MID(STRPATH,1,2)
					X1=INSTRREV(FILE_IN,"\",X1-1)
					PATH_2ND_MID_PART=MID(FILE_IN,X1+1)
					FULL_PATH_OUT=PATH_1ST_MID_PART & "\02_" & PATH_2ND_MID_PART
					FSO.CopyFile FILE_IN, FULL_PATH_OUT
				END IF
				


			' -------------------------------------------------------
			END IF
			END IF
		NEXT
	NEXT
END IF
END IF

' -------------------------------------------------------------------
' -------------------------------------------------------------------

' MSGBOX FILE_TEST_1(1)
' MSGBOX FILE_TEST_2(1)
' WSCRIPT.QUIT


Public Function isValidFileName( ByVal strPath, ByVal strFileName)

	isValidFileName=TRUE
	IF strPath=HOLD_PATH THEN 
	IF MID(strFileName,1,5)=MID(HOLD_NAME,1,5) THEN 
	IF UCASE(MID(strFileName,LEN(strFileName)-3))=".MP3" THEN 
	SET_DEL=FALSE
	' 0#00--*.MP3;0#01--*.MP3;0#02--*.MP3
	IF MID(strFileName,1,5)="0#00-" THEN SET_DEL=TRUE
	IF MID(strFileName,1,5)="0#01-" THEN SET_DEL=TRUE
	IF MID(strFileName,1,5)="0#02-" THEN SET_DEL=TRUE
	IF SET_DEL=TRUE THEN
		IF FSO.FILEEXISTS(strPath+strFileName) = TRUE THEN 
			FSO.DELETEFILE(strPath+strFileName)
			isValidFileName= FALSE
		END IF
	END IF
	END IF
	END IF
	END IF

	HOLD_PATH=strPath
	HOLD_NAME=strFileName

End Function




Public Function ARCHIVE_CODE_isValidFileName( ByVal strPath, ByVal strFileName)

	DIM FILE_TEST_01
	DIM FILE_TEST_02
	
	ON ERROR GOTO 0
	Set FS = CreateObject("Scripting.FileSystemObject")
	Set D = FS.GetDrive(FS.GetDriveName(FS.GetAbsolutePathName(STRPATH)))
	FOR R = 1 TO ARRAY_UBOUND
		IF UCASE(D.VOLUMENAME)=UCASE(FILE_DRIVE_(R)) THEN
			FILE_TEST_01=FILE_TEST_1(R)
			FILE_TEST_02=FILE_TEST_2(R)
			FILE_TEST_INDEX=R
		END IF
	NEXT

	IF XX_INDEX=0 THEN 
		SET FSO = CreateObject("Scripting.FileSystemObject")

		IF FSO.FILEEXISTS(FILE_TEST_01) = TRUE THEN 
			FSO.DELETEFILE(FILE_TEST_01)
		END IF
		IF FILE_TEST_01<>"" THEN 
			Set objFile = FSO.CreateTextFile(FILE_TEST_01,True)
			objFile.Close
		END IF
		IF FSO.FILEEXISTS(FILE_TEST_02) THEN 
			FSO.DELETEFile(FILE_TEST_02)
		END IF
		IF FILE_TEST_02<>"" THEN 
			Set objFile = FSO.CreateTextFile(FILE_TEST_02,True)
			objFile.Close
		END IF
		' -------------------------------------------------------------------
		' NOTEPAD++ HERE SCRIPT HAVE TO RUN AGAIN FROM NOTEPAD++ F5 KEY EXECUTE AFTER SCRIPT RUN FILELOCATOR
		' -----------------------------------------------------------
		Set WSHShell = CreateObject("WScript.Shell")
		SCRIPT_HERE="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 68-FILE LOCATOR -- SCRIPT -  RETEKESS.VBS"
		' FIND CODE TO GET SCRIPT NAME
		WSHShell.Run """C:\Program Files (x86)\Notepad++\notepad++.exe""" + " " + """"+SCRIPT_HERE+""""
		Set WSHShell = Nothing
		
		' Set FILE_1 = FSO.OpenTextFile(FILE_TEST_01, ForAppending)
		' Set FILE_2 = FSO.OpenTextFile(FILE_TEST_02, ForAppending)
		' FILE_1.Write LPad(XX, "0", PAD_LEVEL(FILE_TEST_INDEX)) +"__"+ CURRENT_DIRECTORY + VBCRLF
		' FILE_2.Write LPad(XX, "0", PAD_LEVEL(FILE_TEST_INDEX)) +"__"+ CURRENT_DIRECTORY + VBCRLF
	END IF
	ON ERROR RESUME NEXT

	XX_INDEX=XX_INDEX+1
	
	bIsValid = FALSE
	bIsValid = TRUE
	IF INSTR(strPath & "\" & strFileName,"_gsdata_")>0 THEN bIsValid = FALSE
	IF FSO.FOLDEREXISTS(strPath & "\" & strFileName) THEN bIsValid = FALSE
	
	isValidFileName = bIsValid

	IF isValidFileName=TRUE THEN 
		IF GIVER_ONE_FILE_DONE_NOTEPAD=0 THEN
			GIVER_ONE_FILE_DONE_NOTEPAD=1
		END IF
		Set FILE_1 = FSO.OpenTextFile(FILE_TEST_01, ForAppending)
		Set FILE_2 = FSO.OpenTextFile(FILE_TEST_02, ForAppending)
		XX=XX+1
		PATH_STRING=strPath + strFileName
		X1=INSTRREV(PATH_STRING,"\")
		X1=INSTRREV(PATH_STRING,"\",X1-1)
		PATH_STRING=MID(PATH_STRING,X1)

		PATH_STRING_2=strPath
		X1=INSTRREV(PATH_STRING_2,"\")
		X1=INSTRREV(PATH_STRING_2,"\",X1-1)
		PATH_STRING_2=MID(PATH_STRING_2,X1)
		PATH_STRING_2=MID(PATH_STRING_2,2)
		
		PATH_STRING_PUT=strPath

		IF "FULL_"="FULL" OR 1=1 THEN 
			' ---- WRITE --------------------------------------------
			FILE_1.Write LPad(XX, "0", PAD_LEVEL(FILE_TEST_INDEX)) +"__"+ PATH_STRING + VBCRLF
			' -------------------------------------------------------
		END IF 
		IF PATH_STRING_PUT_OLD<>PATH_STRING_PUT THEN 
			' ---- WRITE --------------------------------------------
			FILE_2.Write LPad(XX, "0", PAD_LEVEL(FILE_TEST_INDEX)) +"__"+ PATH_STRING_2 + VBCRLF
			' -------------------------------------------------------
		END IF 

		PATH_STRING_PUT_OLD=PATH_STRING_PUT
' 		File.Write CSTR(XX) +"  "+ PATH_STRING + VBCRLF

		IF GIVER_ONE_FILE_DONE_NOTEPAD=1 THEN
			GIVER_ONE_FILE_DONE_NOTEPAD=2
			' -------------------------------------------------------------------
			' NOTEPAD++
			' -----------------------------------------------------------
			Set WSHShell = CreateObject("WScript.Shell")
			WSHShell.Run """C:\Program Files (x86)\Notepad++\notepad++.exe""" + " " + """"+FILE_TEST_01+""""
			Set WSHShell = Nothing
			Set WSHShell = CreateObject("WScript.Shell")
			WSHShell.Run """C:\Program Files (x86)\Notepad++\notepad++.exe""" + " " + """"+FILE_TEST_02+""""
			Set WSHShell = Nothing
		END IF 

	END IF
End Function



FUNCTION GET_DATE_FORMATOR
	VAR_NOW=NOW
	YY=Year(VAR_NOW)
	MM2=MONTH(VAR_NOW)
	MMM = MonthName(MONTH(VAR_NOW),TRUE)
	DD=Day(VAR_NOW)
	HH24=Hour(VAR_NOW)
	MM=Minute(VAR_NOW)
	SS=Second(VAR_NOW)
	HH12=Hour(VAR_NOW)
	IF HH12>12 THEN 
		HH12=HH12-12
		AM_PM="PM"
	ELSE
		AM_PM="AM"
	END IF 
	IF HH12=12 THEN 
		AM_PM="PM"
	END IF 

	DDD=WeekdayName(Weekday(VAR_NOW),TRUE)

	i_TIME=""
	i_TIME=i_TIME+LPad(YY, "0", 4)+"-"
	i_TIME=i_TIME+MMM+"-"
	i_TIME=i_TIME+LPad(DD, "0", 2)
	i_TIME=i_TIME+" -- "
	i_TIME=i_TIME+LPad(HH24, "0", 2)+":"
	i_TIME=i_TIME+LPad(MM, "0", 2)+":"
	i_TIME=i_TIME+LPad(SS, "0", 2)
	i_TIME=i_TIME+" - "
	i_TIME=i_TIME+LPad(HH12, "0", 2)+":"
	i_TIME=i_TIME+LPad(MM, "0", 2)+":"
	i_TIME=i_TIME+LPad(SS, "0", 2)+"-"+AM_PM
	i_TIME=i_TIME+" - "
	i_TIME=i_TIME+DDD

	i_TIME=""
	i_TIME=i_TIME+LPad(YY, "0", 4)+"-"
	i_TIME=i_TIME+LPad(MM2, "0", 2)+"-"
	i_TIME=i_TIME+LPad(DD, "0", 2)
	' i_TIME=i_TIME+" -- "
	' i_TIME=i_TIME+LPad(HH24, "0", 2)+"-"
	' i_TIME=i_TIME+LPad(MM, "0", 2)
	' i_TIME=i_TIME+LPad(SS, "0", 2)

	GET_DATE_FORMATOR=i_TIME

End Function

' Wscript.Echo i_TIME
' WScript.Quit(0)

' -------------------------------------------------------------------
' END STOP ----------------------------------------------------------
' -------------------------------------------------------------------

Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

