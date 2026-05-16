Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

' FILE PATH SEARCH STRING
' "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__#NFS_EX__.TXT"
' $(AllLocalDrives)\

' $(AllLocalDrives)\;C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__#NFS_EX__.TXT


	
DIM VAR_ONCE

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")



DIM XXXXXXXXXXXX1
DIM XXXXXXXXXXXX2

Public Function isValidFileName( ByVal strPath, ByVal strFileName )

	' On Error Resume Next
	' bIsValid = False
	' IF EXTNAME=".LOG" AND ISNUMERIC(F4) THEN bIsValid = TRUE
	' IF EXTNAME=".LOG" AND F8="GOODSYNC" THEN bIsValid = TRUE
	' IF EXTNAME=".LOG" AND F8="SENDSMTP" THEN bIsValid = TRUE

	bIsValid = FALSE
	' THS WILL DELETE WAV FILE ALSO *.WAV;*.MP3
	' IF LEN(strFileName)=LEN("2010-04-30-09-44-08.MP3") THEN bIsValid = TRUE
	EXTNAME=UCASE(MID(strFileName,INSTRREV(strFileName,".")))

	' TWO FOLDER DRILL DOWN
	' "O:\0 00 VIDEO SNAPSHOT CCSE HIKVISION\snapshot (58)\20191114\D1\"
	' "20191114\D1\"

	LAST_TWO_PATH1=""
	A=strPath
	If Mid(A, Len(A),1) <> "\" Then A = A + "\"

	' MsgBox A & vbCr & Mid$(A, Len(A))
	XXXXXXXXXXXX1 = InStrRev(A, "\", Len(A) - 1)
	XXXXXXXXXXXX2 = InStrRev(A, "\", XXXXXXXXXXXX1 - 1)
	LAST_TWO_PATH1 = UCase(Right(A, Len(A) - XXXXXXXXXXXX2))
	
	' MsgBox LAST_TWO_PATH1

	IF EXTNAME=".JPG" AND LEN(strFileName)=LEN("DVR_A03_Motion_Detected_20180719122121_1.JPG") THEN
		IF INSTR(UCASE(strPath),"SNAPSHOT")>0  AND  INSTR(UCASE(strFileName),"DVR_")>0  THEN
			IF (LEN(LAST_TWO_PATH1)=LEN("20191114\D1\") OR LEN(LAST_TWO_PATH1)=LEN("20191114 (2)\D1\") ) THEN 
				bIsValid = TRUE 
			END IF
		END IF
	END IF

	
	


	' IF LEN(strFileName)=LEN("1e0c1664ee09b3bb692445ae4d107b91.xml") THEN bIsValid = TRUE
	
		
	IF EXTNAME=".JPG" AND INSTR(UCASE(strPath),"APP AND SCREEN")>0 AND INSTR(UCASE(strPath),"AUTO")>0 THEN 
		bIsValid = TRUE
	END IF
	


	IF EXTNAME=".MP3" AND INSTR(UCASE(strPath),"KAT MP3")>0 THEN 
		bIsValid = FALSE
		IF LEN(strFileName)=LEN("2010-04-30-09-44-08.MP3") THEN bIsValid = TRUE
		IF INSTR(strPath,"8 TEN ELEVEN")>0 THEN bIsValid = FALSE
		IF INSTR(strPath,"POLICE")>0 THEN bIsValid = FALSE
		IF INSTR(strPath,"COMPLAINER")>0 THEN bIsValid = FALSE
	END IF
	
	IF bIsValid = TRUE THEN 
		' FSO.DELETEFILE(strPath & "\" & strFileName)
	END IF
	' bIsValid = TRUE 
	' bIsValid = TRUE
	
	' IF VAR_ONCE=1000 THEN 
	' IF strFileName="FILE_LOCATOR_DUMMY_FILE_TO_RUN_SCRIPT_DO_OTHER.txt" THEN
		' bIsValid = FALSE

		' ' -----------------------------------------------------------
		' ' RUN FROM BAT 
		' ' -----------------------------------------------------------
		' ' XF0="C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.BAT"
		' ' Dim WSHShell
		' ' Set WSHShell = CreateObject("WScript.Shell")
			' ' WSHShell.Run """"+XF0+"""" , 1, TRUE
		' ' EXIT FUNCTION
		
		' ' -----------------------------------------------------------
		' ' RUN DIRECT POWERSHELL
		' ' -----------------------------------------------------------
		' FILE_NAME="""C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1"""
		' Dim WSHShell
		' Set WSHShell = CreateObject("WScript.Shell")
		' WSHShell.run("powershell -file "+FILE_NAME), 0, TRUE
		' EXIT FUNCTION
	' END IF
	' END IF

	' ' DELETE PATH FIND STRING QUICK WAY THAN EXPLORER WHEN LOT OF THEM
	' ' ----------------------------------------------------------------
	' IF INSTR(strPath,            "_saved_")>0 THEN
	' IF INSTR(strPath,            "D:\0 CLOUD\_gsdata_\_saved_\GD-INSYNC\rub.rim@gmail.com")>0 THEN
	' 'IF INSTR(UCASE(strFileName), ".JPG")>0 THEN
		' TIME_GET=DATE & " " & MID(TIME,1,2)
		' IF TIME_GET="03/10/2020 20" THEN 
			' FSO.DELETEFILE(strPath & "\" & strFileName)
		' END IF
	' 'END IF
	' END IF
	' END IF

	
	' IF INSTR(strPath,"_saved_")>0 THEN bIsValid = True
	' IF INSTR(strPath,"_history_")>0 THEN 
		' bIsValid = True
	' END IF
	' IF INSTR(strPath,"_gsdata_")>0 THEN 
		' bIsValid = True
		' EXTNAME=UCASE(MID(strFileName,INSTRREV(strFileName,".")))
		' ' CHECK FOR "._GS" ".GSS" ".GSL" ".LOG"
	' END IF
	' IF INSTR(strPath,"\TEMP\")>0 THEN 
		' DIM COMPNAME(7)
		' COMPNAME(1)="-1-ASUS-X5DIJ"
		' COMPNAME(2)="-2-ASUS-EEE"
		' COMPNAME(3)="-3-LINDA-PC"
		' COMPNAME(4)="-4-ASUS-GL522VW"
		' COMPNAME(5)="-5-ASUS-P2520LA"
		' COMPNAME(6)="-7-ASUS-GL522VW"
		' COMPNAME(7)="-8-MSI-GP62M-7RD"
		' FOR R=1 TO 7
			' IF INSTR(strPath,COMPNAME(R))>0 THEN   
				' bIsValid = True
				' EXIT FOR
			' END IF
		' NEXT 
		' EXTNAME=UCASE(MID(strFileName,INSTRREV(strFileName,".")))
		' ' CHECK FOR ".EXE" -- "insync-upd"
	' END IF
	' IF bIsValid = True THEN 
		' IF INSTR(strFileName,".gdcache")>0    THEN bIsValid = False
		' IF INSTR(strFileName,"insync-upd")>0  THEN bIsValid = False
		' IF EXTNAME="._GS" THEN                     bIsValid = False
		' IF EXTNAME=".GSS" THEN                     bIsValid = False
		' IF EXTNAME=".GSL" THEN                     bIsValid = False
		' IF EXTNAME=".LOG" THEN                     bIsValid = False
		' IF EXTNAME=".EXE" THEN                     bIsValid = False
	' END IF
	
	' ' D:\GOODSYNC__BACK\_gsdata_\_history_\Profile\jobs-groups-options-2020-0627-022318_2020-0627-194941.tic
	' IF INSTR(strFileName,"jobs-groups-options")>0    THEN
	' IF INSTR(strFileName,".tic")>0  THEN 
		' bIsValid = False
	' END IF
	' END IF
	
	' ' D:\GOODSYNC__BACK\_gsdata_\_history_\Profile\accounts-bookmarks_2020-0711-120833.tic
	' IF INSTR(strPath,"\GOODSYNC__BACK\")>0    THEN
		' bIsValid = False
	' END IF
	
	isValidFileName = bIsValid
	' isValidFileName = NOT bIsValid -- INVERT THEM

End Function


PUBLIC FUNCTION TEST_ROUTINE_NOT_RUN_TWO_OF_SOMETHING_IS_COUNTER_HERE_MIGHT_WORK
	bIsValid = TRUE
	' C:\0 00 LINK SET QUICKER MOVER\# 0\FILE_LOCATOR_DUMMY_FILE_TO_RUN_SCRIPT_DO_OTHER\
	
	' IF SearchParms.LookIn="C:\0 00 LINK SET QUICKER MOVER\# 0\;$(AllLocalDrives)\;=C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\PS 04 - GET_ALL_SYSTEM_NETWORK DRIVE.PS1__#NFS_EX__.TXT" THEN
	
	IF INSTR(SearchParms.LookIn,"C:\0 00 LINK SET QUICKER MOVER\# 0\")>0  THEN
		bIsValid = FALSE
		isValidFileName = bIsValid
		EXIT FUNCTION
	END IF

	isValidFileName = bIsValid
	EXIT FUNCTION
	
	' strCustom = 
	
END FUNCTION
