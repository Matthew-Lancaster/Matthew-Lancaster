'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 35-RENAMER VB6 _VBP LCASE.VBS
'# __ 
'# __ VBS 35-RENAMER VB6 _VBP LCASE.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START      TIME [ Thu 18-Oct-2018 13:28:00 ] COPY FROM ANOTHER CODE
'# END LAST EDITOR [ Thu 18-Oct-2018 14:32:04 ] 
'# __ 
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
'# ------------------------------------------------------------------
' ON GOING PROBLEM 
' NOT FOUND ANSWER YET
' BUT SOMETHING KEEP PUT THE EXTENSION FOR _.VBP INTO UPPER CASE
' MAKE THE SYNC TO GOOGLE DRIVE GO WONKY
' -------------------------------------------------------------------
' -------------------------------------------------------------------

'# ------------------------------------------------------------------
' LOCATION ON-LINE 
'# ------------------------------------------------------------------
' Matthew-Lancaster/VBS 30-FUJI CAMERA IMAGE DATE FOLDER CREATE & MOVE.VBS at master · Matthew-Lancaster/Matthew-Lancaster
' https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20VBS/VBS%2030-FUJI%20CAMERA%20IMAGE%20DATE%20FOLDER%20CREATE%20%26%20MOVE.VBS
' -------------------------------------------------------------------

'#-------------------------------------------------------------------
' 001 __ 
'#-------------------------------------------------------------------
' ADDER
' ADD CODE
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' Thu 18-Oct-2018 13:28:00
' Thu 18-Oct-2018 14:32:04 __ 
'#-------------------------------------------------------------------


Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

DIM STOP_PROGRESS
DIM IX_ERROR_MESSENGER

DIM DRIVE_SPACE_INFO

DIM objFolder
DIM VAR_PATH_SCANNER
DIM objSuperFolder

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
DIM objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")

Dim oDict_1 ' Key: Full filename, Item: Filename
Dim oDict_2 ' Key: Full filename, Item: Filename

DIM I_COUNT_1
DIM I_COUNT_2
DIM I_COUNT_3
DIM I_COUNT_4
DIM I_COUNT_5
DIM I_COUNT_7

I_COUNT_1=0
I_COUNT_2=0
I_COUNT_3=0
I_COUNT_4=0
I_COUNT_5=0
I_COUNT_7=0

Dim FileModDate
DIM O_FileModDate
DIM SET_VAR_MODE_ORDER

DIM FOLDER_PATH_VAR_1
DIM FOLDER_PATH_VAR_2
DIM FOLDER_PATH_VAR_3

DIM FindFirstFile_VAR
DIM FindMostRecentFile_VAR

DIM I_SET_DONE

Set args = Wscript.Arguments
For Each arg In args
	Wscript_Arguments = Wscript_Arguments + arg
Next

FOLDER_PATH_VAR_1="D:\VB6\VB-NT\00_Best_VB_01"

DIM FOLDER_PATH_FILTER()
REDIM FOLDER_PATH_FILTER(1)
FOLDER_PATH_FILTER(1)="\#0 AllApi\"

IF FSO.FolderExists(FOLDER_PATH_VAR_1)=FALSE THEN 
	MSGBOX "FOLDER NOT EXIST"+VBCRLF+VBCRLF+FOLDER_PATH_VAR_1,vbMsgBoxSetForeground,Wscript.ScriptName 
	Wscript.QUIT
END IF


' -------------------------------------------------------------------
' -------------------------------------------------------------------

SET_VAR_MODE_ORDER=1
FOLDER_PATH_VAR=FOLDER_PATH_VAR_1
CALL MAIN

' CHANGES MADE ON 1ST RUN

' D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS\Calender\Calender Ace.VBP
' D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.VBP
' D:\VB6\VB-NT\00_Best_VB_01\Drives Gig\Drives_Gig.VBP
' D:\VB6\VB-NT\00_Best_VB_01\Drives Gig_MINIMIZE VERSION\Drives_Gig.VBP
' D:\VB6\VB-NT\00_Best_VB_01\Shell Explorer Loader\Shell Explorer Loader.VBP
' D:\VB6\VB-NT\00_Best_VB_01\Winamp - Volume MINI\VolumeBar MINI.VBP
' D:\VB6\VB-NT\00_Best_VB_01\WinAmp MP3%\WinAmp MP3%-VideoBar.VBP

' D:\VB6\VB-NT\00_Best_VB_01\#0 AllApi\vbedit\VBEDIT.VBP

' D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS\Alarm\ALARM.VBP
' D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS\Calc\CALC.VBP
' D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS\Calc - ORIGINAL\CALC.VBP
' D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS\Calc - TEXT_&_HEX\CALC - TEXT_&_HEX.VBP
' D:\VB6\VB-NT\00_Best_VB_01\METRONOME\METRONOME.VBP
' D:\VB6\VB-NT\00_Best_VB_01\MP3 PLAYED\MP3 PLAYED.VBP
' D:\VB6\VB-NT\00_Best_VB_01\MPH AND DISTANCE\MP3 PLAYED\MP3 PLAYED.VBP
' D:\VB6\VB-NT\00_Best_VB_01\REBOOT CLOSE PROGRAMS\REBOOT CLOSE PROGRAMS.VBP
' D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.VBP
' D:\VB6\VB-NT\00_Best_VB_01\VU_METER_LOGGER\VU METER LOGGER.VBP
' D:\VB6\VB-NT\00_Best_VB_01\Wmi_CPU\WMICPU1 BIG.VBP

I=""
I=I+"RENAME VBASIC __FROM__" + VBCRLF 
I=I+"UCASE _.VBP" + VBCRLF 
I=I+"__TO__" + VBCRLF 
I=I+"LCASE _.vbp" + VBCRLF 
I=I+""+ VBCRLF 
I=I+ FOLDER_PATH_VAR_1 +"\*.VBS"+ VBCRLF 
I=I+"----------------------------------------------------------------"+ VBCRLF 
I=I+"RENAME _.vbs FILE COUNT" + VBCRLF + CSTR(I_COUNT_1) + VBCRLF 
I=I+"TOTAL _.vbs EXTENSION COUNT" + VBCRLF + CSTR(I_COUNT_2) + VBCRLF 
I=I+"ERROR RENAME COUNT" + VBCRLF + CSTR(I_COUNT_3) + VBCRLF 
I=I+""+ VBCRLF 
I=I+I_SET_DONE

IF Wscript_Arguments="" THEN 
	MSGBOX I,vbMsgBoxSetForeground,Wscript.ScriptName 
END IF

SET fso = NOTHING
SET objFSO = NOTHING
SET WSHShell = NOTHING

Wscript.QUIT
' -------------------------------------------------------------------
' -------------------------------------------------------------------

' -------------------------------------------------------------------
Sub Main
    Set oDict_1 = CreateObject("Scripting.Dictionary")
	Set oDict_2 = CreateObject("Scripting.Dictionary")
	IF SET_VAR_MODE_ORDER=1 THEN 
		' Call Recursive Sub
		FindWantedFiles_5(objFSO.GetFolder(FOLDER_PATH_VAR))
		' Display All Findings from Dictionary object
		DisplayFindings_5
	END IF

    Set oDict_1 = Nothing
    Set oDict_2 = Nothing
End Sub

' -------------------------------------------------------------------
Sub FindWantedFiles(sFolder)
    On Error Resume Next
    Dim oFDR, oItem, file
    ' Check if wanted files are in this folder
    For Each file In sFolder.Files
		oDict_1.Add FILE.path, file
    Next
    ' Recurse into it's sub folders
    For Each oFDR In sFolder.SubFolders
		oDict.Add objFSO.GetFolder(oFDR.Path), ""
        FindWantedFiles objFSO.GetFolder(oFDR.Path)
    Next
End Sub

' -------------------------------------------------------------------
Sub FindWantedFiles_2(sFolder)
    On Error Resume Next
    DIM FILE_FOLDER_COUNTER
	Dim oFDR, oItem, file
    ' Check if wanted files are in this folder
    For Each file In sFolder.Files
		oDict_2.Add FILE.path, file
    Next
    ' Recurse into it's sub folders
    For Each oFDR In sFolder.SubFolders
		' -----------------------------------------------------------
		' REMEMBER HERE USE THIS RESULT TWICE ONCE TO STORE FOLDER PATH 
		' AND THEN USE FOLDER PATH FOR RECURSIVE
		' AND THEN GET ALL FILE AND FOLDER _ KINDA TRICKY
		' -----------------------------------------------------------
		' MSGBOX objFSO.GetFolder(oFDR.Path)
		oDict_2.Add objFSO.GetFolder(oFDR.Path), ""
        FindWantedFiles_2 objFSO.GetFolder(oFDR.Path)
    Next
End Sub

' -------------------------------------------------------------------
Sub FindWantedFiles_5(sFolder)
    On Error Resume Next
    DIM FILE_FOLDER_COUNTER
	Dim oFDR, oItem, file
    ' Check if wanted files are in this folder
    For Each file In sFolder.Files
		
		FILE_STRIP_PATH=MID(FILE,INSTRREV(FILE,"\")+1)
		IF UCASE(MID(file,LEN(file)-3))=".VBP" THEN 
			I_COUNT_2=I_COUNT_2+1
		END IF

		SET_FILTER_GO=FALSE
		FOR R_COUNT=1 TO UBOUND(FOLDER_PATH_FILTER)
			IF INSTR(FILE,FOLDER_PATH_FILTER(R_COUNT)) THEN SET_FILTER_GO=TRUE
		NEXT
		
		' IN REM IF FILENAME IS ALL UCASE THEN IGNORE
		'IF UCASE(FILE_STRIP_PATH)<>FILE_STRIP_PATH THEN 
		
		IF SET_FILTER_GO=FALSE THEN 
			IF LCASE(MID(file,LEN(file)-3))= ".vbp" THEN 
				IF MID(file,LEN(file)-3)<>".vbp" THEN 
					I_COUNT_1=I_COUNT_1+1
					oDict_1.Add FILE.path, file
				END IF
			END IF
		END IF
		'END IF
    Next
    ' Recurse into it's sub folders
    For Each oFDR In sFolder.SubFolders
		' -----------------------------------------------------------
		' REMEMBER HERE USE THIS RESULT TWICE ONCE TO STORE FOLDER PATH 
		' AND THEN USE FOLDER PATH FOR RECURSIVE
		' AND THEN GET ALL FILE AND FOLDER _ KINDA TRICKY
		' -----------------------------------------------------------
		' MSGBOX objFSO.GetFolder(oFDR.Path)
		' oDict_1.Add objFSO.GetFolder(oFDR.Path), ""
        FindWantedFiles_5 objFSO.GetFolder(oFDR.Path)
    Next
End Sub


' -------------------------------------------------------------------
' FindFirstFile(sFolder) NOT USED ANYMORE SEEM BETTER TO USE FindMostRecentFile(sFolder)
' -------------------------------------------------------------------
Sub FindFirstFile(sFolder)
    On Error Resume Next
    DIM FILE_FOLDER_COUNTER
	Dim oFDR, oItem, file

	IF FindFirstFile_VAR<>"" THEN EXIT SUB

    For Each file In sFolder.Files
		IF FSO.FILEEXISTS(FILE) = TRUE THEN 
			FindFirstFile_VAR = FILE
			EXIT FOR
		END IF
    Next

	IF FindFirstFile_VAR<>"" THEN EXIT SUB

    ' Recurse into it's sub folders
    For Each oFDR In sFolder.SubFolders
		' -----------------------------------------------------------
		' IF WANT FOLDER ALSO
		' REMEMBER HERE USE THIS RESULT TWICE ONCE TO STORE FOLDER PATH 
		' AND THEN USE FOLDER PATH FOR RECURSIVE
		' AND THEN GET ALL FILE AND FOLDER _ KINDA TRICKY
		' -----------------------------------------------------------
		' oDict_2.Add objFSO.GetFolder(oFDR.Path), ""
        FindFirstFile objFSO.GetFolder(oFDR.Path)
    Next
End Sub

' -------------------------------------------------------------------
' TO USE WITH ---- WSHShell.Run "EXPLORER /select, "
' -------------------------------------------------------------------
Sub FindMostRecentFile(sFolder)
	O_FileModDate=0
	FindMostRecentFile_VAR=""
	FindMostRecentFile_2(objFSO.GetFolder(sFolder))
END SUB

' -------------------------------------------------------------------
Sub FindMostRecentFile_2(sFolder)
    On Error Resume Next
    DIM FILE_FOLDER_COUNTER
	Dim oFDR, oItem, file

    For Each file In sFolder.Files
		Set f = fso.GetFile(file)
		FileModDate = f.DateLastModified
		SET f=NOTHING
		IF FileModDate>O_FileModDate THEN 
			O_FileModDate=FileModDate
			FindMostRecentFile_VAR = FILE
		END IF
    Next

    ' Recurse into it's sub folders
    For Each oFDR In sFolder.SubFolders
		' -----------------------------------------------------------
		' IF WANT FOLDER ALSO
		' REMEMBER HERE USE THIS RESULT TWICE ONCE TO STORE FOLDER PATH 
		' AND THEN USE FOLDER PATH FOR RECURSIVE
		' AND THEN GET ALL FILE AND FOLDER _ KINDA TRICKY
		' -----------------------------------------------------------
		oDict_2.Add objFSO.GetFolder(oFDR.Path), ""
        FindMostRecentFile_2 objFSO.GetFolder(oFDR.Path)
    Next
End Sub

' -------------------------------------------------------------------
Sub DisplayFindings_1()
    Dim oKeys, oKey
    oKeys = oDict_1.Keys
    For Each oKey In oKeys
		VAR_PATH_SCANNER=oKey
		CALL PATH_SCANNER_FOLDER_DATE_CREATE
	Next
End Sub

' -------------------------------------------------------------------
Sub DisplayFindings_2()
    Dim oKeys, oKey
    oKeys = oDict_1.Keys
    For Each oKey In oKeys
		VAR_PATH_SCANNER=oKey
		CALL PATH_SCANNER_FOLDER_MOVE_FROM_IMPORT
	Next
End Sub

' -------------------------------------------------------------------
Sub DisplayFindings_3()
    Dim oKeys, oKey
    oKeys = oDict_2.Keys
    For Each oKey In oKeys
		VAR_PATH_SCANNER=oKey
		PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
		
		IF FSO.FileExists(VAR_PATH_SCANNER) Then 
			ON ERROR RESUME NEXT
			ERR.CLEAR
			FSO.DeleteFile VAR_PATH_SCANNER
			I_COUNT_1=I_COUNT_1+1
			IF ERR.NUMBER>0 THEN 
				MSGBOX ERR.DESCRIPTION & VBCRLF & VBCRLF & "UNABLE TO DELETE" & VBCRLF & VBCRLF & VAR_PATH_SCANNER ,vbMsgBoxSetForeground,Wscript.ScriptName 
			END IF
			ON ERROR GOTO 0
		END IF
		
		IF FSO.FolderExists(VAR_PATH_SCANNER) Then 
			IF FolderEmpty(VAR_PATH_SCANNER) THEN
				ON ERROR RESUME NEXT
				ERR.CLEAR
				FSO.DeleteFolder VAR_PATH_SCANNER
				I_COUNT_2=I_COUNT_2+1
				IF ERR.NUMBER>0 THEN 
					MSGBOX ERR.DESCRIPTION & VBCRLF & VBCRLF & "UNABLE TO DELETE" & VBCRLF & VBCRLF & VAR_PATH_SCANNER ,vbMsgBoxSetForeground,Wscript.ScriptName 
				END IF
				ON ERROR GOTO 0
			END IF
		END IF

	Next
End Sub

' -------------------------------------------------------------------
Sub DisplayFindings_4()
    Dim oKeys, oKey
    oKeys = oDict_2.Keys
    For Each oKey In oKeys
		VAR_PATH_SCANNER=oKey
		PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
		
		IF FSO.FileExists(VAR_PATH_SCANNER) Then 
			Set f = fso.GetFile(VAR_PATH_SCANNER)
			FileModDate = f.DateLastModified
			FileCreateDate = f.DateCreated 
			SET f=NOTHING
			IF FileCreateDate+TIMESERIAL(1,20,0)<NOW THEN 
				' MSGBOX VAR_PATH_SCANNER
				ON ERROR RESUME NEXT
				ERR.CLEAR
				FSO.DeleteFile VAR_PATH_SCANNER
				I_COUNT_1=I_COUNT_1+1
				IF ERR.NUMBER>0 AND Wscript_Arguments="" THEN 
					MSGBOX ERR.DESCRIPTION & VBCRLF & VBCRLF & "UNABLE TO DELETE" & VBCRLF & VBCRLF & VAR_PATH_SCANNER ,vbMsgBoxSetForeground,Wscript.ScriptName 
				END IF
				ON ERROR GOTO 0
			END IF
		END IF
		
		IF FSO.FolderExists(VAR_PATH_SCANNER) Then 
		SET_GO=TRUE
		FOR R_COUNT= 1 TO UBOUND(FOLDER_PATH_FILTER)
			IF VAR_PATH_SCANNER=FOLDER_PATH_FILTER(R_COUNT) THEN
				SET_GO=FALSE
			END IF
		NEXT
			END IF
		
		IF SET_GO=TRUE THEN 
			IF FSO.FolderExists(VAR_PATH_SCANNER) Then 
				IF FolderEmpty(VAR_PATH_SCANNER) THEN
					Set f = fso.GetFolder(VAR_PATH_SCANNER)
					FolderModDate = f.DateLastModified
					FolderCreateDate = f.DateCreated 
					SET f=NOTHING
					IF FolderCreateDate+TIMESERIAL(24,0,0)<NOW THEN 
						' MSGBOX VAR_PATH_SCANNER
						ON ERROR RESUME NEXT
						ERR.CLEAR
						FSO.DeleteFolder VAR_PATH_SCANNER
						I_COUNT_2=I_COUNT_2+1
						IF ERR.NUMBER>0  AND Wscript_Arguments="" THEN 
							MSGBOX ERR.DESCRIPTION & VBCRLF & VBCRLF & "UNABLE TO DELETE" & VBCRLF & VBCRLF & VAR_PATH_SCANNER 
						END IF
						ON ERROR GOTO 0
					END IF
				END IF
			END IF
		END IF
	Next
End Sub


' -------------------------------------------------------------------
Sub DisplayFindings_5()
    Dim oKeys, oKey
    oKeys = oDict_1.Keys
    For Each oKey In oKeys
		VAR_PATH_SCANNER=oKey
		PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
		
		IF FSO.FileExists(VAR_PATH_SCANNER) Then 
			ON ERROR RESUME NEXT
			ERR.CLEAR
			
			VAR_PATH_SCANNER_RENAME= VAR_PATH_SCANNER
			VAR_PATH_SCANNER_RENAME= MID(VAR_PATH_SCANNER_RENAME,1,LEN(VAR_PATH_SCANNER_RENAME)-4)+".vbp"
			
			Fso.MoveFile VAR_PATH_SCANNER, VAR_PATH_SCANNER_RENAME

			' MSGBOX VAR_PATH_SCANNER + VBCRLF + VAR_PATH_SCANNER_RENAME

			I_SET_DONE=I_SET_DONE+VAR_PATH_SCANNER+VBCRLF
			IF ERR.NUMBER>0 AND Wscript_Arguments="" THEN 
				I_COUNT_3=I_COUNT_3+1
				MSGBOX ERR.DESCRIPTION & VBCRLF & VBCRLF & "UNABLE TO RENAME" & VBCRLF & VBCRLF & VAR_PATH_SCANNER ,vbMsgBoxSetForeground,Wscript.ScriptName 
			END IF
			ON ERROR GOTO 0
		END IF
	Next
End Sub


' -------------------------------------------------------------------
Function FolderEmpty(strFolder)
	Dim objFolder
	If FSO.FolderExists(strFolder) Then
		Set objFolder = FSO.GetFolder(strFolder)
		If objFolder.Files.Count = 0 And objFolder.SubFolders.Count = 0 Then
			FolderEmpty=True
		Else
			FolderEmpty=False
		End If
	End If
End Function

' -------------------------------------------------------------------
SUB PATH_SCANNER_FOLDER_MOVE_FROM_IMPORT

	SET_GO_FILE=FALSE
	IF INSTR(Ucase(VAR_PATH_SCANNER),".MOV")>0 then SET_GO_FILE=TRUE
	IF INSTR(Ucase(VAR_PATH_SCANNER),".JPG")>0 then SET_GO_FILE=TRUE
	IF INSTR(Lcase(VAR_PATH_SCANNER),"_gsdata_")>0 then SET_GO_FILE=FALSE
	
	IF SET_GO_FILE=FALSE THEN EXIT SUB

	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)
	
	PATH_VAR_2=REPLACE(VAR_PATH_SCANNER,FOLDER_PATH_VAR,"")
	PATH_VAR_2=MID(PATH_VAR_2,1,INSTRREV(PATH_VAR_2,"\"))

	NewFilename=PATH_VAR_2+FILE_VAR

	If Not FSO.FolderExists(PATH_VAR_2) Then
		Set objFolder = FSO.CreateFolder(PATH_VAR_2)
		Set objFolder = NOTHING	
		I_COUNT_3=I_COUNT_3+1		
	End If

	ON ERROR RESUME NEXT
	ERR.CLEAR
	Fso.MoveFile VAR_PATH_SCANNER, NewFilename
	I_COUNT_4=I_COUNT_4+1		
	IF ERR.NUMBER>0 THEN 
		MSGBOX ERR.DESCRIPTION & VBCRLF & VBCRLF & "FROM" & VBCRLF & VAR_PATH_SCANNER & VBCRLF & VBCRLF & "TO" & VBCRLF & NewFilename,vbMsgBoxSetForeground,Wscript.ScriptName 
		Wscript.QUIT
	END IF
	
End Sub

' -------------------------------------------------------------------
SUB PATH_SCANNER_FOLDER_DATE_CREATE

	SET_GO_FILE=FALSE
	if INSTR(Ucase(VAR_PATH_SCANNER),".JPG")>0 then SET_GO_FILE=TRUE
	if INSTR(Lcase(VAR_PATH_SCANNER),"_gsdata_")>0 then SET_GO_FILE=FALSE
	
	IF SET_GO_FILE=FALSE THEN EXIT SUB

	Set f = fso.GetFile(VAR_PATH_SCANNER)
	FileModDate = f.DateLastModified
	SET f=NOTHING

	i_TIME=""
	i_TIME=i_TIME+LPad(Year(FileModDate), "0", 4)+" "
	i_TIME=i_TIME+LPad(Month(FileModDate), "0", 2)+" "
	i_TIME=i_TIME+LPad(Day(FileModDate), "0", 2)

	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)
	
	PATH_VAR_2=PATH_VAR+i_TIME

	NewFilename=PATH_VAR_2+"\"+FILE_VAR
	
	if INSTR((VAR_PATH_SCANNER),"\"+i_TIME)=0 then
		If Not FSO.FolderExists(PATH_VAR_2) Then
			Set objFolder = FSO.CreateFolder(PATH_VAR_2)
			Set objFolder = NOTHING	
			I_COUNT_1=I_COUNT_1+1		
		End If

		ON ERROR RESUME NEXT
		ERR.CLEAR
		Fso.MoveFile VAR_PATH_SCANNER, NewFilename
		I_COUNT_2=I_COUNT_2+1		
		IF ERR.NUMBER>0 THEN 
			MSGBOX ERR.DESCRIPTION & VBCRLF & VBCRLF & "FROM" & VBCRLF & VAR_PATH_SCANNER & VBCRLF & VBCRLF & "TO" & VBCRLF & NewFilename,vbMsgBoxSetForeground,Wscript.ScriptName 
			Wscript.QUIT
		END IF
	END IF
End Sub

' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' ROUTINES AFTER THIS POINT HERE ARE NOT USED FOR MAIN RUN PROCESS OF 
' THE SCRIPT WRITING
' BUT INCLUDE IF THEY ARE GOING TO BE USEFUL LATER
' IT WAS TAKEN COPY FROM ANOTHER CODE
' C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 29-COPY CAMERA PHOTO IMAGES.VBS
'
' TEMP RUN IF WANT TO JUST A REANME-ER USE ONCE
' -------------------------------------------------------------------
SUB PATH_SCANNER_FOLDER_FUJI_RENAME

	SET_GO_FILE=FALSE
	if INSTR(Ucase(VAR_PATH_SCANNER),".JPG")>0 then SET_GO_FILE=TRUE
	if INSTR(Lcase(VAR_PATH_SCANNER),"_gsdata_")>0 then SET_GO_FILE=FALSE
	
	IF SET_GO_FILE=FALSE THEN EXIT SUB

	NewFilename=REPLACE(VAR_PATH_SCANNER,"FUJIFILM-FinePix XP90 XP91 XP95", "FUJI-FinePix XP90")

	ON ERROR RESUME NEXT
	ERR.CLEAR
	Fso.MoveFile VAR_PATH_SCANNER, NewFilename
	I_COUNT_2=I_COUNT_2+1		
	IF ERR.NUMBER>0 THEN 
		MSGBOX ERR.DESCRIPTION & VBCRLF & VAR_PATH_SCANNER & VBCRLF & NewFilename,vbMsgBoxSetForeground,Wscript.ScriptName 
		Wscript.QUIT
	END IF
End Sub

' -------------------------------------------------------------------
SUB PATH_SCANNER_COPY_CAMERA
	SET_GO_FILE=FALSE
	if mid(ucase(VAR_PATH_SCANNER),3,5)="\DCIM" then    SET_GO_FILE=TRUE
	if mid(ucase(VAR_PATH_SCANNER),3,8)="\MP_ROOT" then SET_GO_FILE=TRUE
	if INSTR(Lcase(VAR_PATH_SCANNER),"_gsdata_")>0 then SET_GO_FILE=FALSE
	
	IF SET_GO_FILE=FALSE THEN EXIT SUB
	
	I_COUNT_2=I_COUNT_2+1
	
	IF STOP_PROGRESS=TRUE THEN EXIT SUB
	
	Set f = fso.GetFile(VAR_PATH_SCANNER)
	FileModDate = f.DateLastModified
	SET f=NOTHING

	IF FileModDate<=FileModDate_2 THEN EXIT SUB

	I_COUNT_1=I_COUNT_1+1

	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)
	
	'PATH_VAR="D:\DSC\2015-2018\2018 CyberShot HX60V"+MID(PATH_VAR,3)
	PATH_VAR=FOLDER_PATH_DEST_VAR+MID(PATH_VAR,3)
	
	PATH_FILE_VAR=PATH_VAR+"\"+FILE_VAR

	If Not FSO.FolderExists(PATH_VAR) Then
		Set objFolder = FSO.CreateFolder(PATH_VAR)
		Set objFolder = NOTHING		
	End If
	
	ON ERROR RESUME NEXT
	ERR.CLEAR
	Fso.CopyFile VAR_PATH_SCANNER, PATH_FILE_VAR
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
	
	If FSO.FileExists(IMAGE_POINTER_NAME) Then 
		FSO.DeleteFile IMAGE_POINTER_NAME
	END IF

	Set objFile = FSO.CreateTextFile(IMAGE_POINTER_NAME)
	objFile.Write fileModDate
	SET objFile=NOTHING
	
END SUB

' -------------------------------------------------------------------
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
			MSGBOX ERR.DESCRIPTION & VBCRLF & VAR_PATH_SCANNER & VBCRLF & NewFilename,vbMsgBoxSetForeground,Wscript.ScriptName 
		END IF
		'Wscript.QUIT
	end if

END SUB

' -------------------------------------------------------------------
SUB PATH_SCANNER_RENAME_3

	NewFilename=REPLACE(VAR_PATH_SCANNER,".JPG",".jpg")
	
	if VAR_PATH_SCANNER <> NewFilename then 
		I_COUNT_1=I_COUNT_1+1
		Fso.MoveFile VAR_PATH_SCANNER, NewFilename
		'msgbox NewFilename
		'Wscript.QUIT
	end if

END SUB

' -------------------------------------------------------------------
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
			msgbox "ERROR CLNG WITH "+VBCRLF+VBCRLF+NewFilename,vbMsgBoxSetForeground,Wscript.ScriptName 
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
			msgbox NewFilename,vbMsgBoxSetForeground,Wscript.ScriptName 
			Wscript.QUIT
			Fso.MoveFile VAR_PATH_SCANNER, NewFilename

			I_COUNT_1=I_COUNT_1+1
		END IF
	END IF
END SUB

' -------------------------------------------------------------------
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

' -------------------------------------------------------------------
function ShowDriveInfo(drvPath)
	'USAGE
	'ShowDriveInfo(MID(FOLDER_PATH_VAR,1,2))
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
		
Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function


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