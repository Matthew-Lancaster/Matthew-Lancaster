'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 34-HUBIC DELETE-ER.VBS
'# __ 
'# __ VBS 34-HUBIC DELETE-ER.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START      TIME [ Thu 18-Oct-2018 03:20:00 ] COPY FROM ANOTHER CODE
'# END LAST EDITOR [ Thu 18-Oct-2018 15:40:00 ] 
'# --- LAST EDITOR [ Mon 29-Oct-2018 03:28:00 ]
'# __ 
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
'# ------------------------------------------------------------------
' WORKS WITH HUBIC CLOUD AND GOODSYNC AND QNAP 
' A LARGE BACKUP OF CLOUD FILING SYSTEM WAS HAD TO DELETE 
' THIS DOES THE JOB BY SYNCING WITH THE QNAP PROGRAMING INTERNALLY TO QNAP 
' CALLED __ CLOUD DRIVE SYNC
' IT DELETE FILE THAT ARE OLDER THAN CERTAIN DATE TIME
' AND DELETE EMPTY FOLDER WITH EVEN LONGER TIME
' QNAP SYNC CLOUD DRIVE SYNC PROGRAM WILL FAIL IF DELETE FILE IN 
' MIDDLE OF COPY
' WAS A SYSTEM TO DO ENTIRELY WITH GOODSYNC
' BUT LESS QUICK AS WENT ON AS SO MANY FILE SYNCER
' NOW INDEPENDENT FOLDER FROM ROOT
' THIS CODE IS CALLED FROM GOODSYNC PRIOR TO EACH ANALYZE
' NOT ANYMORE
' NOW RUN BY Autokey -- 19-SCRIPT_TIMER_UTIL.ahk
' ONCE AN HOUR
'
' IT RUN IN QUITE MODE IF COMMAND LINE HAS SOMETHING
' -------------------------------------------------------------------
' -------------------------------------------------------------------

'# ------------------------------------------------------------------
' LOCATTION ON-LINE 
'# ------------------------------------------------------------------
' -------------------------------------------------------------------
' ----
' Matthew-Lancaster/VBS 30-FUJI CAMERA IMAGE DATE FOLDER CREATE & MOVE.VBS at master · Matthew-Lancaster/Matthew-Lancaster
' https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20VBS/VBS%2030-FUJI%20CAMERA%20IMAGE%20DATE%20FOLDER%20CREATE%20%26%20MOVE.VBS
' ----
' -------------------------------------------------------------------

'#-------------------------------------------------------------------
' 001 __ 
'#-------------------------------------------------------------------
' ADDER
' ADD CODE
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' Thu 18-Oct-2018 03:20:00
' Thu 18-Oct-2018 05:08:57 
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' 002 __ 
'#-------------------------------------------------------------------
' ADDER 
' ADD CODE __ PATH TOO LONG PROBLEM CAME IN ERROR TRAPPER
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' Thu 18-Oct-2018 15:13:45
' Thu 18-Oct-2018 15:40:30 
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' 003 __ 
'#-------------------------------------------------------------------
' ADDER 
' ADD CODE __ MODIFY TO EASIER CHANGE ROOT PATH AND USE 
' ARRAY BETTER FOR SCANNING FOLDER
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' Mon 29-Oct-2018 02:39:19
' Mon 29-Oct-2018 03:28:00
'#-------------------------------------------------------------------

'#-------------------------------------------------------------------
' 004 __ 
'#-------------------------------------------------------------------
' ADDER 
' ADD CODE __ COMPUTER DRIVER SOFTWARE CHANGED
' -------------------------------------------------------------------
' Thu 01-Nov-2018 00:52:13
' Thu 01-Nov-2018 01:30:00 -- IDEA GTTING CLOSE REAL TIME EITOR HERE FOLLOWS
' Thu 01-Nov-2018 01:09:53
' Thu 01-Nov-2018 01:30:00 
'#-------------------------------------------------------------------



Const ForReading = 1, ForWriting = 2, ForAppending = 8
const DontWaitUntilFinished = false, ShowWindow = 1, DontShowWindow = 0, WaitUntilFinished = true

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

DIM ONCE_MESSENGER
DIM ERROR_READ_FOLDER

DIM DELETE_FILE_TIME_SPAN
DIM DELETE_FOLDER_TIME_SPAN

Set args = Wscript.Arguments
For Each arg In args
	Wscript_Arguments = Wscript_Arguments + arg
Next


' -------------------------------------------------------------------
' IF SPECIFY FOLDER PATH LOWER THAN THIS ONE THE DELETE EMPTY 
' WOULD FIND AND DELETE ALL EMPTY FOLDER UP TO HERE
' -------------------------------------------------------------------

DIM FOLDER_PATH_FILTER()
REDIM FOLDER_PATH_FILTER(7)
DRIVE_VAR="O:"
FOLDER_PATH_VAR_1=DRIVE_VAR+"\"
FOLDER_PATH_FILTER(1)=DRIVE_VAR+"\3L"
FOLDER_PATH_FILTER(2)=DRIVE_VAR+"\5P"
FOLDER_PATH_FILTER(3)=DRIVE_VAR+"\7G"
FOLDER_PATH_FILTER(4)=DRIVE_VAR+"\QNAP1"
FOLDER_PATH_FILTER(5)=DRIVE_VAR+"\WHOLE_ARCHIVE"
FOLDER_PATH_FILTER(6)=DRIVE_VAR+"\zz--FOLDER"
FOLDER_PATH_FILTER(7)=DRIVE_VAR+"\zz--NETWORK_FOLDER"

DELETE_FILE_TIME_SPAN=TIMESERIAL(2,0,0)
DELETE_FOLDER_TIME_SPAN=TIMESERIAL(48,0,0)

DIM CUR_USER
DIM ComputerName

CUR_USER = CreateObject("WScript.Network").UserName
Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

' IF ComputerName<>"8-MSI-GP62M-7RD" THEN 
	' Wscript.QUIT
' END IF 
IF ComputerName<>"7-ASUS-GL522VW" THEN 
	Wscript.QUIT
END IF 

IF FSO.FolderExists(FOLDER_PATH_VAR_1)=FALSE THEN
	IF Wscript_Arguments="" THEN
		MSGBOX "VBS 34 - HUBIC DELETER - DRIVE - FOLDER - NOT EXIST - EXIT"+VBCRLF+VBCRLF+FOLDER_PATH_VAR_1
		Wscript.QUIT
	END IF
END IF

' -------------------------------------------------------------------
' 01 OF 03
' -------------------------------------------------------------------
SET_VAR_MODE_ORDER=1
FOLDER_PATH_VAR=FOLDER_PATH_VAR_1
CALL MAIN

I=""
I=I+"DEL FROM HUBIC LEARN"+ VBCRLF 
I=I+"----------------------------------------------------------------"+ VBCRLF 
I=I+""+ VBCRLF 
I=I+"HUBIC DEL FROM FILE   COUNT" + VBCRLF + CSTR(I_COUNT_1) + VBCRLF 
I=I+"HUBIC DEL FROM FOLDER COUNT" + VBCRLF + CSTR(I_COUNT_2) + VBCRLF
I=I+"HUBIC ERROR READ FOLDER COUNT" + VBCRLF + CSTR(ERROR_READ_FOLDER) + VBCRLF
I=I+""+ VBCRLF 
I=I+"DONE...."


IF Wscript_Arguments="" THEN 
	MSGBOX I
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
		FOR R_COUNT= 1 TO UBOUND(FOLDER_PATH_FILTER)
			FOLDER_PATH_VAR_VAR=FOLDER_PATH_FILTER(R_COUNT)
			' Call Recursive Sub
			' MSGBOX FOLDER_PATH_VAR_VAR
			IF FSO.FolderExists(FOLDER_PATH_VAR_VAR)=TRUE THEN 
				FindWantedFiles_2(objFSO.GetFolder(FOLDER_PATH_VAR_VAR))
			END IF
		NEXT


		' Display All Findings from Dictionary object
		DisplayFindings_4
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
				MSGBOX ERR.DESCRIPTION & VBCRLF & VBCRLF & "UNABLE TO DELETE" & VBCRLF & VBCRLF & VAR_PATH_SCANNER 
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
					MSGBOX ERR.DESCRIPTION & VBCRLF & VBCRLF & "UNABLE TO DELETE" & VBCRLF & VBCRLF & VAR_PATH_SCANNER 
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
		
		' MSGBOX VAR_PATH_SCANNER
		
		IF FSO.FileExists(VAR_PATH_SCANNER) Then 
			Set f = fso.GetFile(VAR_PATH_SCANNER)
			FileModDate = f.DateLastModified
			FileCreateDate = f.DateCreated 
			SET f=NOTHING
			IF FileCreateDate+DELETE_FILE_TIME_SPAN<NOW THEN 
				' MSGBOX VAR_PATH_SCANNER
				ON ERROR RESUME NEXT
				ERR.CLEAR
				FSO.DeleteFile VAR_PATH_SCANNER
				I_COUNT_1=I_COUNT_1+1
				IF ERR.NUMBER>0 AND Wscript_Arguments="" THEN 
					MSGBOX ERR.DESCRIPTION & VBCRLF & VBCRLF & "UNABLE TO DELETE" & VBCRLF & VBCRLF & VAR_PATH_SCANNER 
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
					IF FolderCreateDate+DELETE_FOLDER_TIME_SPAN<NOW THEN 
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
Function FolderEmpty(strFolder)
	Dim objFolder
	ON ERROR RESUME NEXT
	' BECAUSE HERE BOTH AT SAME TIME
	' CAN'T GET PATH IF TOO LONG SUPPOSE
	' objFolder.Files.Count
	' objFolder.SubFolders.Count
	' ---------------------------
	' ---------------------------
	' Files.Count D:\DEL HUBIC QNAP\zz--NETWORK_FOLDER\HubiC-DeskBackup_7G--zz--NETWORK--5P--5-ASUS-P2520LA_02_D_DRIVE--FOLDER--0 00 HTTrack\www.lyricsmania.com\http___www.lyricsmania.com_\www.lyricsmania.com\lyrics\adam_lambert_lyrics_52608\for_your_entertainment_lyrics_99113
	' ---------------------------
	' OK   
	' ---------------------------

	' ---------------------------
	' Windows Script Host
	' ---------------------------
	' Script:	C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 34-HUBIC DELETE-ER.VBS
	' Line:	410
	' Char:	3
	' Error:	Path not found
	' Code:	800A004C
	' Source: 	Microsoft VBScript runtime error
	' ---------------------------
	' OK   
	' ---------------------------
	
	If FSO.FolderExists(strFolder) Then
		Set objFolder = FSO.GetFolder(strFolder)
		If objFolder.Files.Count = 0 THEN 
			IF objFolder.SubFolders.Count = 0 Then
				FolderEmpty=True
			Else
				FolderEmpty=False
			End If
		End If
	End If

	IF ERR.NUMBER >0 THEN 
		FolderEmpty=False
		ERROR_READ_FOLDER=ERROR_READ_FOLDER+1
		IF Wscript_Arguments="" THEN 
			IF ONCE_MESSENGER = FALSE THEN 
				ONCE_MESSENGER = TRUE
				MSGBOX "Error Description ----------" & VBCRLF & ERR.DESCRIPTION & VBCRLF & VBCRLF & "UNABLE TO READ FILE AND/OR FOLDER -- MAYBE THE PATH ARE TOO LONG" & VBCRLF & VBCRLF & strFolder
			END IF
		END IF
	END IF
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
		MSGBOX ERR.DESCRIPTION & VBCRLF & VBCRLF & "FROM" & VBCRLF & VAR_PATH_SCANNER & VBCRLF & VBCRLF & "TO" & VBCRLF & NewFilename
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
			MSGBOX ERR.DESCRIPTION & VBCRLF & VBCRLF & "FROM" & VBCRLF & VAR_PATH_SCANNER & VBCRLF & VBCRLF & "TO" & VBCRLF & NewFilename
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
		MSGBOX ERR.DESCRIPTION & VBCRLF & VAR_PATH_SCANNER & VBCRLF & NewFilename
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
			MSGBOX ERR.DESCRIPTION & VBCRLF & VAR_PATH_SCANNER & VBCRLF & NewFilename
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