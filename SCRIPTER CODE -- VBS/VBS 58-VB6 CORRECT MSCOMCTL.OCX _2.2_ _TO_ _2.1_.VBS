'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\
'# __ 
'# __ VBS 58-VB6 CORRECT MSCOMCTL.OCX _2.2_ _TO_ _2.1_.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START      TIME [ Thu 05-Sep-2019 17:09:10 ]
'# END LAST EDITOR [ Thu 05-Sep-2019 20:20:10 ] 
'# __ 
'====================================================================

' -------------------------------------------------------------------
' DESCRIPTION 
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' SESSION 001
' -------------------------------------------------------------------
' CODE SET TAKEN FROM 
' "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 35-RENAMER VB6 _VBP LCASE.VBS" %JOBNAME%
' -------------------------------------------------------------------
' OUR MODIFY
' ALL VBP PROJECT FILE ARE CHECK TO FIND ERROR WITH MSCOMCTL.OCX
' THE PROJECT OF FEW COMPUTER IN MY GROUP NETWORK
' MOSTLY LOW END THING OR SOME KIND MAKE ERROR MORE
' THE REPORT VERSION FOR MSCOMCTL.OCX CHANGE TO 2.2 INSTEAD OF 2.1 
' AND THEN NOT COMPATIBLE
' ONLY REQUIRE RENAME THE TEXT -- END THERE
' MUST FIRST SUB FOLDER SCANNER MANIPULATOR FOR TEXT FILE
' SEARCH AND REPLACE
' GOODSYNC HAD LOT WORK DONE TO SINGLE FILE SEARCH REPLACE
' -------------------------------------------------------------------
' Thu 05-Sep-2019 17:09:10
' Thu 05-Sep-2019 20:20:00 __ 3 HOUR CODE 
'							  AND BEFORE IN DAY ANOTHER CODE 7 HOUR - LONG BASH
'							  FOR FACEBOOK REFRESH SCREEN
' -------------------------------------------------------------------


' -------------------------------------------------------------------
' LOCATION ON-LINE 
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------



'--------------------------------------------------------------------
DIM Version_OS
CALL GET_WINDOWS_VERSION
'--------------------------------------------------------------------
IF Version_OS<10 THEN 
	Wscript.QUIT
END IF
'--------------------------------------------------------------------


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
		IF i1 > 0 THEN ' ANY PROGRAM WSCRIPT 
				' Wscript.Quit
		END IF
		IF i2 > 1 THEN ' MY  PROGRAM WSCRIPT 
				Wscript.Quit
		END IF
    END IF
NEXT


Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

Const vbFromUnicode = 128
Const vbUnicode = 64
Const adTypeBinary = 1
Const adTypeText = 2
Const adLongVarBinary = 205

Const ForReading = 1, ForWriting = 2, ForAppending = 8

' --------------------------------------------------------
' VBSCRIPT SEEMS TO READER IN UNICODE MANIPULATE AND SAVE
' IN VBA VB6 Replace() DOES MULTIPLE OCCUR HER WITH , , 1)
' HERE VBS DOES SINGLE 
' --------------------------------------------------------

' CUR_USER = CreateObject("WScript.Network").UserName
' 
' FILEN_1=
' FILEN_2=

' FSO.COPYFILE FILEN_1, FILEN_2

DIM TT_EDITOR
DIM Counter_VAR
DIM Counter_Change_VAR



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
DIM STR_PATH

' -------------------------------------------------------------------
' SOURCE GIVEN
' -------------------------------------------------------------------
' vbscript - how to check vbs script in windows is running or not? - Stack Overflow
' https://stackoverflow.com/questions/7849699/how-to-check-vbs-script-in-windows-is-running-or-not
' -------------------------------------------------------------------
' [ Wednesday 13:08:00 Pm_20 March 2019 ]
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

DIM FOLDER_PATH_SET_VAR()
DIM FOLDER_PATH_SET_VAR_02()

REDIM FOLDER_PATH_SET_VAR(100)
REDIM FOLDER_PATH_SET_VAR_02(100)

DIM FindFirstFile_VAR
DIM FindMostRecentFile_VAR

DIM I_SET_DONE

Set args = Wscript.Arguments
For Each arg In args
	Wscript_Arguments = Wscript_Arguments + arg
Next

'--------------------------------------------------------------------
' USAGE ONCE A DAY AT WITH QUITE MODE
'--------------------------------------------------------------------
' C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL_1.ahk
'--------------------------------------------------------------------

' -------------------------------------------------------------------
' LOAD ONCE TO GET GO AND AGAIN LATER FOR TWO SHOW
' WHEN SORT ORDER DONE
' -------------------------------------------------------------------
IF Wscript_Arguments="" THEN 
	STR_PATH = """D:\VB6\VB-NT\00_Best_VB_01\Shell VBasic 6 Loader\Shell VBasic 6 Loader.exe"""
	WSHShell.Run STR_PATH
END IF



I=0
I=I+1:FOLDER_PATH_SET_VAR(I)="D:\VB6\VB-NT"
'I=I+1:FOLDER_PATH_SET_VAR(I)=""
'I=I+1:FOLDER_PATH_SET_VAR(I)=""
'I=I+1:FOLDER_PATH_SET_VAR(I)=""
'I=I+1:FOLDER_PATH_SET_VAR(I)=""

REDIM PRESERVE FOLDER_PATH_SET_VAR(I)


DIM FOLDER_PATH_FILTER()
REDIM FOLDER_PATH_FILTER(1)
FOLDER_PATH_FILTER(1)="\_gsdata_\"

FOR RR=1 TO UBOUND(FOLDER_PATH_SET_VAR)
	IF FSO.FolderExists(FOLDER_PATH_SET_VAR(RR))=FALSE THEN 
		MSGBOX "FOLDER NOT EXIST"+VBCRLF+VBCRLF+FOLDER_PATH_SET_VAR(RR)
		Wscript.QUIT
	END IF
NEXT 

' -------------------------------------------------------------------
' -------------------------------------------------------------------
FOR RR=1 TO UBOUND(FOLDER_PATH_SET_VAR)
	SET_VAR_MODE_ORDER=1
	FOLDER_PATH_VAR=FOLDER_PATH_SET_VAR(RR)
	CALL MAIN
NEXT 

I=""
I=I+Wscript.ScriptName + VBCRLF 
I=I+""+ VBCRLF 
FOR RR=1 TO UBOUND(FOLDER_PATH_SET_VAR)
	I=I+ FOLDER_PATH_SET_VAR(RR) +"\*.VBP"+ VBCRLF 
NEXT
I=I+"----------------------------------------------------------------"+ VBCRLF 
I=I+"CORRECTION MADE _.VBP FILE COUNT" + VBCRLF + CSTR(I_COUNT_1) + VBCRLF
I=I+VBCRLF 
I=I+"TOTAL _.VBP COUNT" + VBCRLF + CSTR(I_COUNT_2) + VBCRLF 
I=I+VBCRLF 
I=I+I_SET_DONE

IF Wscript_Arguments="" THEN 
	MSGBOX I,vbMsgBoxSetForeground 
END IF


' -------------------------------------------------------------------
' LOAD ONCE TO GET GO AND AGAIN LATER FOR TWO SHOW
' WHEN SORT ORDER DONE
' -------------------------------------------------------------------
IF Wscript_Arguments="" AND I_COUNT_1>0 THEN 
	STR_PATH = """D:\VB6\VB-NT\00_Best_VB_01\Shell VBasic 6 Loader\Shell VBasic 6 Loader.exe"""
	WSHShell.Run STR_PATH
END IF


SET fso = NOTHING
SET objFSO = NOTHING
SET WSHShell = NOTHING

Wscript.QUIT
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
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
Sub FindWantedFiles_5(sFolder)
    On Error Resume Next
    DIM FILE_FOLDER_COUNTER
	Dim oFDR, oItem, file
    ' Check if wanted files are in this folder
    For Each file In sFolder.Files

		GOT_VBP=FALSE
		FILE_STRIP_PATH=MID(FILE,INSTRREV(FILE,"\")+1)
		IF UCASE(MID(file,LEN(file)-3))=".VBP" THEN 
			I_COUNT_2=I_COUNT_2+1
			GOT_VBP=TRUE
		END IF

		SET_FILTER_GO=TRUE
		FOR R_COUNT=1 TO UBOUND(FOLDER_PATH_FILTER)
			IF INSTR(FILE,FOLDER_PATH_FILTER(R_COUNT)) THEN SET_FILTER_GO=FALSE
		NEXT
		IF SET_FILTER_GO=TRUE THEN 
			IF GOT_VBP=TRUE THEN 
				oDict_1.Add FILE.path, file
				I_COUNT_2=I_COUNT_2+1
			END IF
		END IF
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
Sub DisplayFindings_5()
    Dim oKeys, oKey
    oKeys = oDict_1.Keys
    For Each oKey In oKeys
		VAR_PATH_SCANNER=oKey
		PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
		
		IF FSO.FileExists(VAR_PATH_SCANNER) Then 

			Set f = fso.GetFile(VAR_PATH_SCANNER)
			FileModDate = f.DateLastModified
			SET f=NOTHING
			FILE_DATE_SPAN=40
			IF FileModDate+FILE_DATE_SPAN>NOW THEN
				Set File_1 = fso.OpenTextFile(VAR_PATH_SCANNER, ForReading)
				TT_EDITOR = File_1.ReadAll
				Set File_1 = NOTHING
				IF InStr(UCASE(TT_EDITOR), "#2.2#0; MSCOMCTL.OCX")>0 THEN 
					' MSGBOX VAR_PATH_SCANNER
					CALL USUAL_MAIN
					Set File_2 = fso.OpenTextFile(VAR_PATH_SCANNER, ForWriting)
					File_2.Write TT_EDITOR
					Set File_2 = NOTHING
					' --------------------------------------------------
					STR_PATH = "EXPLORER.EXE /SELECT," & VAR_PATH_SCANNER
					' WSHShell.Run STR_PATH
					' --------------------------------------------------
					' MSGBOX VAR_PATH_SCANNER
					I_COUNT_1=I_COUNT_1+1
				END IF
			END IF
		END IF
	Next
End Sub

SUB USUAL_MAIN

	X = 1
	T1 = InStr(X,UCASE(TT_EDITOR), "#2.2#0; MSCOMCTL.OCX")
	IF T1>0 THEN 
		MODSTR="#2.1#0; MSCOMCTL.OCX"
		TT_EDITOR_2 = Replace(UCASE(TT_EDITOR), "#2.2#0; MSCOMCTL.OCX", MODSTR,1,1)
		TT_EDITOR_3=MID(TT_EDITOR,1,T1-1)+MODSTR+MID(TT_EDITOR,T1+LEN(MODSTR))
		' -----------------------------------------------------------
		' DEBUG DOUBLE CHECKER FOR STRING MANIPULATE TWO METHOD
		' ----------------------------------
		' TT_EDITOR_3=UCASE(TT_EDITOR_3)
		' IF TT_EDITOR_2=TT_EDITOR_3 THEN MSGBOX "MATCH GOOD"
		' IF TT_EDITOR_2<>TT_EDITOR_3 THEN 
			' MSGBOX MID(TT_EDITOR_2,T1-20,40)+VBCRLF+VBCRLF+MID(TT_EDITOR_3,T1-20,40)
		' END IF
		' -----------------------------------------------------------
		TT_EDITOR=TT_EDITOR_3
	END IF
	
END SUB




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

' -------------------------------------------------------------------
SUB GET_WINDOWS_VERSION
' -------------------------------------------------------------------
	Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
	for each System in SystemSet 
	' WScript.Echo System.Caption 
	' WScript.Echo System.Manufacturer 
	' WScript.Echo System.BuildType 
	' WScript.Echo " Version: " + System.Version 
	' WScript.Echo " Locale: " + System.Locale 
	' WScript.Echo " Windows Directory: " + System.WindowsDirectory 
	' WScript.Echo " Total memory: " + System.TotalVisibleMemorySize + " bytes" 
	' WScript.Echo " Serial Number: " + System.SerialNumber 
	' Wscript.Echo "" 
	Version_OS = CINT(MID(System.Version,1,INSTR(System.Version,".")-1))
	'WINDOWS 7 = 6.1.7601
	next 
END SUB


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