' C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 27-RENAME FILE+SUBFOLDER_FRONT_DOOR.VBS
' --------------------------------------------------------------------
' VBS 27-RENAME FILE+SUBFOLDER_FRONT_DOOR.VBS
' --------------------------------------------------------------------
' --------------------------------------------------------------------
' CREATE 01 OF 03....
' --------------------------------------------------------------------
' VBS 27-RENAME FILE+SUBFOLDER_KEY_FOB.VBS
' [ Saturday 22:44:00 Pm_23 Novemeber 2019 ]
' --------------------------------------------------------------------
' CREATE 02 OF 03....
' --------------------------------------------------------------------
' VBS 27-RENAME FILE+SUBFOLDER_MEMORY-BANK-CAMERA.VBS
' [ Monday 15:28:00 Pm_30 December 2019 ]
' --------------------------------------------------------------------
' CREATE 02 OF 03....
' --------------------------------------------------------------------
' VBS 27-RENAME FILE+SUBFOLDER_AUDIO_DEVICE.VBS
' [ Monday 21:28:00 Pm_30 December 2019 ]
' --------------------------------------------------------------------

' --------------------------------------------------------------------
' IS RUN FROM
' --------------------------------------------------------------------
' C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 29-COPY CAMERA PHOTO IMAGES.VBS
' --------------------------------------------------------------------
' HERE ROUTINE HANDLE
' --------------------------------------------------------------------
' "BLUE_1_FRONTDOOR"
' "OUTDOOR_JACKET"
' "MC_HX60V"
' "MC_HX60V1"
' --------------------------------------------------------------------
' --------------------------------------------------------------------
' Sun 19-Jan-2020 18:24:11
' --------------------------------------------------------------------


DIM objStartFolder
DIM objFolder
DIM X_COUNT
DIM I_COUNT

DIM VAR_PATH_SCANNER
DIM objSuperFolder

DIM FSO
SET FSO = CreateObject("Scripting.FileSystemObject")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO_2 = CreateObject("Scripting.FileSystemObject")

Dim arr1
Dim oDict ' Key: Full filename, Item: Filename


' ---------------------------------------------------------------
' NONE REQUIRE TO CHANGE DATE FILENAME THIS DEVICE -- YET
' OKAY NONE FILE RENAME YET
' BUTT 
' EXAMPLE SEE A BIT FURTHER
' ---------------------------------------------------------------
' LATER WILL ARRAY IT TO ALL BLUE AUDIO DEVICE
' SOME MODERN ONE REQUIRE IT RENAME
' ---------------------------------------------------------------
' [ Monday 21:33:30 Pm_30 December 2019 ]
' ---------------------------------------------------------------
' -------------------------------------------------------------------
' EVERYTHING WILL MOVE INTO THIS PATH AND ALL SUBFOLDER WITH FILE
' -------------------------------------------------------------------
' EXAMPLE 
' -------------------------------------------------------------------
' D:\0 00 MOBILE-2\8GB RECORD AUDIO BLUE _ 1 _ FRONT DOOR\RECORD 2019\RECORD
' -------------------------------------------------------------------
' WILL BE MOVE
' -------------------------------------------------------------------
' D:\0 00 MOBILE-2\8GB RECORD AUDIO BLUE _ 1 _ FRONT DOOR\RECORD 2019
' -------------------------------------------------------------------

Set args = Wscript.Arguments
For Each arg In args
	XF0 = XF0 +""""+ arg + """ "
	XF0 = Arg 
	EXIT FOR
Next
' --------------------------------------------------------------
' If Sent-to is Used Sometime Additional Quote Surround are User
' --------------------------------------------------------------
XF0 = UCASE(Replace(XF0,"""""",""""))

FOLDER_PATH="NONE FOLDER PATH GIVEN"
DIM FOLDER_PATH_AR()

REDIM PRESERVE FOLDER_PATH_AR(20)

I=0
I=I+1
FOLDER_PATH_AR(I)="D:\0 00 MOBILE-2\8GB RECORD AUDIO BLUE _ 1 _ FRONT DOOR\RECORD " + CSTR(Year(Now))+"\RECORD"
I=I+1
FOLDER_PATH_AR(I)="D:\0 00 MOBILE-2\8GB RECORD AUDIO BLUE _ 7 _ OUTDOOR JACKET\RECORD " + CSTR(Year(Now))+"\RECORD"
I=I+1
FOLDER_PATH_AR(I)="D:\0 00 MOBILE-2\8GB RECORD AUDIO BLUE _ 7 _ OUTDOOR JACKET\RECORD " + CSTR(Year(Now))+"\RECORD"
I=I+1
FOLDER_PATH_AR(I)="D:\VI_ DSC ME\2010+SONY\" + CSTR(Year(Now)) + " CyberShot HX60V_#"
I=I+1
FOLDER_PATH_AR(I)="D:\VI_ DSC ME\2010+SONY\" + CSTR(Year(Now)) + " CyberShot HX60V_#"
I=I+1
FOLDER_PATH_AR(I)="\\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY\2020 CyberShot HX60V"

REDIM PRESERVE FOLDER_PATH_AR(I)

IF INSTR(XF0,"BLUE_1_FRONTDOOR") THEN 
	FOLDER_PATH=FOLDER_PATH_AR(1)
END IF
IF INSTR(XF0,"OUTDOOR_JACKET") THEN 
	FOLDER_PATH=FOLDER_PATH_AR(2)
END IF
IF INSTR(XF0,"OUTDOOR_JACKET") THEN 
	FOLDER_PATH=FOLDER_PATH_AR(3)
END IF
IF INSTR(XF0,"MC_HX60V1") THEN 
	FOLDER_PATH=FOLDER_PATH_AR(4)
END IF
IF INSTR(XF0,"MC_HX60V") THEN 
	FOLDER_PATH=FOLDER_PATH_AR(5)
END IF

FOLDER_PATH_NAME=FOLDER_PATH

I_COUNT=0
X_COUNT=0

IF XF0<>"" THEN 
IF FOLDER_PATH<>"" THEN 
IF FSO.FolderExists(FOLDER_PATH)=FALSE THEN 
	MSGBOX Wscript.ScriptName+VBCRLF+VBCRLF+"FOLDER NOT EXIST"+VBCRLF+VBCRLF+"FOLDER PATH="+VBCRLF+FOLDER_PATH,vbMsgBoxSetForeground 
	Wscript.QUIT
END IF
END IF
END IF

' ---------------------------------------------------------------
' Call the Search subroutine to start the recursive search.
' ---------------------------------------------------------------

IF FSO.FolderExists(FOLDER_PATH)=TRUE THEN 
	MAIN
END IF

ON ERROR RESUME NEXT
IF FSO.FolderExists(FOLDER_PATH)=FALSE THEN 
	IF ERR.NUMBER=0 THEN 
		FOLDER_PATH_NAME="COUNT OF WOULD BE CHECKER __ "+CSTR(UBOUND(FOLDER_PATH_AR))+" __ ITEMS"+VBCRLF+VBCRLF
		FOR R=1 TO UBOUND(FOLDER_PATH_AR)
			FOLDER_PATH_NAME=FOLDER_PATH_NAME+FOLDER_PATH_AR(R)+VBCRLF+VBCRLF
		NEXT
	END IF
END IF
ON ERROR GOTO 0

IF I_COUNT>0 THEN ' I_COUNT THEN 
	MSGBOX_I=""
	I=""
	I=I+Wscript.ScriptName + VBCRLF + VBCRLF
	I=I+FOLDER_PATH_NAME + VBCRLF + VBCRLF
	I=I+"DONE RENAME COUNT _ " + VBCRLF + VBCRLF + CSTR(I_COUNT)
	I=I+" of "
	I=I+CSTR(X_COUNT)
	I=REPLACE(I, VBCRLF + VBCRLF+ VBCRLF, VBCRLF+ VBCRLF)
	MSGBOX_I=I
	
	MSGBOX MSGBOX_I,vbMsgBoxSetForeground 
END IF

Wscript.QUIT


'-------------------------------------------------------------------
'-------------------------------------------------------------------

Sub Main
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oDict = CreateObject("Scripting.Dictionary")
    ' Call Recursive Sub
    FindWantedFiles(objFSO.GetFolder(FOLDER_PATH))
    ' Display All Findings from Dictionary object
    DisplayFindings
	
	X_COUNT=oDict.count
	
    Set fso = Nothing
    Set oDict = Nothing
End Sub

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
		SET_DO=0
		IF INSTR(XF0,"MC_HX60V1") THEN 
			CALL PATH_SCANNER_RENAME_MC_HX60V1
			SET_DO=1
		END IF
		IF INSTR(XF0,"MC_HX60V") THEN 
			CALL PATH_SCANNER_RENAME_MC_HX60V
			SET_DO=1
		END IF

		IF SET_DO=0 THEN CALL PATH_SCANNER_RENAME
    Next
End Sub



SUB PATH_SCANNER_RENAME_MC_HX60V

	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)
	NewFilename=FILE_VAR
	
	FileModDate_1=0
	IF FSO.FileExists(VAR_PATH_SCANNER) THEN 
		Set F1 = fso.GetFile(VAR_PATH_SCANNER)
		FileModDate_1 = F1.DateLastModified
	END IF
	SET_GO=FALSE
	
	' ---------------------------------------------------------------
	' EXAMPLE
	' ---------------------------------------------------------------
	' 2019-12-09-13-02-18.WAV
	' ---------------------------------------------------------------
	' NONE REQUIRE TO CHANGE DATE FILENAME THIS DEVICE
	' ---------------------------------------------------------------
	
	IF INSTR(UCASE(FILE_VAR),"MP4")>0 THEN 
	IF LEN(FILE_VAR)<41 THEN 
		SET_GO=TRUE 
	END IF
	END IF
	
	' 41
	' 2020_01_03 Jan_Fri 22_47_16__MAH00824.MP4

	
	IF LEN(FILE_VAR)< 41 THEN
		i_TIME=""
		i_TIME=i_TIME+LPad(Year(FileModDate_1), "0", 4)+"_"
		i_TIME=i_TIME+LPad(Month(FileModDate_1), "0", 2)+"_"
		i_TIME=i_TIME+LPad(Day(FileModDate_1), "0", 2)+" "
		i_TIME=i_TIME+WeekdayName(Weekday(FileModDate_1),TRUE)+"_"
		i_TIME=i_TIME+MonthName(MONTH(FileModDate_1),TRUE)+" "
		i_TIME=i_TIME+""
		i_TIME=i_TIME+LPad(Hour(FileModDate_1), "0", 2)+"_"
		i_TIME=i_TIME+LPad(Minute(FileModDate_1), "0", 2)+"_"
		i_TIME=i_TIME+LPad(Second(FileModDate_1), "0", 2)+"__"
	
		NewFilename=i_TIME + UCASE(FILE_VAR)
	END IF
	
	' D:\VI_ DSC ME\2010+SONY\2020 CyberShot HX60V_#\MP_ROOT\100ANV01\MAH00002.MP4
	
	PATH_VAR_2=PATH_VAR ' STAY SAME IF NOT HAVE CHANGE
	
	STRIP_PART="\MP_ROOT\100ANV01"
	STRIP_PART="\MP_ROOT\"
	X_SP_1=INSTRREV(PATH_VAR,STRIP_PART)
	IF SET_GO=TRUE THEN
	IF X_SP_1>0 THEN 
		X_SP_2=INSTRREV(PATH_VAR,"\")
		X_SP_4=MID(PATH_VAR,X_SP_1,X_SP_2-X_SP_1)
		PATH_VAR_2=REPLACE(PATH_VAR,X_SP_4,"")
	END IF
	END IF
	
	ON ERROR RESUME NEXT

	IF SET_GO=TRUE THEN
		IF FSO.FileExists(PATH_VAR_2+NewFilename)=TRUE THEN
			MSGBOX "FILE ALREADY EXISTS "+VBCRLF+VBCRLF+NewFilename,vbMsgBoxSetForeground
			Wscript.QUIT
		END IF
		
		NewFilename=PATH_VAR_2+NewFilename
		IF SET_GO=TRUE THEN
			MSGBOX VAR_PATH_SCANNER+VBCRLF+NewFilename,vbMsgBoxSetForeground
			Fso.MoveFile VAR_PATH_SCANNER, NewFilename
			I_COUNT=I_COUNT+1
		END IF
	END IF


End Sub


SUB PATH_SCANNER_RENAME_MC_HX60V1

	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)
	NewFilename=FILE_VAR
	
	FileModDate_1=0
	IF FSO.FileExists(VAR_PATH_SCANNER) THEN 
		Set F1 = fso.GetFile(VAR_PATH_SCANNER)
		FileModDate_1 = F1.DateLastModified
	END IF
	SET_GO=FALSE
	
	' ---------------------------------------------------------------
	' EXAMPLE
	' ---------------------------------------------------------------
	' 2019-12-09-13-02-18.WAV
	' ---------------------------------------------------------------
	' NONE REQUIRE TO CHANGE DATE FILENAME THIS DEVICE
	' ---------------------------------------------------------------
	
	IF INSTR(UCASE(FILE_VAR),"MP4")>0 THEN 
	IF LEN(FILE_VAR)<41 THEN 
		SET_GO=TRUE 
	END IF
	END IF
	
	IF LEN(FILE_VAR)<41 THEN
		' MSGBOX CSTR(LEN(FILE_VAR)) +VBCRLF + FILE_VAR
	END IF
	
	' 41
	' 2020_01_03 Jan_Fri 22_47_16__MAH00824.MP4

	
	IF LEN(FILE_VAR)< 41 THEN
		i_TIME=""
		i_TIME=i_TIME+LPad(Year(FileModDate_1), "0", 4)+"_"
		i_TIME=i_TIME+LPad(Month(FileModDate_1), "0", 2)+"_"
		i_TIME=i_TIME+LPad(Day(FileModDate_1), "0", 2)+" "
		i_TIME=i_TIME+WeekdayName(Weekday(FileModDate_1),TRUE)+"_"
		i_TIME=i_TIME+MonthName(MONTH(FileModDate_1),TRUE)+" "
		i_TIME=i_TIME+""
		i_TIME=i_TIME+LPad(Hour(FileModDate_1), "0", 2)+"_"
		i_TIME=i_TIME+LPad(Minute(FileModDate_1), "0", 2)+"_"
		i_TIME=i_TIME+LPad(Second(FileModDate_1), "0", 2)+"__"
	
		NewFilename=i_TIME + UCASE(FILE_VAR)
	END IF
	
	' D:\VI_ DSC ME\2010+SONY\2020 CyberShot HX60V_#\MP_ROOT\100ANV01\MAH00002.MP4
	
	
	PATH_VAR_2=PATH_VAR ' STAY SAME IF NOT HAVE CHANGE
	
	STRIP_PART="\MP_ROOT\100ANV01"
	STRIP_PART="\MP_ROOT\"
	X_SP_1=INSTRREV(PATH_VAR,STRIP_PART)
	IF SET_GO=TRUE THEN
	IF X_SP_1>0 THEN 
		X_SP_2=INSTRREV(PATH_VAR,"\")
		X_SP_4=MID(PATH_VAR,X_SP_1,X_SP_2-X_SP_1)
		PATH_VAR_2=REPLACE(PATH_VAR,X_SP_4,"")
	END IF
	END IF
	
	ON ERROR RESUME NEXT

	IF SET_GO=TRUE THEN
		IF FSO.FileExists(PATH_VAR_2+NewFilename)=TRUE THEN
			MSGBOX "FILE ALREADY EXISTS "+VBCRLF+VBCRLF+NewFilename,vbMsgBoxSetForeground
			Wscript.QUIT
		END IF
		
		NewFilename=PATH_VAR_2+NewFilename
		IF SET_GO=TRUE THEN
			' MSGBOX VAR_PATH_SCANNER+VBCRLF+NewFilename,vbMsgBoxSetForeground
			Fso.MoveFile VAR_PATH_SCANNER, NewFilename
			I_COUNT=I_COUNT+1
		END IF
	END IF


End Sub

SUB PATH_SCANNER_RENAME

	PATH_VAR=MID(VAR_PATH_SCANNER,1,INSTRREV(VAR_PATH_SCANNER,"\"))
	FILE_VAR=MID(VAR_PATH_SCANNER,INSTRREV(VAR_PATH_SCANNER,"\")+1)
	NewFilename=FILE_VAR
	
	FileModDate_1=0
	IF FSO.FileExists(VAR_PATH_SCANNER) THEN 
		Set F1 = fso.GetFile(VAR_PATH_SCANNER)
		FileModDate_1 = F1.DateLastModified
	END IF
	
	SET_GO=FALSE
	
	' ---------------------------------------------------------------
	' EXAMPLE
	' ---------------------------------------------------------------
	' 2019-12-09-13-02-18.WAV
	' ---------------------------------------------------------------
	' NONE REQUIRE TO CHANGE DATE FILENAME THIS DEVICE
	' ---------------------------------------------------------------
	
	IF LEN(FILE_VAR)= 23 THEN         SET_GO=TRUE 
	IF MID(FILE_VAR,1,3)="REC" THEN   SET_GO=TRUE    ' -- REC001.WAV
	
	IF LEN(FILE_VAR)<> 23 THEN
		i_TIME=""
		i_TIME=i_TIME+LPad(Year(FileModDate_1), "0", 4)+"-"
		i_TIME=i_TIME+LPad(Month(FileModDate_1), "0", 2)+"-"
		i_TIME=i_TIME+LPad(Day(FileModDate_1), "0", 2)+" "
		i_TIME=i_TIME+""
		i_TIME=i_TIME+LPad(Hour(FileModDate_1), "0", 2)+"-"
		i_TIME=i_TIME+LPad(Minute(FileModDate_1), "0", 2)+"-"
		i_TIME=i_TIME+LPad(Second(FileModDate_1), "0", 2)+""
		' i_TIME=i_TIME+" "
		' i_TIME=i_TIME+MonthName(MONTH(FileModDate_1),TRUE)+" "
		' i_TIME=i_TIME+WeekdayName(Weekday(VAR_NOW),TRUE)	
		i_TIME=i_TIME+" "

		F_TIME=""
		F_TIME=F_TIME+LPad(Year(FileModDate_1), "0", 4)+"-"
		F_TIME=F_TIME+LPad(Month(FileModDate_1), "0", 2)+"-"
		F_TIME=F_TIME+LPad(Day(FileModDate_1), "0", 2)
	
		NewFilename=i_TIME+FILE_VAR
	END IF
	

	PATH_VAR_2=REPLACE(PATH_VAR,"\RECORD " + CSTR(Year(Now))+"\RECORD","\RECORD " + CSTR(Year(Now)))
	
	' ---------------------------------------------------------------
	' EXAMPLE 
	' D:\0 00 MOBILE-2\8GB RECORD AUDIO BLUE _ 1 _ FRONT DOOR\RECORD 2019
	' ---------------------------------------------------------------
	
	' WAS GOING TO SET NOT MAKE PATH IF ROOT LEVEL TO VARIABLE
	' IF LEN(PATH_VAR)>68 THEN

	PATH_VAR=PATH_VAR_2
	BuildFullPath PATH_VAR

	ON ERROR RESUME NEXT
	
	IF SET_GO=TRUE THEN
		IF FSO.FileExists(PATH_VAR+NewFilename)=TRUE THEN
			MSGBOX "FILE ALREADY EXISTS "+VBCRLF+VBCRLF+NewFilename,vbMsgBoxSetForeground
			Wscript.QUIT
		END IF
		
		NewFilename=PATH_VAR+"\"+NewFilename
		IF SET_GO=TRUE THEN
			' MSGBOX VAR_PATH_SCANNER+VBCRLF+NewFilename,vbMsgBoxSetForeground
			Fso.MoveFile VAR_PATH_SCANNER, NewFilename
			I_COUNT=I_COUNT+1
		END IF
	END IF
END SUB


' -------------------------------------------------------------------
' CREATE_FOLDER_PATH CREATE_TREE_PATH CREATE_PATH _ BuildFullPath
' -------------------------------------------------------------------
Sub BuildFullPath(ByVal FullPath)
	If Not fso.FolderExists(FullPath) Then
		ON ERROR RESUME NEXT
		BuildFullPath fso.GetParentFolderName(FullPath)
		fso.CreateFolder FullPath
		ON ERROR GOTO 0
	End If
End Sub


Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function




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