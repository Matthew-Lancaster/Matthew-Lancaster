' C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 27-RENAME FILE+SUBFOLDER_AUDIO_DEVICE.VBS
' --------------------------------------------------------------------
' VBS 27-RENAME FILE+SUBFOLDER_AUDIO_DEVICE.VBS
' --------------------------------------------------------------------
' CREATE 01 OF 02....
' VBS 27-RENAME FILE+SUBFOLDER_KEY_FOB.VBS
' [ Saturday 22:44:00 Pm_23 Novemeber 2019 ]
' --------------------------------------------------------------------
' CREATE 02 OF 02....
' VBS 27-RENAME FILE+SUBFOLDER_MEMORY-BANK-CAMERA.VBS
' [ Monday 15:28:00 Pm_30 December 2019 ]
' --------------------------------------------------------------------
' CREATE 02 OF 02....
' VBS 27-RENAME FILE+SUBFOLDER_AUDIO_DEVICE.VBS
' [ Monday 21:28:00 Pm_30 December 2019 ]
' --------------------------------------------------------------------

DIM objStartFolder
DIM objFolder
DIM I_COUNT
DIM VAR_PATH_SCANNER
DIM objSuperFolder

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
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

Wscript.QUIT

FOLDER_PATH="D:\0 00 MOBILE-2\8GB RECORD AUDIO BLUE _ 1 _ FRONT DOOR\RECORD " + CSTR(Year(Now))+"\RECORD"

I_COUNT=0

IF FSO.FolderExists(FOLDER_PATH)=FALSE THEN 
	MSGBOX Wscript.ScriptName+VBCRLF+VBCRLF+"FOLDER NOT EXIST"+VBCRLF+VBCRLF+FOLDER_PATH,vbMsgBoxSetForeground 
	Wscript.QUIT
END IF

' ---------------------------------------------------------------
' Call the Search subroutine to start the recursive search.
' Search objFSO.GetFolder(FOLDER_PATH)
' ---------------------------------------------------------------
' FindWantedFiles objFSO.GetFolder(FOLDER_PATH)
' DisplayFindings
' ---------------------------------------------------------------

MAIN

'-------------------------------------------------------------------
'-------------------------------------------------------------------

Sub Main
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oDict = CreateObject("Scripting.Dictionary")
    ' Call Recursive Sub
    FindWantedFiles(objFSO.GetFolder(FOLDER_PATH))
    ' Display All Findings from Dictionary object
    DisplayFindings

	IF I_COUNT>0 THEN 
		MSGBOX Wscript.ScriptName+VBCRLF+VBCRLF+"DONE RENAME COUNT _ " & CSTR(I_COUNT) & " of " & CSTR(oDict.count),vbMsgBoxSetForeground 
	END IF
	
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
		CALL PATH_SCANNER_RENAME
    Next
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
	
	IF LEN(FILE_VAR)= 23 THEN SET_GO=TRUE 
	
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
	' IF PATH_VAR_2<>PATH_VAR THEN 
		' MSGBOX PATH_VAR+VBCRLF+CSTR(LEN(PATH_VAR))
	' END IF
	' ---------------------------------------------------------------
	' MSGBOX LEN(PATH_VAR)
	' ---------------------------------------------------------------
	' CHECK PATH LEN AS MIGHT VARY BETWEEN USE OTHER DIRECTORY 
	' AND OTHER CAMERA SHIT
	' ---------------------------------------------------------------
	' EXAMPLE 
	' D:\0 00 MOBILE-2\8GB RECORD AUDIO BLUE _ 1 _ FRONT DOOR\RECORD 2019
	' ---------------------------------------------------------------
	' MSGBOX PATH_VAR+VBCRLF+CSTR(LEN(PATH_VAR))
	' ---------------------------------------------------------------
	
 	IF LEN(PATH_VAR)<=68 THEN 
		MSGBOX PATH_VAR+VBCRLF+CSTR(LEN(PATH_VAR))
		SET_GO=FALSE
	END IF
	IF LEN(PATH_VAR)>68 THEN
		PATH_VAR=PATH_VAR_2
		BuildFullPath PATH_VAR
	END IF

	ON ERROR RESUME NEXT
	
	IF SET_GO=TRUE THEN
		IF FSO.FileExists(PATH_VAR+NewFilename)=TRUE THEN
			MSGBOX "FILE ALREADY EXISTS "+VBCRLF+VBCRLF+NewFilename,vbMsgBoxSetForeground
			Wscript.QUIT
		END IF
		
		NewFilename=PATH_VAR+"\"+NewFilename
		IF SET_GO=TRUE THEN
			' MSGBOX VAR_PATH_SCANNER+VBCRLF+ NewFilename,vbMsgBoxSetForeground
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