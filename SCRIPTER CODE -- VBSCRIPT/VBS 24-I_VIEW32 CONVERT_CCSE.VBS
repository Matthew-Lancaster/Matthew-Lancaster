' ----------------------------------------------------------------
' __ MATTHEW LANCASTER 
' __ MATT.LAN@BTINTERNET.COM
' ----------------------------------------------------------------
' __ FROM ---- Sun 27-May-2018 07:20:00
' __ TO   ---- Sun 27-May-2018 08:30:00 __ 1 HOUR 10 MINUTE
' ----------------------------------------------------------------

'-----------------------------------------------------------------
' HIKVISION CLOUD GOOGLE DRIVE IMAGES SNAPSHOT 
' CONVERT WITH IrfanView TO MAKE GOOGLE ABLE TO DISPLAY
'-----------------------------------------------------------------

' -------------------------------------------
' GOOGLE DRIVE HAS THE ABILITY TO SHOW THESE FILE NOW
' SO THIS PART OF PROJECT IS OVER REGARD TO I_VIEW32
' Thu 03-Oct-2019 11:22:32
' -------------------------------------------
' WScript.Quit(0)

' -------------------------------------------------------------------
' ' http://ss64.com/vb/popup.html
' ----
' Sat 25-Jan-2020 09:00:36
' ----
' VBS MSGBOX AND COUNTDOWN - Google Search
' https://www.google.com/search?q=VBS+MSGBOX+AND+COUNTDOWN
' ----
' VBSCRIPT MsgBox VbYesNo countdown - Stack Overflow
' https://stackoverflow.com/questions/19717496/vbscript-msgbox-vbyesno-countdown
' -------------------------------------------------------------------

Set objShell = CreateObject("WScript.Shell")
X = objShell.Popup("WORK TO DO -- VBS 24-I_VIEW32 CONVERT_CCSE.VBS -- QUIT NOW"+VBCR+VBCR+"4 SECOND GONE MESSENGER", 4,, vbOKOnly+ vbMsgBoxSetForeground)

WScript.Quit


' MsgBox Arguments
' Constant Value Description 
' vbOKOnly 0 OK button only (default) 
' vbOKCancel 1 OK and Cancel buttons 
' vbAbortRetryIgnore 2 Abort, Retry, and Ignore buttons 
' vbYesNoCancel 3 Yes, No, and Cancel buttons 
' vbYesNo 4 Yes and No buttons 
' vbRetryCancel 5 Retry and Cancel buttons 
' vbCritical 16 Critical message 
' vbQuestion 32 Warning query 
' vbExclamation 48 Warning message 
' vbInformation 64 Information message 
' vbDefaultButton1 0 First button is default (default) 
' vbDefaultButton2 256 Second button is default 
' vbDefaultButton3 512 Third button is default 
' vbDefaultButton4 768 Fourth button is default 
' vbApplicationModal 0 Application modal message box (default) 
' vbSystemModal 4096 System modal message box 
' vbMsgBoxHelpButton 16384 Adds Help button to the message box 
' VbMsgBoxSetForeground 65536 Specifies the message box window as the foreground window 
' vbMsgBoxRight 524288 Text is right aligned 
' vbMsgBoxRtlReading 1048576 Specifies text should appear as right-to-left reading on Hebrew and Arabic systems 


' MsgBox Return Values
' Constant Value Description 
' vbOK 1 OK button pressed 
' vbCancel 2 Cancel button pressed 
' vbAbort 3 Abort button pressed 
' vbRetry 4 Retry button pressed 
' vbIgnore 5 Ignore button pressed 
' vbYes 6 Yes button pressed 
' vbNo 7 No button pressed 


Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true
Const ForReading = 1, ForWriting = 2, ForAppending = 8

' -------------------------------------------------------------------
' CHECK IF ANY WSCRIPT ALREADY RUNNER AND QUIT IF ARE
' SOURCE GIVEN
' -------------------------------------------------------------------
' vbscript - how to check vbs script in windows is running or not? - Stack Overflow
' https://stackoverflow.com/questions/7849699/how-to-check-vbs-script-in-windows-is-running-or-not
' -------------------------------------------------------------------
' [ Wednesday 13:08:00 Pm_20 March 2019 ]
' -------------------------------------------------------------------

' -------------------------------------------------------------------
DIM CUR_USER
DIM ComputerName

CUR_USER = CreateObject("WScript.Network").UserName
Set objShell = Wscript.CreateObject("WScript.Shell")
ComputerName = objShell.ExpandEnvironmentStrings("%computername%")

' -------------------------------------------------------------------
' 01 OF 02 
' SOME CONFLICT WAR IN CODE STYLE 
' SOME PEOPLE TALK THIS HAS ERROR 
' THIS ONE IS CORRECT 
' -------------------------------------------------------------------
On Error Resume Next
Dim iMyPID : iMyPID = GetObject("winmgmts:root\cimv2").Get("Win32_Process.Handle='" & CreateObject("WScript.Shell").Exec("mshta.exe").ProcessID & "'").ParentProcessId
On Error Goto 0

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")
' -------------------------------------------------------------------

DIM VBV
VBV=VBCRLF

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")

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
		IF i2 > 0 THEN ' MY  PROGRAM WSCRIPT 
			' -------------------------------------------------------
			' KILL PROCESS NUMBER OF ALL OTHER INSTANCE 
			' HERE SCRIPT WHEN START
			' -------------------------------------------------------
			' IF lngMyProcessId<>iMyPID THEN 
				' I= "ERROR MISMATCH TRUE MYPID RESULT "
				' MSGBOX I +VBV+ CSTR(iMyPID) +VBV+ CSTR(lngMyProcessId)
			' -------------------------------------------------------
			IF iMyPID<>PID_Script THEN 
				ON ERROR RESUME NEXT
				objProcess.Terminate() 
				ON ERROR GOTO 0
				objShell.Run "TASKKILL /F /IM """+CSTR(PID_Script)+"""", DONTShowWindow, WaitUntilFinished
				' MSGBOX CSTR(iMyPID) +VBV+ strScriptName +VBV+ CSTR(PID_Script)  +VBV+ Wscript.ScriptName
			END IF
		END IF
    END IF
NEXT

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
Set OBJFSO = CreateObject("Scripting.FileSystemObject")

Dim oFile
Dim oFSO
DIM COUNT_1
DIM COUNT_2
DIM objBASEFolder

COUNT_1=0
COUNT_2=0

i_view32="""C:\Program Files (x86)\IrfanView\i_view32.exe"""

FOLDER="D:\0 00 VIDEO SNAPSHOT"
objBASEFolder = FOLDER
IF FSO.FOLDEREXISTS(FOLDER) = TRUE THEN
	Call ShowSubfolders (objFSO.GetFolder(FOLDER))
END IF

FOLDER="D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION"
objBASEFolder = FOLDER
IF FSO.FOLDEREXISTS(FOLDER) = TRUE THEN
	' Call ShowSubfolders (objFSO.GetFolder(FOLDER))
END IF


FOR RF1=1 TO 4
	DeleteFolder_HAPPEN=0
	FOLDER="D:\0 00 VIDEO SNAPSHOT"
	IF FSO.FOLDEREXISTS(FOLDER) = TRUE THEN
		Call ShowSubfolders_FOLDER (objFSO.GetFolder(FOLDER))
	END IF

	FOLDER="D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION"
	IF FSO.FOLDEREXISTS(FOLDER) = TRUE THEN
		Call ShowSubfolders_FOLDER (objFSO.GetFolder(FOLDER))
	END IF

	IF DeleteFolder_HAPPEN=0 THEN 
		EXIT FOR
	END IF
NEXT



' =C DRIVE:\Users\MATT 04\Insync\rub.rim@gmail.com\Google Drive

' MSGBOX "DONE __ " ' & COUNT_1 & " _OF_ " & COUNT_2


Set WSHShell = Nothing
WScript.Quit(0)



' -------------------------------------------------------------------
' -------------------------------------------------------------------
' END OF CODE MAIN ROUTINE
' -------------------------------------------------------------------
' -------------------------------------------------------------------


SUB SUB_DO_WITH_FOLDER(FOLDER_PATH)

	FILE_NAME=FOLDER_PATH+"\desktop.ini"
	IF FSO.FILEEXISTS(FILE_NAME)=TRUE THEN 
		' MSGBOX FILE_NAME
		FSO.DELETEFILE(FILE_NAME)
	END IF
	FILE_NAME=FOLDER_PATH+"\thumbs.db"
	IF FSO.FILEEXISTS(FILE_NAME)=TRUE THEN 
		' MSGBOX FILE_NAME
		FSO.DELETEFILE(FILE_NAME)
	END IF
	
	' MSGBOX FOLDER_PATH
	IF INSTR(FOLDER_PATH,"\snapshot")>0 THEN 
	
		ON ERROR RESUME NEXT
		I1=-1
		I2=-1
		ERR.CLEAR 
		Set objFolder = FSO.GetFolder(FOLDER_PATH)
		I1=objFolder.SubFolders.Count
		I2=objFolder.Files.Count
		FILENAME_TO_DEL=objFolder.Path
		If I1=0 AND I2=0 AND ERR.NUMBER=0 THEN  
			' MSGBOX "FOLDER DEL"+VBCRLF+FILENAME_TO_DEL
			FSO.DeleteFolder FILENAME_TO_DEL,TRUE
			DeleteFolder_HAPPEN=TRUE
			If FSO.FolderExists(FILENAME_TO_DEL) Then
				MSGBOX "FOLDER EXIST STILL"+VBCRLF+FILENAME_TO_DEL
			END IF
		END IF
	END IF
	
	
END SUB


' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------

SUB SUB_DO_WITH_FILE(FILE_PATH)

	Dim objFILE
	Set objFILE = objFSO.GetFILE(FILE_PATH)
	
	IF NOT_REQUIRE_ANYMORE=22 THEN
		' INPUT PAST END OF FILE HAPPEN ERROR MESSENGER
		ON ERROR RESUME NEXT
		readBinary=""
		With objFILE.OpenAsTextStream()
			readBinary = .Read(119+4)
			.Close
		End With
		ON ERROR GOTO 0
		Set objFILE = NOTHING

		IF MID(readBinary,119,4)=STRING(4,CHR(&H1E)) THEN 
			' -------------------------------------------
			' jpgq=88 __ UP UNTIL 02 MAR 2019
			' jpgq=92 __       ON 02 MAR 2019
			' QUALITY JPEG CONVERSION LEVEL 0 TO 100%
			' jpgq=94 __       ON 20 SEP 2019
			' AT 94 TYPICAL ONE CHECK
			' TPath: /snapshot/20190813/D1/DVR_D01_Motion_Detected_20190813172414_3.jpg
			' Lnew: [file] size=62,091
			' Rnew: [file] size=36,214
			' SO IS A LOT BIGGER
			' GOT TO TRY DO CONVERSION AND KEEP SAME
			' -------------------------------------------
			' GOOGLE DRIVE HAS THE ABILITY TO SHOW THESE FILE NOW
			' SO THIS PART OF PROJECT IS OVER REGARD TO I_VIEW32
			' Thu 03-Oct-2019 11:22:32
			' -------------------------------------------
			
			' WSHShell.Run i_view32 + " """+FILE_PATH+""" /jpgq=98 /convert=$D$F" , DontShowWindow, WaitUntilFinished
			
			' COUNT_1=COUNT_1+1
		END IF
	END IF 
	COUNT_2=COUNT_2+1
	
	PATH_VAR_2=MID(FILE_PATH,1,INSTRREV(FILE_PATH,"\"))
	STRIP_BASE_FOLDER=REPLACE(PATH_VAR_2,objBASEFolder,"")
	NEW_BASE_FOLDER_1="D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION"+STRIP_BASE_FOLDER
	STRIP_FOLDER_FROM_FILE=MID(FILE_PATH,INSTRREV(FILE_PATH,"\")+1)
	NEW_BASE_FOLDER_2=NEW_BASE_FOLDER_1

	IF INSTRREV(NEW_BASE_FOLDER_1," (")>0 THEN 
		NEW_BASE_FOLDER_2=NEW_BASE_FOLDER_1
		FOR R2=1 TO 100
			NEW_BASE_FOLDER_2=REPLACE(NEW_BASE_FOLDER_2," ("+TRIM(CSTR(R2))+")","")
		NEXT
	END IF
	
	IF NEW_BASE_FOLDER_1<>NEW_BASE_FOLDER_2 THEN
		' MSGBOX NEW_BASE_FOLDER_1+VBCRLF+NEW_BASE_FOLDER_2
	END IF
	
	STRIP_FOLDER_FROM_FILE=MID(FILE_PATH,INSTRREV(FILE_PATH,"\")+1)
	NEW_FILENAME=NEW_BASE_FOLDER_2+STRIP_FOLDER_FROM_FILE

	If Not FSO.FolderExists(NEW_BASE_FOLDER_2) Then
		CreateFolderTree(NEW_BASE_FOLDER_2)
	End If
	ON ERROR GOTO 0

	ERR.CLEAR
	' IF ERROR OF CODE HERE IT HAS JUMP OVER CODE OUT OF SUBROUTINE
	' HEAVIER DEBUG THAN THINK
	' ---------------------------------------------------------------
	IF FSO.FILEEXISTS(NEW_FILENAME)=FALSE THEN 
		' MSGBOX NEW_FILENAME
		FSO.MoveFILE FILE_PATH, NEW_FILENAME
	END IF
	
	ERR.CLEAR
	IF FSO.FILEEXISTS(NEW_FILENAME)=FALSE THEN 
		MSGBOX "MOVE 2ND METHOD "+VBCRLF+FILE_PATH+VBCRLF+NEW_FILENAME
		Set objFILE = objFSO.GetFILE(FILE_PATH)
		objFILE.Move NEW_FILENAME ' NEW_BASE_FOLDER_2
	END IF
	
	IF FSO.FILEEXISTS(NEW_FILENAME)=FALSE THEN 
		IF ERR.NUMBER=0 THEN
			MSGBOX "COPY METHOD "+VBCRLF+FILE_PATH+VBCRLF+NEW_FILENAME
			FSO.COPYFile FILE_PATH, NEW_FILENAME ' NEW_BASE_FOLDER_2
		END IF
	END IF
	
	' DELETE SOURCE IF BOTH EXIST
	IF NEW_FILENAME<>FILE_PATH THEN 
		IF FSO.FILEEXISTS(NEW_FILENAME)=TRUE THEN 
			IF FSO.FILEEXISTS(FILE_PATH)=TRUE THEN 
				' MSGBOX "DELETE SOURCE BOTH EXIST"+VBCRLF+FILE_PATH+VBCRLF+NEW_FILENAME
				FSO.DELETEFILE(FILE_PATH)
			END IF
		END IF
	END IF
	
END SUB
				
				
				


' PRETTY SURE THIS CODE SCRIPT DOES THIS BIT NOW
' Sat 05-Oct-2019 10:52:22
				
' IF COUNT_1>0 THEN 

	' FILE_NAME="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 52-GOODSYNC GOOGLE DRIVE INSYNC HIKVISION DELETE EMPTY FOLDER.VBS"
	' PARAM=""
	' '--------------------------------------------------------------
	' WSHShell.Run """"+FILE_NAME+"""" + " " + ""+PARAM+"" , ShowWindow, DontWaitUntilFinished
	' '--------------------------------------------------------------

' END IF 



Set WSHShell = Nothing
WScript.Quit(0)

' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' -------------------------------------------------------------------

Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

'-----------------------------------------------
FUNCTION FOLDER_SCAN_COUNTER(fFolder,FOLDER_SCAN_COUNTER_VAR)
'-----------------------------------------------
	FILE_SCAN_COUNTER_VAR=0
	Set objFolder = objFSO.GetFolder(fFolder.Path)
    Set colFiles = objFolder.Files
	Dim objFile
	For Each objFile in colFiles
		FILE_NAME=objFile.NAME
		IF MID(FILE_NAME,1,4)="DVR_" THEN 
			'Set oFile = oFSO.GetFile(objFile.PATH)
			FOLDER_SCAN_COUNTER_VAR=FOLDER_SCAN_COUNTER_VAR+objFile.SIZE
		END IF
	Next
End FUNCTION



'-------------------------------------
Sub ShowSubFolders(fFolder)
'-------------------------------------
    Set objFolder = objFSO.GetFolder(fFolder.Path)
    Set colFiles = objFolder.Files
	ON ERROR RESUME NEXT
	Dim objFile
	
	IF INSTR(objFolder.PATH,"\_gsdata_")=0 then 
		For Each objFile in colFiles
		
			EXTENSION_T = UCASE(objFSO.getextensionname(objFile.path))
			IF EXTENSION_T = "JPG" THEN 
				CALL SUB_DO_WITH_FILE(objFile)
			END IF
		NEXT 
	END IF
	
    For Each Subfolder in fFolder.SubFolders
        ShowSubFolders(Subfolder)
    Next
End Sub


'-------------------------------------
Sub ShowSubFolders_FOLDER(fFolder)
'-------------------------------------
    Set objFolder = objFSO.GetFolder(fFolder.PATH)
    Set colFiles = objFolder.Files
	ON ERROR RESUME NEXT
	
	CALL SUB_DO_WITH_FOLDER(objFolder)
	
    For Each Subfolder in fFolder.SubFolders
        ShowSubFolders_FOLDER(Subfolder)
    Next
End Sub


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



'--------------------------------------------------------------------
' SPARE CODE
'--------------------------------------------------------------------
' FullPath=FOLDER_2
' If Not fso.FolderExists(FullPath) Then
	' 'fso.CreateFolder FullPath
' End If
'--------------------------------------------------------------------

'--------------------------------------------------------------------
' Set objSuperFolder = FSO.GetFolder(FOLDER_1)
' CALL FOLDER_SCAN_COUNTER(objSuperFolder,FOLDER_SCAN_COUNTER_1)
' Set objSuperFolder = FSO.GetFolder(FOLDER_2)
' CALL FOLDER_SCAN_COUNTER(objSuperFolder,FOLDER_SCAN_COUNTER_2)

'MSGBOX FOLDER_SCAN_COUNTER_1 & VBCRLF & FOLDER_SCAN_COUNTER_2
'--------------------------------------------------------------------


'--------------------------------------------------------------------
' REF
'--------------------------------------------------------------------

' ----
' scripting - Help To create Folder1/Folder2 in Windows using VBScript ( Both the folders not exists before, i mean to create multilevel folders @ a strech.) - Stack Overflow
' https://stackoverflow.com/questions/4407386/help-to-create-folder1-folder2-in-windows-using-vbscript-both-the-folders-not
' --------
' vba - VBScript - How to make program wait until process has finished? - Stack Overflow
' https://stackoverflow.com/questions/10279404/vbscript-how-to-make-program-wait-until-process-has-finished
' --------
' Read and write binary file in VBscript - Stack Overflow
' https://stackoverflow.com/questions/6060529/read-and-write-binary-file-in-vbscript
' ----