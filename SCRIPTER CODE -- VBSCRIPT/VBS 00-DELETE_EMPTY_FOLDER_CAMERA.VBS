Option Explicit

Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

DIM FILE_NAME , PARAM
Dim STRPATH
DIM ERR_COUNTER
DIM TEST_COUNTER
DIM EXECUTE_LANUCHED
DIM ERROR_WITH_SOME

DIM READY_01
DIM READY_02

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

' If (WScript.Arguments.Count <> 1) Then
    ' WScript.Echo("Usage: cscript DeleteEmptyFolders.vbs {path}")
    ' WScript.Quit(1)
' End If
' STRPATH = WScript.Arguments(0)


STRPATH="G:\"
STRPATH="D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION\snapshot"
STRPATH="T:\VI_ DSC 01 V0 01 MM"
STRPATH="T:\VI_ DSC 01 V0 01 MM22"
'STRPATH=""
'STRPATH=""
STRPATH="X:\"

Dim objFolder
DIM objFolderERR

IF LEN(STRPATH) = 2 AND MID(STRPATH,2,1) = ":" THEN STRPATH = STRPATH + "\"

IF FSO.FOLDEREXISTS(STRPATH+"MP_ROOT") = FALSE THEN 
	WScript.Quit(1)
END IF

IF FSO.FOLDEREXISTS(STRPATH) = TRUE THEN 
		Set objFolder = FSO.GetFolder(STRPATH)
		' WScript.Echo WScript.ScriptName+VBCRLF+VBCRLF+STRPATH & VBCRLF & "TO GO _ KILL PROCESS IF NOT SURE ____"
		DeleteEmptyFolders objFolder
		EXECUTE_LANUCHED = TRUE
	ELSE
		WScript.Echo WScript.ScriptName+VBCRLF+VBCRLF+STRPATH & VBCRLF & "__ NOT A VAILD FOLDER GIVEN __ EXIT END __"
END IF

Sub DeleteEmptyFolders(folder)
    Dim subfolder
	ON ERROR RESUME NEXT
	ERR.CLEAR
    For Each subfolder in folder.SubFolders
			DeleteEmptyFolders subfolder
    Next
	' ---------------------------------------------------------------
	' MORE ERROR SAFE METHOD
	' REMEMBER DECLARE ON AS OPTION EXPLICIT 
	' AND ERROR RESULT FROM DECLARE ERROR MAKE NOT PROPER TRAPP IT
	' ---------------------------------------------------------------
	READY_01=-1
	READY_02=-1
	ERR.CLEAR
	READY_01=folder.SubFolders.Count
	READY_02=folder.Files.Count
	IF INSTR(folder.Path,":\PRIVATE")>0 THEN                      READY_01 = 10000
	IF INSTR(folder.Path,":\AVF_INFO")>0 THEN                     READY_01 = 10000
	IF INSTR(folder.Path,":\System Volume Information")>0 THEN    READY_01 = 10000
	IF INSTR(folder.Path,":\_gsdata_")>0 THEN                     READY_01 = 10000
	
	
	' ---------------------------------------------------------------
	' ---------------------------------------------------------------
    If READY_01 = 0 And READY_02 = 0 AND ERR.NUMBER=0 Then
		ON ERROR GOTO 0
		TEST_COUNTER=TEST_COUNTER+1
		ON ERROR RESUME NEXT
		ERR.CLEAR
		FSO.DeleteFolder folder.Path, TRUE
	End If
End Sub

IF EXECUTE_LANUCHED = TRUE THEN 
	IF TEST_COUNTER="" THEN TEST_COUNTER=0
	IF TEST_COUNTER>0 THEN 
		' -----------------------------------------------------------
		' SUCCESS ACHIEVE ON TEST RUN
		' ALLOW NOT MESSENGER LEAVER
		' Sun 08-Mar-2020 17:20:40
		' -----------------------------------------------------------
		' SURE WOULD LIKE MOVE ROUTINE HERE INTO OTHER CODE SET THAT CALLER
		' ONLY BIT OF DELETER ROUTINE
		' CALL FROM 
		' VBS 29-COPY CAMERA PHOTO IMAGES.VBS
		' OVER HERE
		' VBS 00-DELETE_EMPTY_FOLDER_CAMERA.VBS
		' -----------------------------------------------------------
		' WScript.Echo WScript.ScriptName+VBCRLF+VBCRLF+"DELETE DONE ON _ " & CSTR(TEST_COUNTER) & " _ EMPTY FOLDER"
	END IF
END IF

IF ERROR_WITH_SOME <> "" THEN 
	DIM objShell
	SET objShell = CreateObject("Wscript.Shell")
	' STRPATH = Wscript.Arguments(0)
	' STRPATH = "explorer.exe /e," & ERROR_WITH_SOME
	' objShell.Run STRPATH

	WScript.Echo WScript.ScriptName+VBCRLF+VBCRLF+"ERROR WITH SOME DELETE EMPTY FOLDER HERE _ " & VBCRLF & ERROR_WITH_SOME & VBCRLF & CSTR(ERR_COUNTER) & VBCRLF & "TAKE YOU THERE"
END IF

'IF DIR(folder.Path,VBDIRECTORY) <> "" THEN
'	IF FSO.FOLDEREXISTS(folder.Path) = TRUE THEN
'		IF FSO.ISFOLDER(folder.Path) = TRUE THEN
'			IF FSO.ISFILE(folder.Path) = FALSE THEN


' ---------------------
' Count = 011 -- Tue 05-Sep-2017 07:21:53
' ---------------------
' Form FindWindow ---
' Deleting Empty Folders – Random Musings of Jeremy Jameson - Google Chrome
' ---------------------
' ----
' Deleting Empty Folders – Random Musings of Jeremy Jameson
' https://blogs.msdn.microsoft.com/jjameson/2009/11/03/deleting-empty-folders/
' ----
' ---------------------
' Count = 012 -- Tue 05-Sep-2017 07:21:58
' ---------------------
' Form FindWindow ---
' VBSCRIPT TO DELETE EMPTY FOLDER - Google Search - Google Chrome
' ---------------------
' ----
' VBSCRIPT TO DELETE EMPTY FOLDER - Google Search
' https://www.google.co.uk/search?num=50&rlz=1C1CHBF_en-GBGB759GB759&q=VBSCRIPT+TO+DELETE+EMPTY+FOLDER&spell=1&sa=X&ved=0ahUKEwiu4uKqsY3WAhWCh7QKHfqSAVkQvwUIJSgA&biw=1536&bih=694
' ----
' ---------------------
' ---------------------
' Count = 013 -- Tue 05-Sep-2017 07:22:04
' ---------------------
' Form FindWindow ---
' VBS SCRIPT TO DELETE EMPTY FOLDER - Google Search - Google Chrome
' ---------------------
' ----
' VBS SCRIPT TO DELETE EMPTY FOLDER - Google Search
' https://www.google.co.uk/search?q=VBS+SCRIPT+TO+DELETE+EMPTY+FOLDER&rlz=1C1CHBF_en-GBGB759GB759&oq=VBS+SCRIPT+TO+DELETE+EMPTY+FOLDER&aqs=chrome..69i57j0l2.7816j0j7&sourceid=chrome&ie=UTF-8
' ----