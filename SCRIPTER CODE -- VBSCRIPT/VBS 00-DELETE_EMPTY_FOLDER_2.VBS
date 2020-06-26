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
STRPATH="I:\MUSIC"
STRPATH="T:\VI_ DSC 01 V0 02 V2_I2_MORPHIOUS"
STRPATH="J:\M"
STRPATH="F:\RETEKESS\M"
STRPATH="F:\RETEKESS--M"
STRPATH="D:\0 00 MUSIC ---\04 My Music\02 SINGLE\#007 J BOND\#008 FAVS 02"
STRPATH="D:\0 00 MUSIC ---"
STRPATH="D:\DL"
STRPATH="D:\# MY DOCS"
STRPATH="D:\# MY DOCS _ 02"
STRPATH="D:\My Webs2"
' STRPATH="D:\My Webs"
'STRPATH=""
'STRPATH=""
'STRPATH=""

Dim objFolder
DIM objFolderERR

IF LEN(STRPATH) = 2 AND MID(STRPATH,2,1) = ":" THEN STRPATH = STRPATH + "\"

IF FSO.FOLDEREXISTS(STRPATH) = TRUE THEN 
		Set objFolder = FSO.GetFolder(STRPATH)
		WScript.Echo STRPATH & VBCRLF & "TO GO _ KILL PROCESS IF NOT SURE ____"
		DeleteEmptyFolders objFolder
		EXECUTE_LANUCHED = TRUE
	ELSE
		WScript.Echo STRPATH & VBCRLF & "__ NOT A VAILD FOLDER GIVEN __ EXIT END __"
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
	' ---------------------------------------------------------------
	' ---------------------------------------------------------------
    If READY_01 = 0 And READY_02 = 0 AND ERR.NUMBER=0 Then
		ON ERROR GOTO 0
		TEST_COUNTER=TEST_COUNTER+1
		ON ERROR RESUME NEXT
		ERR.CLEAR
		FSO.DeleteFolder folder.Path, TRUE
		IF ERR.NUMBER>0 THEN 
			ERR_COUNTER = ERR_COUNTER +1
			ERROR_WITH_SOME = folder.Path & VBCRLF & ERR.DESCRIPTION
			ERR.CLEAR
			objFolder.DELETE, TRUE
			IF ERR.NUMBER=0 THEN ERR_COUNTER = ERR_COUNTER - 1
		END IF
	End If
End Sub

IF EXECUTE_LANUCHED = TRUE THEN 
	IF TEST_COUNTER="" THEN TEST_COUNTER=0
	WScript.Echo "DELETE DONE ON _ " & CSTR(TEST_COUNTER) & " _ EMPTY FOLDER"
END IF

IF ERROR_WITH_SOME <> "" THEN 
	DIM objShell
	SET objShell = CreateObject("Wscript.Shell")
	' STRPATH = Wscript.Arguments(0)
	' STRPATH = "explorer.exe /e," & ERROR_WITH_SOME
	' objShell.Run STRPATH

	WScript.Echo "ERROR WITH SOME OF THEM HERE _ " & VBCRLF & ERROR_WITH_SOME & VBCRLF & CSTR(ERR_COUNTER) & VBCRLF & "TAKE YOU THERE"
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