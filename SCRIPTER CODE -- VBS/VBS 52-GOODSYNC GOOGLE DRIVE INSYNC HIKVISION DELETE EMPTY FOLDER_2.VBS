'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 52-GOODSYNC GOOGLE DRIVE INSYNC HIKVISION DELETE EMPTY FOLDER_2.VBS
'# __ 
'# __ VBS 52-GOODSYNC GOOGLE DRIVE INSYNC HIKVISION DELETE EMPTY FOLDER_2.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START     TIME [ Sat 03-Aug-2019 12:38:20 ]
'# END FINAL TIME [ Sat 03-Aug-2019 12:38:20 ]
'# __ 
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
' -------------------------------------------------------------------
'#-------------------------------------------------------------------
' SESSION 001 __ CREATED
' THIS DO MORE THAN DELETE EMPTY FOLDER OF PREDECESSOR HERE
' VBS 52-GOODSYNC GOOGLE DRIVE INSYNC HIKVISION DELETE EMPTY FOLDER.VBS
' THIS ONE REALLY DO DELETE EVERY EMPTY FOLDER AT SUB FOLDER ALSO
' AND IT RUN DOS COMMAND TO DELETE ALL THUMB.DB SO DEL EMPTY WORK WELL
' IT WAS COPY MODIFY FROM
' C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 00-DELETE_EMPTY_FOLDER_2.VBS
'
' -------------------------------------------------------------------
'# START     TIME [ Sat 03-Aug-2019 12:38:20 ]
'# END FINAL TIME [ Sat 03-Aug-2019 12:38:20 ]
'#-------------------------------------------------------------------


Option Explicit

Dim strPath(2)

strPath(1) = "D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION"
strPath(2) = "D:\0 00 VIDEO SNAPSHOT"

DIM FILE_NAME , PARAM

Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

DIM R1
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
FILE_NAME="CMD"
FOR R1=1 TO 2
	PARAM="/C DEL """+strPath(R1)+"\THUMBS.DB"" /S /AH"
	'--------------------------------------------------------------
	WSHShell.Run """"+FILE_NAME+"""" + " " + ""+PARAM+"" , DontShowWindow, WaitUntilFinished
	'--------------------------------------------------------------
NEXT

DIM ERR_COUNTER
DIM TEST_COUNTER
DIM EXECUTE_LANUCHED
DIM ERROR_WITH_SOME

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim objFolder
DIM objFolderERR

DIM ERROR_WITH_SOME_2

FOR R1=1 TO 2
	IF LEN(strPath(R1)) = 2 AND MID(strPath(R1),2,1) = ":" THEN strPath(R1) = strPath(R1) + "\"
	IF fso.FOLDEREXISTS(strPath(R1)) = TRUE THEN 
		Set objFolder = fso.GetFolder(strPath(R1))
		' WScript.Echo strPath(R1) & VBCRLF & "TO GO _ KILL PROCESS IF NOT SURE ____"
		DeleteEmptyFolders objFolder
		EXECUTE_LANUCHED = TRUE
	ELSE
		
		' WScript.Echo strPath(R1) & VBCRLF & "__ NOT A VAILD FOLDER GIVEN __ EXIT END __"
	END IF
NEXT 


IF EXECUTE_LANUCHED = TRUE THEN 
	IF TEST_COUNTER="" THEN TEST_COUNTER=0
	' WScript.Echo "DELETE DONE ON _ " & CSTR(TEST_COUNTER) & " _ EMPTY FOLDER"
END IF

IF ERROR_WITH_SOME_2 <> "" THEN 
	MSGBOX ERROR_WITH_SOME_2
	DIM objShell
	Set objShell = CreateObject("Wscript.Shell")
	strPath = "explorer.exe /e," + ERROR_WITH_SOME_2
	objShell.Run strPath

	' WScript.Echo "ERROR WITH SOME OF THEM HERE _ " & VBCRLF & ERROR_WITH_SOME & VBCRLF & CSTR(ERR_COUNTER) & VBCRLF & "TAKE YOU THERE"
END IF

WScript.Quit(0)


Sub DeleteEmptyFolders(folder)
    Dim subfolder
	ON ERROR RESUME NEXT
    For Each subfolder in folder.SubFolders
		DeleteEmptyFolders subfolder
    Next
	TEST_COUNTER=TEST_COUNTER+1
	I1=-1
	I2=-1
	ERR.CLEAR
	Set objFolder = FSO.GetFolder(folder.Path)
	I1=objFolder.SubFolders.Count
	I2=objFolder.Files.Count
	If I1=0 AND I2=0 AND ERR.NUMBER=0 THEN 
		fso.DeleteFolder folder.Path ' , TRUE
		IF ERR.NUMBER>0 THEN 
			ERR_COUNTER = ERR_COUNTER +1
			ERROR_WITH_SOME_1 = folder.Path & VBCRLF & ERR.DESCRIPTION
			ERROR_WITH_SOME_2 = folder.Path

			MSGBOX "ERROR_WITH_SOME_1 = folder.Path & VBCRLF & ERR.DESCRIPTION"+VBCRLF+ERROR_WITH_SOME_1

			' ERR.CLEAR
			' objFolder.DELETE, TRUE
			' IF ERR.NUMBER=0 THEN ERR_COUNTER = ERR_COUNTER - 1
		END IF
	END IF
End Sub


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