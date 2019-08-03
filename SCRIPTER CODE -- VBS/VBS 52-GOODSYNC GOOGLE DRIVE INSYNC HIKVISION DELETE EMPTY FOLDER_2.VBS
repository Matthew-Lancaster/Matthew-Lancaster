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

Dim strPath

' strPath = WScript.Arguments(0)
strPath = "G:\"
strPath = "D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION\snapshot"

DIM FILE_NAME , PARAM

Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
FILE_NAME="CMD"
PARAM="/C DEL """+strPath+"\THUMBS.DB"" /S /AH"
'--------------------------------------------------------------
WSHShell.Run """"+FILE_NAME+"""" + " " + ""+PARAM+"" , DontShowWindow, WaitUntilFinished
'--------------------------------------------------------------

DIM ERR_COUNTER
DIM TEST_COUNTER
DIM EXECUTE_LANUCHED
DIM ERROR_WITH_SOME

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim objFolder
DIM objFolderERR

IF LEN(strPath) = 2 AND MID(strPath,2,1) = ":" THEN strPath = strPath + "\"

IF fso.FOLDEREXISTS(strPath) = TRUE THEN 

	Set objFolder = fso.GetFolder(strPath)
	' WScript.Echo strPath & VBCRLF & "TO GO _ KILL PROCESS IF NOT SURE ____"
	DeleteEmptyFolders objFolder
	EXECUTE_LANUCHED = TRUE
	ELSE
	
	' WScript.Echo strPath & VBCRLF & "__ NOT A VAILD FOLDER GIVEN __ EXIT END __"
	
END IF

Sub DeleteEmptyFolders(folder)

    Dim subfolder

	ON ERROR RESUME NEXT
	ERR.CLEAR
	
    For Each subfolder in folder.SubFolders
			DeleteEmptyFolders subfolder
    Next

    If folder.SubFolders.Count = 0 And folder.Files.Count = 0 Then

         'WScript.Echo folder.Path & " is empty"

		TEST_COUNTER=TEST_COUNTER+1
		ON ERROR RESUME NEXT
		ERR.CLEAR
		'IF DIR(folder.Path,VBDIRECTORY) <> "" THEN
		'	IF fso.FOLDEREXISTS(folder.Path) = TRUE THEN
		'		IF fso.ISFOLDER(folder.Path) = TRUE THEN
		'			IF fso.ISFILE(folder.Path) = FALSE THEN
		'				ERR.CLEAR
						fso.DeleteFolder folder.Path, TRUE
						IF ERR.NUMBER>0 THEN 
							ERR_COUNTER = ERR_COUNTER +1
							ERROR_WITH_SOME = folder.Path & VBCRLF & ERR.DESCRIPTION
							ERR.CLEAR
							objFolder.DELETE, TRUE
							IF ERR.NUMBER=0 THEN ERR_COUNTER = ERR_COUNTER - 1
						END IF
						
		'			END IF
		'		END IF
		'	END IF
		'END IF
		
    End If

End Sub

IF EXECUTE_LANUCHED = TRUE THEN 
	IF TEST_COUNTER="" THEN TEST_COUNTER=0
	' WScript.Echo "DELETE DONE ON _ " & CSTR(TEST_COUNTER) & " _ EMPTY FOLDER"
END IF

IF ERROR_WITH_SOME <> "" THEN 
	DIM objShell
	' DIM strPath
	Set objShell = CreateObject("Wscript.Shell")
	' strPath = Wscript.Arguments(0)
	strPath = "explorer.exe /e," & ERROR_WITH_SOME
	objShell.Run strPath

	' WScript.Echo "ERROR WITH SOME OF THEM HERE _ " & VBCRLF & ERROR_WITH_SOME & VBCRLF & CSTR(ERR_COUNTER) & VBCRLF & "TAKE YOU THERE"
END IF

WScript.Quit(0)

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