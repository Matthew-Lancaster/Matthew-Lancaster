Option Explicit

' If (WScript.Arguments.Count <> 1) Then

    ' WScript.Echo("Usage: cscript DeleteEmptyFolders.vbs {path}")

    ' WScript.Quit(1)

' End If

Dim strPath

' strPath = WScript.Arguments(0)
 strPath = "G:\"

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
	WScript.Echo strPath & VBCRLF & "__ TO GO _ KILL PROCESS IF NOT SURE __"
	DeleteEmptyFolders objFolder
	EXECUTE_LANUCHED = TRUE
	ELSE
	
	WScript.Echo strPath & VBCRLF & "__ NOT A VAILD FOLDER GIVEN __ EXIT END __"
	
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
	WScript.Echo "DELETE DONE ON _ " & CSTR(TEST_COUNTER) & " _ EMPTY FOLDER/S"
END IF

IF ERROR_WITH_SOME <> "" THEN 
	DIM objShell
	' DIM strPath
	Set objShell = CreateObject("Wscript.Shell")
	' strPath = Wscript.Arguments(0)
	strPath = "explorer.exe /e," & ERROR_WITH_SOME
	objShell.Run strPath

	WScript.Echo "ERROR WITH SOME OF THEM HERE _ " & VBCRLF & ERROR_WITH_SOME & VBCRLF & CSTR(ERR_COUNTER) & VBCRLF & "TAKE YOU THERE"
	
	
END IF

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