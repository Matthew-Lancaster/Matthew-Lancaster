DIM I,FSO
DIM VF1
DIM VF2
DIM COUNTER

Set FSO = CreateObject("scripting.filesystemobject")

Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
strSoundFile = "C:\WINDOWS\Media\Notify.wav"
Set objShell = CreateObject("Wscript.Shell")

strCommand = scriptdir & "\sndrec32 /play /close " & chr(34) & strSoundFile & chr(34)
If FSO.FileExists(scriptdir & "\sndrec32.EXE") = TRUE Then
	objShell.Run strCommand, 1, True
END If

Set fold = FSO.GetFolder(scriptdir)
Set files = fold.files
For Each file In files
	If INSTR(UCASE(file.NAME),".EXE")>0 Then 
		XYZ = scriptdir + "\" + file.NAME

		PATH_FILE_NAME1 = XYZ
		PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE'S\")

		ON ERROR RESUME NEXT
		XTIME=INT(NOW)+TIMER+8
		DO 
			COUNTER=COUNTER+1
			ERR.CLEAR
			Set F1 = FSO.GetFile(PATH_FILE_NAME1)
			Set F2 = FSO.GetFile(PATH_FILE_NAME2)
			IF F1.DateLastModified>F2.DateLastModified THEN 
				FSO.COPYFILE PATH_FILE_NAME1,PATH_FILE_NAME2
				MSGBOX "XYZ1"
			END IF
			IF F1.DateLastModified<F2.DateLastModified THEN 
				FSO.COPYFILE PATH_FILE_NAME2,PATH_FILE_NAME1
				MSGBOX "XYZ2"
			END IF
		LOOP UNTIL ERR.NUMBER=0 OR INT(NOW)+TIMER>XTIME 

		IF ERR.NUMBER>0 THEN
			MSGBOX "1 MINUTE __ COPY FAIL"+VBCRLF+"TRY RETRY"+VBCRLF+PATH_FILE_NAME2
		END IF
	END IF
NEXT

MSGBOX PATH_FILE_NAME1
Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run """"+PATH_FILE_NAME1+""""
Set objShell = Nothing
