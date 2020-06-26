Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true
Const ForReading = 1, ForWriting = 2, ForAppending = 8


Set Sound = CreateObject("WMPlayer.OCX.7")
Sound.URL = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV"
Sound.Controls.play

DIM FILE_TEST
FILE_TEST="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 75-FILE LOCATOR -- SCRIPT -  SONY CAMERA HX60 CHECK FILE TOO LONG.VBS.TXT"
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
IF FSO.FILEEXISTS(FILE_TEST) THEN 
	FSO.DELETEFile(FILE_TEST)
END IF
Set objFile = FSO.CreateTextFile(FILE_TEST,True)
objFile.Close

DIM XX
DIM PATH_STRING_PUT_OLD

Public Function isValidFileName( ByVal strPath, ByVal strFileName)

	bIsValid = FALSE

	IF LEN(strFileName)>12 THEN 
		bIsValid = TRUE
		Set File = FSO.OpenTextFile(FILE_TEST, ForAppending)
		File.Write CSTR(LEN(strFileName)) +"__"+ strFileName + VBCRLF
	END IF

	isValidFileName=bIsValid
	
End Function

