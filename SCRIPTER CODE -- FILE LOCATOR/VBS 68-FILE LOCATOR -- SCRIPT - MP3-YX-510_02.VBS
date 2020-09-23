Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true
Const ForReading = 1, ForWriting = 2, ForAppending = 8


Set Sound = CreateObject("WMPlayer.OCX.7")
Sound.URL = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV"
Sound.Controls.play

DIM FILE_TEST
FILE_TEST="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 68-FILE LOCATOR -- SCRIPT - MP3-YX-510.TXT"
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
IF FSO.FILEEXISTS(FILE_TEST) THEN 
	FSO.DELETEFile(FILE_TEST)
END IF
Set objFile = FSO.CreateTextFile(FILE_TEST,True)
objFile.Close

	

Public Function isValidFileName( ByVal strPath, ByVal strFileName)

	On Error Resume Next

	bIsValid = FALSE

	bIsValid = TRUE
	IF INSTR(strPath,"_gsdata_")>0 THEN bIsValid = FALSE
	
	isValidFileName = bIsValid

	IF isValidFileName=TRUE THEN 
		Set File = FSO.OpenTextFile(FILE_TEST, ForAppending)
		File.Write strPath + strFileName + VBCRLF
	END IF
	
End Function

