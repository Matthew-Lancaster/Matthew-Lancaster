Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true
Const ForReading = 1, ForWriting = 2, ForAppending = 8


Set Sound = CreateObject("WMPlayer.OCX.7")
Sound.URL = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV"
Sound.Controls.play

DIM FILE_TEST
FILE_TEST="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 68-FILE LOCATOR -- SCRIPT -  RETEKESS.TXT"
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

	On Error Resume Next

	bIsValid = FALSE

	bIsValid = TRUE
	IF INSTR(strPath & "\" & strFileName,"_gsdata_")>0 THEN bIsValid = FALSE
	IF FSO.FOLDEREXISTS(strPath & "\" & strFileName) THEN bIsValid = FALSE
	
	isValidFileName = bIsValid

	IF isValidFileName=TRUE THEN 
		Set File = FSO.OpenTextFile(FILE_TEST, ForAppending)
		XX=XX+1
		PATH_STRING=strPath + strFileName
		X1=INSTRREV(PATH_STRING,"\")
		X1=INSTRREV(PATH_STRING,"\",X1-1)
		PATH_STRING=MID(PATH_STRING,X1)

		PATH_STRING_2=strPath
		X1=INSTRREV(PATH_STRING_2,"\")
		X1=INSTRREV(PATH_STRING_2,"\",X1-1)
		PATH_STRING_2=MID(PATH_STRING_2,X1)
		PATH_STRING_2=MID(PATH_STRING_2,2)
		
		PATH_STRING_PUT=strPath
		IF PATH_STRING_PUT_OLD<>PATH_STRING_PUT THEN 
		File.Write CSTR(XX) +"__"+ PATH_STRING_2 + VBCRLF
		END IF 
		PATH_STRING_PUT_OLD=PATH_STRING_PUT
		
' 		File.Write CSTR(XX) +"  "+ PATH_STRING + VBCRLF
	END IF
	
End Function

