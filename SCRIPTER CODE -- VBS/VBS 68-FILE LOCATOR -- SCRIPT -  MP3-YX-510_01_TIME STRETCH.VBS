Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true
Const ForReading = 1, ForWriting = 2, ForAppending = 8


Set Sound = CreateObject("WMPlayer.OCX.7")
Sound.URL = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV"
Sound.Controls.play

DIM FILE_TEST_1
DIM FILE_TEST_2
FILE_TEST_1="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 68-FILE LOCATOR -- SCRIPT - MP3-YX-510_TIME_STRETCH_COUNTER.TXT"
FILE_TEST_2="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 68-FILE LOCATOR -- SCRIPT - MP3-YX-510_FILE_SCRIPT.TXT"
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
IF FSO.FILEEXISTS(FILE_TEST_1) THEN 
	FSO.DELETEFile(FILE_TEST_1)
END IF
IF FSO.FILEEXISTS(FILE_TEST_2) THEN 
	FSO.DELETEFile(FILE_TEST_2)
END IF
Set objFile = FSO.CreateTextFile(FILE_TEST_1,True)
objFile.Close
Set objFile = FSO.CreateTextFile(FILE_TEST_2,True)
objFile.Close

TIME_STRETCH_COUNTER=0	


' -------------------------------------------------------------------
' AFTER FILE WRITE AUOTHOTKEY DO CHANGE DATE 
' -------------------------------------------------------------------


Public Function isValidFileName( ByVal strPath, ByVal strFileName)

	On Error Resume Next

	bIsValid = FALSE
	bIsValid = TRUE

	IF INSTR(strPath,"_gsdata_")>0 THEN bIsValid = FALSE
	IF INSTR(strFileName,".srf")>0 THEN bIsValid = FALSE
	
	isValidFileName = bIsValid
	
	IF isValidFileName=TRUE THEN 
		Set File = FSO.OpenTextFile(FILE_TEST_1, ForReading)
		TT_COUNT = File.ReadAll
		TIME_STRETCH_COUNTER=TT_COUNT+1
		Set File = FSO.OpenTextFile(FILE_TEST_1, ForWriting)
		File.Write TIME_STRETCH_COUNTER
		
		Set F = FSO.GetFile(strPath + strFileName)
        DT1 = F.datelastmodified
		' DT1= TimeSerial(0,0,0)
		DT1= DateSerial(1000+TIME_STRETCH_COUNTER,1,1)

		Set File = FSO.OpenTextFile(FILE_TEST_2, ForAppending)
		File.Write CSTR(DT1) + " .. " + strPath + strFileName + VBCRLF
		
		' DT1= TimeSerial(0,TIME_STRETCH_COUNTER,0)
		' DT1= DT1 + DateSerial(2000,0,0)
		' TT= SetFileDateTime(strPath + strFileName,DT1)
		
	END IF
End Function






' Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
' Private Declare Function SetFileTime Lib "kernel32" _
   ' (ByVal hFile As Long, ByVal MullP As Long, _
    ' ByVal NullP2 As Long, lpLastWriteTime _
    ' As FILETIME) As Long
' Private Declare Function SystemTimeToFileTime Lib _
    ' "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime _
    ' As FILETIME) As Long

' Public Function SetFileDateTime(ByVal Filename As String, _
  ' ByVal TheDate As String) As Boolean
' '************************************************
' 'PURPOSE:    Set File Date (and optionally time)
' '            for a given file)
         
' 'PARAMETERS: TheDate -- Date to Set File's Modified Date/Time
' '            FileName -- The File Name

' 'Returns:    True if successful, false otherwise
' '************************************************
' If Dir(Filename) = "" Then Exit Function
' If Not IsDate(TheDate) Then Exit Function

' Dim lFileHnd As Long
' Dim lRet As Long

' Dim typFileTime As FILETIME
' Dim typLocalTime As FILETIME
' Dim typSystemTime As SYSTEMTIME

' With typSystemTime
    ' .wYear = Year(TheDate)
    ' .wMonth = Month(TheDate)
    ' .wDay = Day(TheDate)
    ' .wDayOfWeek = Weekday(TheDate) - 1
    ' .wHour = Hour(TheDate)
    ' .wMinute = Minute(TheDate)
    ' .wSecond = Second(TheDate)
' End With

' lRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
' lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)

' lFileHnd = CreateFile(Filename, GENERIC_WRITE, _
    ' FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
    ' OPEN_EXISTING, 0, 0)
    
' lRet = SetFileTime(lFileHnd, ByVal 0&, ByVal 0&, _
         ' typFileTime)

' CloseHandle lFileHnd
' SetFileDateTime = lRet > 0

' End Function


