
' -------------------------------------------------------------------
' WORK TIME
' -------------------------------------------------------------------
' SESSION 2 DAY 1
' Sun 26-Apr-2020 11:22:41
' Sun 26-Apr-2020 11:05:46 -- 12 HOUR 56 MINUTE
' -------------------------------------------------------------------
' SESSION 2 DAY 1
' Sun 26-Apr-2020 19:17:03
' Sun 26-Apr-2020 23:58:00 -- 4 HOUR 40 MINUTE
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' TO CREATE EFFORT FOR NEW MP3 UNIT
' -------------------------------------------------------------------
' 
' -------------------------------------------------------------------

' PATH TO LOAD AT FILELOCATOR 
' F:\MP3-YX-510_02_TS_VIDEO
' F:\MP3-YX-510_02_TS_VIDEO

' AFTER HERE RUN AND THEN HERE
' "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 95-SET_DATE_FILE__MP3-YX-510_FILE_SCRIPT.ahk"


Const ShowWindow = 1, DontShowWindow = 0, DontWaitUntilFinished = false, WaitUntilFinished = true
Const ForReading = 1, ForWriting = 2, ForAppending = 8


Set Sound = CreateObject("WMPlayer.OCX.7")
Sound.URL = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET\AKKORD.WAV"
Sound.Controls.play

DIM FILE_TEST_1
DIM FILE_TEST_2
FILE_TEST_1="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 68-FILE LOCATOR -- SCRIPT - MP3-YX-510_TIME_STRETCH_COUNTER.TXT"
FILE_TEST_2="C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPT\VBS 68-FILE LOCATOR -- SCRIPT - MP3-YX-510_FILE_SCRIPT_#NFS_EX_.TXT"
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

	IF INSTR(strPath,"_gsdata_")>0                                                     THEN bIsValid = FALSE
	IF INSTR(strFileName,".srf")>0                                                     THEN bIsValid = FALSE
	IF INSTR(strFileName,"System Volume Information")>0                                THEN bIsValid = FALSE
	IF INSTR(".MP3 .WAV .MP4",UCASE(MID(strFileName,INSTRREV(strFileName,"."))))=FALSE THEN bIsValid = FALSE
	
	
	
	isValidFileName = bIsValid
	
	IF isValidFileName=TRUE THEN 
		Set File = FSO.OpenTextFile(FILE_TEST_1, ForReading)
		TT_COUNT = File.ReadAll
		TIME_STRETCH_COUNTER=TT_COUNT+1
		Set File = FSO.OpenTextFile(FILE_TEST_1, ForWriting)
		File.Write TIME_STRETCH_COUNTER
		
		DT1=DateSerial(2000,1,1)+TIMESERIAL(1,0,0)
		
		TT_COUNT=5000-TT_COUNT
		DT1=DT1+TT_COUNT
		
		DT4=""
		DT4=DT4+LPad(YEAR(DT1), "0", 4)
		DT4=DT4+LPad(MONTH(DT1), "0", 2)
		DT4=DT4+LPad(DAY(DT1), "0", 2)
		Set File = FSO.OpenTextFile(FILE_TEST_2, ForAppending)
		File.Write CSTR(DT4) + " .. " + strPath + strFileName + VBCRLF
		
		' -----------------------------------------------------------
		' TEMP CODE RENAME 
		' TEMP ROUTINE RENAME
		' _T1_C1.mp4
		' -----------------------------------------------------------
		' _T1_C1.mp4
		' SET FSO = WSCRIPT.CREATEOBJECT("SCRIPTING.FILESYSTEMOBJECT")
		' X1=strPath + strFileName
		' X2=MID(X1,1,INSTRREV(X1,"_T")-1)+".MP4"
		' FSO.MOVEFILE X1, X2
		' -----------------------------------------------------------
 		' File.Write X2 + VBCRLF   ' DOUBLE CHECK WHAT RENAME DESTINATION
		' -----------------------------------------------------------
		
		' DT1= TimeSerial(0,TIME_STRETCH_COUNTER,0)
		' DT1= DT1 + DateSerial(2000,0,0)
		' TT= SetFileDateTime(strPath + strFileName,DT1)
		
	END IF
End Function


Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
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


