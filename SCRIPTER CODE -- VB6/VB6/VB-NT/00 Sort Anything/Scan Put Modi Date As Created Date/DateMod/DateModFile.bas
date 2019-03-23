Attribute VB_Name = "DateModFile"
Option Explicit

'THIS TOGGLES A SMALL PASSWORD FEATURE THAT I EMBEDDED
Const MPassword = False

'THIS IS THE CODE THAT MUST BE ENTERED (FINISHED BY PRESSING HASH)
'WHEN THE PASSWORD TOGGLE IS TRUE
Const MSequence = 22515

Dim MEntered As Variant

 Private Type FILETIME
     dwLowDate  As Long
     dwHighDate As Long
 End Type
 
 Private Type SYSTEMTIME
     wYear      As Integer
     wMonth     As Integer
     wDayOfWeek As Integer
     wDay       As Integer
     wHour      As Integer
     wMinute    As Integer
     wSecond    As Integer
     wMillisecs As Integer
 End Type
 
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_WRITE = &H40000000
  
Private Declare Function CreateFile Lib "kernel32" Alias _
   "CreateFileA" (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) _
   As Long

Private Declare Function LocalFileTimeToFileTime Lib _
     "kernel32" (lpLocalFileTime As FILETIME, _
      lpFileTime As FILETIME) As Long

Private Declare Function SetFileTime Lib "kernel32" _
   (ByVal hFile As Long, ByVal MullP As Long, _
    ByVal NullP2 As Long, lpLastWriteTime _
    As FILETIME) As Long

Private Declare Function SystemTimeToFileTime Lib _
    "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime _
    As FILETIME) As Long
    
Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long

Public Function SetFileDateTime(ByVal FileName As String, _
  ByVal TheDate As String) As Boolean
'************************************************
'PURPOSE:    Set File Date (and optionally time)
'            for a given file)
         
'PARAMETERS: TheDate -- Date to Set File's Modified Date/Time
'            FileName -- The File Name

'Returns:    True if successful, false otherwise
'************************************************
If Dir(FileName) = "" Then Exit Function
If Not IsDate(TheDate) Then Exit Function

Dim lFileHnd As Long
Dim lRet As Long

Dim typFileTime As FILETIME
Dim typLocalTime As FILETIME
Dim typSystemTime As SYSTEMTIME

With typSystemTime
    .wYear = Year(TheDate)
    .wMonth = Month(TheDate)
    .wDay = Day(TheDate)
    .wDayOfWeek = Weekday(TheDate) - 1
    .wHour = Hour(TheDate)
    .wMinute = Minute(TheDate)
    .wSecond = Second(TheDate)
End With

lRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)

lFileHnd = CreateFile(FileName, GENERIC_WRITE, _
    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
    OPEN_EXISTING, 0, 0)
    
lRet = SetFileTime(lFileHnd, ByVal 0&, ByVal 0&, _
         typFileTime)

CloseHandle lFileHnd
SetFileDateTime = lRet > 0

End Function


