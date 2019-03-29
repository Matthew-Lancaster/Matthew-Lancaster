Attribute VB_Name = "mCompiler"
Option Explicit
Public LastProfile As String, F, R, Bads, Info()

Private Type SYSTEMTIME
    wYear             As Integer
    wMonth            As Integer
    wDayOfWeek        As Integer
    wDay              As Integer
    wHour             As Integer
    wMinute           As Integer
    wSecond           As Integer
    wMilliseconds     As Integer
End Type

Private Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type

Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public mProjects() As vProjects
'Dim fld As Folder

Public VBPath As String
Public CancelVBSearch As Boolean
Public VBSearchDone As Boolean
Public mCancelFileGet As Boolean

Public Type vProjects
   ProjectName     As String
   ProjectFullPath As String
   EXEName         As String
   ExeFullPath     As String
End Type




Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function TouchFileTimes Lib "imagehlp.dll" (ByVal FileHandle As Long, ByRef pSystemTime As Any) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'right order




Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
'Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'Option Explicit

 Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
 Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
 Const INFINITE = -1&

 
 
 Public Function shellAndWait(ByVal Filename As String) As Long
     Dim executionStatus As Long
     Dim processHandle As Long
     Dim ReturnValue As Long
     'Execute the application/file
     'executionStatus = Shell(fileName, vbNormalFocus)
     executionStatus = Shell(Filename, vbHide)

     'Get the Process Handle
     processHandle = OpenProcess(&H100000, True, executionStatus)

     'Wait till the application is finished
     ReturnValue = WaitForSingleObject(processHandle, INFINITE)

     'Send the Return Value Back
     shellAndWait = ReturnValue

 End Function







Function LastModifiedToCurrent(dufile As String)
    Dim lngHandle As Long
    lngHandle = CreateFile(dufile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If TouchFileTimes(lngHandle, ByVal 0&) = 0 Then
        'messageBox "Error while changing file dates. Access Denied." + vbCrLf + duFile, vbCritical, "Modify Error"
    LastModifiedToCurrent = False
    Else
        'messageBox "File date changed successfully.", vbInformation, "Modify Successful"
    LastModifiedToCurrent = True
    End If
    CloseHandle lngHandle
End Function
'---------------------------




Public Sub LockWindow(hwnd As Long, blnValue As Boolean)
    If blnValue Then
        LockWindowUpdate hwnd
    Else
        LockWindowUpdate 0
    End If
End Sub

Public Function FindFile(sDir As String, sFilename As String, sForm As Form) As String
    
    Dim sCurPath As String
    Dim sName As String
    Dim colFiles As Collection
    Dim sLastDir As String
    
    If Right(sDir, 1) <> "\" Then sDir = sDir & "\"
  
    mCancelFileGet = False
    sCurPath = sDir
    sName = vbNullString
    sName = Dir(sCurPath, vbDirectory)
    
    Do While Not (mCancelFileGet) And sName <> vbNullString  'A nullstring is returned when no files/dirs are left
        
      DoEvents
      If (GetAttr(sCurPath & sName) And vbDirectory) = vbDirectory Then
      
        If sName <> "." And sName <> ".." Then
          sLastDir = sName
          FindFile = FindFile(sCurPath & sName & "\", sFilename, sForm)
          sName = Dir(sCurPath, vbDirectory)
          While sName <> sLastDir And VBPath = ""
              sName = Dir
          Wend
        End If
        
      Else
      
        sForm.Label1 = sCurPath & sName
        If LCase$(sName) = LCase$(sFilename) Then
          mCancelFileGet = True
          VBPath = sCurPath & sName
          FindFile = VBPath
          Exit Function
        End If
        
      End If
      sName = Dir   'Get next file/dir in the loop
      
   Loop
   
End Function

Function GetFileName(vPath As String) As String
  Dim spot As Integer
  
  spot = InStrRev(vPath, "\")
  
  If spot <= 0 Then
    If Len(Trim(vPath)) = 0 Then
      GetFileName = ""
    Else
      GetFileName = vPath
    End If
  Else
    GetFileName = Mid(vPath, spot + 1)
  End If
End Function

Function GetDirectoryName(vPath As String) As String
  Dim spot As Integer
  Dim lSpot As Integer
  
  spot = InStr(spot + 1, vPath, "\")
  Do Until spot = 0
    lSpot = spot
    spot = InStr(spot + 1, vPath, "\")
  Loop
  
  If lSpot <= 0 Then
    GetDirectoryName = vPath
  Else
    GetDirectoryName = Left(vPath, lSpot)
  End If
End Function

Public Sub Pause(Seconds As Single)
    
    Dim T As Single
    Dim T2 As Single
    Dim Num As Single
    
    Num = Seconds * 1000
    
    T = GetTickCount()
    T2 = GetTickCount()
    
    Do Until T2 - T >= Num
        DoEvents
        T2 = GetTickCount()
    Loop
    
End Sub

