Attribute VB_Name = "APublics"
Public TEXT_PATH_1
Public TEXT_PATH_2

Public ProProjects, ProProjects2, TERMINATE_FORMS

Public LoadFolderFile
Public DaysToScan ' The Man
Public DaysToScanYears

Public AStart
Public FontSizez, SetTrueToLoadLast, LoadFolder
Public FS, F

Public A1$, B1$, C1$, OIP$, OIP2$

Public A4$(), B4$(), C4$()

Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

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

'----------------------------------------------------
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
'----------------------------------------------------


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



Public Sub Date_File(Filespec2$, Idate As Date)
'Call Date_File(filespec2$, Idate)


pdate = 0
Filespec = Filespec2$
Set F = FS.getfile((Filespec))
Idate = F.datelastmodified

End Sub


Public Function FindFileSize(Filename As String) As Long
        On Error Resume Next
        Dim filedata As WIN32_FIND_DATA

        filedata = Findfile(Filename)
        
        If filedata.nFileSizeHigh = 0 Then
            FindFileSize = Format$(filedata.nFileSizeLow, "###,###,###")
        Else
            FindFileSize = Format$(filedata.nFileSizeHigh, "###,###,###")
        End If
        On Error GoTo 0

End Function

Public Function Findfile(xstrfilename) As WIN32_FIND_DATA
    Dim Win32Data As WIN32_FIND_DATA
    Dim plngFirstFileHwnd As Long
    Dim plngRtn As Long
    plngFirstFileHwnd = FindFirstFile(xstrfilename, Win32Data)
    ' Get information of file using API call
    If plngFirstFileHwnd = 0 Then
        Findfile.cFileName = "Error"   ' If file was not found, return error as name
    Else
        Findfile = Win32Data    ' Else return results
    End If
    plngRtn = FindClose(plngFirstFileHwnd) ' It is important that you close the handle
                'for FindFirstFile
End Function



