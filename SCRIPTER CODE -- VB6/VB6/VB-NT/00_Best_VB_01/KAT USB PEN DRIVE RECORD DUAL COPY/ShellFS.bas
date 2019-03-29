Attribute VB_Name = "ShellFS"
Dim DRIVESARR(27)


Private Type SHFILEOPSTRUCT
hWnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Long
hNameMappings As Long
lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" _
Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FOF_ALLOWUNDO = &H40&
Private Const FOF_NOCONFIRMATION = &H10&
Private Const FO_COPY = &H2&
Private Const FOF_SILENT As Long = &H4
Private Const FO_MOVE = &H1&
Private Const FO_DELETE = &H3&
Private Const FO_RENAME = &H4&
Private Const FOF_NOCONFIRMMKDIR = &H200&
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4          ' X - Cannot be set
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10      ' X - CreateDirectory to set
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800    ' X -


Dim WinType_SFO As SHFILEOPSTRUCT
Dim lRet As Long
Dim lflags As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
    
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias _
    "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, _
    ByVal dwFlags As Long) As Long

Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4


 Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
 Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
 Const INFINITE = -1&





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
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function

Sub FileInUseDelay(Tx$)
        
    qq = Now + TimeSerial(0, 1, 30)
    Do
        DoEvents
        ii = FileInUse(Tx$)
        If ii = True Then Sleep (200)
    Loop Until ii = False Or qq < Now
    
    If ii = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Tx$ + vbCrLf + "Longer than 1 min 30 secs"
        Stop
    End If
End Sub

Public Function EmptyRecycleBin(ByVal RootPath As String, Optional NoConfirmation As _
    Boolean, Optional NoProgress As Boolean, Optional NoSound As Boolean)
    Dim hWnd As Long, Flags As Long

    ' get the handle of the active window, or zero
    On Error Resume Next
    hWnd = Screen.ActiveForm.hWnd
    ' add a colon and backslash, if missing
    If Len(RootPath) > 0 And Mid$(RootPath, 2, 2) <> ":\" Then
        RootPath = Left$(RootPath, 1) & ":\"
    End If
    ' build the flags argument
    Flags = (NoConfirmation And SHERB_NOCONFIRMATION) Or (NoProgress And _
        SHERB_NOPROGRESSUI) Or (NoSound And SHERB_NOSOUND)
    SHEmptyRecycleBin hWnd, RootPath, Flags
End Function


Public Function ShellFileCopy(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

NoConfirm = True

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or FOF_SILENT
With WinType_SFO
.wFunc = FO_COPY
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileCopy = (lRet = 0)

End Function

Public Function ShellFileMove(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

'BETTER HAVE CHK FILE IN USE
'On Error Resume Next

'NoConfirm = True

'If Mid(src, Len(src) - 1, 1) <> "\" And InStr(src, "*") = 0 Then Form1.FileInUseDelay (src)

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_MOVE
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileMove = (lRet = 0)

End Function

Public Function ShellFileMove2(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

'BETTER HAVE CHK FILE IN USE
'On Error Resume Next

'NoConfirm = True



'If Mid(src, Len(src) - 1, 1) <> "\" And InStr(src, "*") = 0 Then Me.FileInUseDelay (src)

On Error Resume Next

lflags = FOF_ALLOWUNDO

If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or FOF_SILENT
With WinType_SFO
.wFunc = FO_MOVE
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileMove2 = (lRet = 0)

End Function

Public Function ShellFileDelete(src As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_DELETE
.pFrom = src
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileDelete = (lRet = 0)

End Function

Public Function ShellFileRename(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_RENAME
.pFrom = src
.pTo = dest
.fFlags = lflags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileRename = (lRet = 0)

End Function

Public Function FileExists(Fname As String) As Boolean

    If Fname = "" Or Right$(Fname, 1) = "\" Then
      FileExists = False: Exit Function
    End If

    FileExists = (Dir(Fname) <> "")
End Function

Public Function DirExists(ByVal DName As String) As Boolean




Dim sDummy As String
DirExists = False
On Error Resume Next


If Right(DName, 1) <> "\" Then DName = DName & "\"
'YY = Now + TimeSerial(0, 0, 2)
sDummy = Dir$(DName & "*.*", vbDirectory)
'IF YY<NOW THEN DRIVEAR

'P = Err.Description
If Err.Number > 0 Then Exit Function
DirExists = Not (sDummy = "")

End Function

Public Function DeleteOldFiles(DaysOld As Long, FileSpec As _
String, Optional ComparisonDate As Variant) As Boolean

'PURPOSE: DELETES ALL FILES THAT ARE DaysOld Older than
'ComparisonDate, which defaults to now

'RETURNS: True, if succesful
'False otherwise (e.g., user passes non-date argument
'deletion fails because file is in use,
'file doesn't exist, etc.)

'will not delete readonly, hidden or system files

Dim sFileSpec As String
Dim dCompDate As Date
Dim sFileName As String
Dim sFileSplit() As String
Dim iCtr As Integer, iCount As Integer
Dim sDir As String

sFileSpec = FileSpec

If IsMissing(ComparisonDate) Then
    dCompDate = Now
ElseIf Not IsDate(ComparisonDate) Then
    'client passed wrong type
    DeleteOldFiles = False
    Exit Function
Else
    dCompDate = CDate(Format(ComparisonDate, "mm/dd/yyyy"))
End If

sFileName = Dir(sFileSpec)

If sFileName = "" Then
    'returns false is file doesn't exist
    DeleteOldFiles = False
    Exit Function
End If

Do

    If sFileName = "" Then Exit Do

    If InStr(sFileSpec, "\") > 0 Then
        sFileSplit = Split(sFileSpec, "\")
        iCount = UBound(sFileSplit) - 1
        For iCtr = 0 To iCount
            sDir = sDir & sFileSplit(iCtr) & "\"
        Next

        sFileName = sDir & sFileName
    End If

    On Error GoTo errhandler:
    If DateDiff("d", FileDateTime(sFileName), dCompDate) _
       >= DaysOld Then

         Kill sFileName

    End If

    sFileName = Dir
    sDir = ""
Loop

DeleteOldFiles = True

Exit Function

errhandler:
    DeleteOldFiles = False
    Exit Function
End Function






 
 
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
     'ReturnValue = WaitForSingleObject(processHandle, INFINITE)
     
     'DONT WAIT LONG
     ReturnValue = WaitForSingleObject(processHandle, 1000)

     'Send the Return Value Back
     shellAndWait = ReturnValue

 End Function




