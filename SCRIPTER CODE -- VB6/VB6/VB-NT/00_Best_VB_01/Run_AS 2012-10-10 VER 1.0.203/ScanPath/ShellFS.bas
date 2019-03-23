Attribute VB_Name = "ShellFS"

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

Public Function EmptyRecycleBin(ByVal RootPath As String, Optional NoConfirmation As _
    Boolean, Optional NoProgress As Boolean, Optional NoSound As Boolean)
    Dim hWnd As Long, flags As Long

    ' get the handle of the active window, or zero
    On Error Resume Next
    hWnd = Screen.ActiveForm.hWnd
    ' add a colon and backslash, if missing
    If Len(RootPath) > 0 And Mid$(RootPath, 2, 2) <> ":\" Then
        RootPath = Left$(RootPath, 1) & ":\"
    End If
    ' build the flags argument
    flags = (NoConfirmation And SHERB_NOCONFIRMATION) Or (NoProgress And _
        SHERB_NOPROGRESSUI) Or (NoSound And SHERB_NOSOUND)
    SHEmptyRecycleBin hWnd, RootPath, flags
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

Private Function ShellFileDelete(src As String, _
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

If Right(DName, 1) <> "\" Then DName = DName & "\"
sDummy = Dir$(DName & "*.*", vbDirectory)
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






