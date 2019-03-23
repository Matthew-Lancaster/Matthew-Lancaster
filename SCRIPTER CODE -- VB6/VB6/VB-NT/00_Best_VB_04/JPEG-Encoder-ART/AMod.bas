Attribute VB_Name = "AMod"
Public ShowPicFrm2Loaded
Public DD22, DD2, BreakScan, OTimer1 As Boolean, StopScan, NoFiles
Public ReadArray(), ReadArra2(), ReadArra3(), ReadArra4(), ReadArra5(), ReadArra6(), ReadArra7()
Public TotallyFiles, AgWa As Integer, ScanInProgress, OldAgWa2
Public AutoCount, OldCountPic1, FirstRun2, OldAgWa
Public FirstRun3, DONTWARNERROR
Public FS, F, OX As Long, OY As Long, OX1 As Long, OY1 As Long
Public CTCOL(), Control, i, BreakScanNow, ScanFinished, FirstDirScanned, MoreThanOne

Public OCheckQ, T1
Public A1, B1, G1, FF, ScanPathFin



Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPI_GETSCREENSAVEACTIVE = 16
Private Const SPIF_SENDWININICHANGE = &H2
Private Const SPIF_UPDATEINIFILE = &H1

Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
    ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long


'Option Explicit

Private Type SHFILEOPSTRUCT
hwnd As Long
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
Dim lFlags As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
    
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias _
    "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, _
    ByVal dwFlags As Long) As Long

Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4

Public Function EmptyRecycleBin(ByVal RootPath As String, Optional NoConfirmation As _
    Boolean, Optional NoProgress As Boolean, Optional NoSound As Boolean)
    Dim hwnd As Long, Flags As Long

    ' get the handle of the active window, or zero
    On Error Resume Next
    hwnd = Screen.ActiveForm.hwnd
    ' add a colon and backslash, if missing
    If Len(RootPath) > 0 And Mid$(RootPath, 2, 2) <> ":\" Then
        RootPath = Left$(RootPath, 1) & ":\"
    End If
    ' build the flags argument
    Flags = (NoConfirmation And SHERB_NOCONFIRMATION) Or (NoProgress And _
        SHERB_NOPROGRESSUI) Or (NoSound And SHERB_NOSOUND)
    SHEmptyRecycleBin hwnd, RootPath, Flags
End Function


Public Function ShellFileCopy(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lFlags = FOF_ALLOWUNDO
If NoConfirm Then lFlags = lFlags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_COPY
.pFrom = src
.pTo = dest
.fFlags = lFlags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileCopy = (lRet = 0)

End Function

Public Function ShellFileMove(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lFlags = FOF_ALLOWUNDO
If NoConfirm Then lFlags = lFlags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_MOVE
.pFrom = src
.pTo = dest
.fFlags = lFlags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileMove = (lRet = 0)

End Function

Public Function ShellFileDelete(src As String, _
Optional NoConfirm As Boolean = False) As Boolean

lFlags = FOF_ALLOWUNDO
If NoConfirm Then lFlags = lFlags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_DELETE
.pFrom = src
.fFlags = lFlags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileDelete = (lRet = 0)

End Function

Public Function ShellFileRename(src As String, dest As String, _
Optional NoConfirm As Boolean = False) As Boolean

lFlags = FOF_ALLOWUNDO
If NoConfirm Then lFlags = lFlags Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
With WinType_SFO
.wFunc = FO_RENAME
.pFrom = src
.pTo = dest
.fFlags = lFlags
End With

lRet = SHFileOperation(WinType_SFO)
ShellFileRename = (lRet = 0)

End Function

'Put This In Declarations Sub zzLoad_Checks()






' Example of how to call bmp_rotate.
Sub Command1_Click()
    
    Dim ANGLE
    
    Const pi = 3.14159265359

    For ANGLE = pi / 6 To 2 * pi Step pi / 6
'    Picture2.Cls
'    Call bmp_rotate(Picture1, Picture2, angle)
    Next
End Sub


'http://support.microsoft.com/kb/80406


' bmp_rotate(pic1, pic2, theta)
     ' Rotate the image in a picture box.
     '   pic1 is the picture box with the bitmap to rotate
     '   pic2 is the picture box to receive the rotated bitmap
     '   theta is the angle of rotation
     '

     Sub bmp_rotate(pic1 As Control, pic2 As Control, ByVal Theta!)
       Const pi = 3.14159265359
       Dim c1x As Integer  ' Center of pic1.
       Dim c1y As Integer  '   "
       Dim c2x As Integer  ' Center of pic2.
       Dim c2y As Integer  '   "
       Dim A As Single     ' Angle of c2 to p2.
       Dim R As Integer    ' Radius from c2 to p2.
       Dim p1x As Integer  ' Position on pic1.
       Dim p1y As Integer  '   "
       Dim p2x As Integer  ' Position on pic2.
       Dim p2y As Integer  '   "
       Dim n As Integer    ' Max width or height of pic2.

       ' Compute the centers.
       c1x = pic1.ScaleWidth / 2
       c1y = pic1.ScaleHeight / 2
       c2x = pic2.ScaleWidth / 2
       c2y = pic2.ScaleHeight / 2

       ' Compute the image size.
       n = pic2.ScaleWidth
       If n < pic2.ScaleHeight Then n = pic2.ScaleHeight
       n = n / 2 - 1
       ' For each pixel position on pic2.
       For p2x = 0 To n
          For p2y = 0 To n
             ' Compute polar coordinate of p2.
             If p2x = 0 Then
               A = pi / 2
             Else
               A = Atn(p2y / p2x)
             End If
             R = Sqr(1& * p2x * p2x + 1& * p2y * p2y)

             ' Compute rotated position of p1.
             p1x = R * Cos(A + Theta)
             p1y = R * Sin(A + Theta)

             ' Copy pixels, 4 quadrants at once.
             c0& = pic1.POINT(c1x + p1x, c1y + p1y)
             c1& = pic1.POINT(c1x - p1x, c1y - p1y)
             c2& = pic1.POINT(c1x + p1y, c1y - p1x)
             c3& = pic1.POINT(c1x - p1y, c1y + p1x)
             If c0& <> -1 Then pic2.PSet (c2x + p2x, c2y + p2y), c0&
             If c1& <> -1 Then pic2.PSet (c2x - p2x, c2y - p2y), c1&
             If c2& <> -1 Then pic2.PSet (c2x + p2y, c2y - p2x), c2&
             If c3& <> -1 Then pic2.PSet (c2x - p2y, c2y + p2x), c3&
          Next
          ' Allow pending Windows messages to be processed.
          t% = DoEvents()
       Next
     
     'Picture
     
     End Sub
  



'Private Const SPI_SETSCREENSAVEACTIVE = 17
'Private Const SPI_GETSCREENSAVEACTIVE = 16
'Private Const SPIF_SENDWININICHANGE = &H2
'Private Const SPIF_UPDATEINIFILE = &H1

'Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
    ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long




' return the Enabled state of the screen saver

Function GetScreenSaverState() As Boolean
    Dim Result As Long
    SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0, Result, 0
    GetScreenSaverState = (Result <> 0)
End Function

' enable or disable the screen saver
'
' if second argument is true, it writes changes in user's profile
' returns True if the operation was successful, False otherwise

Function SetScreenSaverState(ByVal enabled As Boolean, _
    Optional ByVal PermanentChange As Boolean) As Boolean
    Dim fuWinIni As Long
    If PermanentChange Then
        fuWinIni = SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
    End If
    SetScreenSaverState = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, enabled, _
        ByVal 0&, fuWinIni) <> 0
End Function




