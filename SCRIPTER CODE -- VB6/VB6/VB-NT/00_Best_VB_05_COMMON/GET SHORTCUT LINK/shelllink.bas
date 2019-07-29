Attribute VB_Name = "SHELLARG"
Option Explicit

Public lRet As Long, sLinkPath As String, sTemp As String, sLinkDescript As String
Public txtTargetPath
Public txtTargetDescript

Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const SHGFI_TYPENAME = &H400
Private Const MAX_PATH = 260
Public Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type

Public Declare Function GetLinkInfo Lib "sheltarg.dll" (ByVal hWnd As Long, ByVal lpszLinkFile As String, ByVal lpszPath As String, ByVal lpszDescription As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long


Public Sub GETSHORTLINK(LINKTXT)

'MAKE SURE CHDRIVE APP.PATH AND CHDIR APP.PATH -- ARE SET FOR "sheltarg.dll"
ChDrive App.Path
ChDir App.Path

'Dim lRet As Long, sLinkPath As String, sTemp As String, sLinkDescript As String
txtTargetPath = ""

'Fill the vars that will receive info from DLL
sLinkPath = Space$(MAX_PATH)
sLinkDescript = Space$(MAX_PATH)
'Get the link from the list
sTemp = LINKTXT
'Make the call
lRet = GetLinkInfo(Form1.hWnd, sTemp, sLinkPath, sLinkDescript)
'A null terminated string is returned
txtTargetPath = StripTerminator(sLinkPath)
'Not all have a description
sLinkDescript = StripTerminator(sLinkDescript)

'##MAYBE
'If there is not one get File Type of the target
'If Len(Trim$(sLinkDescript)) = 0 Then
'    sLinkDescript = GetFileType(sLinkPath)
'End If

txtTargetDescript = sLinkDescript


End Sub


Function GetFileType(sFileName As String) As String
Dim himl As Long
Dim shinfo As SHFILEINFO

On Error Resume Next
himl = SHGetFileInfo(sFileName, FILE_ATTRIBUTE_NORMAL Or FILE_ATTRIBUTE_DIRECTORY, shinfo, Len(shinfo), SHGFI_TYPENAME)
If himl <> 0 Then
    GetFileType = shinfo.szTypeName
End If

End Function

Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function


'Declare Function GetShortPathName Lib "kernel32" _
'      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
'      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'
'Public Function GetShortName(ByVal sLongFileName As String) As String
'' --> (All this modules code) Obtained from : -
'' ---> Microsoft's/MSDN's code
'' ->  http://support.microsoft.com/kb/q175512/
'' ---> The original comments were by them :
'
'       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
'       'Set up buffer area for API function call return
'       sShortPathName = Space(255)
'       iLen = Len(sShortPathName)
'
'       'Call the function
'       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
'       'Strip away unwanted characters.
'       GetShortName = Left(sShortPathName, lRetVal)
'End Function


Public Function GetLongName(ByVal sShortName As String) As String

' --> (All this modules code) Obtained from : -
' ---> Microsoft's/MSDN's code
' ->  http://support.microsoft.com/default.aspx?scid=kb;EN-US;154822
' ---> The original comments were by them :
    
    Dim sLongName As String
    Dim sTemp As String
    Dim iSlashPos As Integer
    
    'Add \ to short name to prevent Instr from failing
    
    sShortName = sShortName & "\"
    
    Dim DRIVE_LETTER
    
    'Start from 4 to ignore the "[Drive Letter]:\" characters
    iSlashPos = InStr(4, sShortName, "\")
    DRIVE_LETTER = Left$(sShortName, 2)
    ' --------------------------------------
    ' IF NETWORK FOLDER
    ' [ Monday 09:30:00 Am_29 July 2019 ]
    ' --------------------------------------
    If Mid(sShortName, 1, 2) = "\\" Then
        iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
        DRIVE_LETTER = Left$(sShortName, iSlashPos)
        iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
    End If
    
    DRIVE_LETTER = Mid(DRIVE_LETTER, 1, Len(DRIVE_LETTER) - 1)
    
    'Pull out each string between \ character for conversion
    While iSlashPos
      sTemp = Dir(Left$(sShortName, iSlashPos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
      If sTemp = "" Then
        'Error 52 - Bad File Name or Number
        GetLongName = ""
        Exit Function
      End If
      sLongName = sLongName & "\" & sTemp
      iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
    Wend
    
    'Prefix with the drive letter
    GetLongName = DRIVE_LETTER & sLongName
    
End Function



