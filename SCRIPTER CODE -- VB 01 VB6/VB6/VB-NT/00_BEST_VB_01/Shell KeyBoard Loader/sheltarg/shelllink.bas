Attribute VB_Name = "Module1"
Option Explicit
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const SHGFI_TYPENAME = &H400
Public Const MAX_PATH = 260
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

