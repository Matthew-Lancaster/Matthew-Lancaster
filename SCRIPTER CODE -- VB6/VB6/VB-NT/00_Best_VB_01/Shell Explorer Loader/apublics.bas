Attribute VB_Name = "APublics"
Public NET_C_PATH
Public NET_D_PATH
Public NET_E_PATH

Public COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER
Public NETWORK_2_STEP_JUMPER

Public TEXT_PATH_1
Public TEXT_PATH_2

Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Public FontSizez, SetTrueToLoadLast, LoadFolder

Public A1$, B1$, C1$, OIP$, OIP2$

Public A4$(), B4$(), C4$()

Public FS
Public FSO

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Type FILETIME
   LowDateTime          As Long
   HighDateTime         As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * 260  'MUST be set to 260
   cAlternate           As String * 14
End Type


Public Sub SET_UP_PULIC_FSO()

Set FS = CreateObject("Scripting.FileSystemObject")
Set FSO = CreateObject("Scripting.FileSystemObject")

End Sub


'Public Function CreateFolderTree(ByVal sPath As String) As Boolean
'    Dim nPos As Integer
'
'    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)
'
'    On Error GoTo CreateFolderTreeError
'
'    nPos = InStr(sPath, "\")
'    While nPos > 0
'        If Not FolderExists(Left$(sPath, nPos - 1)) Then
'            MkDir Left$(sPath, nPos - 1)
'        End If
'        nPos = InStr(nPos + 1, sPath, "\")
'    Wend
'    If Not FolderExists(sPath) Then MkDir sPath
'
'    CreateFolderTree = True
'    Exit Function
'
'CreateFolderTreeError:
'    Exit Function
'End Function

Public Function FolderExists(sFolder As String) As Boolean
    '##############################################################################################
    'Returns True if the specified folder exists
    '##############################################################################################
    
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFolder, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FolderExists = True
        Else
            FolderExists = False
        End If
    End If
End Function



