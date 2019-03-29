VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ControlPanelForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   """Remote"" Control Panel"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   Icon            =   "ControlPanelForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8235
   Begin VB.Frame Frame3 
      Caption         =   "Fun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   2850
      TabIndex        =   17
      Top             =   128
      Width           =   5265
      Begin VB.CommandButton NotePadBttn 
         Caption         =   "Run Notepad"
         Height          =   435
         Left            =   2730
         TabIndex        =   27
         Top             =   2370
         Width           =   2145
      End
      Begin VB.CommandButton PaintBttn 
         Caption         =   "Run Paint"
         Height          =   435
         Left            =   2730
         TabIndex        =   26
         Top             =   1860
         Width           =   2145
      End
      Begin VB.CommandButton GameBttn 
         Caption         =   "Run Game [Solitair]"
         Height          =   435
         Left            =   2730
         TabIndex        =   25
         Top             =   1350
         Width           =   2145
      End
      Begin VB.CommandButton CDDoorBttn 
         Caption         =   "Open CD Door"
         Height          =   435
         Left            =   2730
         TabIndex        =   24
         Top             =   840
         Width           =   2145
      End
      Begin VB.CommandButton DBTimeBttn 
         Caption         =   "Set Mouse DblClick Speed"
         Height          =   435
         Left            =   2730
         TabIndex        =   23
         Top             =   330
         Width           =   2145
      End
      Begin VB.CommandButton SSaverBttn 
         Caption         =   "Screen Saver"
         Height          =   435
         Left            =   270
         TabIndex        =   22
         Top             =   2370
         Width           =   2145
      End
      Begin VB.CommandButton NormalMouseBttn 
         Caption         =   "Normal Mouse"
         Height          =   435
         Left            =   270
         TabIndex        =   21
         Top             =   1860
         Width           =   2145
      End
      Begin VB.CommandButton CrezyMouseBttn 
         Caption         =   "Crezy Mouse"
         Height          =   435
         Left            =   270
         TabIndex        =   20
         Top             =   1350
         Width           =   2145
      End
      Begin VB.CommandButton ClipMouseBttn 
         Caption         =   "Clip Mouse Cursor"
         Height          =   435
         Left            =   270
         TabIndex        =   19
         Top             =   840
         Width           =   2145
      End
      Begin VB.CommandButton SwapMouseBttn 
         Caption         =   "Swap Mouse Button [R]"
         Height          =   435
         Left            =   270
         TabIndex        =   18
         Top             =   330
         Width           =   2145
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operations"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3075
      Left            =   2850
      TabIndex        =   9
      Top             =   3165
      Width           =   5265
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1530
         Top             =   1710
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ControlPanelForm.frx":0442
               Key             =   "date&time"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ControlPanelForm.frx":0894
               Key             =   "netword"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ControlPanelForm.frx":0CE6
               Key             =   "display"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ControlPanelForm.frx":24A8
               Key             =   "font"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ControlPanelForm.frx":2902
               Key             =   "keyboard"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ControlPanelForm.frx":2D54
               Key             =   "mouse"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ControlPanelForm.frx":31A6
               Key             =   "regional"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ControlPanelForm.frx":3D78
               Key             =   "users&password"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ControlPanelForm.frx":652A
               Key             =   "system"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ControlPanelForm.frx":8D6C
               Key             =   "phone"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ControlPanelForm.frx":9086
               Key             =   "sound"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1365
         Left            =   120
         TabIndex        =   28
         Top             =   1590
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   2408
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton LogOffBttn 
         Caption         =   "&Log Off"
         Height          =   405
         Left            =   1370
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton ShutDownBttn 
         Caption         =   "Sh&ut Down"
         Height          =   405
         Left            =   3930
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton RestartBttn 
         Caption         =   "&Restart"
         Height          =   405
         Left            =   2650
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton LockBttn 
         Caption         =   "L&ock System"
         Height          =   405
         Left            =   90
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton SysInfoBttn 
         Caption         =   "Get System &Info"
         Height          =   405
         Left            =   90
         TabIndex        =   12
         Top             =   840
         Width           =   1635
      End
      Begin VB.CommandButton GetUserBttn 
         Caption         =   "Get User Name"
         Height          =   405
         Left            =   1800
         TabIndex        =   11
         Top             =   840
         Width           =   1635
      End
      Begin VB.CommandButton LogInBttn 
         Caption         =   "Get Login &Time"
         Height          =   405
         Left            =   3510
         TabIndex        =   10
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Open Control Panel Items:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1350
         Width           =   2235
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6105
      Left            =   120
      TabIndex        =   0
      Top             =   128
      Width           =   2715
      Begin MSWinsockLib.Winsock sckControlPanel 
         Left            =   780
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton EMailBttn 
         Caption         =   "&E-Mail"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   2505
      End
      Begin VB.TextBox txtEMail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Text            =   "email@email.com"
         Top             =   2490
         Width           =   2505
      End
      Begin VB.CommandButton ClipboardBttn 
         Caption         =   "&Set Clipboard Text"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   2105
         Width           =   2505
      End
      Begin VB.TextBox txtClipboard 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Text            =   "Set Clipboard Text"
         Top             =   1720
         Width           =   2505
      End
      Begin VB.CommandButton wwwBttn 
         Caption         =   "&WWW Popup"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1335
         Width           =   2505
      End
      Begin VB.TextBox txtWWW 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "www.yahoo.com"
         Top             =   950
         Width           =   2505
      End
      Begin VB.CommandButton msgBttn 
         Caption         =   "Send &Message"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   565
         Width           =   2505
      End
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "Message Text"
         Top             =   180
         Width           =   2505
      End
   End
End
Attribute VB_Name = "ControlPanelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Sub CDDoorBttn_Click()
    If CDDoorBttn.Caption = "Open CD Door" Then
        sckControlPanel.SendData "CDDoor#Open"
        CDDoorBttn.Caption = "Close CD Door"
    Else
        sckControlPanel.SendData "CDDoor#Close"
        CDDoorBttn.Caption = "Open CD Door"
    End If
End Sub

Private Sub ClipboardBttn_Click()
    sckControlPanel.SendData "SetClipboard#" & txtClipboard.Text
End Sub

Private Sub ClipMouseBttn_Click()
    If ClipMouseBttn.Caption = "Clip Mouse Cursor" Then
        sckControlPanel.SendData "ClipCursor#"
        ClipMouseBttn.Caption = "UnClip Mouse Cursor"
    Else
        sckControlPanel.SendData "UnClipCursor#"
        ClipMouseBttn.Caption = "Clip Mouse Cursor"
    End If
End Sub

Private Sub CrezyMouseBttn_Click()
    sckControlPanel.SendData "CrezyMouse#"
End Sub

Private Sub DBTimeBttn_Click()
Dim txtTime As String
    txtTime = InputBox("Enter Double Click Time:", "Remote Control")
    If Trim(txtTime) = "" Then
        MsgBox "Please enter value"
        Exit Sub
    End If
    sckControlPanel.SendData "MouseSpeed#" & txtTime
End Sub

Private Sub EMailBttn_Click()
    sckControlPanel.SendData "EMail#" & txtEMail.Text
End Sub

Private Sub Form_Load()
    sckControlPanel.RemoteHost = RemoteForm.sckIdentifier.RemoteHost
    sckControlPanel.RemotePort = 8005
    sckControlPanel.Connect
End Sub

Private Sub Form_Resize()
Dim LItem As ListItem
        
    ListView1.ListItems.Add , , "Date and Time", "date&time"
    ListView1.ListItems.Add , , "Display", "display"
    ListView1.ListItems.Add , , "System", "system"
    ListView1.ListItems.Add , , "Fonts", "font"
    ListView1.ListItems.Add , , "Keyboard", "keyboard"
    ListView1.ListItems.Add , , "Mouse", "mouse"
    ListView1.ListItems.Add , , "Sound and Multimedia", "sound"
    ListView1.ListItems.Add , , "Phone and Modem", "phone"
    ListView1.ListItems.Add , , "Network and Dialup Connects", "netword"
End Sub

Private Sub GameBttn_Click()
    sckControlPanel.SendData "Solitaire#"
End Sub

Private Sub GetUserBttn_Click()
    sckControlPanel.SendData "GetUserName#"
End Sub

Private Sub ListView1_DblClick()
Dim dPoint As POINTAPI
Dim X As Single, Y As Single
Dim LItem As ListItem
    
    GetCursorPos dPoint
    
    X = dPoint.X - ScaleX(Me.Left + Frame2.Left + ListView1.Left, vbTwips, vbPixels)
    Y = dPoint.Y - ScaleY(Me.Top + Frame2.Top + ListView1.Top, vbTwips, vbPixels)
    X = ScaleX(X, vbPixels, vbTwips)
    Y = ScaleY(Y, vbPixels, vbTwips)
On Error Resume Next
    Set LItem = ListView1.HitTest(X, Y)
    If LItem Is Nothing Then Exit Sub
    MsgBox LItem.Text
End Sub

Private Sub LockBttn_Click()
    sckControlPanel.SendData "LockSystem#"
End Sub

Private Sub LogInBttn_Click()
    sckControlPanel.SendData "GetTickCount#"
End Sub

Private Sub LogOffBttn_Click()
    sckControlPanel.SendData "Logoff#"
End Sub

Private Sub msgBttn_Click()
    sckControlPanel.SendData "Message#" & txtMsg.Text
End Sub

Private Sub NormalMouseBttn_Click()
    sckControlPanel.SendData "NormalMouse#"
End Sub

Private Sub NotePadBttn_Click()
    sckControlPanel.SendData "Notepad#"
End Sub

Private Sub PaintBttn_Click()
    sckControlPanel.SendData "Paint#"
End Sub

Private Sub RestartBttn_Click()
    sckControlPanel.SendData "Reboot#"
End Sub

Private Sub sckControlPanel_Close()
    'If sckControlPanel.State = sckNotConnected Then
    Unload Me
End Sub

Private Sub sckControlPanel_DataArrival(ByVal bytesTotal As Long)
Dim Command As String, txtData As String
Dim Position As Integer
    
    sckControlPanel.getData txtData
    
    Position = InStr(1, txtData, "#")
    Command = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Select Case Command
        Case "GetUserName"
            MsgBox txtData
        Case "GetTickCount"
            MsgBox Format(txtData, "###")
        Case "SystemInfo"
            MsgBox txtData
    End Select
End Sub

Private Sub ShutDownBttn_Click()
    sckControlPanel.SendData "Shutdown#"
End Sub

Private Sub SSaverBttn_Click()
    sckControlPanel.SendData "ScreenSaver#"
End Sub

Private Sub SwapMouseBttn_Click()
    If SwapMouseBttn.Caption = "Swap Mouse Button [R]" Then
        sckControlPanel.SendData "SwapMouse#Right"
        SwapMouseBttn.Caption = "Swap Mouse Button [L]"
    Else
        sckControlPanel.SendData "SwapMouse#Left"
        SwapMouseBttn.Caption = "Swap Mouse Button [R]"
    End If
End Sub

Private Sub SysInfoBttn_Click()
    sckControlPanel.SendData "SystemInfo#"
End Sub

Private Sub wwwBttn_Click()
    sckControlPanel.SendData "WWW#" & txtWWW.Text
End Sub
