VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm RemoteForm 
   AutoShowChildren=   0   'False
   BackColor       =   &H00000000&
   Caption         =   "Remote Control 1.2"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10830
   Icon            =   "RemotForm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock sckIdentifier 
      Left            =   2430
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   7965
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16034
            Picture         =   "RemotForm.frx":030A
            Text            =   "Not Connected   "
            TextSave        =   "Not Connected   "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:54 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4590
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RemotForm.frx":075C
            Key             =   "connect"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RemotForm.frx":0BAE
            Key             =   "disconnect"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RemotForm.frx":1000
            Key             =   "chat"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RemotForm.frx":131A
            Key             =   "talkchat"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RemotForm.frx":1774
            Key             =   "explorer"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RemotForm.frx":3FB6
            Key             =   "registry"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RemotForm.frx":4E08
            Key             =   "desktop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RemotForm.frx":525A
            Key             =   "taskmanager"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RemotForm.frx":56AC
            Key             =   "controlpanel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "connect"
            Object.ToolTipText     =   "Connect with Remote Computer"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "disconnect"
            Object.ToolTipText     =   "Disconnect with Remote Computer"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "chat"
            Object.ToolTipText     =   "Remote Chat"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "talkchat"
            Object.ToolTipText     =   "Talk Chat"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "explorer"
            Object.ToolTipText     =   "Remote Explorer"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "registry"
            Object.ToolTipText     =   "Remote Registry Editor"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "remotedesktop"
            Object.ToolTipText     =   "Remote Desktop"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "taskmanager"
            Object.ToolTipText     =   "Remote Taskmanager"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "controlpanel"
            ImageKey        =   "controlpanel"
         EndProperty
      EndProperty
   End
   Begin VB.Menu ConnectMenu 
      Caption         =   "&Connect"
      Begin VB.Menu ConnectCmd 
         Caption         =   "&Connect"
         Shortcut        =   {F5}
      End
      Begin VB.Menu DisconnectCmd 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
         Shortcut        =   {F6}
      End
      Begin VB.Menu Sep01 
         Caption         =   "-"
      End
      Begin VB.Menu ExitCmd 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu ToolMenu 
      Caption         =   "&Tools"
      Enabled         =   0   'False
      Begin VB.Menu RChat 
         Caption         =   "&Chat"
      End
      Begin VB.Menu RTalkChat 
         Caption         =   "&Talk Chat"
      End
      Begin VB.Menu RExplorer 
         Caption         =   "&Explorer"
      End
      Begin VB.Menu RRegistry 
         Caption         =   "&Registry"
      End
      Begin VB.Menu RDesktop 
         Caption         =   "D&esktop"
      End
      Begin VB.Menu RTask 
         Caption         =   "Task &Manager"
      End
      Begin VB.Menu RControlPanel 
         Caption         =   "Control &Panel"
      End
   End
   Begin VB.Menu WindowMenu 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu TileHoriz 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu TileVer 
         Caption         =   "Tile Vertical"
      End
      Begin VB.Menu Casecade 
         Caption         =   "Casecade"
      End
      Begin VB.Menu ArrangeIco 
         Caption         =   "Arrange Icons"
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu Help 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "RemoteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub About_Click()
    AboutForm.Show 1
End Sub

Private Sub ArrangeIco_Click()
    RemoteForm.Arrange vbArrangeIcons
End Sub

Private Sub Casecade_Click()
    RemoteForm.Arrange vbCascade
End Sub

Private Sub ConnectCmd_Click()
Dim RemoteHost As String
On Error GoTo ConnectError
    RemoteHost = InputBox("Please enter Remote Computer's Name or IP Address.", "Connect with Remote Computer")
    If Trim(RemoteHost) = "" Then
        MsgBox "Please enter correct Remote Computer's Name or IP Address.", vbInformation
        RemoteHost = ""
        Exit Sub
    End If
    sckIdentifier.RemoteHost = RemoteHost
    sckIdentifier.RemotePort = 8000
    sckIdentifier.Connect
    Exit Sub
ConnectError:
    sckIdentifier.Close
End Sub

Private Sub DisconnectCmd_Click()
    sckIdentifier.Close
    Call sckIdentifier_Close
End Sub

Private Sub ExitCmd_Click()
    End
End Sub

Private Sub Help_Click()
    MsgBox "Sorry, help is under construction.", vbInformation, "Remote Control 1.2"
End Sub

Private Sub MDIForm_Load()
    If App.PrevInstance = True Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub RChat_Click()
    ChatForm.Show
End Sub

Private Sub RControlPanel_Click()
    ControlPanelForm.Show
End Sub

Private Sub RDesktop_Click()
    RemoteDesktop.Show
End Sub

Private Sub RExplorer_Click()
    ExplorerForm.Show
End Sub

Private Sub RRegistry_Click()
    RegForm.Show
End Sub

Private Sub RTalkChat_Click()
    TalkChatForm.Show
End Sub

Private Sub RTask_Click()
    TaskmanagerFrm.Show
End Sub

Private Sub sckIdentifier_Close()
If sckIdentifier.State <> sckNotConnected Then
    ConnectCmd.Enabled = True
    DisconnectCmd.Enabled = False
    ToolMenu.Enabled = False
    Toolbar1.Buttons("connect").Enabled = True
    Toolbar1.Buttons("disconnect").Enabled = False
    Toolbar1.Buttons("chat").Enabled = False
    Toolbar1.Buttons("talkchat").Enabled = False
    Toolbar1.Buttons("explorer").Enabled = False
    Toolbar1.Buttons("registry").Enabled = False
    Toolbar1.Buttons("remotedesktop").Enabled = False
    Toolbar1.Buttons("taskmanager").Enabled = False
    Toolbar1.Buttons("controlpanel").Enabled = False
    StatusBar1.Panels(1).Picture = ImageList1.ListImages(2).Picture
    StatusBar1.Panels(1).Text = "Not Connected   "
End If
End Sub

Private Sub sckIdentifier_Connect()
    ConnectCmd.Enabled = False
    DisconnectCmd.Enabled = True
    ToolMenu.Enabled = True
    Toolbar1.Buttons("connect").Enabled = False
    Toolbar1.Buttons("disconnect").Enabled = True
    Toolbar1.Buttons("chat").Enabled = True
    Toolbar1.Buttons("talkchat").Enabled = True
    Toolbar1.Buttons("explorer").Enabled = True
    Toolbar1.Buttons("registry").Enabled = True
    Toolbar1.Buttons("remotedesktop").Enabled = True
    Toolbar1.Buttons("taskmanager").Enabled = True
    Toolbar1.Buttons("controlpanel").Enabled = True
    StatusBar1.Panels(1).Picture = ImageList1.ListImages(1).Picture
    StatusBar1.Panels(1).Text = " Connected with '" & sckIdentifier.RemoteHostIP & "'   "
End Sub

Private Sub TileHoriz_Click()
    RemoteForm.Arrange vbTileHorizontal
End Sub

Private Sub TileVer_Click()
    RemoteForm.Arrange vbTileVertical
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "connect"
            Call ConnectCmd_Click
        Case "disconnect"
            Call DisconnectCmd_Click
        Case "chat"
            sckIdentifier.SendData "Chat#"
            ChatForm.Show
        Case "talkchat"
            TalkChatForm.Show
        Case "explorer"
            ExplorerForm.Show
        Case "registry"
            RegForm.Show
        Case "remotedesktop"
            RemoteDesktop.Show
        Case "taskmanager"
            TaskmanagerFrm.Show
        Case "controlpanel"
            ControlPanelForm.Show
    End Select
End Sub
