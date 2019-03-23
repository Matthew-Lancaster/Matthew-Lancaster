VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form RemoteDesktop 
   Caption         =   "Remote Desktop"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   Icon            =   "RemoteDesktop.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   6330
      Top             =   2730
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Download in progress, Please Wait..."
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
      Height          =   765
      Left            =   0
      TabIndex        =   3
      Top             =   420
      Visible         =   0   'False
      Width           =   4065
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label ProgressLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0% Complete."
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
         Left            =   150
         TabIndex        =   5
         Top             =   480
         Width           =   1155
      End
   End
   Begin MSWinsockLib.Winsock sckDesktop 
      Left            =   3500
      Top             =   3450
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4260
      Top             =   2070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RemoteDesktop.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RemoteDesktop.frx":0894
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "getremotedesktop"
            Object.ToolTipText     =   "Get Remote Desktop"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cam"
            Object.ToolTipText     =   "Downloading Loop"
            ImageIndex      =   2
            Style           =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3675
      Left            =   450
      ScaleHeight     =   241
      ScaleMode       =   0  'User
      ScaleWidth      =   276
      TabIndex        =   0
      Top             =   990
      Width           =   4200
   End
   Begin VB.PictureBox SourcePicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   2190
      ScaleHeight     =   42
      ScaleMode       =   0  'User
      ScaleWidth      =   44
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "RemoteDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileNumber As Integer
Dim DeskX As Single, DeskY As Single

Private Sub Form_Load()
    sckDesktop.RemoteHost = RemoteForm.sckIdentifier.RemoteHost
    sckDesktop.RemotePort = 8003
    sckDesktop.Connect
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Picture1.Top = Toolbar1.Height + Screen.TwipsPerPixelY * 1
    Picture1.Left = Screen.TwipsPerPixelY * 1
    Picture1.Width = (Me.ScaleWidth - Screen.TwipsPerPixelX * 5)
    Picture1.Height = (Me.ScaleHeight - Screen.TwipsPerPixelY * 33)
    Picture1.Scale (0, 0)-(DeskX + 1, DeskY + 1)
    Picture1.Refresh
    Picture1.PaintPicture SourcePicture.Picture, 0, 0, Picture1.Width, Picture1.Height, 0, 0, SourcePicture.Width, SourcePicture.Height, vbSrcCopy
    Picture1.Refresh
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim txtData As String, L As Single, T As Single
    Select Case Button
        Case vbLeftButton
'            If Y > Picture1.ScaleHeight / 2 Then
'                txtData = "LeftButtonDown#" & Format(X, "000") & "#" & Format(Y - 32, "000") & "#"
'            Else
'                txtData = "LeftButtonDown#" & Format(X, "000") & "#" & Format(Y + 50, "000") & "#"
'            End If
            'L = Me.Top - X
            'T = Me.Left - Y
            'txtData = "LeftButtonDown#" & Format(X, "000") & "#" & Format(Y, "000") & "#"
            txtData = "LeftButtonDown#" & CInt(X) & "#" & CInt(Y) & "#"
            sckDesktop.SendData txtData
        Case vbMiddleButton
        Case vbRightButton
    End Select
End Sub

Private Sub sckDesktop_Close()
    Unload Me
End Sub

Private Sub sckDesktop_Connect()
    sckDesktop.SendData "GetRECT#"
End Sub

Private Sub sckDesktop_DataArrival(ByVal bytesTotal As Long)
Dim txtData As String, Command As String
Dim Position As Integer, FileSize As Long

On Error Resume Next
    sckDesktop.getData txtData
    
    Position = InStr(1, txtData, "#")
    Command = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Select Case Command
        Case "GetRECT"
        Dim Width As Single, Height As Single
            Position = InStr(1, txtData, "#")
            Width = Val(Left(txtData, Position - 1))
            txtData = Mid(txtData, Position + 1)
            Position = InStr(1, txtData, "#")
            Height = Val(Left(txtData, Position - 1))
            DeskX = Width: DeskY = Height
            Picture1.Scale (0, 0)-(Width, Height)
            Picture1.Refresh
        Case "FileSize"
            Toolbar1.Buttons("getremotedesktop").Enabled = False
            FileSize = CLng(txtData) + 512
            ProgressBar1.Max = FileSize
            ProgressBar1.Value = 0
            Frame1.Visible = True
        Case "DownloadFile"
            Put #FileNumber, , txtData
            ProgressBar1.Value = ProgressBar1.Value + 512
            ProgressLabel.Caption = Format(ProgressBar1.Value / ProgressBar1.Max, "###% Complete.")
        Case "DownloadComplete"
            Close (FileNumber)
            SourcePicture.Picture = LoadPicture("C:\desk.bmp")
            ProgressBar1.Value = ProgressBar1.Max
            Frame1.Visible = False
            Toolbar1.Buttons("getremotedesktop").Enabled = True
    End Select
    
End Sub

Private Sub SourcePicture_Change()
    Picture1.PaintPicture SourcePicture.Picture, 0, 0, Picture1.Width, Picture1.Height, 0, 0, SourcePicture.Width, SourcePicture.Height, vbSrcCopy
End Sub

Private Sub Timer1_Timer()
    sckDesktop.SendData "GetDesktop#"
    FileNumber = FreeFile()
    Open "C:\desk.bmp" For Binary As #FileNumber
    DoEvents
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim txtData As String
Select Case Button.Key
    Case "getremotedesktop"
        sckDesktop.SendData "GetDesktop#"
        FileNumber = FreeFile()
        Open "C:\desk.bmp" For Binary As #FileNumber
    Case "cam"
        If Button.Value = tbrUnpressed Then
            Timer1.Enabled = False
        Else
            Timer1.Enabled = True
        End If
End Select
End Sub
