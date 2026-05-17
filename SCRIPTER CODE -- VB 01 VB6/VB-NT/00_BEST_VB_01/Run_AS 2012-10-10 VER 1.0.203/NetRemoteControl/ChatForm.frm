VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ChatForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Chat"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "ChatForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   -45
      TabIndex        =   5
      Top             =   -120
      Width           =   7680
      Begin VB.Image Image1 
         Height          =   825
         Left            =   60
         Picture         =   "ChatForm.frx":030A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   7635
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4875
      Left            =   53
      TabIndex        =   0
      Top             =   840
      Width           =   7485
      Begin MSWinsockLib.Winsock sckChat 
         Left            =   3150
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         Height          =   945
         Left            =   150
         TabIndex        =   4
         Top             =   3810
         Width           =   7215
         Begin VB.CommandButton SendBttn 
            Caption         =   "Send"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5640
            TabIndex        =   2
            Top             =   300
            Width           =   1425
         End
         Begin VB.TextBox SendTxt 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   150
            TabIndex        =   1
            Top             =   330
            Width           =   5415
         End
      End
      Begin VB.TextBox GetTxt 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   3525
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   7215
      End
   End
End
Attribute VB_Name = "ChatForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const SC_CLOSE = &HF060&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_DISABLED = &H2&

Private Sub Form_Load()
    RemoveMenu GetSystemMenu(Me.hwnd, False), SC_CLOSE, MF_BYCOMMAND Or MF_DISABLED
    sckChat.LocalPort = 1002
    sckChat.Listen
End Sub

Private Sub sckChat_Close()
    If sckChat.State <> sckNotConnected Then
        sckChat.Close
        sckChat.Listen
        Unload Me
    End If
End Sub

Private Sub sckChat_ConnectionRequest(ByVal requestID As Long)
    If sckChat.State <> sckClosed Then
        sckChat.Close
        sckChat.Accept requestID
    End If
End Sub

Private Sub sckChat_DataArrival(ByVal bytesTotal As Long)
Dim Data As String, Position As Long
    sckChat.GetData Data
    GetTxt.Text = GetTxt.Text & Data & vbCrLf
    Position = Len(GetTxt.Text)
    GetTxt.SelStart = Position
    SendTxt.SetFocus
End Sub

Private Sub SendBttn_Click()
Dim str As String

    str = SendTxt.Text
    If Trim(str) <> "" Then
        sckChat.SendData str
        SendTxt.Text = ""
        SendTxt.SetFocus
    End If
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call SendBttn_Click
    End If
End Sub
