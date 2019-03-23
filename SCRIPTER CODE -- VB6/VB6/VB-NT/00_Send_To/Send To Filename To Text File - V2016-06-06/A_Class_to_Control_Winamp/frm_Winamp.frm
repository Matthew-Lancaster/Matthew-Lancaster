VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_Winamp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send message to Winamp"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "frm_Winamp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar stbInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10530
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraWinamp 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.PictureBox pichook 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   4320
         Picture         =   "frm_Winamp.frx":08CA
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   480
         Width           =   480
      End
      Begin VB.CommandButton cmdSendMessage 
         Caption         =   "Send Message"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox cmbCommand 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Developed By Federico Bridger (Rosario, Arg.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3960
         TabIndex        =   5
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label lblCommand 
         Caption         =   "Select Command:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frm_Winamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub m_FillCombo()
  
  'Fill the Combo with some of the Messages that we can
  'send to Winamp.
  'If you like you can add more Messages from the Enumeration
  '---------------------------------------------------------

  With cmbCommand
    
    .AddItem "Play"
    .ItemData(.NewIndex) = Enum_WA_Command.WA_Play
    
    .AddItem "Pause / Unpause"
    .ItemData(.NewIndex) = Enum_WA_Command.WA_Pause_Unpause
    
    .AddItem "Next Track"
    .ItemData(.NewIndex) = Enum_WA_Command.WA_Next_track
    
    .AddItem "Previous Track"
    .ItemData(.NewIndex) = Enum_WA_Command.WA_Previous_track
    
    .AddItem "Raise Volume by 1"
    .ItemData(.NewIndex) = Enum_WA_Command.WA_Raise_volume_by_1
    
    .AddItem "Lower Volume By 1"
    .ItemData(.NewIndex) = Enum_WA_Command.WA_Lower_volume_by_1
    
    .AddItem "Fast Forward 5 seconds"
    .ItemData(.NewIndex) = Enum_WA_Command.WA_Fast_forward_5_seconds
    
    .AddItem "Fast Rewind 5 seconds"
    .ItemData(.NewIndex) = Enum_WA_Command.WA_Fast_rewind_5_seconds
    
    .AddItem "Close Winamp"
    .ItemData(.NewIndex) = Enum_WA_Command.WA_Close_Winamp
       
    .ListIndex = 0
    
  End With
  
End Sub

Private Sub cmdSendMessage_Click()
  
  Dim l_Winamp As cls_Winamp
  
  'Create the instance of the Winamp Class
  Set l_Winamp = New cls_Winamp
  
  If l_Winamp.IsWinampRunning Then
    
    'Send Message to the Winamp.exe instance
    l_Winamp.Execute cmbCommand.ItemData(cmbCommand.ListIndex)
    
    stbInfo.Panels(1).Text = "Winamp Instance Found... Message Sent!"

  Else

    stbInfo.Panels(1).Text = "Winamp Instance Not Found, you have to open Winamp.exe"

  End If
  
End Sub

Private Sub Form_Load()

  m_FillCombo
  
  stbInfo.Panels(1).Text = "Developed by Federico D. Bridger (Rosario, Argentina)"

End Sub


