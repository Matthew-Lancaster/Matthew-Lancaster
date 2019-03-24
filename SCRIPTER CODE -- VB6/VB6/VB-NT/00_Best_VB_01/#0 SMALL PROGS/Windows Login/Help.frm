VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "E-Winlog"
   ClientHeight    =   3495
   ClientLeft      =   3765
   ClientTop       =   2010
   ClientWidth     =   6030
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMain 
      Caption         =   "App name"
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5775
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Open the current logfile in Notepad"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1560
         Width           =   5175
      End
      Begin VB.Label Label7 
         Caption         =   "/v"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Info about the location of the current logfile"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label Label5 
         Caption         =   "/l"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "/r"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "/q"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "/i"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "/?"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblI 
         BackStyle       =   0  'Transparent
         Caption         =   "Get information about existing running instance of program or not"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1080
         Width           =   5175
      End
      Begin VB.Label lblQ 
         BackStyle       =   0  'Transparent
         Caption         =   "Kill already running instance of program (this takes 2 seconds)"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label lblHelp 
         BackStyle       =   0  'Transparent
         Caption         =   "Show this dialog"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label lblL 
         BackStyle       =   0  'Transparent
         Caption         =   "Redirect the logfile to somwhere else than the system directory"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   5175
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "Help.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblCopyRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FileName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   75
      Width           =   795
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   6100
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    ' Make caption on the form
    fraMain.Caption = App.ProductName & " v." & App.Major & "." & App.Minor & "." & App.Revision & " commandline help"
    lblCopyRight.Caption = "E-WinLog is Freeware" & Chr(13) & "Copyright © 2000 Trond Sørensen" & Chr(13) & "All Rights Reserved"
End Sub

Private Sub OKButton_Click()
    ' Quit
    End
End Sub
