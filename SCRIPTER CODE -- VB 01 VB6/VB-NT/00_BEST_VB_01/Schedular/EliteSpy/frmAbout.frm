VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About EliteSpy+"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4020
      TabIndex        =   4
      Top             =   2940
      Width           =   1410
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   45
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   2625
      ScaleWidth      =   1575
      TabIndex        =   0
      Top             =   90
      Width           =   1635
   End
   Begin VB.Label Label4 
      Caption         =   "Please vote for me at www.Planet-Source-Code.com/vb"
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1860
      TabIndex        =   5
      Top             =   1920
      Width           =   3195
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   5580
      Y1              =   2865
      Y2              =   2865
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   5580
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright © 2001 Andrea Batina. All Rights Reserved."
      Height          =   465
      Left            =   1845
      TabIndex        =   3
      Top             =   765
      Width           =   2595
   End
   Begin VB.Label Label2 
      Caption         =   "See inside the windows."
      Height          =   240
      Left            =   1845
      TabIndex        =   2
      Top             =   360
      Width           =   2805
   End
   Begin VB.Label Label1 
      Caption         =   "EliteSpy+ v1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1845
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmAbout
'    Project    : EliteSpy
'
'    Description: About box
'
'    Author     : Andrea Batina
'    Modified   : 31/10/2001
'--------------------------------------------------------------------------------
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = "EliteSpy+ v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub
