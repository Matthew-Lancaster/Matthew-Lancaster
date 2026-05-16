VERSION 5.00
Begin VB.Form frmCGWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Generation Wizard"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCGWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPage0 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   180
      Width           =   4995
      Begin VB.OptionButton optFindWindow 
         Caption         =   "Code to find window"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   11
         Top             =   540
         Width           =   3015
      End
      Begin VB.OptionButton optFindWindow 
         Caption         =   "Code to close window"
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   10
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton optFindWindow 
         Caption         =   "Code to minimize window"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   9
         Top             =   1140
         Width           =   3015
      End
      Begin VB.OptionButton optFindWindow 
         Caption         =   "Code to mazimize window"
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   8
         Top             =   1440
         Width           =   3015
      End
      Begin VB.OptionButton optFindWindow 
         Caption         =   "Code to change caption of window"
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   7
         Top             =   1740
         Width           =   3015
      End
      Begin VB.OptionButton optFindWindow 
         Caption         =   "Code to activate window"
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   6
         Top             =   2040
         Width           =   3015
      End
      Begin VB.OptionButton optFindWindow 
         Caption         =   "Code to enable or disable window"
         Height          =   195
         Index           =   6
         Left            =   840
         TabIndex        =   5
         Top             =   2340
         Width           =   3015
      End
      Begin VB.OptionButton optFindWindow 
         Caption         =   "Code to hide or show window"
         Height          =   195
         Index           =   7
         Left            =   840
         TabIndex        =   4
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Select type of code you want to generate:"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   180
         Width           =   3195
      End
      Begin VB.Image Image1 
         Height          =   420
         Left            =   120
         Picture         =   "frmCGWizard.frx":058A
         Stretch         =   -1  'True
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.PictureBox picPage2 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   4995
      TabIndex        =   17
      Top             =   180
      Visible         =   0   'False
      Width           =   4995
      Begin VB.Label Label5 
         Caption         =   "You have successfuly completed the wizard. Pres Finish button to see the generated code!"
         Height          =   795
         Left            =   1200
         TabIndex        =   18
         Top             =   1020
         Width           =   2715
      End
   End
   Begin VB.PictureBox picPage1 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      ScaleHeight     =   3015
      ScaleWidth      =   4995
      TabIndex        =   13
      Top             =   180
      Visible         =   0   'False
      Width           =   4995
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   720
         TabIndex        =   15
         Top             =   1020
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "(Example: NotePad)"
         Height          =   195
         Left            =   780
         TabIndex        =   16
         Top             =   780
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Enter part of the window title you want to change:"
         Height          =   435
         Left            =   720
         TabIndex        =   14
         Top             =   180
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   3360
      Width           =   1155
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<- Back"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2820
      TabIndex        =   1
      Top             =   3360
      Width           =   1155
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next ->"
      Height          =   315
      Left            =   4020
      TabIndex        =   0
      Top             =   3360
      Width           =   1155
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   60
      X2              =   5160
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   60
      X2              =   5160
      Y1              =   3240
      Y2              =   3240
   End
End
Attribute VB_Name = "frmCGWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmCGWizard
'    Project    : EliteSpy
'
'    Description: Generate code wizard.
'
'    Author     : Andrea Batina
'    Modified   : 31/10/2001
'--------------------------------------------------------------------------------
Option Explicit

' Current wizard page
Private m_nPage As Integer
Private m_nCodeType As Integer

Private Sub cmdBack_Click()
    ' Decrease value by one
    m_nPage = m_nPage - 1
    ' Move to previous page
    SetPage
End Sub

Private Sub cmdCancel_Click()
    ' Close wizard
    Unload Me
End Sub

Private Sub cmdNext_Click()
    ' If we finished the wizard
    If cmdNext.Caption = "Finish" Then
        ' If user did not enter window title then
        If Not Len(txtTitle.Text) > 0 Then
            ' Show the error and exit
            MsgBox "You must enter part of the window title!", vbInformation, "EliteSpy+"
            Exit Sub
        End If
        ' Generate code
        GenerateCode m_nCodeType, txtTitle.Text
        ' Close wizard
        Unload Me
    End If
    
    ' Increase value by one
    m_nPage = m_nPage + 1
    ' Move to next page
    SetPage
End Sub

Private Sub SetPage()
    ' Setup pages
    Select Case m_nPage
        Case 0
            cmdBack.Enabled = False
            picPage0.Visible = True
            picPage1.Visible = False
            picPage2.Visible = False
            cmdNext.Caption = "Next ->"
            
        Case 1
            cmdBack.Enabled = True
            picPage0.Visible = False
            picPage1.Visible = True
            picPage2.Visible = False
            cmdNext.Caption = "Next ->"
        Case 2
            cmdBack.Enabled = True
            picPage0.Visible = False
            picPage1.Visible = False
            picPage2.Visible = True
            cmdNext.Caption = "Finish"
    End Select
End Sub

Private Sub Form_Load()
    ' Select first option box
    optFindWindow(0).Value = True
    m_nPage = 0
End Sub
Private Sub Form_Terminate()
    m_nPage = 0
End Sub

Private Sub optFindWindow_Click(Index As Integer)
    m_nCodeType = Index
End Sub
