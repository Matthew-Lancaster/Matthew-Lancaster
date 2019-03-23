VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Draw Regular Filled Polygon"
   ClientHeight    =   6510
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   8265
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3945
      ScaleWidth      =   8025
      TabIndex        =   1
      Top             =   360
      Width           =   8055
      Begin VB.TextBox txtSelCol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6720
         TabIndex        =   22
         Text            =   "16711680"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patterns"
         Height          =   2175
         Left            =   6720
         TabIndex        =   23
         Top             =   240
         Width           =   1215
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   14
            Left            =   840
            MouseIcon       =   "Form1.frx":0000
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":0CCA
            Stretch         =   -1  'True
            Top             =   1800
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   13
            Left            =   480
            MouseIcon       =   "Form1.frx":0DCC
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":1A96
            Stretch         =   -1  'True
            Top             =   1800
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   12
            Left            =   120
            MouseIcon       =   "Form1.frx":1B98
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":2862
            Stretch         =   -1  'True
            Top             =   1800
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   11
            Left            =   840
            MouseIcon       =   "Form1.frx":2964
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":362E
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   10
            Left            =   480
            MouseIcon       =   "Form1.frx":3730
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":43FA
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   9
            Left            =   120
            MouseIcon       =   "Form1.frx":44FC
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":51C6
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   8
            Left            =   840
            MouseIcon       =   "Form1.frx":52C8
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":5F92
            Stretch         =   -1  'True
            Top             =   360
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   7
            Left            =   480
            MouseIcon       =   "Form1.frx":6094
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":6D5E
            Stretch         =   -1  'True
            Top             =   360
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   6
            Left            =   120
            MouseIcon       =   "Form1.frx":6E60
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":7B2A
            Stretch         =   -1  'True
            Top             =   360
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   4
            Left            =   840
            MouseIcon       =   "Form1.frx":7C2C
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":88F6
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   0
            Left            =   840
            MouseIcon       =   "Form1.frx":8A18
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":96E2
            Stretch         =   -1  'True
            Top             =   720
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   1
            Left            =   120
            MouseIcon       =   "Form1.frx":9804
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":A4CE
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   2
            Left            =   480
            MouseIcon       =   "Form1.frx":A5F0
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":B2BA
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   3
            Left            =   120
            MouseIcon       =   "Form1.frx":B3DC
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":C0A6
            Stretch         =   -1  'True
            Top             =   720
            Width           =   255
         End
         Begin VB.Image imgPattern 
            Height          =   255
            Index           =   5
            Left            =   480
            MouseIcon       =   "Form1.frx":C1C8
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":CE92
            Stretch         =   -1  'True
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Selected"
         Height          =   1095
         Left            =   6720
         TabIndex        =   24
         Top             =   2400
         Width           =   1215
         Begin VB.Image imgSelected 
            Height          =   615
            Left            =   240
            Picture         =   "Form1.frx":CFB4
            Stretch         =   -1  'True
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.PictureBox picCol 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1800
         Left            =   120
         Picture         =   "Form1.frx":D0B6
         ScaleHeight     =   1800
         ScaleWidth      =   3345
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   3345
      End
      Begin VB.PictureBox PicSel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   120
         Picture         =   "Form1.frx":20BF8
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   19
         Top             =   120
         Width           =   390
      End
      Begin VB.Timer tmrDemo 
         Interval        =   100
         Left            =   120
         Top             =   120
      End
   End
   Begin VB.CheckBox chkIs 
      Caption         =   "Is Centre"
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox chkDemo 
      Caption         =   "Demo Mode"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dimensions"
      Height          =   1935
      Left            =   4680
      TabIndex        =   2
      Top             =   4440
      Width           =   3495
      Begin VB.TextBox txtSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         TabIndex        =   16
         Text            =   "10"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtY 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         TabIndex        =   12
         Text            =   "2000"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtX 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Text            =   "4000"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtSideLen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         TabIndex        =   8
         Text            =   "2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtSides 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Text            =   "6"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtRot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         TabIndex        =   3
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lbds 
         BackStyle       =   0  'Transparent
         Caption         =   "Rotation  Speed"
         Height          =   480
         Left            =   2640
         TabIndex        =   17
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y Dist"
         Height          =   240
         Left            =   2640
         TabIndex        =   13
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X Dist"
         Height          =   240
         Left            =   840
         TabIndex        =   11
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "cm           Side Length "
         Height          =   600
         Left            =   840
         TabIndex        =   9
         Top             =   1320
         Width           =   1110
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sides"
         Height          =   240
         Left            =   840
         TabIndex        =   7
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rotation"
         Height          =   240
         Left            =   2640
         TabIndex        =   4
         Top             =   960
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdControll 
      Caption         =   "< Manual >  "
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fill Style"
      Height          =   1935
      Left            =   1560
      TabIndex        =   20
      Top             =   4440
      Width           =   3015
      Begin VB.ListBox lstDrawStyle 
         Height          =   1500
         ItemData        =   "Form1.frx":212FA
         Left            =   120
         List            =   "Form1.frx":21319
         TabIndex        =   21
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R <<  |  >> T"
      Height          =   240
      Left            =   240
      TabIndex        =   14
      Top             =   6120
      Width           =   1020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
'The form code is only to show the function calling.
'It looks a littlelong but there is 'Nothing' in the form code.
'Please don't get trouble with them
'The 'StdPicture' is only used in the FillStyle 'FillBitmap'
'We can only use 8 * 8 Bitmap as pattern
'===============================================================
Option Explicit
Dim FillPattern As StdPicture

Private Sub chkDemo_Click()
    tmrDemo.Enabled = chkDemo.Value
End Sub

Private Sub cmdControll_KeyPress(KeyAscii As Integer)
If tmrDemo.Enabled = True Then MsgBox "Please disable 'Demo Mode'", vbInformation
    If KeyAscii = vbKeyR Then
        txtRot = txtRot - 20: tmrDemo_Timer
    ElseIf KeyAscii = vbKeyT Then
        tmrDemo_Timer
    End If
End Sub

Private Sub Form_Load()
    lstDrawStyle.Selected(8) = True
    PicView.FillColor = Val(txtSelCol)
    Set FillPattern = imgPattern(12).Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Is it Satisfactory?", vbQuestion + vbYesNo, "Please tell Me") = vbYes Then
        MsgBox "(  PLEASE 'RATE' THIS CODE  ).I want to know how do you rate this code.The site address will be copied to your clipboard", vbInformation, "ThankYou"
    Else
        MsgBox "( PLEASE GIVE FEEDBACK ) to improve this code.The site address will be copied to your clipboard", vbInformation, "Please Give FeedBack"
    End If
    Clipboard.SetText ("http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57807&lngWId=1")
End Sub

Private Sub imgPattern_Click(Index As Integer)
    Set FillPattern = imgPattern(Index).Picture
    imgSelected.Picture = imgPattern(Index).Picture
End Sub

Private Sub picCol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicView.FillColor = picCol.Point(X, Y)
    txtSelCol = picCol.Point(X, Y)
    txtSelCol.BackColor = picCol.Point(X, Y)
    picCol.Visible = False: tmrDemo.Enabled = True
End Sub

Private Sub PicSel_Click()
    picCol.Visible = True: tmrDemo.Enabled = False
End Sub

Private Sub tmrDemo_Timer()
Dim Cnt  As Boolean

    If chkIs = 1 Then Cnt = True    'Checks whether the 'IsCentre' is True
    txtRot = Val(txtRot) + Val(txtSpeed)
    
    'I can't add an error message when 'DrwPlgn=False' in a 'timer'
    PicView.Cls
    DrawPolygon PicView, Val(txtX), Val(txtY), Cnt, Val(txtSides), Val(txtSideLen) * 567, lstDrawStyle.ListIndex, Val(txtRot), FillPattern
End Sub
