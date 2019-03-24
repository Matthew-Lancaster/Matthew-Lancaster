VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fill Non-Regular Polygons - Timer Enabled"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Select Fill Mode"
      Height          =   4815
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Width           =   2655
      Begin VB.PictureBox picSource 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   1920
         Picture         =   "frmMain.frx":0000
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   11
         Top             =   4380
         Width           =   135
      End
      Begin VB.OptionButton optFillMode 
         Caption         =   "Bitmap Pattern"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   10
         Top             =   4320
         Width           =   2000
      End
      Begin VB.OptionButton optFillMode 
         Caption         =   "Solid"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   9
         Top             =   3840
         Value           =   -1  'True
         Width           =   2000
      End
      Begin VB.OptionButton optFillMode 
         Caption         =   "DiagonalCross"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   8
         Top             =   3360
         Width           =   2000
      End
      Begin VB.OptionButton optFillMode 
         Caption         =   "Cross"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   2000
      End
      Begin VB.OptionButton optFillMode 
         Caption         =   "DownwardDiogonal"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   2235
      End
      Begin VB.OptionButton optFillMode 
         Caption         =   "UpwardDiagonal"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   2000
      End
      Begin VB.OptionButton optFillMode 
         Caption         =   "VerticalLine"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   2000
      End
      Begin VB.OptionButton optFillMode 
         Caption         =   "HorizontalLine"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2000
      End
      Begin VB.OptionButton optFillMode 
         Caption         =   "Transparent"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2000
      End
   End
   Begin VB.PictureBox PicView2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   240
      ScaleHeight     =   4635
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   360
      Width           =   5175
      Begin VB.Timer tmrUpdate 
         Interval        =   100
         Left            =   120
         Top             =   120
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Drawing random polygons
Private Sub tmrUpdate_Timer()
Dim X As Integer
Dim fMode As Integer
    'Preparing the picturebox
    PicView.Cls
    'Getting the selected FillMode
    For X = 0 To optFillMode.Count - 1
        If optFillMode(X).Value = True Then fMode = X: Exit For
    Next X

'Randomly defining the vertieces
Dim Pnt() As POINTAPI
X = Rnd * 10 + 1
ReDim Pnt(1 To X) As POINTAPI
    For X = 1 To UBound(Pnt)
        Pnt(X).X = Rnd * 300 + 1
        Pnt(X).Y = Rnd * 300 + 1
    Next X
    'Filling the defined region / polygon
    FillRegion PicView.hDC, Pnt, Rnd * 999999, fMode, picSource.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Integer
    If MsgBox("Is it Satisfactory?", vbQuestion + vbYesNo, "Please tell Me") = vbYes Then
        X = MsgBox("(  PLEASE 'RATE' THIS CODE  ).Click 'Ok' to copy the site address  to your clipboard", vbInformation + vbOKCancel, "ThankYou")
    Else
        X = MsgBox("( PLEASE GIVE FEEDBACK ) to improve this code.Click 'Ok' to copy the site address  to your clipboard", vbInformation + vbOKCancel, "Please Give FeedBack")
    End If
    If X = vbOK Then Clipboard.SetText ("http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=58667&lngWId=1")
End Sub
