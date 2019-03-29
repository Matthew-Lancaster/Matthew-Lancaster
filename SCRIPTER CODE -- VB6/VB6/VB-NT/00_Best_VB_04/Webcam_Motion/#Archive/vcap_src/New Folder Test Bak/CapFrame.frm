VERSION 5.00
Begin VB.Form frmCapFrame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capture Frames"
   ClientHeight    =   1740
   ClientLeft      =   2460
   ClientTop       =   3900
   ClientWidth     =   3555
   Icon            =   "CapFrame.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2085
      TabIndex        =   4
      Top             =   1215
      Width           =   960
   End
   Begin VB.CommandButton cmdCapture 
      Caption         =   "C&apture"
      Height          =   360
      Left            =   645
      TabIndex        =   3
      Top             =   1215
      Width           =   960
   End
   Begin VB.Label lblFrames 
      Alignment       =   2  'Center
      Caption         =   "0 Frames"
      Height          =   225
      Left            =   1065
      TabIndex        =   2
      Top             =   780
      Width           =   1560
   End
   Begin VB.Label lblCapFile 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   15
      TabIndex        =   1
      Top             =   435
      Width           =   3480
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Select Capture to capture an image to"
      Height          =   225
      Left            =   405
      TabIndex        =   0
      Top             =   135
      Width           =   2820
   End
End
Attribute VB_Name = "frmCapFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCapture_Click()
    If capCaptureSingleFrame(frmMain.capwnd) Then
        lblFrames.Caption = Val(lblFrames.Caption) + 1 & " Frames"
        cmdCancel.Caption = "Close"
    Else
        MsgBox "Could not capture frame", App.Title, vbInformation
    End If
End Sub

Private Sub Form_Load()
lblCapFile.Caption = capFileGetCaptureFile(frmMain.capwnd)
If lblCapFile.Caption = "" Then
    lblCapFile.Caption = "<error: no cap file>"
End If
Call capCaptureSingleFrameOpen(frmMain.capwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call capCaptureSingleFrameClose(frmMain.capwnd)
End Sub
