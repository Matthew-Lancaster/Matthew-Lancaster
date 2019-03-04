VERSION 5.00
Begin VB.Form frmCopyFile2ClipBoard 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy file to clipboard"
   ClientHeight    =   1164
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   9744
   Icon            =   "frmCopyFile2ClipBoard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1164
   ScaleWidth      =   9744
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelectFile 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Text            =   "C:\temp\TestFile.txt"
      Top             =   120
      Width           =   8655
   End
   Begin VB.CommandButton cmdCopyToClipboard 
      Caption         =   "Copy to ClipBoard"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   612
      Width           =   2415
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frmCopyFile2ClipBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ************************************************************************************************************
' * File input dialog
' ************************************************************************************************************

' ************************************************************************************************************
' * Copyright ©2016, Frank Donckers
' * All source code in this file is licensed under a modified BSD license.
' * This means you may use the code in your own projects IF you provide attribution.
' ************************************************************************************************************

Dim cCmDlg      As cCommonDialogCopyFileClipboard

'Copy file to clipboard
Public Sub cmdCopyToClipboard_Click()
    'frmCopyFile2ClipBoard.cmdCopyToClipboard
    sReturn = File_To_ClipBoard(txtFile.Text)
    If sReturn <> "" Then MsgBox sReturn, , "Copy file to ClipBoard"
End Sub

'Select file dialog
Private Sub cmdSelectFile_Click()
    Set cCmDlg = New cCommonDialogCopyFileClipboard
    'Open File Dialog
    With cCmDlg
        'Set some Dialog properties
        If txtFile.Text <> "" Then
            .InitialDir = GetPathForFile(txtFile.Text)
        Else
            .InitialDir = "c:\temp\"
        End If
        .Filter = "Text Files (*.txt;*.tmp;*.xml;*.csv)|*.txt;*.tmp;*.xml;*.csv"
        .ShowOpen Me.hWnd
        If .Canceled = False Then
            txtFile.Text = .FileName
        End If
    End With
    Set cCmDlg = Nothing

End Sub

