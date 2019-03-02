VERSION 5.00
Begin VB.Form frmReceiver 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "String Receiver"
   ClientHeight    =   1035
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReceiver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtTarget 
      Height          =   330
      Left            =   135
      TabIndex        =   1
      Top             =   465
      Width           =   4140
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Destination text box (for best effect, this can be hidden):"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   4155
   End
End
Attribute VB_Name = "frmReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------'
' Using SendMessage to send a string between two '
'  VB-created applications                       '
'                                                '
' This is the receiver form.                     '

' By Penagate (April 2005)                       '
'------------------------------------------------'

Option Explicit
'


Private Sub Form_Load()
    ' Pick a window caption to use.
    Dim strCaption      As String
    
    'strCaption = InputBox("Enter the caption of the receiver window:", "String Receiver", "String Receiver")
    strCaption = "String Receiver CID_RUN_ME"
    
    If (LenB(strCaption) = 0) Then
        strCaption = "String Receiver"
    End If
    
    Me.Caption = strCaption
End Sub


' Received message handler

Private Sub txtTarget_Change()
    If (LenB(txtTarget.Text)) Then
        
        ' Be aware that, in this example, until the message
        ' box is dismissed, the sending app will be hung. So
        ' the message box is only used to demonstrate that the
        ' message is sent instantaneously. In a real-life
        ' example we would simply handle the string normally.
        
    '   MsgBox txtTarget.Text, vbExclamation, "Received message"
    globalhot$ = txtTarget.Text
    txtTarget.Text = ""
    Call CID_Run_Me.mystery
    End If
End Sub
