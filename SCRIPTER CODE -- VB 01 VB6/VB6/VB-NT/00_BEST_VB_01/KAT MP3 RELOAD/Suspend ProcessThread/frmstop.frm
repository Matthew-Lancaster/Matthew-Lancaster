VERSION 5.00
Begin VB.Form frmstop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suspend Process"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   Icon            =   "frmstop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdpomos 
      Caption         =   "Help"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtid 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdstop 
      Caption         =   "Suspend"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Process PID:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmstop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////'
' This code were explicitly developed for PSC(Planet Source Code) Users,
' as Open Source Project. This code are property of their author.
' The code is provided "as is" WITHOUT any warranty.

' You may use any of this code in you're own application(s).

' Code by Janevski Blagoj
' Please vote for me on planet-source-code.com
' e-mail: blagoj_bl@yahoo.com for comments,help or anything else.
' (c) XbXan 2006
'//////////////////////////////////////////////////////////////////////////////'


Option Explicit
Private Sub cmdpomos_Click()
MsgBox "Made by Janevski Blagoj - blagoj_bl@yahoo.com" _
& vbNewLine & "To find the process's PID press CTRL-ALT-DELETE,click on the menu View->Select Columns..." _
& vbNewLine & "Select PID (Process Identifier) and click OK", _
vbInformation + vbSystemModal, "Help and About"
End Sub

Private Sub cmdstop_Click()

If Trim(txtid.Text) = "" Then
    txtid.Text = ""
    txtid.SetFocus
End If

'Accept only numeric values
If Not IsNumeric(Trim(txtid.Text)) Then
    txtid.Text = ""
    Exit Sub
End If

If CLng(Trim(txtid.Text)) = GetCurrentProcessId() Then
    MsgBox "Error while suspending the process.", vbInformation, "Error"
    Exit Sub
End If

If cmdstop.Caption = "Suspend" Then
    cmdstop.Enabled = False
    txtid.Enabled = False
    If SuspendResumeProcess(CLng(Trim(txtid.Text)), True) Then
        MsgBox "The process is suspended.", vbInformation, "Suspended"
        cmdstop.Caption = "Resume"
    Else
        MsgBox "Error while suspending the process.", vbInformation, "Error"
        txtid.Text = ""
    End If
    txtid.Enabled = True
    cmdstop.Enabled = True
    txtid.SetFocus
Else 'if cmdstop.caption=Resume
    cmdstop.Enabled = False
    
    If SuspendResumeProcess(CLng(Trim(txtid.Text)), False) Then
        MsgBox "The process is resumed.", vbInformation, "Resumed"
        cmdstop.Caption = "Suspend"
    Else
        MsgBox "Error while resuming the process.", vbInformation, "Error"
    End If
    txtid.Enabled = True
    cmdstop.Enabled = True
    txtid.SetFocus
End If 'cmdstop.caption
End Sub

Private Sub Form_Activate()
txtid.SetFocus
End Sub

Private Sub Form_Load()
'if another instance of the program is already running
If App.PrevInstance Then
    MsgBox "Another instance of the program is already running.", vbInformation, "Error"
    End
End If

'Hide the program from tasklist
App.TaskVisible = False
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)
'if the user press Enter(return) key
If KeyAscii = vbKeyReturn Then cmdstop_Click
End Sub
