VERSION 5.00
Begin VB.Form frmSender 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send string message"
   ClientHeight    =   3210
   ClientLeft      =   180
   ClientTop       =   480
   ClientWidth     =   4680
   Icon            =   "frmSender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   510
      TabIndex        =   6
      Text            =   "String Receiver Caller_ID"
      Top             =   795
      Width           =   3645
   End
   Begin VB.TextBox txtDestCaption 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   510
      TabIndex        =   3
      Text            =   "String Receiver Caller_ID"
      Top             =   450
      Width           =   3645
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send it! "
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      TabIndex        =   2
      Top             =   2655
      Width           =   1215
   End
   Begin VB.TextBox txtMsg 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   495
      TabIndex        =   1
      Top             =   2100
      Width           =   3645
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Remember that the receiver app must be open."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   630
      TabIndex        =   5
      Top             =   1260
      Width           =   3435
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Destination window caption:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   4
      Top             =   255
      Width           =   2055
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Type the text you want to send here:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1035
      TabIndex        =   0
      Top             =   1785
      Width           =   2730
   End
End
Attribute VB_Name = "frmSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------'
' Using SendMessage to send a string between two '
'  VB-created applications                       '
'                                                '
' By Penagate (April 2005)                       '
'------------------------------------------------'

Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, _
     ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszTitle As String) _
    As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByRef lParam As Any) _
    As Long
    
Private Const WM_SETTEXT            As Long = &HC


Const GW_CHILD = 5


Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142




Public Sub cmdSend_Click()
    
    Dim hWndForm        As Long
    Dim hWndTextBox     As Long
    
    Dim ash$
    
    ' Find the destination window
    hWndForm = FindWindow(vbNullString, txtDestCaption.Text)
    
    If (hWndForm = 0) Then
      '  MsgBox "Couldn't find a destination window with that caption.", vbExclamation
        Exit Sub
    End If
    
    ' Find the text box as a child of the destination window
    hWndTextBox = FindWindowEx(hWndForm, 0&, "ThunderTextBox", vbNullString)
    
    If (hWndTextBox = 0) Then
        hWndTextBox = FindWindowEx(hWndForm, 0&, "ThunderTextBox", vbNullString)
    End If
    
    If (hWndTextBox = 0) Then
        hWndTextBox = FindWindowEx(hWndForm, 0&, "Thunder", vbNullString)
    End If
    
    If (hWndTextBox = 0) Then
        hWndTextBox = GetWindow(hWndForm, GW_CHILD)
    End If
    
    
    If hWndTextBox > 0 Then
        'ash$ = GetWindowTitle(hWndTextBox)
        'If ash$ = "" Then
        '    ash$ = GetWindowText(hWndTextBox)
        'End If
        
        'Text1.Text = ash$
    End If
    
    
    If (hWndTextBox = 0) Then
    '    MsgBox "Couldn't find target VB text box in the destination window.", vbExclamation
        Exit Sub
    End If
    
    ' If all's well, send the message.
    '
    ' Be aware that, in this example, until the message
    ' box is dismissed, this app will be hung.
    SendMessage hWndTextBox, WM_SETTEXT, 0&, ByVal CStr(txtMsg.Text)
    
End Sub
