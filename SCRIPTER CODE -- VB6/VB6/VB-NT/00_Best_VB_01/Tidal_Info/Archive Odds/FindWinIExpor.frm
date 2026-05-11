VERSION 5.00
Begin VB.Form FindWinIExplor 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   570
      Left            =   4080
      TabIndex        =   7
      Top             =   825
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   4035
      TabIndex        =   6
      Top             =   180
      Width           =   975
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
      Left            =   0
      TabIndex        =   2
      Top             =   1845
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
      Left            =   1260
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
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
      Left            =   15
      TabIndex        =   0
      Text            =   "String Receiver CID_RUN_ME"
      Top             =   180
      Width           =   3645
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
      Left            =   540
      TabIndex        =   5
      Top             =   1530
      Width           =   2730
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
      Left            =   825
      TabIndex        =   4
      Top             =   0
      Width           =   2055
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
      Left            =   135
      TabIndex        =   3
      Top             =   1005
      Width           =   3435
   End
End
Attribute VB_Name = "FindWinIExplor"
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
Private Const GW_CHILD = 5




'Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142




Public Sub cmdSend_Click()
    
    Dim hWndForm        As Long
    Dim hWndTextBox     As Long
    
    Dim ash$
    
    'FindWinIExplor.txtDestCaption.Text=
    
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
        ash$ = GetWindowTitle(hWndTextBox)
    End If
    
    If (hWndTextBox = 0) Or Trim(ash$ = "") Then
        hWndTextBox = FindWindowEx(hWndForm, 0&, "Thunder", vbNullString)
        ash$ = GetWindowTitle(hWndTextBox)
    End If
    
    If (hWndTextBox = 0) Or Trim(ash$ = "") Then
        hWndTextBox = GetWindow(hWndForm, GW_CHILD)
        ash$ = GetWindowTitle(hWndTextBox)
    End If
    
    
    If hWndTextBox > 0 Then
        ash$ = GetWindowTitle(hWndTextBox)
        If ash$ = "" Then
            ash$ = GetWindowTitle(hWndTextBox)
        End If
        
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



      Public Sub Command1_Click()
        Stop 'chkin if used
         
         Dim lRet As Long, lParam As Long
         Dim lhWnd As Long

         'lhWnd = Me.hwnd  ' Find the Form's Child Windows
         ' Comment the line above and uncomment the line below to
         ' enumerate Windows for the DeskTop rather than for the Form
         lhWnd = GetDesktopWindow()  ' Find the Desktop's Child Windows
         ' enumerate the list
         'lRet = EnumChildWindows(lhWnd, AddressOf EnumChildProc, lParam)
      
         'enumerate the list
         
        Open "D:\# MY DOCS\# 01 My Documents\LogLog.txt" For Output As #1
       
         lRet = EnumWindows(AddressOf EnumWinProc, lParam)
         
        Close #1
         ' How many Windows did we find?
'         Debug.Print TopCount; " Total Top level Windows"
'         Debug.Print ChildCount; " Total Child Windows"
'         Debug.Print ThreadCount; " Total Thread Windows"
'         Debug.Print "For a grand total of "; TopCount + _
'         ChildCount + ThreadCount; " Windows!"
      
      
      End Sub

      Public Sub Command2_Click()
         Dim lRet As Long
         Dim lParam As Long

         'enumerate the list
         lRet = EnumWindows(AddressOf EnumWinProc, lParam)
         ' How many Windows did we find?
'         Debug.Print TopCount; " Total Top level Windows"
'         Debug.Print ChildCount; " Total Child Windows"
'         Debug.Print ThreadCount; " Total Thread Windows"
'         Debug.Print "For a grand total of "; TopCount + _
'         ChildCount + ThreadCount; " Windows!"
      End Sub


Private Sub Form_Load()

End Sub
