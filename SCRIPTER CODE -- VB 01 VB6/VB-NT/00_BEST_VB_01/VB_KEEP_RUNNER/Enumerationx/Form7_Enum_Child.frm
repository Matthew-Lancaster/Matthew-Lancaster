VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form7_Enum_Child 
   Caption         =   "Enumeration Goes On"
   ClientHeight    =   5340
   ClientLeft      =   1656
   ClientTop       =   2208
   ClientWidth     =   8472
   LinkTopic       =   "Form7_Enum_Child"
   ScaleHeight     =   5340
   ScaleWidth      =   8472
   Begin VB.CheckBox Check1 
      Caption         =   "&Windows with captions"
      Height          =   195
      Left            =   1920
      TabIndex        =   8
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Patch'em"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      ToolTipText     =   "Change text of any window :)"
      Top             =   4800
      Width           =   1215
   End
   Begin MSComctlLib.ListView View2 
      Height          =   4215
      Left            =   4080
      TabIndex        =   5
      ToolTipText     =   "Child windows"
      Top             =   480
      Width           =   3855
      _ExtentX        =   6795
      _ExtentY        =   7430
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView View1 
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Parent windows"
      Top             =   480
      Width           =   3855
      _ExtentX        =   6795
      _ExtentY        =   7430
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&numThem"
      Height          =   375
      Left            =   6708
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Leave"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "http://go.to/abubakar"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Left or Right click the Handles to see what happens"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Enumerating to the Max"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Show 
         Caption         =   "&Show Window using ShowWindow API"
      End
      Begin VB.Menu Show_BWTT 
         Caption         =   "Show &Window using BringWindowToTop API"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu Max 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu Min 
         Caption         =   "Mi&nimize"
      End
      Begin VB.Menu Restore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu Hide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu Close 
         Caption         =   "&Close this Window"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu SpyMenu 
         Caption         =   "Spy the &Menus"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "menu2"
      Visible         =   0   'False
      Begin VB.Menu BnClick 
         Caption         =   "&Click"
      End
   End
End
Attribute VB_Name = "Form7_Enum_Child"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------
'           Author: Muhammad Abubakar
'           http://go.to/abubakar
'           <joehacker@yahoo.com>
'------------------------------------
'------------------------------------
' [ Wednesday 20:39:50 Pm_14 November 2018 ]
'----
'FreeVBCode code snippet: Enumerate All Open Windows (Parent and Children)
'http://www.freevbcode.com/ShowCode.asp?ID=701
'----
'------------------------------------

Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private ClassResize As New CResize

'API to open the browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
    String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const WM_CLOSE = &H10


Private Sub BnClick_Click()
    SendMessage Val(View2.SelectedItem), BM_CLICK, 0, 0
    
End Sub

Private Sub Close_Click()
    'close window code goes here:
    Dim LHWND As Long
    
    On Error Resume Next
    LHWND = Val(View1.SelectedItem)
    SendMessage LHWND, WM_CLOSE, 0, 0

End Sub

Private Sub Command1_Click()
    'Free the memory occupied by the Object
    Set ClassResize = Nothing
    Unload Me

End Sub
Private Sub Command2_Click()
    Command2.Caption = "&Refresh"
    View1.ListItems.Clear
    View2.ListItems.Clear
    View1.GridLines = True
    Dim myLong As Long
    VCount = 1
    myLong = EnumWindows(AddressOf WndEnumProc, View1)

End Sub

Private Sub Command3_Click()
    Form8_Enum_Child.Show vbModal
    
End Sub

Private Sub Form_Load()
    With ClassResize
        .hParam = Form7_Enum_Child.height
        .wParam = Form7_Enum_Child.width
        .Map Command1, RS_Top_Left
        .Map Command2, RS_Top_Left
        .Map Command3, RS_Top_Left
        .Map Label2, RS_TopOnly
        .Map Label3, RS_LeftOnly
        .Map View1, RS_HeightOnly
        .Map View2, RS_HeightOnly
        .Map Check1, RS_Top_Left
    End With
    Form7_Enum_Child.width = 11000
    
    Me.Left = (Screen.width - Me.width) / 2
    Me.Top = (Screen.height - Me.height) / 2
    
    View1.View = lvwReport
    With View1.ColumnHeaders
        .Add , , "Handle", 1000
        .Add , , "Class Name", 1500
        .Add , , "Text", 4500
    End With
    VCount = 1
    View2.View = lvwReport
    With View2.ColumnHeaders
        .Add , , "Handle", 1000
        .Add , , "Class Name", 1500
        .Add , , "Text", 4500
        .Add , , "IsPassword field", 1000
        
    End With
    ICount = 1
    Options.Visible = False
    
    Dim hWndForm
    
    hWndForm = FindWindow("CabinetWClass", vbNullString)
    
    ' GotoChildhWnd (hWndForm)
    
End Sub

Private Sub Form_Resize()
    ClassResize.rSize Form7_Enum_Child
    
    'OK now resize if you must!
     View2.Left = Int(Form7_Enum_Child.width / 2)
     View1.width = View2.Left - 255
     View2.width = Int(Form7_Enum_Child.width / 2) - 255
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
End Sub

Private Sub Hide_Click()
    ShowWindow Val(View1.SelectedItem), SW_HIDE
End Sub

Private Sub Label2_Click()
    Dim ret As Long
    ret = ShellExecute(Me.hWnd, "Open", "http://go.to/abubakar", "", App.Path, 1)

End Sub

Private Sub Max_Click()
    ShowWindow Val(View1.SelectedItem), SW_MAXIMIZE
    
End Sub

Private Sub Min_Click()
    ShowWindow Val(View1.SelectedItem), SW_MINIMIZE
End Sub

Private Sub Restore_Click()
    ShowWindow Val(View1.SelectedItem), SW_RESTORE
End Sub

Private Sub Show_BWTT_Click()
    Dim LHWND As Long
    
    On Error GoTo bugging
    LHWND = Val(View1.SelectedItem)
    'ShowWindow LHWND, SW_SHOW
    BringWindowToTop LHWND
    
    Exit Sub
bugging:
    Rem Do Nothing
    
End Sub

Private Sub Show_Click()
    'show window code goes here:
    Dim LHWND As Long
    On Error Resume Next

    LHWND = Val(View1.SelectedItem)
    ShowWindow LHWND, SW_SHOW
End Sub

Private Sub SpyMenu_Click()
    Dim st As RECT
    
    Spy_Form.Show
    SpyhWnd = Val(View1.SelectedItem)
    Spy_Form.Tree.Nodes.Clear
    'If its a MDI type window and its child windows are maximized
    'then 'GetMenuItemInfo' crashes the 'EnumerationX'.
    'I tried to cascade the windows of other app but that doesnt
    'happen, do you know how I can do this?
    'MsgBox CascadeWindows(SpyhWnd, MDITILE_SKIPDISABLED, st, 0, 0)
    'SendMessage SpyhWnd, WM_MDICASCADE, MDITILE_SKIPDISABLED, 0
    'SendMessage SpyhWnd, WM_MDITILE, MDITILE_HORIZONTAL, 0
    
    SMenu GetMenu(SpyhWnd), Spy_Form.Tree
        
End Sub

Private Sub View1_Click()
    GotoChild
End Sub

Private Sub View1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then GotoChild
                'So that you are able to see child windows easily by
                'scrolling through up-down arrow keys instead of
                'clicking the parent window handle every time.
    
End Sub

Private Sub GotoChildhWnd(hWnd_VAR)
    On Error GoTo HandleErrorPlz
    
    Dim num As Long
    Dim myLong As Long
    'Num = Val(View1.SelectedItem)
    num = hWnd_VAR
    View2.ListItems.Clear
    View2.GridLines = True
    ICount = 1
    myLong = EnumChildWindows(num, AddressOf WndEnumChildProc, View2)

HandleErrorPlz:
    'Exit Sub ' As simple as that :)
End Sub

Private Sub GotoChild()
    On Error GoTo HandleErrorPlz
    
    Dim num As Long
    Dim myLong As Long
    num = Val(View1.SelectedItem)
    View2.ListItems.Clear
    View2.GridLines = True
    ICount = 1
    myLong = EnumChildWindows(num, AddressOf WndEnumChildProc, View2)

HandleErrorPlz:
    'Exit Sub ' As simple as that :)
End Sub

Private Sub View1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And View1.ListItems.Count > 0 Then
        If GetMenu(Val(View1.SelectedItem)) > 0 Then
            SpyMenu.Enabled = True
        Else
            SpyMenu.Enabled = False
        End If
               
        PopupMenu Options
    End If
    
End Sub
Private Sub View2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And View2.ListItems.Count > 0 Then
        PopupMenu menu2
    End If

End Sub
