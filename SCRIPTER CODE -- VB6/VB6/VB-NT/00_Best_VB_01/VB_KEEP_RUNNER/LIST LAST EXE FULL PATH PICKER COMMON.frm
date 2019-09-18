VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form LINE_EXE_PICKER_COMMON 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   5892
   ClientLeft      =   132
   ClientTop       =   780
   ClientWidth     =   12864
   LinkTopic       =   "Form1"
   ScaleHeight     =   5892
   ScaleWidth      =   12864
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer_SET_FOCUS 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9276
      Top             =   300
   End
   Begin MSComctlLib.ListView LISTVIEW1 
      Height          =   576
      Left            =   816
      TabIndex        =   1
      Top             =   2124
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   995
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "SPEED LINE PICKER EXECUTE EXE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   288
      TabIndex        =   0
      Top             =   132
      Width           =   7752
   End
   Begin VB.Menu MNU_UNLOAD_FORM 
      Caption         =   "HIDE FORM"
   End
   Begin VB.Menu MNU_ALWAYS_ON_TOP 
      Caption         =   "ALWAYS ON TOP"
   End
End
Attribute VB_Name = "LINE_EXE_PICKER_COMMON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RESIZE_DONE_ONCE
Dim RESIZE_WINDOWSTATE_CHANGE
Public EXIT_TRUE
Dim RESULT_API
Dim AlwaysOnTop_MODE

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVS_EX_FULLROWSELECT = &H20


Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Sub Form_Resize()


If RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True Then
    Exit Sub
End If
If Me.WindowState <> RESIZE_WINDOWSTATE_CHANGE Then
    RESIZE_DONE_ONCE = False
End If
If RESIZE_WINDOWSTATE_CHANGE = vbMinimized And (Me.WindowState = vbMaximized Or Me.WindowState = vbNormal) Then
    LINE_EXE_PICKER_COMMON.SetFocus
    Timer_SET_FOCUS.Enabled = True
    DoEvents
End If
RESIZE_WINDOWSTATE_CHANGE = Me.WindowState

If RESIZE_DONE_ONCE = True Then
    Exit Sub
End If

RESIZE_DONE_ONCE = True
'Me.Visible = True

If Me.WindowState = vbNormal Then
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
    Me.Top = Form1.Top + 200
    Me.Left = Form1.Left + 200
    Me.width = Form1.width
    Me.height = Form1.height
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
End If
LISTVIEW1.Top = 500
LISTVIEW1.Left = 100
LISTVIEW1.width = Me.width - 300
H1 = Me.height - (LISTVIEW1.Top + 800 + 20)
If H1 > 0 Then
    LISTVIEW1.height = H1
    Else
    LISTVIEW1.height = Me.height
End If


If Me.WindowState = vbNormal Then
    If LISTVIEW1.width = Me.width - 300 Then
        RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
        Me.height = LISTVIEW1.height + LISTVIEW1.Top + 900
        RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
    End If
End If

'LISTVIEW1.FontSize = 12
'LISTVIEW1.FontBold = True
'LISTVIEW1.FontName = "Arial"

LISTVIEW1.Left = 100

LISTVIEW1.Font.size = 14
' LISTVIEW1.width = 2400


Call LV_AutoSizeColumn(LISTVIEW1, Nothing)


Label1.BackColor = RGB(255, 255, 255)
Label1.Top = 80
Label1.Left = LISTVIEW1.Left + 40
Label1.height = 400
Label1.Caption = "SPEED LINE PICKER SEND TO CLIPBOARD"
Label1.width = LISTVIEW1.width - 80
Label1.FontSize = 14
Label1.FontBold = True
Label1.FontName = "Arial"


End Sub



Public Sub LOAD_INITAL_FORM_VALUE()


Dim M()
ReDim M(100)
i = 0
i = i + 1: M(i) = App.Path + "\" + App.EXEName + ".exe"
i = i + 1: M(i) = "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
i = i + 1: M(i) = "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\URL Logger.exe"
i = i + 1: M(i) = "C:\Program Files (x86)\Notepad++\notepad++.exe"
i = i + 1: M(i) = "D:\VB6\VB-NT\00_Best_VB_01\CLIPBOARD_VIEWER\ClipBoard Viewer.exe"
i = i + 1: M(i) = "C:\Program Files (x86)\Siber Systems\AI RoboForm\identities.exe"
i = i + 1: M(i) = "C:\Program Files\Process Lasso\ProcessLasso.exe"
i = i + 1: M(i) = "C:\Windows\explorer.exe"
i = i + 1: M(i) = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe"
i = i + 1: M(i) = "D:\GoodSync\x64\GoodSync2Go.exe"
i = i + 1: M(i) = "C:\Program Files\Norton Security\Engine\22.18.0.213\NortonSecurity.exe"
i = i + 1: M(i) = "C:\Program Files (x86)\Duplicate Cleaner Pro\DuplicateCleaner.exe"
i = i + 1: M(i) = ""
i = i + 1: M(i) = ""
i = i + 1: M(i) = ""
i = i + 1: M(i) = ""

Dim LV1

' LISTVIEW1.Sorted = True
For R3 = 1 To UBound(M)
    If M(R3) <> "" Then
        I_1 = M(R3)
        I_2 = ""
        'ADDITEM
        ' LINE_EXE_PICKER_COMMON.LISTVIEW1.ListItems.Add
        With LINE_EXE_PICKER_COMMON.LISTVIEW1
            Set LV1 = .ListItems.Add(, , I_2)
                LV1.SubItems(1) = I_1
        End With
            '------------------------
            'PAIN TO GET THE FORMULAR
            'EVEN THEN FAR OUT AGAIN
            '------------------------
        End If
Next


End Sub

Sub Form_Load()
    
'    Me.Show
    Me.Caption = Form1.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)

    WORKER_WITH_PICKER = False
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
    
    LINE_EXE_PICKER_COMMON.WindowState = vbMinimized
    RESULT_API = NotAlwaysOnTop(Me.hWnd)
    Form1.WindowState = vbMinimized

    LINE_EXE_PICKER_COMMON.Visible = False
    If WORKER_WITH_PICKER = True Then
        Form1.Visible = False
        Else
        Form1.Visible = True
    End If
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False

'If IsIDE = False Then
    If Me.WindowState <> vbMinimized And Me.EXIT_TRUE = False Then
        Cancel = True
        Exit Sub
    End If
'End If

End Sub

Private Sub LISTVIEW1_Click()
    
    
    ITEMGET = LISTVIEW1.ListItems(LISTVIEW1.SelectedItem.Index).SubItems(1)
    ' ITEMGET_PID = lstProcess_2_ListView.ListItems(lstProcess_2_ListView.SelectedItem.Index)

    ' ITEMGET = Mid(ITEMGET, InStr(ITEMGET, " -- ") + 5)
    
    Dim VAR_ST_1
    Dim VAR_ST_2
    Dim VAR_ST_4
    
    VAR_ST_1 = ITEMGET
    
    Me.Hide

    Shell VAR_ST_1, vbNormalFocus

    ' Form1.Visible = False

'    If AlwaysOnTop_MODE = False Then
'        Form1.WindowState = vbMinimized
'        LINE_EXE_PICKER_COMMON.WindowState = vbMinimized
'        LINE_EXE_PICKER_COMMON.Visible = False
'        RESULT_API = NotAlwaysOnTop(Me.hWnd)
'    End If


    Beep
End Sub

Private Sub MNU_ALWAYS_ON_TOP_Click()
    RESULT_API = AlwaysOnTop(Me.hWnd)
    AlwaysOnTop_MODE = True
End Sub

Private Sub MNU_UNLOAD_FORM_Click()

    Me.Visible = False
    
    Exit Sub

    Unload Me

End Sub

Public Sub Timer_SET_FOCUS_Timer()

    Me.SetFocus
    Timer_SET_FOCUS.Enabled = False
    
End Sub


Private Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Private Function NotAlwaysOnTop(ByVal hWnd As Long)
    Dim flags
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, flags
End Function


Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column As ColumnHeader = Nothing)
    Dim c As ColumnHeader
    If Column Is Nothing Then
    For Each c In LV.ColumnHeaders
        SendMessage LV.hWnd, LVM_FIRST + 30, c.Index - 1, -1
    Next
    Else
        SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
    End If
    LV.Refresh
End Sub

