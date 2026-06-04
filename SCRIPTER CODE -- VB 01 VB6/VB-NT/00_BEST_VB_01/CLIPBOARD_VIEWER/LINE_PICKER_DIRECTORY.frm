VERSION 5.00
Begin VB.Form LINE_PICKER_DIRECTORY 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   5892
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   12864
   LinkTopic       =   "Form1"
   ScaleHeight     =   5892
   ScaleWidth      =   12864
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer TIMER_GET_PATH_CLIPBOARD 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8208
      Top             =   408
   End
   Begin VB.DirListBox Dir1 
      Height          =   720
      Left            =   10344
      TabIndex        =   2
      Top             =   108
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Timer Timer_SET_FOCUS 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9276
      Top             =   300
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   168
      TabIndex        =   0
      Top             =   876
      Width           =   12504
   End
   Begin VB.Label Label1 
      Caption         =   "SPEED LINE PICKER TO CLIPBOARD"
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
      TabIndex        =   1
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
Attribute VB_Name = "LINE_PICKER_DIRECTORY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OLD_CLIP_BOARD_STRING
Dim DIR_FOLDER
Dim FSO

Dim M()


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

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Sub Form_Load()
    Me.Show
    Me.Caption = FRM_ClipTest.Caption
    
    ReDim M(100)
    ' LINE_PICKER_DIRECTORY_LIST_SET
        
    On Error Resume Next
    '-------------------------------------------------------------------------------------
    CLIP_BOARD_STRING = Trim(Replace(Trim(Clipboard.GetText), """", ""))
    On Error GoTo 0
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    DIR_FOLDER = ""
    If FSO.FolderExists(CLIP_BOARD_STRING) = True Then
        DIR_FOLDER = CLIP_BOARD_STRING
    End If
    If FSO.FileExists(CLIP_BOARD_STRING) = True Then
        DIR_FOLDER = Mid(CLIP_BOARD_STRING, 1, InStrRev(CLIP_BOARD_STRING, "\") - 1)
    End If
            
    If DIR_FOLDER = "" Then
        TIMER_GET_PATH_CLIPBOARD.Enabled = True
    End If
            
    
    i = 0
    i = i + 1: M(i) = Dir1.Path

    If DIR_FOLDER <> "" Then
        Dir1.Path = DIR_FOLDER
        Dir1.Refresh
        If Dir1.ListCount = 0 Then
            Dir1.Path = Mid(Dir1.Path, 1, InStrRev(Dir1.Path, "\") - 1)
            i = i + 1: M(i) = Dir1.Path
            Dir1.Refresh
        End If
        For R1 = 0 To Dir1.ListCount - 1
           i = i + 1: M(i) = Mid(Dir1.List(R1), InStrRev(Dir1.List(R1), "\"))
        Next
        For R1 = 0 To Dir1.ListCount - 1
           i = i + 1: M(i) = Dir1.List(R1)
        Next
    End If

    I_BOOKMARK = i
    i = i + 1: M(i) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE"
    i = i + 1: M(i) = "\\2-ASUS-EEE\2_ASUS_EEE_01_C_DRIVE"
    i = i + 1: M(i) = "\\3-LINDA-PC\3_LINDA_PC_01_C_DRIVE"
    i = i + 1: M(i) = "\\4-ASUS-GL522VW\4_Asus_Gl522Vw_C_Drive"
    i = i + 1: M(i) = "\\5-ASUS-P2520LA\5_ASUS_P2520LA_01_C_DRIVE"
    i = i + 1: M(i) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW_01_C_DRIVE"
    i = i + 1: M(i) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_01_C_DRIVE"
    i = i + 1: M(i) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_02_D_DRIVE"
    i = i + 1: M(i) = "\\2-ASUS-EEE\2_ASUS_EEE_02_D_DRIVE"
    i = i + 1: M(i) = "\\3-LINDA-PC\3_LINDA_PC_02_D_DRIVE"
    i = i + 1: M(i) = "\\4-ASUS-GL522VW\4_Asus_Gl522Vw_D_Drive"
    i = i + 1: M(i) = "\\5-ASUS-P2520LA\5_ASUS_P2520LA_02_D_DRIVE"
    i = i + 1: M(i) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW_02_D_DRIVE"
    i = i + 1: M(i) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_02_D_DRIVE"
    i = i + 1: M(i) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_03_FAT32_4GB"
    i = i + 1: M(i) = "\\2-ASUS-EEE\2_ASUS_EEE_03_FAT32_4GB"
    i = i + 1: M(i) = "\\3-LINDA-PC\3_LINDA_PC_03_FAT32_4GB"
    i = i + 1: M(i) = "\\4-ASUS-GL522VW\4_Asus_Gl522Vw_E_Drive"
    i = i + 1: M(i) = "\\5-ASUS-P2520LA\5_ASUS_P2520LA_03_FAT32_4GB"
    i = i + 1: M(i) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW_03_FAT32_4GB"
    i = i + 1: M(i) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_03_FAT32_4GB"
    For R3 = I_BOOKMARK To i
        If InStr(M(R3), "_D_DRIVE") > 0 Then
            i = i + 1: M(i) = M(R3) + "\VB6\VB-NT"
        End If
    Next
    
    
    
    
    i = i + 1: M(i) = ""
    i = i + 1: M(i) = ""
    i = i + 1: M(i) = ""
    i = i + 1: M(i) = ""
    
    List1.Clear
    ' List1.Sorted = False
    For R3 = 1 To UBound(M)
        If M(R3) <> "" Then
            AR = AR + 1
            List1.AddItem Format(AR, "000") + " --  " + M(R3)
        End If
    Next

End Sub



Private Sub Form_Resize()
If RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True Then
    Exit Sub
End If
If Me.WindowState <> RESIZE_WINDOWSTATE_CHANGE Then
    RESIZE_DONE_ONCE = False
End If
If RESIZE_WINDOWSTATE_CHANGE = vbMinimized And (Me.WindowState = vbMaximized Or Me.WindowState = vbNormal) Then
    LINE_PICKER_DIRECTORY.SetFocus
    Timer_SET_FOCUS.Enabled = True
    DoEvents
End If
RESIZE_WINDOWSTATE_CHANGE = Me.WindowState

If RESIZE_DONE_ONCE = True Then
    Exit Sub
End If

RESIZE_DONE_ONCE = True
Me.Visible = True

If Me.WindowState = vbNormal Then
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
    Me.Top = FRM_ClipTest.Top + 200
    Me.Left = FRM_ClipTest.Left + 200
    Me.Width = FRM_ClipTest.Width
    Me.Height = FRM_ClipTest.Height
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
End If
List1.Top = 500
List1.Left = 100
List1.Width = Me.Width - 300
H1 = Me.Height - (List1.Top + 800 + 20)
If H1 > 0 Then
    List1.Height = H1
    Else
    List1.Height = Me.Height
End If

If Me.WindowState = vbNormal Then
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
    Me.Height = List1.Height + List1.Top + 900
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
End If
Dim CONT As Control
On Error Resume Next
For Each CONT In Form1.Controls
    CONT.Font.Name = "SOURCE CODE PRO SEMIBOLD" ' -- "ARIAL"
    CONT.FontSize = 12
    CONT.FontBold = True
Next

Label1.FontSize = 14

Label1.BackColor = RGB(255, 255, 255)
Label1.Top = 80
Label1.Left = List1.Left + 40
Label1.Height = 400
Label1.Caption = "SPEED LINE PICKER SEND TO CLIPBOARD"
Label1.Width = List1.Width - 80





End Sub

Private Sub Form_Unload(Cancel As Integer)

    WORKER_WITH_DIR_PICKER = False
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
    
    LINE_PICKER_DIRECTORY.WindowState = vbMinimized
    RESULT_API = NotAlwaysOnTop(Me.hWnd)
    FRM_ClipTest.WindowState = vbMinimized
    LINE_PICKER_DIRECTORY.Visible = False
    If WORKER_WITH_DIR_PICKER = True Then
        FRM_ClipTest.Visible = False
        Else
        FRM_ClipTest.Visible = True
    End If
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False

'If IsIDE = False Then
    If Me.WindowState <> vbMinimized And Me.EXIT_TRUE = False Then
        Cancel = True
        Exit Sub
    End If
'End If

End Sub

Private Sub List1_Click()

ITEMGET_1 = List1.List(List1.ListIndex)
ITEMGET_1 = Mid(ITEMGET_1, InStr(ITEMGET_1, " -- ") + 5)

Dim VAR_ST_1
Dim VAR_ST_2
Dim VAR_ST_4

VAR_ST_1 = ITEMGET_1

Do
    On Error Resume Next
    Clipboard.Clear
    VAR_ST_2 = Clipboard.GetText
    On Error GoTo 0
    If VAR_ST_2 <> "" Then
        DoEvents
        Sleep 500
    End If
Loop Until VAR_ST_2 = ""
Sleep 100
Do
    On Error Resume Next
    Clipboard.SetText VAR_ST_1, vbCFText
    Sleep 100
    VAR_ST_2 = Clipboard.GetText
    On Error GoTo 0
    If VAR_ST_2 <> VAR_ST_1 Then
        DoEvents
        Sleep 500
    End If
Loop Until VAR_ST_2 = VAR_ST_1

If Err.Number = 0 Then
    
    WORKER_WITH_DIR_PICKER = True
    If WORKER_WITH_DIR_PICKER = True Then
        FRM_ClipTest.Visible = False
        Else
        FRM_ClipTest.Visible = True
    End If

    If AlwaysOnTop_MODE = False Then
        FRM_ClipTest.WindowState = vbMinimized
        LINE_PICKER_DIRECTORY.WindowState = vbMinimized
        
        RESULT_API = NotAlwaysOnTop(Me.hWnd)
    End If
End If

Beep
End Sub

Private Sub MNU_ALWAYS_ON_TOP_Click()
    RESULT_API = AlwaysOnTop(Me.hWnd)
    AlwaysOnTop_MODE = True
End Sub

Private Sub MNU_UNLOAD_FORM_Click()

    Unload Me

End Sub

Public Sub Timer_SET_FOCUS_Timer()

Me.SetFocus
Timer_SET_FOCUS.Enabled = False

'
'If WORKER_WITH_DIR_PICKER = True Then
'    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
'    FRM_ClipTest.Visible = True
'    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
'End If

End Sub


Private Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Private Function NotAlwaysOnTop(ByVal hWnd As Long)
    Dim flags
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function



Private Sub TIMER_GET_PATH_CLIPBOARD_Timer()

    On Error Resume Next
    '-------------------------------------------------------------------------------------
    If Clipboard.GetFormat(vbCFText) = False Then Exit Sub

    CLIP_BOARD_STRING = Mid(FRM_ClipTest.Text2.Text, 1, 1000)
    On Error GoTo 0
    CLIP_BOARD_STRING = Trim(Replace(Trim(CLIP_BOARD_STRING), """", ""))
    On Error GoTo 0
    
    If OLD_CLIP_BOARD_STRING = CLIP_BOARD_STRING Then
        Exit Sub
    End If
    
    
    If OLD_CLIP_BOARD_STRING <> CLIP_BOARD_STRING Then
        DIR_FOLDER = ""
    End If
    
    OLD_CLIP_BOARD_STRING = CLIP_BOARD_STRING
    
    DIR_FOLDER = ""
    If FSO.FolderExists(CLIP_BOARD_STRING) = True Then
        DIR_FOLDER = CLIP_BOARD_STRING
    End If
    If FSO.FileExists(CLIP_BOARD_STRING) = True Then
        DIR_FOLDER = Mid(CLIP_BOARD_STRING, 1, InStrRev(CLIP_BOARD_STRING, "\") - 1)
    End If

    If DIR_FOLDER <> "" Then
        ' TIMER_GET_PATH_CLIPBOARD.Enabled = False
        Call Form_Load
    End If

    


End Sub
