VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form LINE_PICKER_EXE_SCREEN 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   192
   ClientTop       =   240
   ClientWidth     =   15456
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   15456
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer_RUN_ONCE_RESIZE 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5100
      Top             =   840
   End
   Begin VB.ListBox LIST2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   2340
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1392
      Visible         =   0   'False
      Width           =   11004
   End
   Begin VB.Timer Timer_MONITOR_FORM_SIZE_WIDTH_HEIGHT 
      Interval        =   100
      Left            =   4008
      Top             =   900
   End
   Begin VB.Timer TIMER_TO_RESIZE 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4368
      Top             =   900
   End
   Begin MSComctlLib.ListView LISTVIEW1 
      Height          =   696
      Left            =   480
      TabIndex        =   2
      Top             =   864
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1207
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Timer Timer_SET_FOCUS 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4740
      Top             =   900
   End
   Begin VB.ListBox LIST1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   2340
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   888
      Width           =   1476
   End
   Begin VB.Label Label_CLIPBOARD_SELECTOR_1 
      Caption         =   "CLIPBOARD SELECTOR   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   10176
      TabIndex        =   6
      Top             =   540
      Width           =   2904
   End
   Begin VB.Label Label_FOLDER_SELECTOR 
      Caption         =   "FOLDER SELECTOR   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   7404
      TabIndex        =   5
      Top             =   540
      Width           =   2664
   End
   Begin VB.Label Label_EXE 
      Caption         =   "EXECUTE / URL SELECTION   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   4608
      TabIndex        =   4
      Top             =   540
      Width           =   2664
   End
   Begin VB.Label Label1 
      Caption         =   "SPEED LINE PICKER SEND TO CLIPBOARD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   504
      TabIndex        =   1
      Top             =   240
      Width           =   4992
   End
   Begin VB.Menu MNU_ENDER 
      Caption         =   "ENDER -- ISIDE ONE"
   End
   Begin VB.Menu MNU_UNLOAD_FORM 
      Caption         =   "HIDE FORM"
   End
   Begin VB.Menu MNU_ALWAYS_ON_TOP 
      Caption         =   "ALWAYS ON TOP"
   End
   Begin VB.Menu MNU_EDITOR_NOTEPAD_01 
      Caption         =   "EDITOR NOTEPAD++ FILE DATA 01"
   End
   Begin VB.Menu MNU_ADD_CLIPBOARD 
      Caption         =   "ADD CLIBOARD"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_SHOW_MAIN_FORM 
      Caption         =   "MAIN FORM BACK"
   End
   Begin VB.Menu MNU_MINIMIZE_MAIN_FORM 
      Caption         =   "MAIN FORM -- MINIMIZE"
   End
End
Attribute VB_Name = "LINE_PICKER_EXE_SCREEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------

Public FORM_UNLOAD_INITIATE
Dim NOT_RESIZE_ROUTINE
Public LINE_PICKER_EXE_SCREEN_FROM_START

Dim OLD_LIST1_WIDTH
Dim OLD_LIST1_HEIGHT
Dim MOUSE_UP
Dim MOUSE_DOWN
Dim GETWINNT_VER_VAR
Dim TIMER_TO_RESIZE_VAR
Dim OLD_HEIGHT_WIDTH
Dim MNU_SORT_A_VAR
Dim FILENAME_LOAD_01
Dim FILENAME_LOAD_02

Dim MNU_DOUBLE_LINE_SELECTOR_ON_VAR
Dim M()

Dim RESIZE_DONE_ONCE

Public EXIT_TRUE
Dim RESULT_API
Dim AlwaysOnTop_MODE


Const hWnd_TOPMOST = -1
Const hWnd_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Dim HY, WX

Dim ARCHIVE_Menu_Height

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type MENUBARINFO
  cbSize As Long
  rcBar As RECT
  hMenu As Long
  hWndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer

' ---------------------------------------
' HAPPEN MINIMIZE AFTER AND THEN END FORM
' ---------------------------------------
Dim VAR_WINSTATE_BEEN_HIGH


Dim FORM_LOAD_STATE

'Private Type MENUBARINFO
'  cbSize As Long
'  rcBar As RECT
'  hMenu As Long
'  hWndMenu As Long
'  fBarFocused As Boolean
'  fFocused As Boolean
'End Type
'Private MenuInfo As MENUBARINFO
'Private Const OBJID_MENU As Long = &HFFFFFFFD
'Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
' Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
' Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
 




Sub GOT_FOCUS()

'    If LINE_PICKER_FORM_LIST_CLICK_VALUE > 0 Then
'        PLUS_10 = LINE_PICKER_FORM_LIST_CLICK_VALUE + 10
'        NOT__10 = LINE_PICKER_FORM_LIST_CLICK_VALUE - 10
'        LIST_COUNT = LISTVIEW1.ListItems.Count
'        If NOT__10 < 1 Then NOT__10 = 1
'        If PLUS_10 > LIST_COUNT Then PLUS_10 = LIST_COUNT
'        LISTVIEW1.ListItems(NOT__10).EnsureVisible
'        LISTVIEW1.ListItems(PLUS_10).EnsureVisible
'        LISTVIEW1.ListItems(LINE_PICKER_FORM_LIST_CLICK_VALUE).EnsureVisible
'        LISTVIEW1.ListItems(LINE_PICKER_FORM_LIST_CLICK_VALUE).Selected = True
'        On Error Resume Next
'        ' FORM NOT VISIBLE OR SOMETHING
'        LISTVIEW1.SetFocus
'        On Error GoTo 0
'    End If
'
    ' -----------------------------------
    ' For i = 0 To Screen.FontCount - 1
    '     A = A + Screen.Fonts(i) + vbCrLf
    ' Next i
    ' -----------------------------------
    ' Clipboard.Clear
    ' Clipboard.SetText A
    ' -----------------------------------
    ' Fri 25-Sep-2020 19:51:41
    ' -----------------------------------
    ' Change Fontstyle Using VB - VB6 | Dream.In.Code
    ' https://www.dreamincode.net/forums/topic/69063-change-fontstyle-using-vb/
    ' -----------------------------------
    Dim CONT As Control
    On Error Resume Next
    For Each CONT In Me.Controls
        FONT_SIZE = 12
'        Select Case UCase(Cont.Name)
'            Case "MMM_SHOW_THE_TIME_ARR": FONT_SIZE = 9
'        End Select
'        If InStr(UCase(Cont.Name), "TEXT") > 0 Then
'            FONT_SIZE = 12
'        End If
'        If InStr(UCase(Cont.Name), "MMM_SHOW_THE_TIME_ARR") > 0 Then
'        If InStr(UCase(Cont.Caption), "000") > 0 Then FONT_SIZE = 13
'        If InStr(UCase(Cont.Caption), "TIME SLICE") > 0 Then FONT_SIZE = 8.5
'        If Cont.Index = 5 Then FONT_SIZE = 9
'        If Cont.Index = 6 Then FONT_SIZE = 9.4
'        If Cont.Index = 7 Then FONT_SIZE = 9.4
'        End If
        FONT_NAME = "SOURCE CODE PRO SEMIBOLD"
        FONT_NAME = "SOURCE CODE PRO BLACK"
        FONT_NAME = "ARIAL"
        FONT_NAME = "ARIAL BLACK"
        FONT_NAME = "SOURCE CODE PRO BLACK"
        CONT.Font.Name = FONT_NAME
        CONT.FontBold = True
        CONT.FontSize = FONT_SIZE
        CONT.Font.Size = FONT_SIZE
        ' TEXT BOX REQUIRE HERE STYLE __ CONT.FONT.Size = FONT_SIZE
    Next

End Sub




Private Sub Form_Activate()
    
    AlwaysOnTop_MODE = True
    If AlwaysOnTop_MODE = True Then
        MNU_ALWAYS_ON_TOP.Caption = "ALWAYS ON TOP __ CORRECT"
    Else
        MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __]"
    End If
    
    Call GOT_FOCUS
    
End Sub

Private Sub Form_GotFocus()

    Call GOT_FOCUS

End Sub

Sub Form_Load()
    If IsIDE = False Then
        MNU_ENDER.Visible = False
    End If
    Set FSO = CreateObject("Scripting.FileSystemObject")

    RESIZE_DONE_ONCE = False
    NOT_RESIZE_ROUTINE = False

    MNU_SORT_A_VAR = True
    
    
    If LINE_PICKER_EXE_SCREEN.LINE_PICKER_EXE_SCREEN_FROM_START = False Then
        Me.Show
        Me.Caption = FRM_ClipTest.Caption
        LINE_PICKER_EXE_SCREEN.LINE_PICKER_EXE_SCREEN_FROM_START = False
    End If
    
    GETWINNT_VER_VAR = GETWinNT_Ver
    
    ReDim M(9000)
    ' LINE_PICKER_EXE_SCREEN_LIST_SET
    
    i = 0
    
    ' ---------------------------------------------------------------------
    ' LOAD FILE INTO ARRAY 01 OF 02 BEGIN
    ' ---------------------------------------------------------------------
    ' BEGIN HEADER LINE ENTER IN READER AS PER TEXT FILE AND DETECTOR
    ' ---------------------------------------------------------------------
    ' ! CLIPBOARD CHUNK BELOW -- ADBLOCK FOR CHROME
    ' ---------------------------------------------------------------------
    ' ENDER LINE
    ' ---------------------------------------------------------------------
    ' ! END CHUNK -- CLIPBOARD CHUNK BELOW -- ADBLOCK.COM CHROME
    ' ---------------------------------------------------------------------
    
    ' ---------------------------------------------------------------------
    ' AND THE HARDCOER INTO ARRAY
    ' ---------------------------------------------------------------------
'    i = i + 1: M(i) = "name *.EXE"
'    i = i + 1: M(i) = "name *#NFS_EX*"
'    i = i + 1: M(i) = "name *#NFS_EX*; name *"
'    i = i + 1: M(i) = "----------"
    
    
    
    
    REMLINE_2 = 0
    CreateFolderTree App.Path + "\# DATA"
    FILENAME_LOAD_01 = App.Path + "\# DATA\VB__" + UCase(App.EXEName) + "_LINE_PICKER_EXE_SCREEN.TXT"
    If Dir(FILENAME_LOAD_01) <> "" Then
        FR1 = FreeFile
        Open FILENAME_LOAD_01 For Input As #FR1
        Do
            Line Input #FR1, VALUE_CHUNK
            If Mid(VALUE_CHUNK, 1, 1) = "#" Then REMLINE = True Else REMLINE = False
                If REMLINE = False And REMLINE_2 = 1 Then REMLINE_2 = -2
                If REMLINE = True Then REMLINE_2 = 1
                If REMLINE = True And REMLINE_2 = 1 Then BLOCK_REMLINE = True
                If REMLINE_2 = -2 Then BLOCK_REMLINE = False
                If Trim(VALUE_CHUNK) = "" Then BLOCK_REMLINE = True
                If Mid(VALUE_CHUNK, 1, 3) = "E:\" Then BLOCK_REMLINE = True
                If BLOCK_REMLINE = False Then
                    If Dir(VALUE_CHUNK) = "" Then
                    If InStr(UCase(VALUE_CHUNK), "(X86)") Then
                        VALUE_CHUNK = Replace(VALUE_CHUNK, " (X86)", "")
                        VALUE_CHUNK = Replace(VALUE_CHUNK, " (x86)", "")
                    End If
                    End If
                    If Dir(VALUE_CHUNK) = "" Then
                    If InStr(UCase(VALUE_CHUNK), "Program Files\") > 0 Then
                        VALUE_CHUNK = Replace(VALUE_CHUNK, "Program Files\", "Program Files (x86)\")
                    End If
                    End If
                    
                    i = i + 1: M(i) = VALUE_CHUNK
                End If
        Loop Until EOF(FR1)
        Close #FR1
    End If
    ' ---------------------------------------------------------------------
    ' LOAD FILE INTO ARRAY 01 OF 02 DONE
    ' ---------------------------------------------------------------------
    
'    ' ---------------------------------------------------------------------
'    ' LOAD FILE INTO ARRAY 02 OF 02 BEGIN
'    ' ---------------------------------------------------------------------
'    FILENAME_LOAD_02 = "C:\SCRIPTER\SCRIPT 08 IN_NET_ ADBLOCK\TEXT_ADBLOCK_01_MAIN.TXT"
'    If Dir(FILENAME_LOAD_02) <> "" Then
'        FR1 = FreeFile
'        Open FILENAME_LOAD_02 For Input As #FR1
'        Do
'            Line Input #FR1, VALUE_CHUNK
'            If Mid(VALUE_CHUNK, 1, 1) = "#" Then REMLINE = True Else REMLINE = False
'                If REMLINE = False And REMLINE_2 = 1 Then REMLINE_2 = -2
'                If REMLINE = True Then REMLINE_2 = 1
'                If REMLINE = True And REMLINE_2 = 1 Then BLOCK_REMLINE = True
'                If REMLINE_2 = -2 Then BLOCK_REMLINE = False
'                If Trim(VALUE_CHUNK) = "" Then BLOCK_REMLINE = True
'                If BLOCK_REMLINE = False Then
'                    i = i + 1: M(i) = VALUE_CHUNK
'                End If
'        Loop Until EOF(FR1)
'        Close #FR1
'    End If
'    ' ---------------------------------------------------------------------
'    ' LOAD FILE INTO ARRAY 02 OF 02 DONE
'    ' ---------------------------------------------------------------------
    

    
    ' ---------------------------------------------------------------------
    ' THE ARRAY BUILD ENDER HERE
    ' ---------------------------------------------------------------------
    
    ReDim Preserve M(i)
    LIST1.Clear
    AR = 0
    ' List1.Sorted = False
    For r3 = 1 To UBound(M)
        If M(r3) <> "" Then
            AR = AR + 1
            LIST1.AddItem Format(AR, "000") + " --  " + M(r3)
        End If
    Next
    
    
    
    
    
    
    
    ' LISTVIEW1.
    With LISTVIEW1
        .ColumnHeaders.Add , "ID", "ID", 800, lvwColumnLeft
        .ColumnHeaders.Add , "IDX", "IDX", 800, lvwColumnLeft
        .ColumnHeaders.Add , "ITEM", "ITEM", 24000, lvwColumnLeft
        .View = lvwReport
        .Height = LIST1.Height
        .Top = LIST1.Top
        .Left = LIST1.Left
        .Width = LIST1.Width
'        .Font.Bold = True
'        .Font.Name = LIST1.FontNaizeme
'        .Font.Size = LIST1.Fonts

        .FullRowSelect = True
    End With
    
    LISTVIEW1.ListItems.Clear
    LIST1.Visible = False
    
    For r3 = 1 To UBound(M)
        If M(r3) <> "" Then
            AR = AR + 1
            With LISTVIEW1
                Set LV1 = .ListItems.Add(, , Format(AR, "000"))  ' AR VARIABLE COUNTER
                LV1.SubItems(2) = M(r3)
            End With
        End If
    Next
    
    If MNU_SORT_A_VAR = True Then
        LISTVIEW1.SortOrder = lvwAscending
        LISTVIEW1.SortKey = 2
        LISTVIEW1.Sorted = True
    End If
    
    For r = LISTVIEW1.ListItems.Count To 2 Step -1
            ITEMGET_1 = LISTVIEW1.ListItems(r - 1).SubItems(2)
            ITEMGET_2 = LISTVIEW1.ListItems(r).SubItems(2)
            If UCase(ITEMGET_1) = UCase(ITEMGET_2) Then
                LISTVIEW1.ListItems.Remove (r)
            End If
    Next
    
    AR = 0
    For r3 = 1 To LISTVIEW1.ListItems.Count
        If LISTVIEW1.ListItems.Item(r3).SubItems(2) <> "" Then
            AR = AR + 1
            FILENAME_VAR = LISTVIEW1.ListItems.Item(r3).SubItems(2)
            If FSO.FileExists(FILENAME_VAR) = True Then
                ADD_STAR = "1 +"
            Else
                ADD_STAR = "3 -"
            End If
            If FSO.FileExists(FILENAME_VAR) = False And FSO.FolderExists(FILENAME_VAR) = True Then
                ADD_STAR = "2 D"
            End If
            FOLDERNAME_1 = Mid(FILENAME_VAR, 1, InStr(FILENAME_VAR, "\") - 1)
            If FSO.FileExists(FILENAME_VAR) = False And FSO.FolderExists(FILENAME_VAR) = False And FSO.FolderExists(FOLDERNAME_1) = False Then
                ADD_STAR = "4 X"
            End If
            LISTVIEW1.ListItems.Item(r3) = (Format(AR, "000"))
            LISTVIEW1.ListItems.Item(r3).SubItems(1) = ADD_STAR
        End If
    Next
    
    
    
    If MNU_SORT_A_VAR = True Then
        LISTVIEW1.SortOrder = lvwAscending
        LISTVIEW1.SortKey = 1
        LISTVIEW1.Sorted = True
    End If
    
    AR = 0
    For r3 = 1 To LISTVIEW1.ListItems.Count
        If LISTVIEW1.ListItems.Item(r3).SubItems(2) <> "" Then
            AR = AR + 1
            LISTVIEW1.ListItems.Item(r3) = (Format(AR, "000"))
        End If
    Next
    
    LISTVIEW1.Sorted = False
    
    
    
    
    RESIZE_DONE_ONCE = True
    NOT_RESIZE_ROUTINE = True

    
'     AlwaysOnTop_MODE = True
    If AlwaysOnTop_MODE = True Then
        MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __]"
        MNU_ALWAYS_ON_TOP.Caption = "ALWAYS ON TOP __ CORRECT"
    Else
        MNU_ALWAYS_ON_TOP.Caption = "ALWAYS ON TOP __ CORRECT"
        MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __]"
    End If
    ' ------------------------------------------------
    ' SET TO INVERSE WHAT WANT AS CLICK TURN IT AROUND
    ' ------------------------------------------------
    ' Call MNU_ALWAYS_ON_TOP_Click

    Call GOT_FOCUS
    
    FORM_LOAD_STATE = True
    
    Timer_RUN_ONCE_RESIZE.Enabled = True

End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MOUSE_DOWN = True
MOUSE_UP = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MOUSE_DOWN = False
MOUSE_UP = True
End Sub


Public Sub Label_CLIPBOARD_SELECTOR_1_Click()
' ---------------------------------------------------------------
' ---------------------------------------------------------------
' MAIN ----------------------------------------------------------
' ---------------------------------------------------------------

ITEMGET_1 = LISTVIEW1.ListItems(LISTVIEW1.SelectedItem.Index).SubItems(2)

On Error Resume Next
Err.Clear
Clipboard.Clear
Err.Clear
Clipboard.SetText ITEMGET_1
If Err.Number > 0 Then
    VAR_TIMER_CLIPBOARD_TIMER_RETRY = "Label_CLIPBOARD_SELECTOR_Click"
    TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
    Exit Sub
End If

Me.Visible = False
Beep
Unload Me
Exit Sub

End Sub

Public Sub Label_CLIPBOARD_SELECTOR_2_Click()
' LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click
' ---------------------------------------------------------------
' ---------------------------------------------------------------
' MAIN ----------------------------------------------------------
' ---------------------------------------------------------------

ITEMGET_1 = PASS_ME_THE_STRING_FOR_THE_CLIPBOARD

On Error Resume Next
Err.Clear
Clipboard.Clear
Err.Clear
Clipboard.SetText ITEMGET_1
If Err.Number > 0 Then
    VAR_TIMER_CLIPBOARD_TIMER_RETRY = "Label_CLIPBOARD_SELECTOR_2_Click"
    TIMER_CLIPBOARD_TIMER_RETRY.Enabled = True
    Exit Sub
End If

Exit Sub

End Sub


Private Sub Label_EXE_Click()

' ---------------------------------------------------------------
' ---------------------------------------------------------------
' MAIN ----------------------------------------------------------
' ---------------------------------------------------------------

ITEMGET_1 = LISTVIEW1.ListItems(LISTVIEW1.SelectedItem.Index).SubItems(2)

FILENAME_VAR = ITEMGET_1
FOLDERNAME_1 = FILENAME_VAR
FOLDERNAME_2 = FILENAME_VAR
If InStr(FILENAME_VAR, "\") > 0 Then
    Do
        FOLDERNAME_2 = Mid(FOLDERNAME_2, 1, InStrRev(FOLDERNAME_2, "\") - 1)
    Loop Until InStr(FOLDERNAME_2, "\") = 0 Or FSO.FolderExists(FOLDERNAME_2) = True
End If

If FSO.FileExists(FILENAME_VAR) = True Then
    On Error Resume Next
    Err.Clear
    If UCase(Right(FILENAME_VAR, 4)) = ".EXE" Then
        Shell FILENAME_VAR, vbNormalFocus
        If Err.Number > 0 Then MsgBox "ERROR TRY TO EXECUTE FILE" + vbCrLf + vbCrLf + Err.Description + vbCrLf + vbCrLf + FILENAME_VAR, vbCritical + vbInformation + vbOKOnly
        PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = FILENAME_VAR
        LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click
    Else
        Shell "Explorer.exe " + FILENAME_VAR, vbMaximizedFocus
        PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = FILENAME_VAR
        LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click
    End If
    
    If FSO.FileExists(FILENAME_VAR) = True Then
        Shell "Explorer.exe /SELECT," + FILENAME_VAR, vbMaximizedFocus
    End If
    
    If Err.Number = 0 Then
        Beep
        Me.WindowState = vbMinimized
        Me.Visible = False
        Unload Me
    End If
    On Error GoTo 0
    Exit Sub
End If

If FSO.FolderExists(FILENAME_VAR) = True Then
    Me.Visible = False
    Shell "Explorer.exe /e," + FILENAME_VAR, vbMaximizedFocus
    PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = FILENAME_VAR
    LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click
    Beep
    Unload Me
    Exit Sub
End If

If FSO.FolderExists(FOLDERNAME_1) = True Then
    Me.Visible = False
    Shell "Explorer.exe /e," + FOLDERNAME_1, vbMaximizedFocus
    PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = FILENAME_VAR
    LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click
    PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = FOLDERNAME_1
    LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click
    
    Beep
    Unload Me
    Exit Sub
End If

If FSO.FileExists(FILENAME_VAR) = False And FSO.FolderExists(FILENAME_VAR) = False Then
    ' FOLDERNAME_1 = Mid(FILENAME_VAR, 1, InStr(FILENAME_VAR, "\") - 1)
    If FSO.FolderExists(FOLDERNAME_2) = True Then
        Me.Visible = False
        Shell "Explorer.exe /e," + FOLDERNAME_2, vbMaximizedFocus
        AA_STRING = ""
        AA_STRING = AA_STRING + "ITEM NOT EXIST" + vbCrLf
        AA_STRING = AA_STRING + "LOAD NEAREST MODIFY FOLDER INSTEAD" + vbCrLf + vbCrLf
        AA_STRING = AA_STRING + FILENAME_VAR + vbCrLf + vbCrLf
        AA_STRING = AA_STRING + FOLDERNAME_1 + vbCrLf + vbCrLf
        AA_STRING = AA_STRING + FOLDERNAME_2
        
        PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = AA_STRING
        LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click
        
        MsgBox AA_STRING, vbCritical + vbMsgBoxSetForeground + vbOKOnly
        
        Beep
        Unload Me
        Exit Sub
    End If
End If

Beep
AA_STRING = ""
AA_STRING = AA_STRING + "FILE OR FOLDER ITEM NOT EXIST" + vbCrLf
AA_STRING = AA_STRING + "NOTHING TO DO FROM SELECTION -- OR" + vbCrLf
AA_STRING = AA_STRING + "NOTHING TO DO COPY OF TEXTBOX" + vbCrLf + vbCrLf
AA_STRING = AA_STRING + FILENAME_VAR + vbCrLf + vbCrLf
AA_STRING = AA_STRING + FOLDERNAME_1 + vbCrLf + vbCrLf
AA_STRING = AA_STRING + FOLDERNAME_2

PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = AA_STRING
LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click

MsgBox AA_STRING, vbCritical + vbMsgBoxSetForeground + vbOKOnly

End Sub

Private Sub Label_FOLDER_SELECTOR_Click()
' ---------------------------------------------------------------
' ---------------------------------------------------------------
' MAIN ----------------------------------------------------------
' ---------------------------------------------------------------

ITEMGET_1 = LISTVIEW1.ListItems(LISTVIEW1.SelectedItem.Index).SubItems(2)

FILENAME_VAR = ITEMGET_1

FOLDERNAME_1 = FILENAME_VAR
FOLDERNAME_2 = FILENAME_VAR

If InStr(FILENAME_VAR, "\") > 0 Then
    Do
        FOLDERNAME_2 = Mid(FOLDERNAME_2, 1, InStrRev(FOLDERNAME_2, "\") - 1)
    Loop Until InStr(FOLDERNAME_2, "\") = 0 Or FSO.FolderExists(FOLDERNAME_2) = True
End If

'If FSO.FileExists(FILENAME_VAR) = True Then
'    Me.Visible = False
'    Shell FILENAME_VAR, vbNormalFocus
'    Beep
'    Unload Me
'    Exit Sub
'End If

If FSO.FileExists(FILENAME_VAR) = True Then
    Shell "Explorer.exe /SELECT," + FILENAME_VAR, vbMaximizedFocus
    PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = FILENAME_VAR
    Call LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click
    Beep
    Unload Me
    Exit Sub
End If

If FSO.FolderExists(FILENAME_VAR) = True Then
    Me.Visible = False
    Shell "Explorer.exe /e," + FILENAME_VAR, vbMaximizedFocus
    PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = FILENAME_VAR
    Call LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click
    Beep
    Unload Me
    Exit Sub
End If

If FSO.FolderExists(FOLDERNAME_1) = True Then
    Me.Visible = False
    Shell "Explorer.exe /e," + FOLDERNAME_1, vbMaximizedFocus
    PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = FILENAME_VAR
    Call LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click
    PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = FOLDERNAME_1
    Call LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click
    Beep
    Unload Me
    Exit Sub
End If

If FSO.FileExists(FILENAME_VAR) = False And FSO.FolderExists(FILENAME_VAR) = False Then
    ' FOLDERNAME_1 = Mid(FILENAME_VAR, 1, InStr(FILENAME_VAR, "\") - 1)
    If FSO.FolderExists(FOLDERNAME_2) = True Then
        Me.Visible = False
        Shell "Explorer.exe /e," + FOLDERNAME_2, vbMaximizedFocus
        AA_STRING = ""
        AA_STRING = AA_STRING + "ITEM NOT EXIST" + vbCrLf
        AA_STRING = AA_STRING + "LOAD NEAREST MODIFY FOLDER INSTEAD" + vbCrLf + vbCrLf
        AA_STRING = AA_STRING + FILENAME_VAR + vbCrLf + vbCrLf
        AA_STRING = AA_STRING + FOLDERNAME_1 + vbCrLf + vbCrLf
        AA_STRING = AA_STRING + FOLDERNAME_2
        
        PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = AA_STRING
        LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click
        
        MsgBox AA_STRING, vbCritical + vbMsgBoxSetForeground + vbOKOnly
        Beep
        Unload Me
        Exit Sub
    End If
End If

Beep
AA_STRING = ""
AA_STRING = AA_STRING + "FILE OR FOLDER ITEM NOT EXIST" + vbCrLf
AA_STRING = AA_STRING + "NOTHING TO DO FROM SELECTION -- OR" + vbCrLf
AA_STRING = AA_STRING + "NOTHING TO DO COPY OF TEXTBOX" + vbCrLf + vbCrLf
AA_STRING = AA_STRING + FILENAME_VAR + vbCrLf + vbCrLf
AA_STRING = AA_STRING + FOLDERNAME_1 + vbCrLf + vbCrLf
AA_STRING = AA_STRING + FOLDERNAME_2

PASS_ME_THE_STRING_FOR_THE_CLIPBOARD = AA_STRING
LINE_PICKER_EXE_SCREEN.Label_CLIPBOARD_SELECTOR_2_Click

MsgBox AA_STRING, vbCritical + vbMsgBoxSetForeground + vbOKOnly

End Sub

Private Sub Timer_RUN_ONCE_RESIZE_Timer()

' -------------------------------------------------
' 03-OCT-2026 11:45:54 SAT
' TO
' 04-OCT-2026 06:30:00 SUN
' NOW THE FORMS LOAD PROPER AND RESIZING
' -------------------------------------------------

Call frmListMenu.SET_MENU_PADD_WORK

Me.Height = Screen.Height - 4000
Me.Width = Screen.Width
DoEvents
Me.Top = (Screen.Height / 2 - Me.Height / 2) - 1000
Me.Left = 0 'Screen.Width / 2 - Me.Width / 2 - 1000

RESIZE_DONE_ONCE = False
NOT_RESIZE_ROUTINE = False
' CALL UPON ONE REIZE RUN ROUTINE
' TO FINAL SET
' -------------------------------

'IF ME.WindowState
'Call Form_Resize
Timer_RUN_ONCE_RESIZE.Enabled = False
RESIZE_DONE_ONCE = True
NOT_RESIZE_ROUTINE = True

End Sub


Private Sub Form_Resize()

'If NOT_RESIZE_ROUTINE = True Then
'    ' NOT_RESIZE_ROUTINE = False
'    Exit Sub
'End If

If Me.Visible = False Then Exit Sub

Me.WindowState = vbNormal
Me.Visible = True
' DoEvents
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
'I2 = 0: KASC = 0
'WIN 10 IS GIVE -- GetAsyncKeyState(255) ALWAYS PRESSED
'For I2 = 1 To 254
'    BDF = GetAsyncKeyState(I2)
'    If BDF < 0 Then
'        X = X
'    End If
'    If BDF < -300 Then
'        KCODE = Now
'        ' I2 -- 1 IS MOUSE LEFT BUTTON
'    End If
'Next

If GetAsyncKeyState(1) < 0 Then
    MOUSE_UP = False
    MOUSE_DOWN = True
Else
    MOUSE_UP = True
    MOUSE_DOWN = False
End If

'If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
'    VAR_WINSTATE_BEEN_HIGH = Me.WindowState ' vbMaximized
'End If
If Me.WindowState = vbMinimized Then

'    VAR_WINSTATE_BEEN_HIGH = "UNLOAD ME"
    
    ' IT UNLAOD ME AFTER A SELECTION -- UNLESS -- OPTION NOT MOVE
    ' -----------------------------------------------------------
    Unload Me
    
    Exit Sub
End If

If RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True Then
    If TIMER_TO_RESIZE_VAR <> True Then
        Exit Sub
    End If
End If
If Me.WindowState <> RESIZE_WINDOWSTATE_CHANGE Then
    RESIZE_DONE_ONCE = False
End If
If RESIZE_WINDOWSTATE_CHANGE = vbMinimized And (Me.WindowState = vbMaximized Or Me.WindowState = vbNormal) Then
    ' REQUIRE BETTER ABOUT ERROR TRAP 01 OF 02
    On Error Resume Next
    LINE_PICKER_EXE_SCREEN.SetFocus
    Timer_SET_FOCUS.Enabled = True
    On Error GoTo 0
    DoEvents
End If
RESIZE_WINDOWSTATE_CHANGE = Me.WindowState

If RESIZE_DONE_ONCE = True Then
    If TIMER_TO_RESIZE_VAR <> True Then
        Exit Sub
    End If
End If

' Me.Visible = True

If Me.WindowState = vbNormal Then
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
    If TIMER_TO_RESIZE_VAR = False Then
'        Me.Top = Screen.Height / 2 - Me.Height / 2
'        Me.Left = Screen.Width / 2 - Me.Width / 2
        If FORM_SET_WIDTH_ONE = False Then
            FORM_SET_WIDTH_ONE = True
            'Me.Width = Screen.Width - 1000
            'SET HIEGHT ONCE ALSO
            'Me.Height = FRM_ClipTest.Height
        End If

    End If
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
End If

If Me.WindowState = vbMaximized Then
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
    If TIMER_TO_RESIZE_VAR = False Then
'        Me.Top = FRM_ClipTest.Top + 200
'        Me.Left = FRM_ClipTest.Left + 200
        If FORM_SET_WIDTH_ONE = False Then
            FORM_SET_WIDTH_ONE = True
'            Me.Width = FRM_ClipTest.Width
'            ' SET HIEGHT ONCE ALSO
'            Me.Height = FRM_ClipTest.Height
        End If

    End If
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
End If

HEIGHT_ADJUST = 250
WIDTH_ADJUST = 250

If GETWINNT_VER_VAR = "WIN XP" Then WIDTH_ADJUST = 250
If GETWINNT_VER_VAR = "WIN 7" Then WIDTH_ADJUST = 250
If GETWINNT_VER_VAR = "WIN 10" Then WIDTH_ADJUST = 250
If GETWINNT_VER_VAR = "WIN 11" Then WIDTH_ADJUST = 250

'HEIGHT_ADJUST = 70 ' FOR WIN XP
If GETWINNT_VER_VAR = "WIN XP" Then HEIGHT_ADJUST = 1500
If GETWINNT_VER_VAR = "WIN 7" Then HEIGHT_ADJUST = 1500
If GETWINNT_VER_VAR = "WIN 10" Then HEIGHT_ADJUST = 1500
If GETWINNT_VER_VAR = "WIN 11" Then HEIGHT_ADJUST = 1500

''Bigger adjust number smaller inner form
'If GetComputerName = "1-ASUS-EEE" Then HEIGHT_ADJUST = 70
'If GetComputerName = "1-ASUS-X5DIJ" Then HEIGHT_ADJUST = 70
'If GetComputerName = "3-LINDA-PC" Then HEIGHT_ADJUST = 70
'If GetComputerName = "4-ASUS-GW522VW" Then HEIGHT_ADJUST = 70
'If GetComputerName = "5-ASUS-P2520LA" Then HEIGHT_ADJUST = 0
''HIGHER NUMBER SMALLER INNER BOX

If RESIZE_DONE_ONCE = False Then
    Label1.BackColor = RGB(255, 255, 255)
    Label1.Left = 20
    Label1.Height = 400
    Label1.Caption = "SPEED EXE/FOLDER/URL PICKER....1+=FILE EXIST...2D=FOLDER....3-=FILE NOT EXIST BUT FOLDER DOES....4X=FILE AND FOLDER NOT EXIST"
'    Label1.FontSize = 14
'    Label1.FontBold = True
'    Label1.FontName = "Arial"
End If

If RESIZE_DONE_ONCE = False Then
    ' Label1.Width = Me.Width - WIDTH_ADJUST + 20
    Label1.AutoSize = True
    Label_EXE.AutoSize = True
    Label_FOLDER_SELECTOR.AutoSize = True
    Label_CLIPBOARD_SELECTOR_1.AutoSize = True
    
    Label1.AutoSize = True
    Label_EXE.AutoSize = True
    Label_FOLDER_SELECTOR.AutoSize = True
    Label_CLIPBOARD_SELECTOR_1.AutoSize = True
    DoEvents
    Label_EXE.Left = 10
    Label_FOLDER_SELECTOR.Left = Label_EXE.Left + Label_EXE.Width + 80
    Label_CLIPBOARD_SELECTOR_1.Left = Label_FOLDER_SELECTOR.Left + Label_FOLDER_SELECTOR.Width + 80
    
    'Label1.Width = 4000 ' Me.Width - WIDTH_ADJUST + 20
    LB1H = (Menu_Height_EXE * Screen.TwipsPerPixelY) - 500
    If LB1H < 0 Then LB1H = 0
    Label1.Top = 20 ' LB1H
    Label_EXE.Top = Label1.Top + Label1.Height
    Label_FOLDER_SELECTOR.Top = Label_EXE.Top
    Label_CLIPBOARD_SELECTOR_1.Top = Label_EXE.Top
    LIST1.Top = Label_EXE.Top + Label_EXE.Top
    LIST1.Left = 10
    LIST1.Visible = False
    LIST2.Visible = False
    
End If

'LIST1.Height = Me.Height - (Menu_Height_EXE * Screen.TwipsPerPixelY) - 600
'LIST1.Width = Me.Width - 4000 'WIDTH_ADJUST - 2000
'DoEvents

' Call HEIGHT_BAR_TOP_ROUTINE
' ---------------------------
'MAKE FORM TALLER OR TEXT BOX
'FORM IS PRIOITY OVER TEXT BOX
If Me.WindowState <> vbMaximized Then
If Me.WindowState <> vbMinimized Then
    XXDD = Me.Height - (Menu_Height_EXE * Screen.TwipsPerPixelY) - (Label1.Top + Label1.Height) - 0
End If
If XXDD - HEIGHT_ADJUST < 0 Then
    XXDD = 0 + HEIGHT_ADJUST + 100
End If

HEIGHT_ADJUST_22 = Me.Height - HEIGHT_ADJUST - Menu_Height
If OLD_LIST1_HEIGHT <> HEIGHT_ADJUST_22 Then
    LISTVIEW1.Height = HEIGHT_ADJUST_22
End If
OLD_LIST1_HEIGHT = HEIGHT_ADJUST_22

WIDTH_ADJUST_VAR = 200
If OLD_LIST1_WIDTH <> Me.Width - WIDTH_ADJUST_VAR Then
    LISTVIEW1.Width = Me.Width - WIDTH_ADJUST_VAR
End If
OLD_LIST1_WIDTH = Me.Width - WIDTH_ADJUST_VAR
End If


'H1 = Me.Height - (List1.Top + 800 + 20)
'If H1 > 0 Then
'    List1.Height = H1
'    Else
'    List1.Height = Me.Height
'End If

If RESIZE_DONE_ONCE = False Then
    If Me.WindowState = vbNormal Then
        RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
        DoEvents
        Me.Height = LISTVIEW1.Height + LISTVIEW1.Top - 200 ' THIS LINE TAKE PRIORTY THAN ABOVE LINE
        RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
    End If
    If Me.WindowState = vbMaximized Then
        ' RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
        ' Me.Height = List1.Height + List1.Top + 1200 ' THIS LINE TAKE PRIORTY THAN ABOVE LINE
        ' RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
    End If
End If



'If RESIZE_DONE_ONCE = False Then
'    With LISTVIEW1
'        .Height = LIST1.Height
'        .Top = LIST1.Top
'        .Left = LIST1.Left
'        .Width = LIST1.Width
'    End With
'End If

RESIZE_DONE_ONCE = True
TIMER_TO_RESIZE_VAR = False


End Sub





Private Sub Form_Unload(Cancel As Integer)




'    Exit Sub

    FORM_UNLOAD_INITIATE = True
    WORKER_WITH_PICKER = False
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
    
    LINE_PICKER_EXE_SCREEN.WindowState = vbMinimized
    NOT_RESIZE_ROUTINE = True
    RESULT_API = NotAlwaysOnTop(Me.hWnd)
    ' WHEN LINE_PICKER_EXE_SCREEN.GONE HERE MINIMIZE TOO
    NOT_RESIZE_ROUTINE = True
    FRM_ClipTest.WindowState = vbMinimized
    
    LINE_PICKER_EXE_SCREEN.Visible = False
    If WORKER_WITH_PICKER = True Then
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

'Public Sub ME_POSITION()
'    ' LINE_PICKER_EXE_SCREEN.ME_POSITION
'    Dim OFFSET_WX, OFFSET_HY
'
'    OFFSET_WX = 180
'    OFFSET_HY = -280
'    HY = HY + ((Menu_Height * Screen.TwipsPerPixelY) - Menu_Height) + OFFSET_HY
'    WX = WX + (Screen.TwipsPerPixelX) + OFFSET_WX
'    ' On Error Resume Next
'    ' WHEN CALL BY HALIFAX ROUTINE
'    ' A FORM NOT POSITION MOVE WHEN MINIMIZE MAXIMIZE
'    Me.Width = Text3.Left + Text3.Width + 200  ' WX + Text3.Left + Text3.Width + 200
'    On Error GoTo 0
'    WX = ((Screen.Width) / 2 - ((Me.Width / 2)))
'    Me.Left = WX
'    Me.Top = 200
'    Me.Visible = True
'
'End Sub


Private Sub List1_Click()

ITEMGET_1 = LIST1.List(LIST1.ListIndex)
ITEMGET_1 = Mid(ITEMGET, InStr(ITEMGET, " -- ") + 5)

If MNU_DOUBLE_LINE_SELECTOR_ON_VAR = True Then
    ITEMGET_1 = LIST1.List(LIST1.ListIndex)
    ITEMGET_1 = Mid(ITEMGET_1, InStr(ITEMGET_1, " -- ") + 5)

    ITEMGET_2 = LIST1.List(LIST1.ListIndex + 1)
    ITEMGET_2 = Mid(ITEMGET_2, InStr(ITEMGET_2, " -- ") + 5)

    ITEMGET_1 = ITEMGET_1 + vbCrLf + ITEMGET_2
End If

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
    WORKER_WITH_PICKER = True
    If WORKER_WITH_PICKER = True Then
        FRM_ClipTest.Visible = False
        Else
        FRM_ClipTest.Visible = True
    End If

    If AlwaysOnTop_MODE = False Then
        FRM_ClipTest.WindowState = vbMinimized
        LINE_PICKER_EXE_SCREEN.WindowState = vbMinimized
        
        RESULT_API = NotAlwaysOnTop(Me.hWnd)
    End If
End If

Beep
End Sub


Private Sub ListView1_DBLClick()


' ---------------------------------------------------------------
' ---------------------------------------------------------------
' MAIN ----------------------------------------------------------
' ---------------------------------------------------------------

Call Label_EXE_Click

'FILENAME_VAR = ITEMGET_1
'FOLDERNAME_1 = FILENAME_VAR
'If InStr(FILENAME_VAR, "\") > 0 Then
'    Do
'        FOLDERNAME_1 = Mid(FOLDERNAME_1, 1, InStrRev(FOLDERNAME_1, "\") - 1)
'    Loop Until InStr(FOLDERNAME_1, "\") = 0 Or FSO.FolderExists(FOLDERNAME_1) = True
'End If
'
'If FSO.FileExists(FILENAME_VAR) = True Then
'    Me.Visible = False
'    Shell FILENAME_VAR, vbNormalFocus
'    Beep
'    Unload Me
'    Exit Sub
'End If
'
'If FSO.FolderExists(FILENAME_VAR) = True Then
'    Me.Visible = False
'    Shell "Explorer.exe /e," + FOLDERNAME_1, vbMaximizedFocus
'    Beep
'    Unload Me
'    Exit Sub
'End If
'
'If FSO.FileExists(FILENAME_VAR) = False And FSO.FolderExists(FILENAME_VAR) = False Then
'    ' FOLDERNAME_1 = Mid(FILENAME_VAR, 1, InStr(FILENAME_VAR, "\") - 1)
'    If FSO.FolderExists(FOLDERNAME_1) = True Then
'        Me.Visible = False
'        Shell "Explorer.exe /e," + FOLDERNAME_1, vbMaximizedFocus
'        MsgBox "ITEM NOT EXIST" + vbCrLf + "LOAD NEAREST MODIFY FOLDER INSTEAD" + vbCrLf + vbCrLf + FILENAME_VAR + vbCrLf + vbCrLf + FOLDERNAME_1, vbCritical + vbMsgBoxSetForeground + vbOKOnly
'        Beep
'        Unload Me
'        Exit Sub
'    End If
'End If
'
'' ---------------------------------------------------------------
'' ---------------------------------------------------------------
'' ---------------------------------------------------------------
'' ---------------------------------------------------------------
'
'
'Beep
'If Dir(ITEMGET_1) <> "" Then
'    Unload Me
'End If
'
'
'
'
'Exit Sub
'
'
'
'
'
'
''VAR_ST_1 = ITEMGET_1
''Do
''    On Error Resume Next
''    Clipboard.Clear
''    VAR_ST_2 = Clipboard.GetText
''    On Error GoTo 0
''    If VAR_ST_2 <> "" Then
''        DoEvents
''        Sleep 500
''    End If
''Loop Until VAR_ST_2 = ""
''
''Sleep 100
''Do
''    On Error Resume Next
''    Clipboard.SetText VAR_ST_1, vbCFText
''    Sleep 100
''    VAR_ST_2 = Clipboard.GetText
''    On Error GoTo 0
''    If VAR_ST_2 <> VAR_ST_1 Then
''        DoEvents
''        Sleep 500
''    End If
''Loop Until VAR_ST_2 = VAR_ST_1
'
'If Err.Number = 0 Then
'    WORKER_WITH_PICKER = True
'    If AlwaysOnTop_MODE = False Then
'    If WORKER_WITH_PICKER = True Then
'        FRM_ClipTest.Visible = False
'    Else
'        FRM_ClipTest.Visible = True
'    End If
'    End If
'
'    SELECT_SHOW = LISTVIEW1.SelectedItem.Index - 10
'    If SELECT_SHOW < 1 Then SELECT_SHOW = 1
'    If SELECT_SHOW < 15 Then
'        LISTVIEW1.ListItems(1).EnsureVisible
'    End If
'    LISTVIEW1.ListItems(SELECT_SHOW).EnsureVisible
'
'    If AlwaysOnTop_MODE = False Then
'        FRM_ClipTest.WindowState = vbMinimized
'        LINE_PICKER_EXE_SCREEN.WindowState = vbMinimized
'        Unload Me
'    End If
'    If AlwaysOnTop_MODE = True Then
'        If SELECT_SHOW < 10 Then
'            LISTVIEW1.ListItems(1).EnsureVisible
'        End If
'        LISTVIEW1.ListItems(SELECT_SHOW).EnsureVisible
'
'        ' RESULT_API = NotAlwaysOnTop(Me.hWnd)
'    End If
'End If
'
'

End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 And IsIDE = True Then End

End Sub

Private Sub MNU_ADD_CLIPBOARD_Click()

    
    MX = Split(FRM_ClipTest.Text2.Text, vbLf)
    
    AR = LIST1.ListCount
    For r3 = 0 To UBound(MX)
        If MX(r3) <> "" Then
            AR = AR + 1
            LIST1.AddItem Format(AR, "000") + " --  " + MX(r3)
            With LISTVIEW1
                Set LV1 = .ListItems.Add(, , Format(AR, "000"))
                LV1.SubItems(1) = MX(r3)
            End With
        End If
    Next
    
    FILENAME_LOAD_01 = App.Path + "\# DATA\VB__" + UCase(App.EXEName) + "_LINE_PICKER_EXE_SCREEN.TXT"
    If Dir(FILENAME_LOAD_01) <> "" Then
        FR1 = FreeFile
        Open FILENAME_LOAD_01 For Append As #FR1
            Print #FR1, "---- CLIPBOARD INFO -- " + Format(Now, "YYYY MM DD  HH MM SS")
            Print #FR1, "-------------------------------------------"
            For r3 = 0 To UBound(MX)
                If MX(r3) <> "" Then
                    Print #FR1, MX(r3)
                End If
            Next
            Print #FR1, "-------------------------------------------"
        Close #FR1
    End If
    
    
    If MNU_SORT_A_VAR = True Then
        LISTVIEW1.SortOrder = lvwAscending
        LISTVIEW1.SortKey = 1
        LISTVIEW1.Sorted = True
    End If

    LISTVIEW1.ListItems(LISTVIEW1.ListItems.Count).EnsureVisible

End Sub

Private Sub MNU_ALWAYS_ON_TOP_Click()
    
    ' MENU FLIPPER
    If InStr(MNU_ALWAYS_ON_TOP.Caption, "ON TOP __ CORRECT") > 0 Then
        MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __]"
        AlwaysOnTop_MODE = False
        RESULT_API = NotAlwaysOnTop(Me.hWnd)
    Else
        MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __ CORRECT __]"
        AlwaysOnTop_MODE = True
        RESULT_API = AlwaysOnTop(Me.hWnd)
    End If
    Me.SetFocus

End Sub

Public Sub MNU_ALWAYS_ON_TOP_SET()
    
    ' SET CORRECT AND RESULT REQUIRE THAT
    '
    ' DO WHAT MENU GOT
    '
    If InStr(MNU_ALWAYS_ON_TOP.Caption, "ON TOP __ CORRECT") > 0 Then
        AlwaysOnTop_MODE = True
        RESULT_API = AlwaysOnTop(Me.hWnd)
    Else
        AlwaysOnTop_MODE = False
        RESULT_API = NotAlwaysOnTop(Me.hWnd)
    End If

End Sub

Public Sub MNU_ALWAYS_ON_TOP_SET_TRUE()
    
    ' FORCE ON
    MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __ CORRECT __]"
    AlwaysOnTop_MODE = True
    RESULT_API = AlwaysOnTop(Me.hWnd)

End Sub


Private Sub MNU_BUTTON_01_REMOVE_EMPTY_LINE_AND_REM_02_SORT_03_REMOVE_DUPER_Click()

    MX = Split(FRM_ClipTest.Text2.Text, vbLf)
    
    LIST2.Clear
    For r3 = 0 To UBound(MX)
        MX(r3) = Replace(MX(r3), vbCr, "")
        MX(r3) = Replace(MX(r3), vbLf, "")
        If Mid(Trim(MX(r3)), 1, 1) <> "!" Then
        If Trim(MX(r3)) <> "" Then
            LIST2.AddItem MX(r3)
        End If
        End If
    Next

    ' LIST2.Sorted = True
    LIST2.Refresh
    DoEvents
    
    'DUPE COPY REMOVED
    For R1 = LIST2.ListCount - 1 To 1 Step -1
        If LIST2.List(R1) = LIST2.List(R1 + 1) Then LIST2.RemoveItem (R1)
    Next

    'GATHER TO STRING
    STING_GATHER = ""
    For R1 = 1 To LIST2.ListCount - 1
        STING_GATHER = STING_GATHER + LIST2.List(R1) + vbCrLf
    Next

    ' TO CLIPBOARD

    

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
VAR_ST_1 = STING_GATHER
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
    Call NotAlwaysOnTop(LINE_PICKER_EXE_SCREEN.hWnd)
    AlwaysOnTop_MODE = False
    Unload Me
End If

Beep

End Sub

Private Sub MNU_CLIPBOARD_ALL_COMPUTER_DATA_FOLDER_BLOCK_FOR_GOODSYNC_Click()

Dim MSI

MSI = MSI + "name *1-ASUS-*; name *" + vbCrLf
MSI = MSI + "name *2-ASUS-*; name *" + vbCrLf
MSI = MSI + "name *1-ASUS-*; name *" + vbCrLf
MSI = MSI + "name *3-LINDA-*; name *" + vbCrLf
MSI = MSI + "name *4-ASUS-*; name *" + vbCrLf
MSI = MSI + "name *5-ASUS-*; name *" + vbCrLf
MSI = MSI + "name *7-ASUS-*; name *" + vbCrLf
MSI = MSI + "name *8-MSI-*; name *" + vbCrLf
MSI = MSI + "name *9-ASUS-*; name *" + vbCrLf
MSI = MSI + "name *55-88-HAPPY*; name *" + vbCrLf
MSI = MSI + "name *MATT-555ROIDS*; name *" + vbCrLf
MSI = MSI + "name *ASUS-X5DIJ*; name *" + vbCrLf
MSI = MSI + "name *LINDA-PC*; name *" + vbCrLf
MSI = MSI + "name *MATT-555ROIDS*; name *" + vbCrLf
MSI = MSI + "name *MININT-0G58L6G*; name *" + vbCrLf
MSI = MSI + "name *MININT-82SO8CF*; name *" + vbCrLf
MSI = MSI + "name *MININT-GNNFCHO*; name *" + vbCrLf
MSI = MSI + "name *MININT-HLFAN8N*; name *" + vbCrLf
MSI = MSI + "name *55-88-HAPPY*; name *" + vbCrLf
MSI = MSI + "name *ASUS-4-GL522VW*; name *" + vbCrLf
MSI = MSI + "name *ASUS-BIGGER*; name *" + vbCrLf
MSI = MSI + "name *ASUSEEE*; name *" + vbCrLf
MSI = MSI + "name *ASUS-EEE*; name *" + vbCrLf
MSI = MSI + "name *ASUS-EEE-XPHOME*; name *" + vbCrLf
MSI = MSI + "name *ASUS-MACHINE*; name *" + vbCrLf
MSI = MSI + "name *#NFS*" + vbCrLf
MSI = MSI + "name name *#NFS*; name *" + vbCrLf
MSI = MSI + "name 0TextData" + vbCrLf
MSI = MSI + "name # DATA" + vbCrLf
MSI = MSI + "name Day-Data" + vbCrLf
MSI = MSI + "name # DATA" + vbCrLf
MSI = MSI + "name Text Loggs" + vbCrLf
MSI = MSI + "name 00_Text_Data_02" + vbCrLf
MSI = MSI + "name 00_Text_Data_KeyLogg" + vbCrLf
MSI = MSI + "name # DATA" + vbCrLf


VAR_ST_1 = MSI

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
    WORKER_WITH_PICKER = True
    If WORKER_WITH_PICKER = True Then
        FRM_ClipTest.Visible = False
        Else
        FRM_ClipTest.Visible = True
    End If

    If AlwaysOnTop_MODE = False Then
        FRM_ClipTest.WindowState = vbMinimized
        LINE_PICKER_EXE_SCREEN.WindowState = vbMinimized
        
        RESULT_API = NotAlwaysOnTop(Me.hWnd)
    End If
End If

Beep


End Sub

Private Sub MNU_DOUBLE_LINE_SELECTOR_ON_Click()
MNU_DOUBLE_LINE_SELECTOR_ON_VAR = Not MNU_DOUBLE_LINE_SELECTOR_ON_VAR

If MNU_DOUBLE_LINE_SELECTOR_ON_VAR = True Then
    MNU_DOUBLE_LINE_SELECTOR_ON.Caption = "DOUBLE LINE SELECTOR ON"
Else
    MNU_DOUBLE_LINE_SELECTOR_ON.Caption = "DOUBLE LINE SELECTOR OFF"
End If


End Sub

Private Sub MNU_EDITOR_NOTEPAD_01_Click()

APP_NAME = "C:\Program Files (x86)\Notepad++\notepad++.exe"
If Dir(APP_NAME) = "" Then
    APP_NAME = "C:\Program Files\Notepad++\notepad++.exe"
End If
Me.WindowState = vbMinimized

Shell APP_NAME + " """ + FILENAME_LOAD_01 + """", vbMaximizedFocus

End Sub

Private Sub MNU_EDITOR_NOTEPAD_02_Click()

APP_NAME = "C:\Program Files (x86)\Notepad++\notepad++.exe"
If Dir(APP_NAME) = "" Then
    APP_NAME = "C:\Program Files\Notepad++\notepad++.exe"
End If
Me.WindowState = vbMinimized

Shell APP_NAME + " """ + FILENAME_LOAD_02 + """", vbMaximizedFocus

End Sub

Private Sub MNU_ENDER_Click()
End
End Sub

Public Sub MNU_HALIFAX_Click()
'     LINE_PICKER_EXE_SCREEN.MNU_HALIFAX_Click
    FIND_IT = 0
    For R_1 = 1 To LISTVIEW1.ListItems.Count
        
        ITEMGET_2 = LISTVIEW1.ListItems(R_1).SubItems(1)
        If InStr(ITEMGET_2, "HALIFAX") > 0 Then
            FIND_IT = R_1
            Exit For
        End If
    Next

    If FIND_IT > 0 Then
        LINE_PICKER_FORM_LIST_CLICK_VALUE = FIND_IT - 2
        
        NOT__10 = LISTVIEW1.ListItems.Count
        LISTVIEW1.ListItems(NOT__10).EnsureVisible
        LISTVIEW1.ListItems(LINE_PICKER_FORM_LIST_CLICK_VALUE).EnsureVisible
        LISTVIEW1.ListItems(LINE_PICKER_FORM_LIST_CLICK_VALUE).Selected = True
        LISTVIEW1.SetFocus
    End If

End Sub

Private Sub MNU_JUMP_LINE_ADBLOCK_Click()

    ' ---------------------------------------------------------------------
    ' ! CLIPBOARD CHUNK BELOW -- ADBLOCK FOR CHROME
    ' ---------------------------------------------------------------------
    ' ENDER LINE
    ' ---------------------------------------------------------------------
    ' ! END CHUNK -- CLIPBOARD CHUNK BELOW -- ADBLOCK.COM CHROME
    ' ---------------------------------------------------------------------
    FIND_IT = 0
    For R_1 = 1 To LISTVIEW1.ListItems.Count
        
        ITEMGET_2 = LISTVIEW1.ListItems(R_1).SubItems(1)
        If InStr(ITEMGET_2, "! CLIPBOARD CHUNK BELOW -- ADBLOCK") > 0 Then
            FIND_IT = R_1
            Exit For
        End If
    Next

    If FIND_IT > 0 Then
        LINE_PICKER_FORM_LIST_CLICK_VALUE = FIND_IT
        PLUS_10 = LINE_PICKER_FORM_LIST_CLICK_VALUE + 10
        NOT__10 = LINE_PICKER_FORM_LIST_CLICK_VALUE - 10
        LIST_COUNT = LISTVIEW1.ListItems.Count
        If NOT__10 < 1 Then NOT__1 = 1
        If PLUS_10 > LIST_COUNT Then PLUS_1 = LIST_COUNT
        LISTVIEW1.ListItems(NOT__10).EnsureVisible
        LISTVIEW1.ListItems(PLUS_10).EnsureVisible
        LISTVIEW1.ListItems(LINE_PICKER_FORM_LIST_CLICK_VALUE).EnsureVisible
        LISTVIEW1.ListItems(LINE_PICKER_FORM_LIST_CLICK_VALUE).Selected = True
        LISTVIEW1.SetFocus
    End If

End Sub


Private Sub MNU_UMP_DRIVE_NETWORK_Click()

    FIND_IT = 0
    For R_1 = 1 To LISTVIEW1.ListItems.Count
        
        ITEMGET_2 = LISTVIEW1.ListItems(R_1).SubItems(1)
        If InStr(ITEMGET_2, "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01") > 0 Then
            FIND_IT = R_1
            Exit For
        End If
    Next

    If FIND_IT > 0 Then
        LINE_PICKER_FORM_LIST_CLICK_VALUE = FIND_IT
        PLUS_10 = LINE_PICKER_FORM_LIST_CLICK_VALUE + 10
        NOT__10 = LINE_PICKER_FORM_LIST_CLICK_VALUE - 10
        LIST_COUNT = LISTVIEW1.ListItems.Count
        If NOT__10 < 1 Then NOT__1 = 1
        If PLUS_10 > LIST_COUNT Then PLUS_1 = LIST_COUNT
        LISTVIEW1.ListItems(NOT__10).EnsureVisible
        LISTVIEW1.ListItems(PLUS_10).EnsureVisible
        LISTVIEW1.ListItems(LINE_PICKER_FORM_LIST_CLICK_VALUE).EnsureVisible
        LISTVIEW1.ListItems(LINE_PICKER_FORM_LIST_CLICK_VALUE).Selected = True
        LISTVIEW1.SetFocus
    End If

End Sub


Private Sub MNU_MINIMIZE_MAIN_FORM_Click()
FRM_ClipTest.WindowState = vbMinimized
End Sub

Private Sub MNU_SHOW_MAIN_FORM_Click()

    Unload Me
    
    FRM_ClipTest.Show
    FRM_ClipTest.WindowState = vbNormal
End Sub

Private Sub MNU_SORT_A_Click()

MNU_SORT_A_VAR = Not MNU_SORT_A_VAR
If MNU_SORT_A_VAR = True Then
    LISTVIEW1.SortOrder = lvwAscending
    LISTVIEW1.SortKey = 1
    LISTVIEW1.Sorted = True
End If
If MNU_SORT_A_VAR = False Then

End If


End Sub
Public Sub MNU_UNLOAD_FORM_Click()
    
'    MNU_ALWAYS_ON_TOP.Caption = "ALWAYS ON TOP __ FALSE" ' --AlwaysOnTop_MODE
'    AlwaysOnTop_MODE = False

    ' Call NotAlwaysOnTop(Me.hWnd)
    ' NOT_RESIZE_ROUTINE = True
    
    Call NotAlwaysOnTop(LINE_PICKER_EXE_SCREEN.hWnd)
    Unload Me

End Sub

Private Sub Timer_MONITOR_FORM_SIZE_WIDTH_HEIGHT_Timer()

    If Me.WindowState = vbMinimized And VAR_WINSTATE_BEEN_HIGH = "UNLOAD ME" Then
        VAR_WINSTATE_BEEN_HIGH = ""
        Unload Me
        Exit Sub
    End If
    
    If GetAsyncKeyState(1) < 0 Then
        MOUSE_UP = False
        MOUSE_DOWN = True
    Else
        MOUSE_UP = True
        MOUSE_DOWN = False
    End If

    ' If MOUSE_DOWN = True Then Stop
    If MOUSE_DOWN = True Then Exit Sub
    
    If OLD_HEIGHT_WIDTH = 0 Then
        OLD_HEIGHT_WIDTH = Me.Height + Me.Width
    End If
    If Me.Height + Me.Width <> OLD_HEIGHT_WIDTH Then
        OLD_HEIGHT_WIDTH = Me.Height + Me.Width
        TIMER_TO_RESIZE.Enabled = False
        TIMER_TO_RESIZE.Interval = 100
        TIMER_TO_RESIZE.Enabled = True
        ' TIMER_TO_RESIZE_VAR = True ' USER WITHIN ---- TIMER_TO_RESIZE
        ' RESIZE_DONE_ONCE = False
        ' Exit Sub
    End If

End Sub


Public Sub Timer_SET_FOCUS_Timer()

' REQUIRE BETTER ABOUT ERROR TRAP 02 OF 02
On Error Resume Next
    Me.SetFocus
    Timer_SET_FOCUS.Enabled = False
On Error GoTo 0

'
'If WORKER_WITH_PICKER = True Then
'    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
'    FRM_ClipTest.Visible = True
'    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
'End If

End Sub


Public Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    'If Me.Height < 2000 Then
    '    Call ME_POSITION
    'End If
    SetWindowPos hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Public Function NotAlwaysOnTop(ByVal hWnd As Long)
    ' VAR WILL FLIPPER FALSE WHEN -- SETWINDOWPOS HWND_NOTOPMOST -- GO THROUGH IT ROUTINE
    ' HWND_TOPMOST NOT REQUIRE NOT OPERATE RESIZE
    ' -----------------------------------------------------------------------------------
'    If Me.WindowState <> vbMinimized Then
'    End If
'    If Me.WindowState = vbMinimized Then
    If Me.Height < 200 Then
        Call ME_POSITION
    End If
    
    NOT_RESIZE_ROUTINE = True
    SetWindowPos hWnd, hWnd_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function


Private Function Menu_Height()
 
    MenuInfo.cbSize = Len(MenuInfo)
    
    If GetMenuBarInfo(Me.hWnd, OBJID_MENU, 0, MenuInfo) Then
   
        With MenuInfo.rcBar
       
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            Menu_Height = CStr(.Bottom) - CStr(.Top)

        End With
       
    End If
    
    If Menu_Height <> ARCHIVE_Menu_Height Then
        TIMER_TO_RESIZE.Enabled = False
        TIMER_TO_RESIZE.Interval = 100
        TIMER_TO_RESIZE.Enabled = True
    End If
    
    ARCHIVE_Menu_Height = Menu_Height
   
End Function

Private Sub TIMER_TO_RESIZE_Timer()
    TIMER_TO_RESIZE.Enabled = False
    ' RESIZE_DONE_ONCE = False
    TIMER_TO_RESIZE_VAR = True
    'NOT_RESIZE_EVENTER = False
    Call Form_Resize
    'NOT_RESIZE_EVENTER = True

End Sub

Public Sub ME_POSITION()

    Dim OFFSET_WX, OFFSET_HY

    OFFSET_WX = 180
    OFFSET_HY = -280
    HY = HY + ((Menu_Height * Screen.TwipsPerPixelY) - Menu_Height) + OFFSET_HY
    WX = WX + (Screen.TwipsPerPixelX) + OFFSET_WX

    If HY <> Me.Height Then
        Me.Height = HY
    End If
    Me.Width = WX ' LISTVIEW1.Left + LISTVIEW1.Width + 200 ' WX (+) LISTVIEW1.Left + LISTVIEW1.Width + 200
    WL = ((Screen.Width) / 2 - ((Me.Width / 2)))
    Me.Left = WL
    Me.Top = 200

    Me.Visible = True

End Sub


Private Function Menu_Height_EXE()
 
    If Me.Visible = True Then Exit Function
    If Me.WindowState <> 0 Then
    End If
    MenuInfo.cbSize = Len(MenuInfo)
    
    If GetMenuBarInfo(Me.hWnd, OBJID_MENU, 0, MenuInfo) Then
   
        With MenuInfo.rcBar
       
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            Menu_Height = CStr(.Bottom) - CStr(.Top)

        End With
       
    End If
    
    If Menu_Height <> ARCHIVE_Menu_Height Then
        TIMER_TO_RESIZE.Enabled = True
    End If
    
    ARCHIVE_Menu_Height = Menu_Height
   
End Function


