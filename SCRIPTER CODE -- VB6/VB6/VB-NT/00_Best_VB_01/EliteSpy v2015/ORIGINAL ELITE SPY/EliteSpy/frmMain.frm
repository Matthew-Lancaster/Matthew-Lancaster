VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EliteSpy+ by Andrea Batina"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCodeGeneration 
      Caption         =   "Show Code Generation Wizard"
      Height          =   315
      Left            =   60
      TabIndex        =   44
      Top             =   3060
      Width           =   3255
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy to Clipboard"
      Height          =   315
      Left            =   3720
      TabIndex        =   43
      Top             =   1380
      Width           =   1515
   End
   Begin VB.TextBox txtRect 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   1560
      Width           =   2115
   End
   Begin VB.TextBox txtParentClass 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   2640
      Width           =   2115
   End
   Begin VB.CommandButton cmdMemInfo 
      Caption         =   "Memory Info"
      Height          =   315
      Left            =   4080
      TabIndex        =   38
      Top             =   2340
      Width           =   1155
   End
   Begin VB.CheckBox chkOnTop 
      Appearance      =   0  'Flat
      Caption         =   "Always On Top"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1740
      TabIndex        =   37
      Top             =   5520
      Width           =   1395
   End
   Begin VB.TextBox txtParentText 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   2280
      Width           =   2115
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process ->"
      Height          =   315
      Left            =   4200
      TabIndex        =   34
      Top             =   5460
      Width           =   1155
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   315
      Left            =   7920
      TabIndex        =   33
      Top             =   5460
      Width           =   1155
   End
   Begin VB.CommandButton cmdTerminateProcess 
      Caption         =   "Terminate"
      Height          =   315
      Left            =   9120
      TabIndex        =   32
      Top             =   5460
      Width           =   1155
   End
   Begin VB.ListBox lstProcess 
      Height          =   5100
      IntegralHeight  =   0   'False
      Left            =   5520
      TabIndex        =   30
      Top             =   300
      Width           =   4815
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5100
      Top             =   120
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   315
      Left            =   4080
      TabIndex        =   28
      Top             =   2700
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Left            =   4080
      TabIndex        =   27
      Top             =   3060
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Manage Window"
      Height          =   1875
      Left            =   60
      TabIndex        =   12
      Top             =   3540
      Width           =   5355
      Begin VB.CommandButton cmdTerminate 
         Caption         =   "Terminate"
         Height          =   315
         Left            =   4080
         TabIndex        =   26
         Top             =   1380
         Width           =   1155
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show"
         Height          =   315
         Left            =   4080
         TabIndex        =   25
         Top             =   1020
         Width           =   1155
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide"
         Height          =   315
         Left            =   4080
         TabIndex        =   24
         Top             =   660
         Width           =   1155
      End
      Begin VB.CommandButton cmdNotOnTop 
         Caption         =   "Not OnTop"
         Height          =   315
         Left            =   2760
         TabIndex        =   23
         Top             =   1380
         Width           =   1155
      End
      Begin VB.CommandButton cmdOnTop 
         Caption         =   "Always OnTop"
         Height          =   315
         Left            =   2760
         TabIndex        =   22
         Top             =   1020
         Width           =   1155
      End
      Begin VB.CommandButton cmdSetTitle 
         Caption         =   "Set Title"
         Height          =   315
         Left            =   2760
         TabIndex        =   21
         Top             =   660
         Width           =   1155
      End
      Begin VB.CommandButton cmdDisable 
         Caption         =   "Disable"
         Height          =   315
         Left            =   1440
         TabIndex        =   20
         Top             =   1380
         Width           =   1155
      End
      Begin VB.CommandButton cmdEnable 
         Caption         =   "Enable"
         Height          =   315
         Left            =   1440
         TabIndex        =   19
         Top             =   1020
         Width           =   1155
      End
      Begin VB.CommandButton cmdFlash 
         Caption         =   "Flash"
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   660
         Width           =   1155
      End
      Begin VB.CommandButton cmdMaximize 
         Caption         =   "Maximize"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   660
         Width           =   1155
      End
      Begin VB.CommandButton cmdMinimize 
         Caption         =   "Minimize"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   1020
         Width           =   1155
      End
      Begin VB.CommandButton cmdNormal 
         Caption         =   "Normal"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   1380
         Width           =   1155
      End
      Begin VB.TextBox txtMhWnd 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label3 
         Caption         =   "hWnd:"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   2115
   End
   Begin VB.TextBox txtParent 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   2115
   End
   Begin VB.TextBox txtClass 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   2115
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   2115
   End
   Begin VB.TextBox txthWnd 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2115
   End
   Begin VB.PictureBox picCrossHair 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4260
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   60
      Width           =   495
   End
   Begin VB.Line Line10 
      X1              =   3480
      X2              =   3660
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Line Line9 
      X1              =   3360
      X2              =   3480
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Line Line8 
      X1              =   3360
      X2              =   3480
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line Line7 
      X1              =   3360
      X2              =   3480
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Line Line6 
      X1              =   3360
      X2              =   3480
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Line Line5 
      X1              =   3360
      X2              =   3480
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line Line4 
      X1              =   3360
      X2              =   3480
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Line Line3 
      X1              =   3360
      X2              =   3480
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line2 
      X1              =   3360
      X2              =   3480
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Line Line1 
      X1              =   3480
      X2              =   3480
      Y1              =   2790
      Y2              =   270
   End
   Begin VB.Label Label6 
      Caption         =   "Rectangle:"
      Height          =   195
      Left            =   120
      TabIndex        =   42
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label Label5 
      Caption         =   "Parent Class:"
      Height          =   195
      Left            =   120
      TabIndex        =   40
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label lblParentText 
      Caption         =   "Parent Text:"
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Running Process:"
      Height          =   255
      Left            =   5520
      TabIndex        =   31
      Top             =   60
      Width           =   1515
   End
   Begin VB.Label lblCordi 
      Caption         =   "X: 1043  Y: 0032"
      Height          =   255
      Left            =   60
      TabIndex        =   29
      Top             =   5520
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3660
      Picture         =   "frmMain.frx":0D0C
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Style:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1260
      Width           =   675
   End
   Begin VB.Label lblParent 
      Caption         =   "Parent:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   675
   End
   Begin VB.Label lblClass 
      Caption         =   "Class:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   900
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Drag this icon over the window you want to spy"
      Height          =   435
      Left            =   3660
      TabIndex        =   5
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label lblTitle 
      Caption         =   "Title:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   555
   End
   Begin VB.Label lblHwnd 
      Caption         =   "hWnd:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   675
   End
   Begin VB.Image imgCursor 
      Height          =   315
      Left            =   4740
      MouseIcon       =   "frmMain.frx":114E
      Top             =   60
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmMain
'    Project    : EliteSpy
'
'    Description: Main form
'
'    Author     : Andrea Batina
'    Modified   : 31/10/2001
'--------------------------------------------------------------------------------
Option Explicit

' Dragging window?
Private m_bDragging As Boolean

Private Sub chkOnTop_Click()
    If chkOnTop.Value = 1 Then
        ' Put window on top of all others
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        SaveSetting "EliteSpy+", "Settings", "AlwaysOnTop", "1"
    Else
        ' Remove window from top
        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        SaveSetting "EliteSpy+", "Settings", "AlwaysOnTop", "0"
    End If
End Sub

'////////////////////////////////////////////////////////////////////
'//// BUTTON EVENTS
'////////////////////////////////////////////////////////////////////
Private Sub cmdMaximize_Click()
    ' Maximize window
    ShowWindow txtMhWnd.Text, SW_MAXIMIZE
End Sub
Private Sub cmdMinimize_Click()
    ' Minimize window
    ShowWindow txtMhWnd.Text, SW_MINIMIZE
End Sub
Private Sub cmdNormal_Click()
    ' Show window
    ShowWindow txtMhWnd.Text, SW_NORMAL
End Sub

Private Sub cmdFlash_Click()
    ' Flash window
    FlashWindow txtMhWnd.Text, 3
End Sub
Private Sub cmdEnable_Click()
    ' Enable window
    EnableWindow txtMhWnd.Text, 1
End Sub
Private Sub cmdDisable_Click()
    ' Disable window
    EnableWindow txtMhWnd.Text, 0
End Sub

Private Sub cmdSetTitle_Click()
    Dim sTitle As String
    ' Ask user for new window title
    sTitle = InputBox("Enter new window title:", "EliteSpy +")
    ' Set new window title
    SetWindowText txtMhWnd.Text, sTitle
End Sub
Private Sub cmdOnTop_Click()
    ' Put window on top of all others
    SetWindowPos txtMhWnd.Text, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
Private Sub cmdNotOnTop_Click()
    ' Remove window from top
    SetWindowPos txtMhWnd.Text, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub cmdHide_Click()
    ' Hide window
    ShowWindow txtMhWnd.Text, SW_HIDE
End Sub
Private Sub cmdShow_Click()
    ' Show window
    ShowWindow txtMhWnd.Text, SW_SHOW
End Sub
Private Sub cmdTerminate_Click()
    ' Close window
    SendMessage txtMhWnd.Text, WM_CLOSE, 0, 0
End Sub

Private Sub cmdTerminateProcess_Click()
    ' Terminate process
    EnumProcess lstProcess.List(lstProcess.ListIndex)
End Sub
Private Sub cmdRefresh_Click()
    ' Enumerate open processes
    EnumProcess
End Sub

Private Sub cmdProcess_Click()
    ' If we are not showing the process bar then
    If cmdProcess.Caption = "Process ->" Then
        ' Show it
        Me.Width = 10530
        ' And set button caption
        cmdProcess.Caption = "Process <-"
        
    Else
        ' Hide process bar
        Me.Width = 5580
        ' And set button caption
        cmdProcess.Caption = "Process ->"
    End If
End Sub

Private Sub cmdCopy_Click()
    Dim sText As String
    
    ' Setup window information's
    sText = sText & "Window Handle:  " & txthWnd.Text & vbCrLf
    sText = sText & "Window Caption: " & txtTitle.Text & vbCrLf
    sText = sText & "Window Class:   " & txtClass.Text & vbCrLf
    sText = sText & "Window Style:   " & txtStyle.Text & vbCrLf
    sText = sText & "Rectangle:      " & txtRect.Text & vbCrLf
    sText = sText & "Parent Handle:  " & txtParent.Text & vbCrLf
    sText = sText & "Parent Caption: " & txtParentText.Text & vbCrLf
    sText = sText & "Parent Class:   " & txtParentClass.Text & vbCrLf
    
    ' Clear clipboard
    Clipboard.Clear
    ' Copy text to clipboard
    Clipboard.SetText sText
End Sub

Private Sub cmdMemInfo_Click()
    ' Remove window from top
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    chkOnTop.Value = 0
    frmMemInfo.Show , Me
End Sub

Private Sub cmdAbout_Click()
    ' Remove window from top
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    chkOnTop.Value = 0
    ' Show about box
    frmAbout.Show vbModal
End Sub
Private Sub cmdClose_Click()
    ' Close program
    Unload Me
End Sub

Private Sub cmdCodeGeneration_Click()
    ' Remove window from top
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    chkOnTop.Value = 0
    frmCGWizard.Show , Me
End Sub

'////////////////////////////////////////////////////////////////////
'//// FORM EVENTS
'////////////////////////////////////////////////////////////////////
Private Sub Form_Load()
    ' Set form width to default width (without process)
    Me.Width = 5580
    
    ' Make textboxes flat
    MakeFlat txthWnd.hwnd
    MakeFlat txtTitle.hwnd
    MakeFlat txtClass.hwnd
    MakeFlat txtRect.hwnd
    MakeFlat txtParent.hwnd
    MakeFlat txtParentText.hwnd
    MakeFlat txtParentClass.hwnd
    MakeFlat txtStyle.hwnd
    MakeFlat txtMhWnd.hwnd
    
    ' Get value from registry
    If GetSetting("EliteSpy+", "Settings", "AlwaysOnTop", "0") = "1" Then
        ' Put window on top of all others
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        Me.chkOnTop.Value = 1
    End If
        
    ' Enumerate open processes
    EnumProcess
End Sub

'////////////////////////////////////////////////////////////////////
'//// CROSSHAIR EVENTS
'////////////////////////////////////////////////////////////////////
Private Sub picCrossHair_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If user pressed left mouse button and we are not dragging
    If Button = vbLeftButton And Not m_bDragging Then
        ' Set dragging flag to true
        m_bDragging = True
        ' Set mouse pointer
        Me.MouseIcon = imgCursor.MouseIcon
        Me.MousePointer = 99
        ' Erase picture from picCrossHair
        picCrossHair.Picture = Nothing
    End If
End Sub

Private Sub picCrossHair_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If user pressed left mouse button and we are dragging
    If Button = vbLeftButton And m_bDragging Then
        Dim tPA As POINTAPI
        Dim lhWnd As Long
        Dim sTitle As String * 255
        Dim sClass As String * 255
        Dim tRC As RECT
        Dim sParentTitle As String * 255
        Dim sParentClass As String * 255
        Dim lhWndParent As Long
        Dim sStyle As String
        Dim lRetVal As Long
        
        ' Get cursor position
        GetCursorPos tPA
        ' Get window handle from point
        lhWnd = WindowFromPoint(tPA.X, tPA.Y)
        ' Get window caption
        lRetVal = GetWindowText(lhWnd, sTitle, 255)
        ' Get window class name
        lRetVal = GetClassName(lhWnd, sClass, 255)
        ' Get window style
        sStyle = GetWindowStyle(lhWnd)
        ' Get window rect
        GetWindowRect lhWnd, tRC
        ' Get window parent
        lhWndParent = GetParent(lhWnd)
        ' Get parent window caption
        lRetVal = GetWindowText(lhWndParent, sParentTitle, 255)
        ' Get parent window class name
        lRetVal = GetClassName(lhWndParent, sParentClass, 255)
        
        ' Set values to textboxes
        txthWnd.Text = lhWnd
        txtTitle.Text = sTitle
        txtClass.Text = sClass
        txtStyle.Text = sStyle
        txtRect.Text = "(" & tRC.Left & ", " & tRC.Top & ") - (" & tRC.Right & ", " & tRC.Bottom & ")"
        txtParent.Text = lhWndParent
        txtParentText.Text = sParentTitle
        txtParentClass.Text = sParentClass
        txtMhWnd.Text = lhWnd
    End If
End Sub

Private Sub picCrossHair_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If user pressed left mouse button and we are dragging
    If Button = vbLeftButton And m_bDragging Then
        ' Set dragging flag to true
        m_bDragging = False
        ' Restore mouse pointer to normal (arrow)
        Me.MousePointer = vbNormal
        ' Load picture into picCrossHair
        picCrossHair.Picture = imgCursor.MouseIcon
    End If
End Sub

'////////////////////////////////////////////////////////////////////
'//// PRIVATE FUNCTIONS
'////////////////////////////////////////////////////////////////////
' Get window styles
Private Function GetWindowStyle(ByVal lhWnd As Long) As String
    Dim lStyle As Long
        
    ' Get window styles
    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
    
    ' Get window styles
    If lStyle And WS_BORDER Then GetWindowStyle = GetWindowStyle & "WS_BORDER "
    If lStyle And WS_CAPTION Then GetWindowStyle = GetWindowStyle & "WS_CAPTION "
    If lStyle And WS_CHILD Then GetWindowStyle = GetWindowStyle & "WS_CHILD "
    If lStyle And WS_CLIPCHILDREN Then GetWindowStyle = GetWindowStyle & "WS_CLIPCHILDREN "
    If lStyle And WS_CLIPSIBLINGS Then GetWindowStyle = GetWindowStyle & "WS_CLIPSIBLINGS "
    If lStyle And WS_DLGFRAME Then GetWindowStyle = GetWindowStyle & "WS_DLGFRAME "
    If lStyle And WS_GROUP Then GetWindowStyle = GetWindowStyle & "WS_GROUP "
    If lStyle And WS_HSCROLL Then GetWindowStyle = GetWindowStyle & "WS_HSCROLL "
    If lStyle And WS_MAXIMIZEBOX Then GetWindowStyle = GetWindowStyle & "WS_MAXIMIZEBOX "
    If lStyle And WS_MINIMIZEBOX Then GetWindowStyle = GetWindowStyle & "WS_MINIMIZEBOX "
    If lStyle And WS_SYSMENU Then GetWindowStyle = GetWindowStyle & "WS_SYSMENU "
    If lStyle And WS_POPUPWINDOW Then GetWindowStyle = GetWindowStyle & "WS_POPUPWINDOW "
    If lStyle And WS_TABSTOP Then GetWindowStyle = GetWindowStyle & "WS_TABSTOP "
    If lStyle And WS_THICKFRAME Then GetWindowStyle = GetWindowStyle & "WS_THICKFRAME "
    If lStyle And WS_VISIBLE Then GetWindowStyle = GetWindowStyle & "WS_VISIBLE "
    If lStyle And WS_VSCROLL Then GetWindowStyle = GetWindowStyle & "WS_VSCROLL "

End Function

' Make textboxes flat
Private Sub MakeFlat(lhWnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
    ' Setup window styles
    lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    ' Set window style
    SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
    RemoveBorder lhWnd
End Sub
Private Sub RemoveBorder(lhWnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
    ' Setup window styles
    lStyle = lStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)
    ' Set window style
    SetWindowLong lhWnd, GWL_STYLE, lStyle
    ' Update window
    SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

' Get current mouse cordinates
Private Sub Timer1_Timer()
    Dim tPA As POINTAPI
    
    ' Get cursor cordinates
    GetCursorPos tPA
    ' Set label caption to cursor cordinates
    lblCordi.Caption = "X: " & tPA.X & "  Y: " & tPA.Y
End Sub

' Enumerate open processes
Private Sub EnumProcess(Optional ByVal sExeName As String = "")
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As PROCESSENTRY32
    
    ' Create snapshot
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If lSnapShot <> -1 Then
        ' Clear list
        lstProcess.Clear
        ' Length of the structure
        tPE.dwSize = Len(tPE)
        
        ' Find first process
        lNextProcess = Process32First(lSnapShot, tPE)
        Do While lNextProcess
            ' Found specified process
            If sExeName = Left$(tPE.szExeFile, Len(sExeName)) And Len(sExeName) > 0 Then
                Dim lProcess As Long
                Dim lExitCode As Long
                ' Open process
                lProcess = OpenProcess(0, False, tPE.th32ProcessID)
                ' Terminate process
                TerminateProcess lProcess, lExitCode
                ' Close handle
                CloseHandle lProcess
            Else
                ' Add exe to list
                lstProcess.AddItem tPE.szExeFile
            End If
            ' Get next process
            lNextProcess = Process32Next(lSnapShot, tPE)
        Loop
        
        ' Close handle
        CloseHandle (lSnapShot)
        
    Else
        lstProcess.AddItem "Cannot enumerate running process!"
    End If
End Sub
