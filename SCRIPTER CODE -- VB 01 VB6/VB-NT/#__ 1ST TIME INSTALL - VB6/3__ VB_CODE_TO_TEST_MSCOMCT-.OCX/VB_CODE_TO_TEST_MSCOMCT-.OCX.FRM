VERSION 5.00
Begin VB.Form TEST_MSCOMCT_OCX 
   Caption         =   "VB CODE TO TEST MSCOMCTL.OCX -- MSCOMCT2.OCX"
   ClientHeight    =   8988
   ClientLeft      =   2652
   ClientTop       =   1860
   ClientWidth     =   12384
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8988
   ScaleWidth      =   12384
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   11484
      Top             =   1260
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4440
      Left            =   516
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "VB_CODE_TO_TEST_MSCOMCT-.OCX.frx":0000
      Top             =   1872
      Visible         =   0   'False
      Width           =   11610
   End
   Begin VB.Timer Timer_FORM_RUNNING 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11472
      Top             =   828
   End
   Begin VB.TextBox Text1 
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4440
      Left            =   -45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "VB_CODE_TO_TEST_MSCOMCT-.OCX.frx":0008
      Top             =   15
      Width           =   11610
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB_ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_RUN_ME_AGAIN 
      Caption         =   "RUN_RESTART"
   End
   Begin VB.Menu MNU_SPACER_12 
      Caption         =   "|"
   End
   Begin VB.Menu MNU_COMCTL32 
      Caption         =   "CORRECT LEGACY COMCTL32.OCX"
   End
   Begin VB.Menu MNU_SPACER 
      Caption         =   "|"
   End
   Begin VB.Menu MNU_MSCOMCT2 
      Caption         =   "CORRECT LEGACY MSCOMCT2.OCX"
   End
   Begin VB.Menu MNU_SPACER11 
      Caption         =   "|"
   End
   Begin VB.Menu MNU_MSCOMCTL 
      Caption         =   "CORRECT LEGACY MSCOMCTL.OCX"
   End
   Begin VB.Menu MNU_SPACER2 
      Caption         =   "|"
   End
   Begin VB.Menu MNU_ENABLE_ALL_MENU 
      Caption         =   "ENABLE ALL MENU -- NOT REQUIRING IF NOT FAULT -- TEST PROCESS RUN IF  WANTING"
   End
   Begin VB.Menu MNU_SPACER3 
      Caption         =   "|"
   End
   Begin VB.Menu MNU_COPY_CLIPBOARD 
      Caption         =   "CLIPBOARD ALL"
   End
   Begin VB.Menu MNU_SPACER_4 
      Caption         =   "|"
   End
   Begin VB.Menu MNU_HELP_ABOUT 
      Caption         =   "HELP IS AT HARP"
   End
   Begin VB.Menu MNU_URL 
      Caption         =   "URL OF SOURCE MESSENGER BOX"
   End
   Begin VB.Menu MNU_CMD_LINE 
      Caption         =   "CMD_LINE"
   End
End
Attribute VB_Name = "TEST_MSCOMCT_OCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
'               Copyright (C) 1994 Microsoft Corporation
'
' You have a royalty-free right to use, modify, reproduce and distribute
' the Sample Application Files (and/or any modified version) in any way
' you find useful, provided that you agree that Microsoft has no warranty,
' obligations or liability for any Sample Application Files.
' ------------------------------------------------------------------------
Option Explicit

Dim UNLOAD_REQUEST_TO_FULFILLER

Dim WINDOWS_PATH
Dim i
Dim FS, F, FR1
Dim ADATE1, ADATE2, ADATE3, ADATE4, ADATE5, ADATE7
Dim ASIZE1 As Long
Dim ASIZE2 As Long
Dim ASIZE3 As Long
Dim ASIZE4 As Long
Dim ASIZE5 As Long
Dim ASIZE7 As Long

Dim FILE_NAME_PATH
Dim FILE_NAME_PATH2
Dim F1_MSCOMCT2_LEGACY
Dim F2_MSCOMCTL_LEGACY
Dim F3_COMCTL32_LEGACY

Dim FILE_MSCOMCT2_OCX
Dim FILE_MSCOMCTL_OCX
Dim FILE_COMCTL32_OCX

Dim MATCH_MADE_IN_HEAVEN_00
Dim MATCH_MADE_IN_HEAVEN_1
Dim MATCH_MADE_IN_HEAVEN_2
Dim MATCH_MADE_IN_HEAVEN_3
Dim STORE_1
Dim STORE_2
Dim SEPERATER_STRING
Dim FILE_SIZE_VAR As Long
Dim FILE_1
Dim FILE_2

'-------------------
Dim FILE_1_DATE
Dim FILE_2_DATE
Dim FILE_3_DATE
'-------------------
Dim FILE_1_SIZE As Long
Dim FILE_2_SIZE As Long
Dim FILE_3_SIZE As Long
'-------------------
Dim MATCH_TRUE
'-------------------
Dim I_ARRAY

Dim FILE_STRIPPER_1_
Dim FILE_STRIPPER_11
Dim FILE_STRIPPER_2_
Dim FILE_STRIPPER_22


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then End
End Sub

' Initialization routine for the form.
' Set all variables to initial values.
Private Sub Form_Load()

'Text1.Top = 0
'Text1.Left = 0
'Text1.Width = Me.Width
'Text1.Height = Me.Height
'Text2.Top = 0
'Text2.Left = 0
'Text2.Width = Me.Width
'Text2.Height = Me.Height

Text1.FontSize = 7
Text1.FontBold = True

Me.Show

Timer_FORM_RUNNING.Enabled = True
   
Call MNU_HELP_ABOUT_Click
   
End Sub


Private Sub Form_Resize()

Text1.Top = 0
Text1.Left = 0
Text1.Width = Me.Width
Text1.Height = Me.Height
Text2.Top = 0
Text2.Left = 0
Text2.Width = Me.Width
Text2.Height = Me.Height
Me.WindowState = vbMaximized
Me.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
'End

UNLOAD_REQUEST_TO_FULFILLER = True

End Sub

Private Sub MNU_CMD_LINE_Click()
    
    Dim I2
    
    I2 = I2 + "COMMAND LINE OPTION ANYTHING ADDED ON " + vbCrLf
    I2 = I2 + "RUN AS QUITE MODE ALERT IF INTERESTING MODE" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "IF COMMAND$<>"""" THEN -- RUN AS A REBOOT MODE " + vbCrLf
    I2 = I2 + "CHECK AT STARTUP -- EXIT WHEN DONE" + vbCrLf
    I2 = I2 + vbCrLf

    MsgBox I2, vbMsgBoxSetForeground

End Sub

Private Sub MNU_COMCTL32_Click()

    FILE_NAME_PATH = FILE_COMCTL32_OCX
    FILE_NAME_PATH2 = F3_COMCTL32_LEGACY
    Call SHELL_THE_MODIFIER_BATCH
    
End Sub

Private Sub MNU_COPY_CLIPBOARD_Click()
Clipboard.Clear
Clipboard.SetText Text1.Text
Beep
End Sub

Private Sub MNU_URL_Click()
    
    'Microsoft ActiveX Common Control MSCOMCTL.OCX Security Update Problem with Registration Affects Treeview and ListBox in Microsoft Access, Office, VB6
    'http://www.fmsinc.com/microsoftaccess/controls/mscomctl/
    
    Dim AA

    AA = "Microsoft ActiveX Common Control MSCOMCTL.OCX Security Update Problem with Registration Affects Treeview and ListBox in Microsoft Access, Office, VB6" + vbCrLf
    AA = AA + "http://www.fmsinc.com/microsoftaccess/controls/mscomctl/" + vbCrLf + vbCrLf
    AA = AA + "CTRL C KEY TO CLIPBOARD MESGBOX LEARN"

    MsgBox AA, vbMsgBoxSetForeground
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then End

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then End

End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button <> 1 Then Exit Sub
Text2.Visible = False
Text1.Visible = True


End Sub

Private Sub Timer_FORM_RUNNING_Timer()

'-------------------------------------
Timer_FORM_RUNNING.Enabled = False
'-------------------------------------
Call RUN_CHECK_STATE
'-------------------------------------
   
'-----------------------------------------
'COMMAND$ FOR TEST MODE AND EXIT AT BOOTER
'-----------------------------------------
If UNLOAD_REQUEST_TO_FULFILLER Then Unload Me: Exit Sub
i = Command$
'I = "TEST"
If i <> "" Then
    If MNU_MSCOMCTL.Enabled = False And MNU_MSCOMCT2.Enabled = False Then
        Unload Me
        Beep
    End If
End If

End Sub



Sub RUN_CHECK_STATE()

On Error Resume Next

Set FS = CreateObject("Scripting.FileSystemObject")
    
If FS.FOLDEREXISTS("C:\Windows\SysWOW64\") Then
    WINDOWS_PATH = "C:\Windows\SysWOW64\"
Else
    WINDOWS_PATH = "C:\Windows\System32\"
End If

Set F = FS.GETFILE(WINDOWS_PATH + "MSCOMCT2.OCX")
ADATE1 = F.datelastmodified
ASIZE1 = F.Size
    
Set F = FS.GETFILE(WINDOWS_PATH + "MSCOMCTL.OCX")
ADATE2 = F.datelastmodified
ASIZE2 = F.Size

Set F = FS.GETFILE(WINDOWS_PATH + "COMCTL32.OCX")
ADATE5 = F.datelastmodified
ASIZE5 = F.Size

F1_MSCOMCT2_LEGACY = App.Path + "\MSCOMCT2.OCX.LEGACY"
Set F = FS.GETFILE(F1_MSCOMCT2_LEGACY)
ADATE3 = F.datelastmodified
ASIZE3 = F.Size

F2_MSCOMCTL_LEGACY = App.Path + "\MSCOMCTL.OCX.LEGACY"
Set F = FS.GETFILE(F2_MSCOMCTL_LEGACY)
ADATE4 = F.datelastmodified
ASIZE4 = F.Size

F3_COMCTL32_LEGACY = App.Path + "\COMCTL32.OCX.LEGACY"
Set F = FS.GETFILE(F3_COMCTL32_LEGACY)
ADATE7 = F.datelastmodified
ASIZE7 = F.Size

FILE_MSCOMCT2_OCX = WINDOWS_PATH + "MSCOMCT2.OCX"
FILE_MSCOMCTL_OCX = WINDOWS_PATH + "MSCOMCTL.OCX"
FILE_COMCTL32_OCX = WINDOWS_PATH + "COMCTL32.OCX"


Dim SEPERATER_STRING_2
Dim SEPERATER_STRING_3
Dim SEPERATER_STRING_4_CENTER
Dim SEPERATER_STRING_5_CENTER

SEPERATER_STRING = String(110, "-") + vbCrLf
SEPERATER_STRING_2 = String(110, "-") + vbCrLf
SEPERATER_STRING_3 = String(110, "=") + vbCrLf
SEPERATER_STRING_4_CENTER = String(28, "_")
SEPERATER_STRING_5_CENTER = String(15, "_")

i = ""
i = i + SEPERATER_STRING_2
i = i + SEPERATER_STRING_4_CENTER + " WORKER APP PROGRAM PATH AND WITHER THE 3 LEGACY .OCX FILE LOCATER " + SEPERATER_STRING_5_CENTER + vbCrLf
i = i + SEPERATER_STRING_2
i = i + "CODE VBP __ APP.PATH _____ = " + App.Path + "\" + vbCrLf
i = i + "CODE VBP __ APP.EXENAME __ = " + App.EXEName + ".EXE" + vbCrLf
i = i + SEPERATER_STRING_3
'----------------------------------------------------

'#01 ------------------------------------------------
'----------------------------------------------------
'----------------------------------------------------
STORE_1 = "": STORE_2 = ""
'----------------------------------------------------
FILE_SIZE_VAR = ASIZE1
'----------------------------------------------------
FILE_1 = WINDOWS_PATH + "MSCOMCT2.OCX"
FILE_2 = F1_MSCOMCT2_LEGACY
    
FILE_STRIPPER_1_ = Mid(FILE_1, InStrRev(FILE_1, "\")) + " ______"
FILE_STRIPPER_2_ = Mid(FILE_2, InStrRev(FILE_2, "\")) + " __"

If ASIZE1 = ASIZE3 Then
    '----------------------------------------------------
    'i = i + SEPERATER_STRING
    
    'REMOVE HERE 2026
    'i = i + "## 01 OF 02 -- TEST COMPARE REQUIRED OF -- " + FILE_1 + " --  WITHER -- LEGACY .OCX" + vbCrLf
    ''i = i + "LOADING" + vbCrLf
    'i = i + SEPERATER_STRING
    'Text1.Text = i
    'Text1.Refresh: DoEvents
    '----------------------------------------------------
    If UNLOAD_REQUEST_TO_FULFILLER Then Unload Me: Exit Sub
    Call COMPARE_ROUTINE
    If UNLOAD_REQUEST_TO_FULFILLER Then Unload Me: Exit Sub
    MATCH_MADE_IN_HEAVEN_1 = MATCH_MADE_IN_HEAVEN_00
    '----------------------------------------------------
End If

'#02 ------------------------------------------------
'----------------------------------------------------
'----------------------------------------------------
STORE_1 = "": STORE_2 = ""
'----------------------------------------------------
FILE_SIZE_VAR = ASIZE2
'----------------------------------------------------
FILE_1 = WINDOWS_PATH + "MSCOMCTL.OCX"
FILE_2 = F2_MSCOMCTL_LEGACY

FILE_STRIPPER_1_ = Mid(FILE_1, InStrRev(FILE_1, "\")) + " ______"
FILE_STRIPPER_2_ = Mid(FILE_2, InStrRev(FILE_2, "\")) + " __"

If ASIZE2 = ASIZE4 Then
    '----------------------------------------------------
    'REMOVE 2026
    'i = i + SEPERATER_STRING
    'i = i + "## 02 OF 02 -- TEST COMPARE REQUIRED OF -- " + FILE_1 + " --  WITHER -- LEGACY .OCX" + vbCrLf
    ''i = i + "LOADING" + vbCrLf
    'i = i + SEPERATER_STRING
    'Text1.Text = i
    'Text1.Refresh: DoEvents
    '----------------------------------------------------
    Call COMPARE_ROUTINE
    If UNLOAD_REQUEST_TO_FULFILLER Then Unload Me: Exit Sub
    MATCH_MADE_IN_HEAVEN_2 = MATCH_MADE_IN_HEAVEN_00
    '----------------------------------------------------
Else
    If FILE_SIZE_VAR <> 0 Then
        i = "02 OF 02 -- TEST COMPARE **NOT** REQUIRED OF -- " + FILE_1 + "--  WITHER -- LEGACY .OCX" + vbCrLf
        i = "02 OF 02 -- SIZE'S =" + SIZE_VAR(ASIZE2, True) + " -- " + SIZE_VAR(ASIZE4, True) + vbCrLf
        i = "LOADING" + vbCrLf
    End If
End If

'#03 ------------------------------------------------
'----------------------------------------------------
'----------------------------------------------------
STORE_1 = "": STORE_2 = ""
'----------------------------------------------------
FILE_SIZE_VAR = ASIZE5
'----------------------------------------------------
FILE_1 = WINDOWS_PATH + "COMCTL32.OCX"
FILE_2 = F3_COMCTL32_LEGACY

FILE_STRIPPER_1_ = Mid(FILE_1, InStrRev(FILE_1, "\")) + " ______"
FILE_STRIPPER_2_ = Mid(FILE_2, InStrRev(FILE_2, "\")) + " __"

If ASIZE5 = ASIZE7 Then
    '----------------------------------------------------
    ' REMOVE 2026
    'i = i + SEPERATER_STRING
    'i = i + "## 02 OF 02 -- TEST COMPARE REQUIRED OF -- " + FILE_1 + " --  WITHER -- LEGACY .OCX" + vbCrLf
    ''i = i + "LOADING" + vbCrLf
    'i = i + SEPERATER_STRING
    'Text1.Text = i
    'Text1.Refresh: DoEvents
    '----------------------------------------------------
    Call COMPARE_ROUTINE
    If UNLOAD_REQUEST_TO_FULFILLER Then Unload Me: Exit Sub
    MATCH_MADE_IN_HEAVEN_3 = MATCH_MADE_IN_HEAVEN_00
    '----------------------------------------------------
Else
    If FILE_SIZE_VAR <> 0 Then
        i = "02 OF 02 -- TEST COMPARE **NOT** REQUIRED OF -- " + FILE_1 + "--  WITHER -- LEGACY .OCX" + vbCrLf
        i = "02 OF 02 -- SIZE'S =" + SIZE_VAR(ASIZE5, True) + " -- " + SIZE_VAR(ASIZE7, True) + vbCrLf
        i = "LOADING" + vbCrLf
    End If
End If


'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------

'#01
'------------------------------------------
FILE_1 = WINDOWS_PATH + "MSCOMCT2.OCX"
FILE_2 = F1_MSCOMCT2_LEGACY

FILE_1_DATE = ADATE1
FILE_2_DATE = ADATE3
FILE_3_DATE = ADATE5

FILE_1_SIZE = ASIZE1
FILE_2_SIZE = ASIZE3
FILE_3_SIZE = ASIZE7

MATCH_TRUE = MATCH_MADE_IN_HEAVEN_1

i = i + SEPERATER_STRING
i = i + SEPERATER_STRING
I_ARRAY = "## ARRAY 01 ##"
i = i + I_ARRAY + " 01 OF 02 __ " + FILE_1 + vbCrLf

Call OUTPUT_TEXTING
'------------------------------------------

'#02
'------------------------------------------
FILE_1 = WINDOWS_PATH + "MSCOMCTL.OCX"
FILE_2 = F2_MSCOMCTL_LEGACY

FILE_1_DATE = ADATE2
FILE_2_DATE = ADATE4

FILE_1_SIZE = ASIZE2
FILE_2_SIZE = ASIZE4

MATCH_TRUE = MATCH_MADE_IN_HEAVEN_2

i = i + SEPERATER_STRING
i = i + SEPERATER_STRING
I_ARRAY = "## ARRAY 02 ##"
i = i + I_ARRAY + " 01 OF 02 __ " + FILE_1 + vbCrLf

Call OUTPUT_TEXTING
'------------------------------------------

'#03
'------------------------------------------
FILE_1 = WINDOWS_PATH + "COMCTL32.OCX"
FILE_2 = F3_COMCTL32_LEGACY

FILE_1_DATE = ADATE5
FILE_2_DATE = ADATE7

FILE_1_SIZE = ASIZE5
FILE_2_SIZE = ASIZE7

MATCH_TRUE = MATCH_MADE_IN_HEAVEN_3

i = i + SEPERATER_STRING
i = i + SEPERATER_STRING
I_ARRAY = "## ARRAY 03 ##"
i = i + I_ARRAY + " 01 OF 02 __ " + FILE_1 + vbCrLf

Call OUTPUT_TEXTING
'------------------------------------------

If ADATE1 <> ADATE3 Then
    MNU_MSCOMCTL.Enabled = True
Else
    MNU_MSCOMCTL.Enabled = False
End If

If ADATE2 <> ADATE4 Then
    MNU_MSCOMCT2.Enabled = True
Else
    MNU_MSCOMCT2.Enabled = False
End If

If ADATE5 <> ADATE7 Then
    MNU_COMCTL32.Enabled = True
Else
    MNU_COMCTL32.Enabled = False
End If
If MATCH_MADE_IN_HEAVEN_1 = True Then MNU_MSCOMCT2.Enabled = False
If MATCH_MADE_IN_HEAVEN_2 = True Then MNU_MSCOMCTL.Enabled = False
If MATCH_MADE_IN_HEAVEN_3 = True Then MNU_COMCTL32.Enabled = False

i = i + SEPERATER_STRING

Text1.Text = Replace(i, "-", "_")

End Sub


Function SIZE_VAR(SIZE_INPUT As Long, PAD_BEFORE As Boolean) As String

    Dim R
    Dim X22, X33
    
    SIZE_VAR = Format(SIZE_INPUT, "0")
    
    X22 = Len(SIZE_VAR)
    X33 = 20 - X22
    
    If PAD_BEFORE = True Then
        SIZE_VAR = String(X33, "_") + " " + SIZE_VAR
    Else
        SIZE_VAR = SIZE_VAR + " " + String(X33, "_")
    End If

End Function

Sub OUTPUT_TEXTING()

    
    FILE_STRIPPER_1_ = Mid(FILE_1, InStrRev(FILE_1, "\")) + " ________________________"
    FILE_STRIPPER_11 = Mid(FILE_1, InStrRev(FILE_1, "\")) + " ____________"
    FILE_STRIPPER_2_ = Mid(App.Path, 1, 3) + "__APP.PATH__" + Mid(FILE_2, InStrRev(FILE_2, "\"))
    FILE_STRIPPER_22 = Mid(App.Path, 1, 3) + "__APP.PATH__" + Mid(FILE_2, InStrRev(FILE_2, "\")) + " __"
    i = i + SEPERATER_STRING
    '--------------------------------------------------
    Dim UNDER_SCORE_PADD
    UNDER_SCORE_PADD = "_______"
    If FS.FILEEXISTS(FILE_1) Then
        i = i + FILE_STRIPPER_1_ + " FILE DATE __ " + Format(FILE_1_DATE, "DDD DD MMM YYYY -- HH:MM:SS") + vbCrLf
        
        i = i + FILE_STRIPPER_1_ + " FILE SIZE __" + UNDER_SCORE_PADD + SIZE_VAR(FILE_1_SIZE, True) + vbCrLf
        i = i + SEPERATER_STRING
    Else
        i = i + FILE_STRIPPER_1_ + " FILE DATE __ THE FILE DON'T EXIST" + vbCrLf
        i = i + FILE_STRIPPER_1_ + " FILE SIZE  __ THE FILE DON'T EXIST" + vbCrLf
    End If

    '--------------------------------------------------
    If FS.FILEEXISTS(FILE_2) Then
        i = i + I_ARRAY + " 02 OF 02 __ " + FILE_STRIPPER_2_ + vbCrLf
        i = i + SEPERATER_STRING
        i = i + FILE_STRIPPER_22 + " FILE DATE __ " + Format(FILE_2_DATE, "DDD DD MMM YYYY -- HH:MM:SS") + vbCrLf
        i = i + FILE_STRIPPER_22 + " FILE SIZE __" + UNDER_SCORE_PADD + SIZE_VAR(FILE_2_SIZE, True) + vbCrLf
        i = i + String(80, "-") + vbCrLf
    
    Else
        i = i + FILE_STRIPPER_22 + " FILE DATE __ THE FILE DON'T EXIST" + vbCrLf
        i = i + FILE_STRIPPER_22 + " FILE SIZE __ THE FILE DON'T EXIST" + vbCrLf
        i = i + "THE LEGACY FILE IN THE APP FOLDER HERE IS NOT EXIST ANYMORE " + vbCrLf + FILE_2
        'MsgBox "THE LEGACY FILE IN THE APP FOLDER HERE IS NOT EXIST ANYMORE " + vbCrLf + FILE_2, vbMsgBoxSetForeground
    End If
        
    If MATCH_TRUE = True Then
        i = i + FILE_STRIPPER_11 + " FILE VERIFIED __________ COMPARE WITH LEGACY FILE AS AN PERFECT MATCH *********" + vbCrLf
    Else
        i = i + FILE_STRIPPER_11 + " FILE VERIFIED **NOT** __ COMPARE WITH LEGACY FILE AS AN **NOT** PERFECT MATCHER" + vbCrLf
    End If

End Sub

Sub COMPARE_ROUTINE()
    MATCH_MADE_IN_HEAVEN_00 = False

    FILE_STRIPPER_1_ = Mid(FILE_1, InStrRev(FILE_1, "\")) + " ______"
    FILE_STRIPPER_2_ = Mid(FILE_2, InStrRev(FILE_2, "\")) + " __"

    '----------------------------------------------------
    ' REMOVE 2026
    'i = i + "01 OF 02 -- FILE PAIR COMPARE LOAD SIZE -- " + SIZE_VAR(FILE_SIZE_VAR, True) + " -- " + FILE_1 + vbCrLf
    'Text1.Text = i
    'Text1.Refresh: DoEvents
    '----------------------------------------------------
    FR1 = FreeFile
    Open FILE_1 For Binary As #FR1
        STORE_1 = Input(FILE_SIZE_VAR, FR1)
    Close #FR1
    '----------------------------------------------------
    ' REMOVE 2026
    'i = i + "02 OF 02 -- FILE PAIR COMPARE LOAD SIZE -- " + SIZE_VAR(FILE_SIZE_VAR, True) + " -- " + FILE_STRIPPER_2_ + vbCrLf
    'Text1.Text = i
    'Text1.Refresh: DoEvents
    '----------------------------------------------------
    If UNLOAD_REQUEST_TO_FULFILLER = True Then Unload Me: Exit Sub
    FR1 = FreeFile
    Open FILE_2 For Binary As #FR1
        STORE_2 = Input(FILE_SIZE_VAR, FR1)
    Close #FR1
    '----------------------------------------------------
    Dim PADD_LINE
    PADD_LINE = String(42, "_") + " "
    If STORE_1 = STORE_2 Then
        MATCH_MADE_IN_HEAVEN_00 = True
        ' REMOVE 2026
        ' i = i + PADD_LINE + FILE_STRIPPER_1_ + "____ RESULT MATCH PERFECT _____________________" + vbCrLf
    Else
        MATCH_MADE_IN_HEAVEN_00 = False
        ' REMOVE 2026
        ' i = i + PADD_LINE + FILE_STRIPPER_1_ + "____ RESULT **NOT** MATCH PERFECT **NOT**" + vbCrLf
    End If

End Sub


Private Sub MNU_ENABLE_ALL_MENU_Click()
    MNU_MSCOMCTL.Enabled = True
    MNU_MSCOMCT2.Enabled = True
End Sub

Private Sub MNU_EXIT_Click()
    Unload Me
End Sub

Private Sub MNU_HELP_ABOUT_Click()


    '----
    'Microsoft ActiveX Common Control MSCOMCTL.OCX Security Update Problem with Registration Affects Treeview and ListBox in Microsoft Access, Office, VB6
    'http://www.fmsinc.com/microsoftaccess/controls/mscomctl/
    '----


    Dim I2

    I2 = ""
    '----------------------------------------------------------------------------
    'EXTRA FOUR _ AT END THIS TOP LINE GET STRIPPED OUT BY END REPLACE OPERATION
    '@ SYMBOL CONVER REPLACE TO VBCRLF
    'NOT USED @ YET
    
    Dim IDATE_VB
    On Error Resume Next
    Set FS = CreateObject("Scripting.FileSystemObject")
    If FS.FILEEXISTS(App.Path + "\" + App.EXEName + ".VBP") Then
        Set F = FS.GETFILE(App.Path + "\" + App.EXEName + ".VBP")
        IDATE_VB = F.datelastmodified
    End If
    
    Dim I_NAME, INPUT_DATE
    
    I_NAME = App.Path + "\" + App.EXEName + "__APP_LAST_VERSIN_REVISION_CHANGER.TXT"
    If Dir(I_NAME) <> "" Then
        FR1 = FreeFile
        Open I_NAME For Input As #FR1
            Line Input #FR1, INPUT_DATE
        Close #FR1
    End If
    If DateValue(INPUT_DATE) + TimeValue(INPUT_DATE) <> IDATE_VB Then
        FR1 = FreeFile
        Open I_NAME For Output As #FR1
            Print #FR1, IDATE_VB
            Print #FR1, "PROJECT UPDATED LAST WORK ON DATE AND VERSION " + App.Major + "_" + App.Minor + "_" + App.Revision
            
        Close #FR1
    End If
    
    Dim VER_MAJOR_MINOR
    VER_MAJOR_MINOR = Format(IDATE_VB, "DDD DD MMM YYYYY HH:MM:SSap") + " __ Version " + Format(App.Major, "0") + "_" + Format(App.Minor, "0") + "_" + Format(App.Revision, "0") + " Rev. " + Format(App.Revision, " 0") + " Revision Process Done"
    
    '----------------------------------------------------------------------------
    I2 = I2 + "________ LAST CODING -- LAST SUNDAY 04 DECEMBER 2K SIXTEENTH __ " + VER_MAJOR_MINOR + " -- COUPLE OF DAY AND INITAL ONE MONTH BEFORE UP AND RUNNING ________" + "____" + vbCrLf
    I2 = I2 + "_____________________________________________________________________________________________________________________________________@"
    I2 = I2 + "OPERATION HAS COMMAND LINE OPTION ANYTHING ADDED ON RUN AS QUITE MODE ALERT IF INTERESTING MODE" + vbCrLf
    I2 = I2 + vbCrLf
    
    
    I2 = I2 + "THE LEGACY METHOD -- ARE ABOUT -- THE INITIAL INSTALL HAPPEN AND SUBSEQUENT MOST RECENT UPDATE IMPROVER LEFT A BUG ONCE -- " + vbCrLf
    I2 = I2 + "NOT SURE IF WAS FIXED AT A LATER UPDATE -- THE PROBLEM IT STILL LINGERING" + vbCrLf
    'I2 = I2 + vbCrLf
    
    I2 = I2 + "THE LEGACY ARE THE MOST UP_TO_DATE BEFORE THE ERROR UPDATE" + vbCrLf
    I2 = I2 + "MAYBE WINDOWS UPDATE HAS A PROBLEM FOR THIS AS THAT WHERE HAPPEN MOST TIME" + vbCrLf
    I2 = I2 + vbCrLf
    
    I2 = I2 + "YOU WILL NOTICE THE ERROR WITH THESE TWO FILE IN THE FORM DESIGNER " + vbCrLf
    I2 = I2 + "THE CONTROLS IN THE FORM WILL CHANGE TO PICTURE BOX AND VALUE WILL BE LOST " + vbCrLf
    I2 = I2 + "EVEN IF REQEUST IS ASKING THAT A FAULT WITHIN AND WHEN EXIT ASKING SAVE __ YES NO " + vbCrLf
    I2 = I2 + "ANSWER NOT WITH NO WILL STILL HARM-ER IT WITH LOSS OF DATA VIA PICTURE BOX REPLACE" + vbCrLf
    I2 = I2 + vbCrLf
    
    I2 = I2 + "THIS CODE WILL DO THE WORK OF FIX THE PROBLEM AS DETAILED BY ANOTHER PERSON" + vbCrLf
    I2 = I2 + "INITAL TO DO IN A BATCH FILE -- THIS MAKE IT DIFFERCULT CHECKING DATE AND COMPARE" + vbCrLf
    I2 = I2 + "MY LEGACY TWO FILE ONE IS AN EXACT HOUR OUT FROM A FAT32 COPY" + vbCrLf
    I2 = I2 + "OPERATION ARE ABLE TO GO BACK TO THE LAST WORKABLE LEGACY VERSION OF YOUR OWN BACKUP-ER" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "AS MY RESULT SHOW THE LEGACY VERSION IS 2011 AND A COMPARE VERIFY MATCH TRUE SHOW AN NEW INSTALL IS A 2004 AS MATCHING LEGACY AT 2011" + vbCrLf
    I2 = I2 + "REPEAT SOME THE OTHER 2ND -- .OCX IS 2004 AND MATCHER AND HAS TIME OUT BY ONE HOUR FAT32" + vbCrLf
    I2 = I2 + vbCrLf

    
    I2 = I2 + "THIS CODE WILL DO ALL THE REQUIRED WORK TO MAKE THE OPERATION SMOOTH AS DESCRIBED" + vbCrLf
    I2 = I2 + "IT WILL BE DONE BY A CREATED BATCH FILE AND RUN HAPPEN -- THIS PROGRAM WILL REMAIN -- " + vbCrLf
    I2 = I2 + "RERUN AFTER IF YOU WANTER CHECK THE CHANGE -- IT IS NOT A HASSLE TO RUN UN-REQUIRED "
    I2 = I2 + "IF ENABLE THE OPTION TO TEST IT GOING" + vbCrLf
    'I2 = I2 + vbCrLf
    
    I2 = I2 + "THE PROCESS HOWEVER NOT HASSLE FOR THIS CODE -- AS NOT USE THE .OCX FILE IN QUESTION" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "THE .OCX TWO FILE ARE FOR THE LISTVIEW TEXT BOX QUITE AN IMPORTANT USED MANY THING IN OFFICE AND WINDOW" + vbCrLf
    I2 = I2 + "THE OTHER ALTERNATIVE TO LIST VIEW IS LISTBOX MUCH INFERIOR AND LESS CAPACITY" + vbCrLf
    I2 = I2 + "ALSO MAYBE THE OTHER ONE .OCX FILE THE DATE PICKER GET LOST WITH THE LEGACY ERROR" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "THE ALSO MAYBE MORE AS - BEEN NOTED IN THE PAST BUT NOT DOCUMENT TO HAND -- SO HERE THE BEGIN CODE" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "OF ALL THE TALK RECENTLY WITH WIN 10" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "VB 6 DOES A MOST SCETCHY INSTALL SOMETIME WITH FAIL AT END OR WITH FREEZER AT END" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "CLICK THE STARMENU PROGRAM AND FIND IT IS WORKING" + vbCrLf
    I2 = I2 + "SOME FAULT FOUND NOT A HISTORY OF WORKED PROJECT -- BUT GO OVER THAT ONE" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "USE OF THE AUTOMATED INSTALL IS USEFUL -- BUT NOT ALLOWED TO RUN FIRST TIME" + vbCrLf
    I2 = I2 + "RUN IT THE SECOND TIME" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "USE OF OTHER CODE TO DO THE INSTAL PROCESS HAS BEEN TALK BY TWO PERSON" + vbCrLf
    I2 = I2 + "ONE IS NOT MAKE THE CODE AVAILABLE THAT HAS BEEN OUT THERE" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "THAT IS NOT REQUIRING" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "IS TALK OF A CERTAIN FILE DRIVER THAT MAKE THE INSTALL NOT HAPPENING OF WIN 10" + vbCrLf
    I2 = I2 + "NOT A PROBLEM" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "ONE TIP FROM COPY THE FOLDER OVER FROM ALL DISK SET IS PUT IN A ROOT AND ENTER DISK 1 AND SETUP" + vbCrLf
    I2 = I2 + "PUT IN A ROOT AND MAP A DRIVE AS A NETOWRK DRIVE BETTER THAN SUBST AS WANT INSTALL CONTINUE AFTER REBOOT DO" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "ONLY PROBLEM WHEN GET TO THE MSDN DISK WON'T INSTALL IT " + vbCrLf
    I2 = I2 + "IN MY WORKING NOT HAVING MSDN IS A PROBLEM" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "ONE OF THE INSTALL OF TWO TACKLE THE PROBLEM OF MSDN BUT THAT IS THE ONE REMOVER" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "A 3RD SOURCE -- WWW.PLANET_SOURCE_CODE.COM HAS A WRITE UP ABOUT WINDOW 10 INSTALL WITH VB" + vbCrLf
    I2 = I2 + "BEEN SEE A CODE EXAMPLE OF MICROSOFT CODING FROM SOURCE WEB THAT ENABLE THE MOUSE WHEEL HOWEVER COMPLEX SOURCER" + vbCrLf
    I2 = I2 + "WITH A .DLL DRIVER ALREADY COMPILED -- IT NOT WORKING FOR ME -- SO KATMOUSE IS STILL ON IS ANOTHER 2ND METHOD ALTERNATIVE" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "DO A CLICK ON HERE TEXT BOX TO RETURN PREVIOUS RESULT SCRIPT" + vbCrLf
    I2 = I2 + "ESC KEY EXIT ANYWHERE INSTANT" + vbCrLf
    I2 = I2 + vbCrLf
    I2 = I2 + "___________________________________________________________________" + vbCrLf

    'STRIPE VBCRLF -- BUT NOT DOUBLE ONE SET
    '---------------------------------------
    I2 = Replace(I2, " " + vbCrLf, vbCrLf)
    I2 = Replace(I2, vbCrLf + "____", "****__****__")
    I2 = Replace(I2, vbCrLf + vbCrLf, "--__--__")
    I2 = Replace(I2, vbCrLf, " ____ ")
    I2 = Replace(I2, "--__--__", vbCrLf + vbCrLf)
    I2 = Replace(I2, "****__****__", vbCrLf + "____")
    'i2 = Replace(i2, "-- " + vbCrLf, vbCrLf)
    I2 = Replace(I2, "-", "_")
    I2 = Replace(I2, "  ", " ")
    I2 = Replace(I2, "____", " ____ ")
    I2 = Replace(I2, "  ", " ")
    I2 = Replace(I2, "____ __", "______")
    I2 = Replace(I2, vbCrLf + " __", vbCrLf + "__")
    I2 = Replace(I2, "__ ____", "______")
    I2 = Replace(I2, "____ " + vbCrLf, vbCrLf)
    I2 = Replace(I2, " __ ", " ____ ")
    I2 = Replace(I2, "_ _", "__")
    I2 = Replace(I2, "@", vbCrLf)
    
    'Text2 = i2 'Trim(Mid(i2, 1, Len(i2) - 3))
    Text2 = Trim(Mid(I2, 1, Len(I2) - 5))

    Text2.Visible = True
    Text2.FontName = "ARIAL"
    Text2.FontBold = True
    Text2.FontSize = 8
    Text1.Visible = False

End Sub

Private Sub MNU_MSCOMCT2_Click()

    FILE_NAME_PATH = FILE_MSCOMCT2_OCX
    FILE_NAME_PATH2 = F1_MSCOMCT2_LEGACY
    Call SHELL_THE_MODIFIER_BATCH

End Sub

Private Sub MNU_MSCOMCTL_Click()
    
    FILE_NAME_PATH = FILE_MSCOMCTL_OCX
    FILE_NAME_PATH2 = F2_MSCOMCTL_LEGACY
    Call SHELL_THE_MODIFIER_BATCH

End Sub


Sub SHELL_THE_MODIFIER_BATCH()

i = ""
i = i + "CD\" + vbCrLf
i = i + "Regsvr32.exe /u """ + FILE_NAME_PATH + """" + vbCrLf
i = i + "COPY """ + FILE_NAME_PATH2 + """ """ + FILE_NAME_PATH + """" + vbCrLf
i = i + "Regsvr32.exe """ + FILE_NAME_PATH + """" + vbCrLf
i = i + "Regsvr32.exe /u """ + FILE_NAME_PATH + """" + vbCrLf
i = i + "DEL """ + FILE_NAME_PATH + """" + vbCrLf
i = i + "COPY """ + FILE_NAME_PATH2 + """ """ + FILE_NAME_PATH + """" + vbCrLf
i = i + "Regsvr32.exe """ + FILE_NAME_PATH + """" + vbCrLf
i = i + "@ECHO ------------------------------------------------" + vbCrLf
i = i + "@ECHO." + vbCrLf
i = i + "@ECHO THE BATCH FILE HAS COMPLETED ---- WAIT TO HAPPEN" + vbCrLf
i = i + "@ECHO ------------------------------------------------" + vbCrLf
i = i + "@ECHO SUCCESS WOULD HAPPEN WHATEVER INCLUDING IN IDE" + vbCrLf
i = i + "@ECHO AS LONG AS NOT ANOTHER PROGRAM IS HOGG THE FILE ACCESS FOR" + vbCrLf
i = i + "@ECHO THESE TWO .OCX FILE" + vbCrLf
i = i + "@ECHO A FAULT ABOUT FILE ACCESS WOULD REPORT IN THE COMMAND WINDOW" + vbCrLf
i = i + "@ECHO ------------------------------------------------" + vbCrLf
i = i + "@ECHO IF FAULT LIKE FILE COPY NONE THEN RUN IN SAFE MODE" + vbCrLf
i = i + "@ECHO ------------------------------------------------" + vbCrLf
i = i + "@ECHO." + vbCrLf
i = i + "PAUSE" + vbCrLf

FR1 = FreeFile
Open App.Path + "\RUN_BATCH_CUSTOM_CREATOR.BAT" For Output As #FR1
    Print #FR1, i
Close #FR1

' ----------------------------------------
Debug.Print "CMD.EXE /K START /D """ + App.Path + """ /B " + "RUN_BATCH_CUSTOM_CREATOR.BAT"
' ----------------------------------------
' RESULT FROM DEBUG.PRINT
' ----------------------------------------
' CMD.EXE /K START /D "D:\VB6\VB-NT\#__ 1ST TIME INSTALL INFO\3__ VB_CODE_TO_TEST_MSCOMCT-.OCX" /B RUN_BATCH_CUSTOM_CREATOR.BAT
' ----------------------------------------

' ----------------------------------------
' THE PATH HAS QUOTES AROUND HER
' BUTTER THE PROGRAM TO RUN MUST NOT
' SO MOT SURE IF ALLOW PROGRAM WITH SPACER
' ----------------------------------------
Dim CMD_STR
CMD_STR = "CMD.EXE /K START """" /D """ + App.Path + """ /B " + "RUN_BATCH_CUSTOM_CREATOR.BAT"
Shell CMD_STR, vbMaximizedFocus

'Unload Me

End Sub


Private Sub MNU_RUN_ME_AGAIN_Click()
    Unload Me
    DoEvents
    Load Me
    Me.Show
End Sub

Private Sub MNU_VB_FOLDER_Click()
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
    Unload Me
End Sub

Private Sub MNU_VB_ME_Click()
'GetSpecialfolder(38) OR
'GetSpecialfolder(42)
'38 -- PROGRAM FILES (X86)

On Error Resume Next

Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
Beep

If Err.Number > 0 Then MsgBox "ERROR HERE ::" + vbCrLf + """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbMsgBoxSetForeground

Unload Me

End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = False

Text2.Visible = False
Text1.Visible = True

End Sub
