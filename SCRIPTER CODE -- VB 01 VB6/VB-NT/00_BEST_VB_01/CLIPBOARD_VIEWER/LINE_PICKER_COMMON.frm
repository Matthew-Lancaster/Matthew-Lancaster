VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form LINE_PICKER_COMMON 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   225
   ClientTop       =   1695
   ClientWidth     =   13755
   Icon            =   "LINE_PICKER_COMMON.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   13755
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox LIST2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "LINE_PICKER_COMMON.frx":1272
      Left            =   2340
      List            =   "LINE_PICKER_COMMON.frx":1279
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1380
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
      Top             =   888
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1217
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
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2340
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   888
      Width           =   1476
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
      Height          =   396
      Left            =   504
      TabIndex        =   1
      Top             =   240
      Width           =   7752
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
   Begin VB.Menu MNU_HALIFAX 
      Caption         =   "JUMP HALIFAX"
   End
   Begin VB.Menu MNU_DOUBLE_LINE_SELECTOR_ON 
      Caption         =   "DOUBLE LINE SELECTOR ON"
   End
   Begin VB.Menu MNU_EDITOR_NOTEPAD_01 
      Caption         =   "EDITOR NOTEPAD++ FILE DATA 01"
   End
   Begin VB.Menu MNU_EDITOR_NOTEPAD_02 
      Caption         =   "EDITOR NOTEPAD++ FILE DATA 02 -- ADBLOCK"
   End
   Begin VB.Menu MNU_BUTTON_01_REMOVE_EMPTY_LINE_AND_REM_02_SORT_03_REMOVE_DUPER 
      Caption         =   "BUTTON 01 REMOVE EMPTY LINE && REM 02 SORT 03 REMOVE DUPER"
   End
   Begin VB.Menu MNU_ADD_CLIPBOARD 
      Caption         =   "ADD CLIBOARD"
   End
   Begin VB.Menu MNU_SORT_A 
      Caption         =   "SORT ACENDING"
   End
   Begin VB.Menu MNU_JUMP_LINE_ADBLOCK 
      Caption         =   "JUMP TO LINE __ ADBLOCK __ ARE"
   End
   Begin VB.Menu MNU_UMP_DRIVE_NETWORK 
      Caption         =   "JUMP DRIVE NETWORK"
   End
   Begin VB.Menu MNU_SHOW_MAIN_FORM 
      Caption         =   "MAIN FORM BACK"
   End
   Begin VB.Menu MNU_MINIMIZE_MAIN_FORM 
      Caption         =   "MAIN FORM -- MINIMIZE"
   End
End
Attribute VB_Name = "LINE_PICKER_COMMON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------
Public FORM_UNLOAD_INITIATE
Public NOT_RESIZE_ROUTINE
Public LINE_PICKER_COMMON_FROM_START

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

' PUBLIC RESIZE_DONE_ONCE

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



Sub GOT_FOCUS()

    If LINE_PICKER_FORM_LIST_CLICK_VALUE > 0 Then
        PLUS_10 = LINE_PICKER_FORM_LIST_CLICK_VALUE + 10
        NOT__10 = LINE_PICKER_FORM_LIST_CLICK_VALUE - 10
        LIST_COUNT = LISTVIEW1.ListItems.Count
        If NOT__10 < 1 Then NOT__10 = 1
        If PLUS_10 > LIST_COUNT Then PLUS_10 = LIST_COUNT
        LISTVIEW1.ListItems(NOT__10).EnsureVisible
        LISTVIEW1.ListItems(PLUS_10).EnsureVisible
        LISTVIEW1.ListItems(LINE_PICKER_FORM_LIST_CLICK_VALUE).EnsureVisible
        LISTVIEW1.ListItems(LINE_PICKER_FORM_LIST_CLICK_VALUE).Selected = True
        On Error Resume Next
        ' FORM NOT VISIBLE OR SOMETHING
        LISTVIEW1.SetFocus
        On Error GoTo 0
    End If
    
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
        FONT_SIZE = 14
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
    
    Call GOT_FOCUS
    
End Sub

Private Sub Form_GotFocus()

    Call GOT_FOCUS

End Sub

Sub Form_Load()
    If IsIDE = False Then
        MNU_ENDER.Visible = False
    End If
    
    RESIZE_DONE_ONCE = False
    
    MNU_SORT_A = False
    
    If LINE_PICKER_COMMON.LINE_PICKER_COMMON_FROM_START = False Then
        Me.Show
        Me.Caption = FRM_ClipTest.Caption
        LINE_PICKER_COMMON.LINE_PICKER_COMMON_FROM_START = False
    End If
    
    GETWINNT_VER_VAR = GETWinNT_Ver
    
    ReDim M(9000)
    ' LINE_PICKER_COMMON_LIST_SET
    
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
    i = i + 1: M(i) = "name *.EXE"
    i = i + 1: M(i) = "name *#NFS_EX*"
    i = i + 1: M(i) = "name *#NFS_EX*; name *"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "*1-ASUS*"
    i = i + 1: M(i) = "*2-ASUS*"
    i = i + 1: M(i) = "*3-LINDA*"
    i = i + 1: M(i) = "*4-ASUS*"
    i = i + 1: M(i) = "*5-ASUS*"
    i = i + 1: M(i) = "*7-ASUS*"
    i = i + 1: M(i) = "*8-MSI*"
    i = i + 1: M(i) = "*9-ASUS*"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "name *1-ASUS-*; name *"
    i = i + 1: M(i) = "name *2-ASUS-*; name *"
    i = i + 1: M(i) = "name *1-ASUS-*; name *"
    i = i + 1: M(i) = "name *3-LINDA-*; name *"
    i = i + 1: M(i) = "name *4-ASUS-*; name *"
    i = i + 1: M(i) = "name *5-ASUS-*; name *"
    i = i + 1: M(i) = "name *7-ASUS-*; name *"
    i = i + 1: M(i) = "name *8-MSI-*; name *"
    i = i + 1: M(i) = "name *9-ASUS-*; name *"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "*#NFS*"
    i = i + 1: M(i) = "name *#NFS*; name *"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "*-ASUS-*"
    i = i + 1: M(i) = "*-LINDA-*"
    i = i + 1: M(i) = "*-MSI-G*"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "name *-ASUS-*; name *"
    i = i + 1: M(i) = "name *-LINDA-*; name *"
    i = i + 1: M(i) = "name *-MSI-G*; name *"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "name *.MP3;all size<1"
    i = i + 1: M(i) = "----------"

    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "MATT.LAN@BTINTERNET.COM"
    i = i + 1: M(i) = "matt.lan@btinternet.com"
    i = i + 1: M(i) = "RUB.RIM@GMAIL.COM"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "all size>0 ---- USE WITH EXCLUDE "
    i = i + 1: M(i) = "all size>0 ---- BOTH EXIST AR WITH INCLUDE "
    i = i + 1: M(i) = "all size>0"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "any time>2020/01/01"
    i = i + 1: M(i) = "any time>2020/01/01 FOR EXCLUDE"
    i = i + 1: M(i) = "all time<2015/01/01"
    i = i + 1: M(i) = "all time<2m"
    i = i + 1: M(i) = "all time<2d"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "nowait: ""C:\SCRIPTER\SCRIPTER CODE -- VB 02 VBSCRIPT\VBS 35-RENAMER VB6 _VBP LCASE.VBS"" %JOBNAME%"""
    i = i + 1: M(i) = """C:\SCRIPTER\SCRIPTER CODE -- VB 02 VBSCRIPT\VBS 24-I_VIEW32 CONVERT_CCSE.AHK"""
    i = i + 1: M(i) = """C:\SCRIPTER\SCRIPTER CODE -- VB 02 VBSCRIPT\VBS 35-RENAMER VB6 _VBP LCASE.VBS"" MUTE"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "https://www.goodsync.com/php/gsss/main/showticket/?act=change"
    i = i + 1: M(i) = "name .tmp.drivedownload"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "path /Local/Google/Chrome/User Data"
    i = i + 1: M(i) = "path /Local/Google/Chrome/User Data/Default/Extension State"
    i = i + 1: M(i) = "path /Local/Google/Drive/user_default"
    i = i + 1: M(i) = "path /Local/Microsoft/Internet Explorer/Tiles"
    i = i + 1: M(i) = "path /Local/Microsoft/OneDrive"
    i = i + 1: M(i) = "path /Local/Nero"
    i = i + 1: M(i) = "path /Local/Temp"
    i = i + 1: M(i) = "path /Local/Packages"
    i = i + 1: M(i) = "path /Local/Amazon Drive"
    i = i + 1: M(i) = "path /Local/Microsoft/WindowsApps"
    i = i + 1: M(i) = "path /Local/Microsoft/Windows/WER/ReportQueue"
    i = i + 1: M(i) = "path /Local/Grammarly"
    i = i + 1: M(i) = "name opentabs-*.txt"
    i = i + 1: M(i) = "----------"
    
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "---------- NET PATH NOT LEAD SLASH"
    i = i + 1: M(i) = "1-ASUS-X5DIJ"
    i = i + 1: M(i) = "2-ASUS-EEE"
    i = i + 1: M(i) = "3-LINDA-PC"
    i = i + 1: M(i) = "4-ASUS-GL522VW"
    i = i + 1: M(i) = "5-ASUS-P2520LA"
    i = i + 1: M(i) = "7-ASUS-GL522VW"
    i = i + 1: M(i) = "8-MSI-GP62M-7RD"
    i = i + 1: M(i) = "9-ASUS-G815LM"
    i = i + 1: M(i) = "----------"
    
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "---------- NET PATH WITHER SLASH"
    i = i + 1: M(i) = "\\1-ASUS-X5DIJ"
    i = i + 1: M(i) = "\\2-ASUS-EEE"
    i = i + 1: M(i) = "\\3-LINDA-PC"
    i = i + 1: M(i) = "\\4-ASUS-GL522VW"
    i = i + 1: M(i) = "\\5-ASUS-P2520LA"
    i = i + 1: M(i) = "\\7-ASUS-GL522VW"
    i = i + 1: M(i) = "\\8-MSI-GP62M-7RD"
    i = i + 1: M(i) = "\\9-ASUS-G815LM"
    i = i + 1: M(i) = "----------"
    
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "---------- NET PATH C DRIVE"
    i = i + 1: M(i) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE"
    i = i + 1: M(i) = "\\2-ASUS-EEE\2_ASUS_EEE_01_C_DRIVE"
    i = i + 1: M(i) = "\\3-LINDA-PC\3_LINDA_PC_01_C_DRIVE"
    i = i + 1: M(i) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_01_C_DRIVE"
    i = i + 1: M(i) = "\\5-ASUS-P2520LA\5_ASUS_P2520LA_01_C_DRIVE"
    i = i + 1: M(i) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW_01_C_DRIVE"
    i = i + 1: M(i) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_01_C_DRIVE"
    i = i + 1: M(i) = "\\9-ASUS-G815LM\9_ASUS_G815LM_01_C_DRIVE"
    
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_02_D_DRIVE"
    i = i + 1: M(i) = "\\2-ASUS-EEE\2_ASUS_EEE_02_D_DRIVE"
    i = i + 1: M(i) = "\\3-LINDA-PC\3_LINDA_PC_02_D_DRIVE"
    i = i + 1: M(i) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_02_D_DRIVE"
    i = i + 1: M(i) = "\\5-ASUS-P2520LA\5_ASUS_P2520LA_02_D_DRIVE"
    i = i + 1: M(i) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW_02_D_DRIVE"
    i = i + 1: M(i) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_02_D_DRIVE"
    i = i + 1: M(i) = "\\9-ASUS-G815LM\9_ASUS_G815LM_02_D_DRIVE"
    
    
    I_BOOKMARK_TOP = i
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "---------- NET PATH -- GROUP COMPUTER NAME"
    i = i + 1: M(i) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE"
    i = i + 1: M(i) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_02_D_DRIVE"
    i = i + 1: M(i) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_03_FAT32_4GB"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "\\2-ASUS-EEE\2_ASUS_EEE_01_C_DRIVE"
    i = i + 1: M(i) = "\\2-ASUS-EEE\2_ASUS_EEE_02_D_DRIVE"
    i = i + 1: M(i) = "\\2-ASUS-EEE\2_ASUS_EEE_03_FAT32_4GB"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "\\3-LINDA-PC\3_LINDA_PC_01_C_DRIVE"
    i = i + 1: M(i) = "\\3-LINDA-PC\3_LINDA_PC_02_D_DRIVE"
    i = i + 1: M(i) = "\\3-LINDA-PC\3_LINDA_PC_03_FAT32_4GB"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_01_C_DRIVE"
    i = i + 1: M(i) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_02_D_DRIVE"
    i = i + 1: M(i) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_03_FAT32_4GB"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "\\5-ASUS-P2520LA\5_ASUS_P2520LA_01_C_DRIVE"
    i = i + 1: M(i) = "\\5-ASUS-P2520LA\5_ASUS_P2520LA_02_D_DRIVE"
    i = i + 1: M(i) = "\\5-ASUS-P2520LA\5_ASUS_P2520LA_03_FAT32_4GB"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW_01_C_DRIVE"
    i = i + 1: M(i) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW_02_D_DRIVE"
    i = i + 1: M(i) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW_03_FAT32_4GB"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_01_C_DRIVE"
    i = i + 1: M(i) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_02_D_DRIVE"
    i = i + 1: M(i) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_03_FAT32_4GB"
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "\\9-ASUS-G815LM\9_ASUS_G815LM_01_C_DRIVE"
    i = i + 1: M(i) = "\\9-ASUS-G815LM\9_ASUS_G815LM_02_D_DRIVE"
    i = i + 1: M(i) = "\\9-ASUS-G815LM\9_ASUS_G815LM_03_FAT32_4GB"
    I_BOOKMARK_BOTTOM = i
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "---------- NET PATH -- ---- \VB6\VB-NT ----"
    For R3 = I_BOOKMARK_TOP To I_BOOKMARK_BOTTOM
        If InStr(M(R3), "_D_DRIVE") > 0 Then
            i = i + 1: M(i) = M(R3) + "\VB6\VB-NT"
        End If
    Next
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "---------- NET PATH -- ---- \VB6-EXE\VB-NT ----"
    For R3 = I_BOOKMARK_TOP To I_BOOKMARK_BOTTOM
        If InStr(M(R3), "_D_DRIVE") > 0 Then
            i = i + 1: M(i) = M(R3) + "\VB6-EXE\VB-NT"
        End If
    Next
    i = i + 1: M(i) = "----------"
    i = i + 1: M(i) = "---------- NET PATH -- ---- \SCRIPTER ----"
    i = i + 1: M(i) = "\\1-ASUS-X5DIJ\1_ASUS_X5DIJ_01_C_DRIVE\SCRIPTER"
    i = i + 1: M(i) = "\\2-ASUS-EEE\2_ASUS_EEE_01_C_DRIVE\SCRIPTER"
    i = i + 1: M(i) = "\\3-LINDA-PC\3_LINDA_PC_01_C_DRIVE\SCRIPTER\"
    i = i + 1: M(i) = "\\4-ASUS-GL522VW\4_ASUS_GL522VW_01_C_DRIVE\SCRIPTER\"
    i = i + 1: M(i) = "\\5-ASUS-P2520LA\5_ASUS_P2520LA_01_C_DRIVE\SCRIPTER\"
    i = i + 1: M(i) = "\\7-ASUS-GL522VW\7_ASUS_GL522VW_01_C_DRIVE\SCRIPTER\"
    i = i + 1: M(i) = "\\8-MSI-GP62M-7RD\8_MSI_GP62M_7RD_01_C_DRIVE\SCRIPTER\"
    i = i + 1: M(i) = "\\9-ASUS-G815LM\9_ASUS_G815LM_01_C_DRIVE\SCRIPTER\"
    
    i = i + 1: M(i) = "----------"
    
    
    
    
    REMLINE_2 = 0
    CreateFolderTree App.Path + "\# DATA"
    FILENAME_LOAD_01 = App.Path + "\# DATA\VB__" + UCase(App.EXEName) + "_LINE_PICKER_COMMON.TXT"
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
                If BLOCK_REMLINE = False Then
                    i = i + 1: M(i) = VALUE_CHUNK
                End If
        Loop Until EOF(FR1)
        Close #FR1
    End If
    ' ---------------------------------------------------------------------
    ' LOAD FILE INTO ARRAY 01 OF 02 DONE
    ' ---------------------------------------------------------------------
    
    ' ---------------------------------------------------------------------
    ' LOAD FILE INTO ARRAY 02 OF 02 BEGIN
    ' ---------------------------------------------------------------------
    FILENAME_LOAD_02 = "C:\SCRIPTER\SCRIPT 08 IN_NET_ ADBLOCK\TEXT_ADBLOCK_01_MAIN.TXT"
    If Dir(FILENAME_LOAD_02) <> "" Then
        FR1 = FreeFile
        Open FILENAME_LOAD_02 For Input As #FR1
        Do
            Line Input #FR1, VALUE_CHUNK
            If Mid(VALUE_CHUNK, 1, 1) = "#" Then REMLINE = True Else REMLINE = False
                If REMLINE = False And REMLINE_2 = 1 Then REMLINE_2 = -2
                If REMLINE = True Then REMLINE_2 = 1
                If REMLINE = True And REMLINE_2 = 1 Then BLOCK_REMLINE = True
                If REMLINE_2 = -2 Then BLOCK_REMLINE = False
                If Trim(VALUE_CHUNK) = "" Then BLOCK_REMLINE = True
                If BLOCK_REMLINE = False Then
                    i = i + 1: M(i) = VALUE_CHUNK
                End If
        Loop Until EOF(FR1)
        Close #FR1
    End If
    ' ---------------------------------------------------------------------
    ' LOAD FILE INTO ARRAY 02 OF 02 DONE
    ' ---------------------------------------------------------------------
    

    
    ' ---------------------------------------------------------------------
    ' THE ARRAY BUILD ENDER HERE
    ' ---------------------------------------------------------------------
    
    ReDim Preserve M(i)
    List1.Clear
    AR = 0
    ' List1.Sorted = False
    For R3 = 1 To UBound(M)
        If M(R3) <> "" Then
            AR = AR + 1
            List1.AddItem Format(AR, "000") + " --  " + M(R3)
        End If
    Next
    
    
    
    
    
    
    
    ' LISTVIEW1.
    With LISTVIEW1
        .ColumnHeaders.Add , "ID", "ID", 800, lvwColumnLeft
        .ColumnHeaders.Add , "ITEM", "ITEM", 18000, lvwColumnLeft
        .View = lvwReport
        .Height = List1.Height
        .Top = List1.Top
        .Left = List1.Left
        .Width = List1.Width
'        .Font.Bold = True
'        .Font.Name = LIST1.FontNaizeme
'        .Font.Size = LIST1.Fonts

        .FullRowSelect = True
    End With
    
    LISTVIEW1.ListItems.Clear
    List1.Visible = False
    
    For R3 = 1 To UBound(M)
        If M(R3) <> "" Then
            AR = AR + 1
            With LISTVIEW1
                Set LV1 = .ListItems.Add(, , Format(AR, "000"))
                LV1.SubItems(1) = M(R3)
            End With
        End If
    Next
    
    If MNU_SORT_A_VAR = True Then
        LISTVIEW1.SortOrder = lvwAscending
        LISTVIEW1.SortKey = 1
        LISTVIEW1.Sorted = True
    End If
    
    AR = 0
    For R3 = 1 To UBound(M)
        If M(R3) <> "" Then
            AR = AR + 1
            LISTVIEW1.ListItems.Item(R3) = (Format(AR, "000"))
        End If
    Next
    
    Call frmListMenu.SET_MENU_PADD_WORK
    
    
    FORM_LOAD_STATE = True
    
    If AlwaysOnTop_MODE = True Then
        MNU_ALWAYS_ON_TOP.Caption = "ALWAYS ON TOP __ CORRECT"
    Else
        MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __]"
    End If
    
    Call GOT_FOCUS
    
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MOUSE_DOWN = True
MOUSE_UP = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MOUSE_DOWN = False
MOUSE_UP = True
End Sub

Private Sub Form_Resize()

If NOT_RESIZE_ROUTINE = True Then
    NOT_RESIZE_ROUTINE = False
    Exit Sub
End If

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

If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
    VAR_WINSTATE_BEEN_HIGH = vbMaximized
End If
If Me.WindowState = vbMinimized And VAR_WINSTATE_BEEN_HIGH = vbMaximized Then
    VAR_WINSTATE_BEEN_HIGH = "UNLOAD ME"
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
    LINE_PICKER_COMMON.SetFocus
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
        Me.Top = FRM_ClipTest.Top + 200
        Me.Left = FRM_ClipTest.Left + 200
        If FORM_SET_WIDTH_ONE = False Then
            FORM_SET_WIDTH_ONE = True
            Me.Width = FRM_ClipTest.Width
            ' SET HIEGHT ONCE ALSO
            Me.Height = FRM_ClipTest.Height
        End If

    End If
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
End If

WIDTH_ADJUST = 70 ' FOR WIN XP
If GETWINNT_VER_VAR = "WIN XP" Then WIDTH_ADJUST = 70
If GETWINNT_VER_VAR = "WIN 7" Then WIDTH_ADJUST = 170
If GETWINNT_VER_VAR = "WIN 10" Then WIDTH_ADJUST = 250

If GetComputerName = "1-ASUS-EEE" Then WIDTH_ADJUST = 70
If GetComputerName = "1-ASUS-X5DIJ" Then WIDTH_ADJUST = 70
If GetComputerName = "3-LINDA-PC" Then WIDTH_ADJUST = 170
If GetComputerName = "4-ASUS-GW522VW" Then WIDTH_ADJUST = 250
If GetComputerName = "5-ASUS-P2520LA" Then WIDTH_ADJUST = 220 '100

HEIGHT_ADJUST = 70 ' FOR WIN XP
If GETWINNT_VER_VAR = "WIN XP" Then HEIGHT_ADJUST = 70
If GETWINNT_VER_VAR = "WIN 7" Then HEIGHT_ADJUST = 70
If GETWINNT_VER_VAR = "WIN 10" Then HEIGHT_ADJUST = 200

'Bigger adjust number smaller inner form
If GetComputerName = "1-ASUS-EEE" Then HEIGHT_ADJUST = 70
If GetComputerName = "1-ASUS-X5DIJ" Then HEIGHT_ADJUST = 70
If GetComputerName = "3-LINDA-PC" Then HEIGHT_ADJUST = 70
If GetComputerName = "4-ASUS-GW522VW" Then HEIGHT_ADJUST = 100
If GetComputerName = "5-ASUS-P2520LA" Then HEIGHT_ADJUST = 0
'HIGHER NUMBER SMALLER INNER BOX

If RESIZE_DONE_ONCE = False Then
    Label1.BackColor = RGB(255, 255, 255)
    Label1.Left = 20
    Label1.Height = 400
    Label1.Caption = "SPEED LINE PICKER -- CLIPBOARD"
'    Label1.FontSize = 14
'    Label1.FontBold = True
'    Label1.FontName = "Arial"
End If

If RESIZE_DONE_ONCE = False Then
    Label1.Width = Me.Width - WIDTH_ADJUST + 20
    LB1H = (Menu_Height * Screen.TwipsPerPixelY) - 500
    If LB1H < 0 Then LB1H = 0
    Label1.Top = 20 ' LB1H
    List1.Top = Label1.Top + Label1.Height
    List1.Left = 10
End If
' List1.Height = Me.Height - (Menu_Height * Screen.TwipsPerPixelY)
List1.Width = Me.Width - WIDTH_ADJUST + 20
' Call HEIGHT_BAR_TOP_ROUTINE
' ---------------------------
'MAKE FORM TALLER OR TEXT BOX
'FORM IS PRIOITY OVER TEXT BOX
XXDD = Me.Height - (Menu_Height * Screen.TwipsPerPixelY) - (Label1.Top + Label1.Height) - 500
If XXDD - HEIGHT_ADJUST < 0 Then
    XXDD = 0 + HEIGHT_ADJUST + 100
End If
If OLD_LIST1_HEIGHT <> XXDD - HEIGHT_ADJUST Then
    List1.Height = XXDD - HEIGHT_ADJUST
    LISTVIEW1.Height = List1.Height
End If
OLD_LIST1_HEIGHT = XXDD - HEIGHT_ADJUST

If OLD_LIST1_WIDTH <> List1.Width Then
    LISTVIEW1.Width = List1.Width
End If
OLD_LIST1_WIDTH = List1.Width



'H1 = Me.Height - (List1.Top + 800 + 20)
'If H1 > 0 Then
'    List1.Height = H1
'    Else
'    List1.Height = Me.Height
'End If

If RESIZE_DONE_ONCE = False Then
    If Me.WindowState = vbNormal Then
        RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
        Me.Height = List1.Height + List1.Top + 900
        RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = False
    End If
End If



If RESIZE_DONE_ONCE = False Then
    With LISTVIEW1
        .Height = List1.Height
        .Top = List1.Top
        .Left = List1.Left
        .Width = List1.Width
    End With
End If

RESIZE_DONE_ONCE = True
TIMER_TO_RESIZE_VAR = False


End Sub


Private Sub Form_Unload(Cancel As Integer)

    FORM_UNLOAD_INITIATE = True
    WORKER_WITH_PICKER = False
    RESIZE_WINDOWSTATE_CHANGE_WORKAROUND = True
    
    LINE_PICKER_COMMON.WindowState = vbMinimized
    NOT_RESIZE_ROUTINE = True
    RESULT_API = NotAlwaysOnTop(Me.hWnd)
    ' WHEN LINE_PICKER_COMMON.GONE HERE MINIMIZE TOO
    NOT_RESIZE_ROUTINE = True
    FRM_ClipTest.WindowState = vbMinimized
    
    LINE_PICKER_COMMON.Visible = False
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
'    ' LINE_PICKER_COMMON.ME_POSITION
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

ITEMGET_1 = List1.List(List1.ListIndex)
ITEMGET_1 = Mid(ITEMGET, InStr(ITEMGET, " -- ") + 5)

If MNU_DOUBLE_LINE_SELECTOR_ON_VAR = True Then
    ITEMGET_1 = List1.List(List1.ListIndex)
    ITEMGET_1 = Mid(ITEMGET_1, InStr(ITEMGET_1, " -- ") + 5)

    ITEMGET_2 = List1.List(List1.ListIndex + 1)
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
        LINE_PICKER_COMMON.WindowState = vbMinimized
        
        RESULT_API = NotAlwaysOnTop(Me.hWnd)
    End If
End If

Beep
End Sub


Private Sub ListView1_Click()

ITEMGET_1 = LISTVIEW1.ListItems(LISTVIEW1.SelectedItem.Index).SubItems(1)
LINE_PICKER_FORM_LIST_CLICK_VALUE = LISTVIEW1.SelectedItem.Index

GOT_SOME = False
If InStr(ITEMGET_1, "! CLIPBOARD CHUNK BELOW -- ADBLOCK") > 0 Then
    ' ---------------------------------------------------------------------
    ' ! CLIPBOARD CHUNK BELOW -- ADBLOCK FOR CHROME
    ' ---------------------------------------------------------------------
    ' ENDER LINE
    ' ---------------------------------------------------------------------
    ' ! END CHUNK -- CLIPBOARD CHUNK BELOW -- ADBLOCK.COM CHROME
    ' ---------------------------------------------------------------------
    GOT_SOME = True
    GET_ = "01"
    For r = 1 To LISTVIEW1.ListItems.Count
        If GET_ = "01" Then
            ITEMGET_2 = LISTVIEW1.ListItems(r).SubItems(1)
            If InStr(ITEMGET_2, "! CLIPBOARD CHUNK BELOW -- ADBLOCK") > 0 Then
                GET_ = "02"
                ITEMGET_1 = ""
            End If
        End If
        If GET_ = "02" Then
            ITEMGET_2 = LISTVIEW1.ListItems(r).SubItems(1)
            If InStr(ITEMGET_2, "! END CHUNK -- CLIPBOARD CHUNK BELOW") > 0 Then
                Exit For
            End If
            If Trim(ITEMGET_2) <> "" Then
            If Mid(ITEMGET_2, 1, 4) <> "! ----" Then
                ITEMGET_1 = ITEMGET_1 + ITEMGET_2 + vbCrLf
            End If
            End If
        End If
    Next
End If

If MNU_DOUBLE_LINE_SELECTOR_ON_VAR = True And GOT_SOME = False Then
    ITEMGET_1 = LISTVIEW1.ListItems(LISTVIEW1.SelectedItem.Index).SubItems(1)
    ITEMGET_2 = LISTVIEW1.ListItems(LISTVIEW1.SelectedItem.Index + 1).SubItems(1)
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
    If AlwaysOnTop_MODE = False Then
    If WORKER_WITH_PICKER = True Then
        FRM_ClipTest.Visible = False
    Else
        FRM_ClipTest.Visible = True
    End If
    End If
    
    SELECT_SHOW = LISTVIEW1.SelectedItem.Index - 10
    If SELECT_SHOW < 1 Then SELECT_SHOW = 1
    If SELECT_SHOW < 15 Then
        LISTVIEW1.ListItems(1).EnsureVisible
    End If
    LISTVIEW1.ListItems(SELECT_SHOW).EnsureVisible
    
    If AlwaysOnTop_MODE = False Then
        FRM_ClipTest.WindowState = vbMinimized
        LINE_PICKER_COMMON.WindowState = vbMinimized
        Unload Me
    End If
    If AlwaysOnTop_MODE = True Then
        If SELECT_SHOW < 10 Then
            LISTVIEW1.ListItems(1).EnsureVisible
        End If
        LISTVIEW1.ListItems(SELECT_SHOW).EnsureVisible
        
        ' RESULT_API = NotAlwaysOnTop(Me.hWnd)
    End If
End If

Beep
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 And IsIDE = True Then End

End Sub

Private Sub MNU_ADD_CLIPBOARD_Click()

    
    MX = Split(FRM_ClipTest.Text2.Text, vbLf)
    
    AR = List1.ListCount
    For R3 = 0 To UBound(MX)
        If MX(R3) <> "" Then
            AR = AR + 1
            List1.AddItem Format(AR, "000") + " --  " + MX(R3)
            With LISTVIEW1
                Set LV1 = .ListItems.Add(, , Format(AR, "000"))
                LV1.SubItems(1) = MX(R3)
            End With
        End If
    Next
    
    FILENAME_LOAD_01 = App.Path + "\# DATA\VB__" + UCase(App.EXEName) + "_LINE_PICKER_COMMON.TXT"
    If Dir(FILENAME_LOAD_01) <> "" Then
        FR1 = FreeFile
        Open FILENAME_LOAD_01 For Append As #FR1
            Print #FR1, "---- CLIPBOARD INFO -- " + Format(Now, "YYYY MM DD  HH MM SS")
            Print #FR1, "-------------------------------------------"
            For R3 = 0 To UBound(MX)
                If MX(R3) <> "" Then
                    Print #FR1, MX(R3)
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

Public Sub MNU_ALWAYS_ON_TOP_Click()
    
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
    For R3 = 0 To UBound(MX)
        MX(R3) = Replace(MX(R3), vbCr, "")
        MX(R3) = Replace(MX(R3), vbLf, "")
        If Mid(Trim(MX(R3)), 1, 1) <> "!" Then
        If Trim(MX(R3)) <> "" Then
            LIST2.AddItem MX(R3)
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
    Call NotAlwaysOnTop(LINE_PICKER_COMMON.hWnd)
    AlwaysOnTop_MODE = False
    Unload Me
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
'     LINE_PICKER_COMMON.MNU_HALIFAX_Click
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
    NOT_RESIZE_ROUTINE = True
    Call NotAlwaysOnTop(LINE_PICKER_COMMON.hWnd)
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

Public Sub ME_POSITION_HALIFAX()

    Dim OFFSET_WX, OFFSET_HY

    OFFSET_HY = 10100
    HY = ((Menu_Height * Screen.TwipsPerPixelY) - Menu_Height) + OFFSET_HY

    NOT_RESIZE_ROUTINE = True
    Me.Height = HY
    
    NOT_RESIZE_ROUTINE = True
    Me.Width = 7000
    Me.Left = 15500
    Me.Top = 400
    Me.Visible = True
    DoEvents
    Call LINE_PICKER_COMMON.MNU_HALIFAX_Click
    Me.SetFocus
    

End Sub

