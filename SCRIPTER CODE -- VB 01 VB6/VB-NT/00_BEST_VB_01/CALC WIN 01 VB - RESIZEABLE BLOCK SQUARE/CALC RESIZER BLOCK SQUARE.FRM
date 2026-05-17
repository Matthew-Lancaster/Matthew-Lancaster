VERSION 5.00
Begin VB.Form Calculator 
   Caption         =   "Calculator"
   ClientHeight    =   3312
   ClientLeft      =   2652
   ClientTop       =   1860
   ClientWidth     =   3444
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
   Icon            =   "CALC RESIZER BLOCK SQUARE.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3312
   ScaleWidth      =   3444
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Number 
      Caption         =   "7"
      Height          =   480
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      Height          =   480
      Index           =   8
      Left            =   720
      TabIndex        =   8
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      Height          =   480
      Index           =   9
      Left            =   1320
      TabIndex        =   9
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "C"
      Height          =   480
      Left            =   2040
      TabIndex        =   10
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton CancelEntry 
      Caption         =   "CE"
      Height          =   480
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      Height          =   480
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      Height          =   480
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      Height          =   480
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "+"
      Height          =   480
      Index           =   1
      Left            =   2040
      TabIndex        =   12
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "-"
      Height          =   480
      Index           =   3
      Left            =   2640
      TabIndex        =   13
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      Height          =   480
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      Height          =   480
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      Height          =   480
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "X"
      Height          =   480
      Index           =   2
      Left            =   2040
      TabIndex        =   14
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "/"
      Height          =   480
      Index           =   0
      Left            =   2640
      TabIndex        =   15
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      Height          =   480
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1080
   End
   Begin VB.CommandButton Decimal 
      Caption         =   "."
      Height          =   480
      Left            =   1320
      TabIndex        =   18
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "="
      Height          =   480
      Index           =   4
      Left            =   2040
      TabIndex        =   16
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Percent 
      Caption         =   "%"
      Height          =   480
      Left            =   2640
      TabIndex        =   17
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Readout 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   105
      Width           =   3000
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_ON_TOP 
      Caption         =   "ON TOP"
   End
End
Attribute VB_Name = "Calculator"
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
' Option Explicit
Dim Op1, Op2                ' Previously input operand.
Dim DecimalFlag As Integer  ' Decimal point present yet?
Dim NumOps As Integer       ' Number of operands.
Dim LastInput               ' Indicate type of last keypress event.
Dim OpFlag                  ' Indicate pending operation.
Dim TempReadout


Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'----------------
'## MODIFY BATCH SHORT CUTS VARS
'Option Explicit

'---------------------------
'Skrol 29
'skrol29@freesurf.fr
'http://www.rezo.net/dir/skrol29/
'---------------------------
'Version 1.00, on 02/13/1999
'Version 1.01, on 04/19/1999
'---------------------------
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_COMMON_STARTMENU = &H16
Private Const CSIDL_COMMON_PROGRAMS = &H17
Private Const CSIDL_COMMON_STARTUP = &H18
Private Const CSIDL_COMMON_FAVORITES = &H1F

Private Declare Function api_SHAddToRecentDocs Lib _
   "shell32.dll" Alias "SHAddToRecentDocs" (ByVal dwFlags As _
   Long, ByVal dwData As String) As Long

Private Declare Function api_SHGetSpecialFolderLocation Lib _
  "shell32.dll" Alias "SHGetSpecialFolderLocation" (ByVal _
  hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long

Private Declare Function api_SHGetPathFromIDList Lib _
   "shell32.dll" Alias "SHGetPathFromIDList" _
   (ByVal pidl As Long, ByValsPath As String) _
   As Long

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function GetComputerNameA Lib "Kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long

' Public constants for set window on top
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

' Public constants for SetWindowPos API declaration
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2


Private Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Private Function NotAlwaysOnTop(ByVal hWnd As Long)
    Dim flags
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function



' Click event procedure for C (cancel) key.
' Reset the display and initializes variables.
Private Sub Cancel_Click()
    Readout = Format(0, "0.")
    Op1 = 0
    Op2 = 0
    Form_Load
End Sub

' Click event procedure for CE (cancel entry) key.
Private Sub CancelEntry_Click()
    Readout = Format(0, "0.")
    DecimalFlag = False
    LastInput = "CE"
End Sub

' Click event procedure for decimal point (.) key.
' If last keypress was an operator, initialize
' readout to "0." Otherwise, append a decimal
' point to the display.
Private Sub Decimal_Click()
    If LastInput = "NEG" Then
        Readout = Format(0, "-0.")
    ElseIf LastInput <> "NUMS" Then
        Readout = Format(0, "0.")
    End If
    DecimalFlag = True
    LastInput = "NUMS"
End Sub

' Initialization routine for the form.
' Set all variables to initial values.
Private Sub Form_Load()
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    Readout = Format(0, "0.")
    'Decimal.Caption = Format(0, ".")
    
    'End
    
    RESULT_API = AlwaysOnTop(Me.hWnd)
    
End Sub

Private Sub MNU_EXIT_Click()
End
End Sub

Private Sub MNU_ON_TOP_Click()

        RESULT_API = AlwaysOnTop(Me.hWnd)

Exit Sub

    If InStr(MNU_ALWAYS_ON_TOP.Caption, "ALWAYS ON TOP __") > 0 Then
        MNU_ALWAYS_ON_TOP.Caption = "[__ ON TOP __]"
        AlwaysOnTop_MODE = False
        RESULT_API = NotAlwaysOnTop(Me.hWnd)
    Else
        MNU_ALWAYS_ON_TOP.Caption = "[__ ALWAYS ON TOP __ CORRECT ___]"
        AlwaysOnTop_MODE = True
        RESULT_API = AlwaysOnTop(Me.hWnd)
    End If


End Sub

' Click event procedure for number keys (0-9).
' Append new number to the number in the display.
Private Sub Number_Click(Index As Integer)
    If LastInput <> "NUMS" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + Number(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + Number(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then Readout = "-" & Readout
    LastInput = "NUMS"
End Sub

' Click event procedure for operator keys (+, -, x, /, =).
' If the immediately preceeding keypress was part of a
' number, increments NumOps. If one operand is present,
' set Op1. If two are present, set Op1 equal to the
' result of the operation on Op1 and the current
' input string, and display the result.
Private Sub Operator_Click(Index As Integer)
    TempReadout = Readout
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
        If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
            Readout = "-" & Readout
            LastInput = "NEG"
        End If
        Case 1
        Op1 = Readout
        If Operator(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
            Readout = "-"
            LastInput = "NEG"
        End If
        Case 2
        Op2 = TempReadout
        Select Case OpFlag
            Case "+"
                Op1 = CDbl(Op1) + CDbl(Op2)
            Case "-"
                Op1 = CDbl(Op1) - CDbl(Op2)
            Case "X"
                Op1 = CDbl(Op1) * CDbl(Op2)
            Case "/"
                If Op2 = 0 Then
                   MsgBox "Can't divide by zero", 48, "Calculator"
                Else
                   Op1 = CDbl(Op1) / CDbl(Op2)
                End If
            Case "="
                Op1 = CDbl(Op2)
            Case "%"
                Op1 = CDbl(Op1) * CDbl(Op2)
            End Select
        Readout = Op1
        NumOps = 1
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
End Sub

' Click event procedure for percent key (%).
' Compute and display a percentage of the first operand.
Private Sub Percent_Click()
    Readout = Readout / 100
    LastInput = "Ops"
    OpFlag = "%"
    NumOps = NumOps + 1
    DecimalFlag = True
End Sub

Private Sub Mnu_VB_ME_Click()
    
    Dim CODER_VBP_FILE_NAME_2
    Dim VB_1, VB_2, VB_3
    VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
    VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
    If Dir(VB_1) <> "" Then VB_3 = VB_1
    If Dir(VB_2) <> "" Then VB_3 = VB_2
    '------------------------------------------------------
    'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
    '------------------------------------------------------
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
    Call CHECK_INTEGRITY_OF_VISUAL_BASIC_PROJECT_VBP(CODER_VBP_FILE_NAME_2)
    
    objShell.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set objShell = Nothing
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub
    
    
    Beep
    
'    i = MsgBox(GetSpecialfolder(38))
    
    If GetSpecialfolder(38) = "" Then
        MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__<" + GetSpecialfolder(38) + ">__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
        Exit Sub
    End If
    
    If IsIDE = True Then
        If IsIDE = True Then
            MSGQ = vbYesNoCancel
        Else
            MSGQ = vbOKOnly
        End If
        IR = MsgBox("NOT IN IDE ARE YOU SURE", MSGQ, vbMsgBoxSetForeground)
        If IR <> vbYes Then Exit Sub
    End If
    
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    End
End Sub
Private Sub MNU_VB_FOLDER_Click()
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub

Sub CHECK_INTEGRITY_OF_VISUAL_BASIC_PROJECT_VBP(CODER_VBP_FILE_NAME_2)
    ' --------------------------------------------------------------------
    ' FIND IF THE IS NOT CORRECT VERSION AND SIMPLE PUT CORRECT
    ' ONE MY COMPUTER SEEM PROBLEM WITH BASIC
    ' TALK HAS WRONG VERISON ONCE EVERY FEW DAY
    ' AND EDIT IT TO WHAT NORM VERISON SUPPOSED TO BE AND FINE
    ' DISCOVERY -- IT SEEM TO HAPPEN JUST AFTER A COMPILE SOMETIME_
    ' [ Monday 15:37:30 Pm_17 June 2019 ]
    ' --------------------------------------------------------------------
    ' --------------------------------------------------------------------
    Dim r, FR
    Dim VAR_STRING As String

    FR = FreeFile
    Open CODER_VBP_FILE_NAME_2 For Binary As FR
        VAR_STRING = Space(LOF(FR))
        Get #FR, , VAR_STRING
    Close FR

    ' --------------------------------------------------------------
    ' # .VBP # .VBP" -- # MSCOMC -- # OCX -- # MSCOMCTL # CTL # .CTL
    ' --------------------------------------------------------------
    ' FOUND TO BE ERROR
    ' --------------------------------------------------------------
    ' Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0; MSCOMCTL.OCX
    ' --------------------------------------------------------------
    ' CORRECT VALUE
    ' --------------------------------------------------------------
    ' Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0; MSCOMCTL.OCX
    ' --------------------------------------------------------------

    XR1 = InStr(UCase(VAR_STRING), UCase("A1}#2.1; MSCOMCTL.OCX"))
    XR2 = InStr(UCase(VAR_STRING), UCase("A1}#2.#")) ' ---- "A1}#2.#0; mscomctl.OCX"
    XR3 = InStr(UCase(VAR_STRING), UCase("#0; MSCOMCTL.OCX"))
    
    If XR1 = 0 And XR2 > 0 And XR3 > 0 Then
        GET_01 = InStr(UCase(VAR_STRING), UCase("MSCOMCTL.OCX")) - XR2
        Mid(VAR_STRING, XR2, GET_01 + Len("MSCOMCTL.OCX")) = "A1}#2.1#0; MSCOMCTL.OCX"
        ' MsgBox "MSCOMCTL.OCX" + vbCrLf + "WRONG VERSION -- CHANGE TO" + vbCrLf + "#2.1# -- MSCOMCTL.OCX"), vbMsgBoxSetForeground
        GO_NEXT_IN = True
    End If
    
    If GO_NEXT_IN = True Then
        If Dir(CODER_VBP_FILE_NAME_2) <> "" Then
            Kill CODER_VBP_FILE_NAME_2
        End If
        FR = FreeFile
        Open CODER_VBP_FILE_NAME_2 For Binary As FR
            Put #FR, , VAR_STRING
        Close FR
    End If

End Sub



Private Sub MNU_FOLDER_DIGITAL_STILL_CAMERA_Click()
    ' Beep
    Me.WindowState = vbMinimized
    Dim FOLDER_Path
    FOLDER_Path = "D:\DSC\2015+SONY\2020 CyberShot HX60V\JPG"
    Dir1.Path = FOLDER_Path
    Dim TD
    For r = 0 To Dir1.ListCount - 1
        FIND_PATH_1 = Dir1.List(r)
        FIND_PATH_2 = Mid(FIND_PATH_1, InStrRev(FIND_PATH_1, "\"))
        If Mid(FIND_PATH_2, 1, 3) = "\20" Then
            TD = Dir1.List(r)
        End If
    Next
    FOLDER_Path = TD
    
    Shell "EXPLORER /SELECT, """ + FOLDER_Path + """", vbNormalFocus
    EXIT_TRUE = True
    ' Unload Me
    'End
End Sub

Private Sub MNU_CAMERA_VIDEO_FOLDER_4G_Click()
    Me.WindowState = vbMinimized
    Shell "EXPLORER \\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY_MP4\2020 CyberShot HX60V __ MP4", vbMaximizedFocus
End Sub



Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    ' Beep
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
    End If
End Sub
'-------------------------------------------
'-------------------------------------------
'---------------------------------------------------
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
'Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
'Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Private Type SHITEMID
'    cb As Long
'    abID As Byte
'End Type
'Private Type ITEMIDLIST
'    mkid As SHITEMID
'End Type
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
'Function GetSpecialfolder(ByVal CSIDL As Long) As String
'    Dim R As Long
'    Dim IDL As ITEMIDLIST
'    R = SHGetSpecialFolderLocation(100, CSIDL, IDL)
'    If R = NOERROR Then
'        Path$ = Space$(512)
'        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
'        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
'        Exit Function
'    End If
'    GetSpecialfolder = ""
'End Function
'---------------------------------------------------



Private Function GetSpecialfolder(CSIDL As Long) As String
    '##############################################################################################
    'Returns the Path to a "Special" Folder (i.e. Internet History)
    '##############################################################################################
    
    Dim IDL As ITEMIDLIST
    Dim lResult As Long
    Dim sPath As String
    
    lResult = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If lResult = 0 Then
        sPath = Space$(512)
        lResult = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        
        GetSpecialfolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End Function


'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    i = CODER_VBP_FILE_NAME_2
    CODER_EXE_FILE_NAME_1 = i
    If InStr(i, " - 64 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 64 BIT.vbp", ".vbp")
    If InStr(i, " - 32 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 32 BIT.vbp", ".vbp")
    
    CODER_EXE_FILE_NAME_1 = Replace(CODER_EXE_FILE_NAME_1, ".vbp", ".exe")
    
    FileSpec = CODER_VBP_FILE_NAME_2
    If Dir(FileSpec) <> "" Then
        FILE_SPEC_GO = FileSpec
    Else
        If LCase(GetSpecialfolder(&H29)) = LCase("C:\Windows\SysWOW64") Then
            FileSpec = Replace(FileSpec, ".vbp", " - 64 BIT.vbp")
        Else
            FileSpec = Replace(FileSpec, ".vbp", " - 32 BIT.vbp")
        End If
    End If
    
    If Dir(FileSpec) <> "" Then FILE_SPEC_GO = FileSpec
    If Dir(FileSpec) = "" Then FILE_SPEC_GO = CODER_VBP_FILE_NAME_2
    
    If Dir(FILE_SPEC_GO) = "" Then MsgBox "ERROR NOT EXISTOR _ " + FILE_SPEC_GO
    
    CODER_VBP_FILE_NAME_2 = FILE_SPEC_GO
End Sub


