VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   5136
   ClientLeft      =   192
   ClientTop       =   2376
   ClientWidth     =   15228
   Icon            =   "#0 Send To Multi Menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5136
   ScaleWidth      =   15228
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RTB5 
      Height          =   435
      Left            =   13560
      TabIndex        =   13
      Top             =   2850
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   762
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"#0 Send To Multi Menu.frx":0ECA
   End
   Begin RichTextLib.RichTextBox RTB4 
      Height          =   435
      Left            =   13545
      TabIndex        =   12
      Top             =   2370
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   762
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"#0 Send To Multi Menu.frx":0F55
   End
   Begin RichTextLib.RichTextBox RTB3 
      Height          =   435
      Left            =   13560
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   762
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"#0 Send To Multi Menu.frx":0FE0
   End
   Begin RichTextLib.RichTextBox RTB2 
      Height          =   435
      Left            =   13560
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   762
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"#0 Send To Multi Menu.frx":106B
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   405
      Left            =   13530
      TabIndex        =   9
      Top             =   1020
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   720
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"#0 Send To Multi Menu.frx":10F6
   End
   Begin VB.Timer Timer_1_MINUTE 
      Interval        =   1000
      Left            =   4908
      Top             =   468
   End
   Begin VB.Timer Timer_Form_Load_Opperation_Do_Time_Manager 
      Interval        =   1
      Left            =   4560
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   4260
      Top             =   480
   End
   Begin VB.FileListBox File1 
      Height          =   648
      Left            =   5220
      TabIndex        =   5
      Top             =   1476
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.DirListBox Dir1 
      Height          =   720
      Left            =   3168
      TabIndex        =   4
      Top             =   1452
      Visible         =   0   'False
      Width           =   1836
   End
   Begin VB.Timer Timer1 
      Interval        =   40000
      Left            =   3960
      Top             =   480
   End
   Begin VB.Label Label5 
      Caption         =   "PATH_FILE __ SOURCE __ "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   3180
      Width           =   13104
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "AVAILABLE CLIPBOARD __ PATH_FILE __ EXIST  __ REQUEST IF YOU WANT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   108
      TabIndex        =   7
      Top             =   2856
      Width           =   13116
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5_CLIPBOARD 
      Caption         =   "AVAILABLE CLIPBOARD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   132
      TabIndex        =   6
      Top             =   2532
      Width           =   13092
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label_TEXT_LIST_OF_FILE_WITH_LINE_NUMBER 
      Caption         =   "Text List of JPG Files With Number Bullet Point -- Camera"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   108
      TabIndex        =   3
      Top             =   3924
      Width           =   13116
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label_2_FOLDER 
      Caption         =   "Label_2_FOLDER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   96
      TabIndex        =   2
      Top             =   1296
      Width           =   13128
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Create a DOC Folder Here -- LABEL SET IN CODE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   108
      TabIndex        =   1
      Top             =   3516
      Width           =   13116
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label_1_FOLDER_FILE 
      Caption         =   "Label_1_FOLDER_FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   96
      TabIndex        =   0
      Top             =   48
      Width           =   13128
      WordWrap        =   -1  'True
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu Mnu_VB_ME 
      Caption         =   "VB_ME"
   End
   Begin VB.Menu Mnu_VB_Folder 
      Caption         =   "VB Folder"
   End
   Begin VB.Menu MNU_TIMER_PROJECT 
      Caption         =   "TIMER PROJECT"
   End
   Begin VB.Menu MNU_FILEDATE_CLIPBOARD 
      Caption         =   "FILE DATE TO CLIPBOARD"
   End
   Begin VB.Menu MNU_FILEDATE_WHOLE_FOLDER 
      Caption         =   "FILEDATE WHOLE FOLDER"
   End
   Begin VB.Menu MNU_MAKE_FOLDER_FILE_DATE 
      Caption         =   "MAKE FOLDER OF FILE DATE MENU__"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_MAKE_FOLDER_FILE_DATE_YYYY_MM_DD 
      Caption         =   "MAKE FOLDER OF FILE DATE HERE __ YYYY-MM-DD"
   End
   Begin VB.Menu MNU_MAKE_FOLDER_FILE_NVMS_DATE_YYYY_MM_DD 
      Caption         =   "MAKE FOLDER OF FILE WITHER NVMS __ YYYY-MM-DD"
   End
   Begin VB.Menu MNU_TREESIZE 
      Caption         =   "TREESIZE FREE"
   End
   Begin VB.Menu MNU_CLIPBOARD 
      Caption         =   "CLIPBOARD CONVERTOR TOOL MENU"
      Begin VB.Menu MNU_CLIPBOARD_VBCRLF_VBCRLF_IN_ONE 
         Caption         =   "CLIPBOARD MAKE DOUBLE (VBCRLF VBCRLF) INTO ONE VBCRLF __ ACCIDENT HAPPEN"
      End
   End
   Begin VB.Menu MNU_EXPLORER_FOLDER 
      Caption         =   "EXPLORER FOLDER"
   End
   Begin VB.Menu MNU_EXPLORER_FILE 
      Caption         =   "EXPLORER FILE"
   End
   Begin VB.Menu MNU_AUTO_RENAMER 
      Caption         =   "AUTO RENAMER APP"
   End
   Begin VB.Menu MNU_MEDIA_PLAYER_CLASSIC 
      Caption         =   "MPC MEDIAPLAYER CLASSIC __ RUN BY __ #0 Send To Text List of Files and Sub Folders IRFAN.exe"
   End
   Begin VB.Menu MNU_MEDIA_PLAYER_CLASSIC_LINE 
      Caption         =   "MPC MEDIAPLAYER CLASSIC"
   End
   Begin VB.Menu MNU_INFORAPID 
      Caption         =   "INFORAPID SEARCH AND REPLACE"
   End
   Begin VB.Menu MNU_INFORAPID_EBAY 
      Caption         =   "INFORAPID EBAY"
   End
   Begin VB.Menu MNU_NOTEPAD_EBAY 
      Caption         =   "NOTEPAD_ EBAY"
   End
   Begin VB.Menu MNU_NOTEPAD_FILE 
      Caption         =   "NOTEPAD FILE"
   End
   Begin VB.Menu MNU_NOTEPAD_PLUS_PLUS_FILE 
      Caption         =   "NOTEPAD++ FILE"
   End
   Begin VB.Menu MNU_PROCESS_LIST_CONVERT_RENAME_EDITOR 
      Caption         =   "PROCESS LIST CONVERT RENAME EDITOR"
   End
   Begin VB.Menu MNU_PROCESS_LIST_CONVERT_RENAME_01 
      Caption         =   "PROCESS BUTTON RENAME_01"
   End
   Begin VB.Menu MNU_PROCESS_LIST_CONVERT_RENAME_02 
      Caption         =   "PROCESS BUTTON RENAME_02 _ MAQ"
   End
   Begin VB.Menu MNU_FFMPEG_VERIFY 
      Caption         =   "FFMPEG VERIFY"
   End
   Begin VB.Menu MNU_TAKE_OWNERSHIP 
      Caption         =   "TAKE OWNERSHIP"
   End
   Begin VB.Menu MNU_REGISTER 
      Caption         =   "REGISTER"
   End
   Begin VB.Menu MNU_VB_SYNCRONIZER 
      Caption         =   "VB_SYNCRONIZER"
   End
   Begin VB.Menu MNU_LAUNCH_BATCH_COMPILER 
      Caption         =   "VB LAUNCH_BATCH_COMPILER"
   End
   Begin VB.Menu MNU_LAUNCH_Shell_VBasic_6_Loader 
      Caption         =   "VB LAUNCH_Shell_VBasic_6_Loader"
   End
   Begin VB.Menu MNU_VB_Create_Dir_Another_Drive 
      Caption         =   "VB Create Dir Another Drive"
   End
   Begin VB.Menu MNU_JUMP_TO_FOLDER_ON_NETWORK_DRIVE 
      Caption         =   "Jump to Folder On Network Drive"
   End
   Begin VB.Menu MNU_WINMERGE 
      Caption         =   "WIN_MERGE"
   End
   Begin VB.Menu MNU_HASH_MY_FILES 
      Caption         =   "HASH_MY_FILES"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SEARCH
'MNU_FILEDATE_CLIPBOARD_Click

Private Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long



Dim XVB_DATE
Dim MAQ
Dim FILE_NAME
Dim FILE_NAME_2
Dim NAME_PART_02
Dim PATH_FILE
Dim OCBGT
Dim FS, FSO
Dim DONT_WANT_CLIPBOARD_AT_START
Dim USE_CLIPBOARD
Dim FIRST_RUN_DONE
Dim OCBGT_DOUBLE_CLIPBOARD_VAR
Dim AT_TEMP
Dim FIRST_RUN_PATH_FILE_EXIST_DOUBLE
Dim FIRST_RUN_PATH_FILE_EXIST

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type MENUBARINFO
  cbSize As Long
  rcBar As rect
  hMenu As Long
  hwndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hwnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long


Private Sub Form_Activate()



'-----------------------------------------------
'--------------------------
'--------------------------
'--------------------------MAQ
'-----------------------------------------------

VAR_STRIP_PATH = Mid(Label_1_FOLDER_FILE.Caption, InStrRev(Label_1_FOLDER_FILE.Caption, "\") + 1)

If IsNumeric(Mid(Label_1_FOLDER_FILE.Caption, 1, 4)) Then Exit Sub
If IsNumeric(Mid(VAR_STRIP_PATH, 1, 4)) Then Exit Sub


If Label_1_FOLDER_FILE.Caption <> "" Then
    If InStr(Label_1_FOLDER_FILE.Caption, "MAQ0") > 0 Then
        Call MNU_FILEDATE_CLIPBOARD_Click
    End If
End If

If Label_1_FOLDER_FILE.Caption <> "" Then
    If InStr(GetLongName(Label_1_FOLDER_FILE.Caption), "MAQ0") > 0 Then
        Call MNU_FILEDATE_CLIPBOARD_Click
    End If
End If

If Label_1_FOLDER_FILE.Caption <> "" Then
    If InStr(Label_1_FOLDER_FILE.Caption, "MAH0") > 0 Then
        Call MNU_FILEDATE_CLIPBOARD_Click
    End If
End If

If Label_1_FOLDER_FILE.Caption <> "" Then
    If InStr(GetLongName(Label_1_FOLDER_FILE.Caption), "MAH0") > 0 Then
        Call MNU_FILEDATE_CLIPBOARD_Click
    End If
End If

If Label_1_FOLDER_FILE.Caption <> "" Then
    If InStr(GetLongName(Label_1_FOLDER_FILE.Caption), "M4V0") > 0 Then
        Call MNU_FILEDATE_CLIPBOARD_Click
    End If
End If

If Label_1_FOLDER_FILE.Caption <> "" Then
    If InStr(GetLongName(Label_1_FOLDER_FILE.Caption), "DSCF") > 0 Then
        If InStr(GetLongName(Label_1_FOLDER_FILE.Caption), ".MOV") > 0 Then
            Call MNU_FILEDATE_CLIPBOARD_Click
        End If
    End If
End If

' DOUBLE SCREEN CAMERA
If Label_1_FOLDER_FILE.Caption <> "" Then
    If InStr(GetLongName(Label_1_FOLDER_FILE.Caption), "DSCF") > 0 Then
        If InStr(GetLongName(Label_1_FOLDER_FILE.Caption), ".AVI") > 0 Then
            Call MNU_FILEDATE_CLIPBOARD_Click
        End If
    End If
End If


'Sleep 4000

End Sub


Private Sub MNU_FILEDATE_WHOLE_FOLDER_Click()

    '---------------------------------------------
    'VIDEO MP4
    '---------------------------------------------

    ' If FILE_NAME = "" Then Beep: MsgBox "NOT A FILE NAME PIPED": End
    
    File1.Path = Label_2_FOLDER
    
    Beep
    
    ' Me.Hide

'    On Error Resume Next
'    Me.Visible = True
'    DoEvents
'    Me.WindowState = vbMinimized
'    DoEvents
'    MsgBox FILE_NAME, vbMsgBoxSetForeground
'    On Error GoTo 0

'SOMETIME THE PASS BY CONTEXT MENU ISN;T LONG NAME BUT SHORTNAME
    
R_C_COUNTER_MAX = 0
For R_C_2 = 1 To 3
    
    A_M = ""
    
    R_C_COUNTER = 0
        
    For R_C = 0 To File1.ListCount - 1

        FILE_NAME = File1.Path + "\" + File1.List(R_C)
    
        If Mid(FILE_NAME, 1, 2) <> "\\" Then
            FILE_NAME = GetLongName(FILE_NAME)
        End If
        
        Set F = FS.GetFile((FILE_NAME))
        ADATE1 = F.DateLastModified
        
        'Clipboard.Clear
        'Sleep 200
        'If Year(ADATE1) = 2016 Then at1 = "-- 2k Sixteenth"
        'If Year(ADATE1) = 2015 Then at1 = "-- 2k Fifteenth"
        'If Year(ADATE1) <= 2015 Then at1 = "-- " + Format(ADATE1, "YYYY")
        'If Year(ADATE1) > 2016 Then at1 = "-- " + Format(ADATE1, "YYYY")
        
        'at1 = "-- " + Format(ADATE1, "YYYY")
        
        
        '------------------------
        ' LONG DATE FOR YOUTUBING
        '------------------------
        'If InStr(FILE_NAME, "D:\DSC\") > 0 Then
        'Clipboard.SetText " " + at1 + Format(ADATE1, " MMMM DD DDDD HH-MM-SS Am/Pm")
        FILE_NAME_2 = Mid(FILE_NAME, InStrRev(FILE_NAME, "\") + 1)
        FILE_NAME_2_EXT = Mid(FILE_NAME, InStrRev(FILE_NAME, "."))
        
        MAQ = ""
        START_MAQ = 0
        
        Call MNU_FILEDATE_CLIPBOARD_ROUTINE_CHECKER
        
        
        If File1.List(R_C) <> Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02 + FILE_NAME_2_EXT Then
            If MAQ <> "" Then
            
                R_C_COUNTER = R_C_COUNTER + 1
                If R_C_2 = 1 Then
                    R_C_COUNTER_MAX = R_C_COUNTER_MAX + 1
                
                End If
                
                A_M = A_M + "RENAMER --" + Str(R_C_COUNTER) + " OF" + Str(R_C_COUNTER_MAX) + " OF" + Str(File1.ListCount) + vbCrLf + File1.Path + "\" + vbCrLf + File1.List(R_C) + vbCrLf + File1.Path + "\" + vbCrLf + Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02 + FILE_NAME_2_EXT
                A_M = A_M + vbCrLf
                
                
                If R_C_2 = 3 Then
                
                    If R_C_COUNTER < 10 Then
                    
                        A_M_2 = "RENAMER -- VERIFY 1ST 10 AND THEN AUTO" + vbCrLf + Trim(Str(R_C + 1)) + " OF" + Str(File1.ListCount) + vbCrLf + File1.Path + "\" + vbCrLf + File1.List(R_C) + vbCrLf + File1.Path + "\" + vbCrLf + Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02 + FILE_NAME_2_EXT
                        
                        X_M = MsgBox(A_M_2, vbYesNo)
                        If X_M = vbNo Then
                            R_C_COUNTER = R_C_COUNTER - 1
                            Exit For
                        End If
                    End If
                    
                    Name File1.Path + "\" + File1.List(R_C) As File1.Path + "\" + Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02 + FILE_NAME_2_EXT
                
                End If
            End If
        End If
    Next
        
    If R_C_2 = 2 Then
        R_C_COUNTER = 0
        X_M = MsgBox("RENAME CONVERSION AS FOLLOW PROCESS YES / NO" + vbCrLf + A_M, vbYesNo)
        If X_M = vbNo Then Exit For
    End If
Next
    
If R_C_COUNTER > 0 Then

MsgBox "DONE " + Trim(Str(R_C_COUNTER)) + " CHANGES"

End If

End
    
End Sub



Private Sub MNU_FILEDATE_CLIPBOARD_Click()

    ' MNU_FILEDATE_CLIPBOARD
    ' [ Sunday 09:46:40 Am_18 November 2018 ]

    '---------------------------------------------
    'VIDEO MP4
    '---------------------------------------------

    If FILE_NAME = "" Then Beep: MsgBox "NOT A FILE NAME PIPED": End
    
    Beep
    
    Me.Hide

'    On Error Resume Next
'    Me.Visible = True
'    DoEvents
'    Me.WindowState = vbMinimized
'    DoEvents
'    MsgBox FILE_NAME, vbMsgBoxSetForeground
'    On Error GoTo 0

'    SOMETIME THE PASS BY CONTEXT MENU ISN;T LONG NAME BUT SHORTNAME

    If Mid(FILE_NAME, 1, 2) <> "\\" Then
        FILE_NAME = GetLongName(FILE_NAME)
    End If
    
    Set F = FS.GetFile((FILE_NAME))
    ADATE1 = F.DateLastModified
    
    Clipboard.Clear
    Sleep 200
    'If Year(ADATE1) = 2016 Then at1 = "-- 2k Sixteenth"
    'If Year(ADATE1) = 2015 Then at1 = "-- 2k Fifteenth"
    'If Year(ADATE1) <= 2015 Then at1 = "-- " + Format(ADATE1, "YYYY")
    'If Year(ADATE1) > 2016 Then at1 = "-- " + Format(ADATE1, "YYYY")
    
    'at1 = "-- " + Format(ADATE1, "YYYY")
    
    
    '------------------------
    ' LONG DATE FOR YOUTUBING
    '------------------------
    'If InStr(FILE_NAME, "D:\DSC\") > 0 Then
    'Clipboard.SetText " " + at1 + Format(ADATE1, " MMMM DD DDDD HH-MM-SS Am/Pm")
    FILE_NAME_2 = Mid(FILE_NAME, InStrRev(FILE_NAME, "\") + 1)
    'MsgBox FILE_NAME2
    MAQ = ""
    START_MAQ = 0
    

    Call MNU_FILEDATE_CLIPBOARD_ROUTINE_CHECKER
    
    'Sleep 4000
    
    Clipboard.SetText Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02
    'Else
    '------------------------
    ' OTHER DATE
    '------------------------
    'Clipboard.SetText Format(ADATE1, "YYYY-MM-DD")
    'End If

    End
    
End Sub

Sub MNU_FILEDATE_CLIPBOARD_ROUTINE_CHECKER()
        TEE = "MAQ"
        If InStr(FILE_NAME_2, TEE) > 0 Then
            START_MAQ_LEN = 8
            MAQ = Mid(FILE_NAME_2, InStr(FILE_NAME_2, TEE), START_MAQ_LEN)
            START_MAQ = InStr(FILE_NAME_2, TEE)
        End If
        
        TEE = "MAH"
        If InStr(FILE_NAME_2, TEE) > 0 Then
            START_MAQ_LEN = 8
            MAQ = Mid(FILE_NAME_2, InStr(FILE_NAME_2, TEE), START_MAQ_LEN)
            START_MAQ = InStr(FILE_NAME_2, TEE)
        End If
        
        TEE = "M4V"
        If InStr(FILE_NAME_2, TEE) > 0 Then
            START_MAQ_LEN = 8
            MAQ = Mid(FILE_NAME_2, InStr(FILE_NAME_2, TEE), START_MAQ_LEN)
            START_MAQ = InStr(FILE_NAME_2, TEE)
        End If
        
        TEE = "DSCF"
        If InStr(GetLongName(FILE_NAME), TEE) > 0 Then
            If InStr(GetLongName(FILE_NAME), ".MOV") > 0 Then
                START_MAQ_LEN = 8
                MAQ = Mid(FILE_NAME_2, InStr(FILE_NAME_2, TEE), START_MAQ_LEN)
                START_MAQ = InStr(FILE_NAME_2, TEE)
            End If
        End If

        'DOUBLE SCREEN CAMERA
        TEE = "DSCF"
        If InStr(GetLongName(FILE_NAME), TEE) > 0 Then
            If InStr(GetLongName(FILE_NAME), ".AVI") > 0 Then
                START_MAQ_LEN = 8
                MAQ = Mid(FILE_NAME_2, InStr(FILE_NAME_2, TEE), START_MAQ_LEN)
                START_MAQ = InStr(FILE_NAME_2, TEE)
            End If
        End If
        
        TEE = "image"
        If InStr(GetLongName(FILE_NAME), TEE) > 0 Then
            If InStr(GetLongName(FILE_NAME), ".AVI") > 0 Then
                START_MAQ_LEN = 7
                MAQ = Mid(FILE_NAME_2, InStr(FILE_NAME_2, TEE), START_MAQ_LEN)
                START_MAQ = InStr(FILE_NAME_2, TEE)
            End If
        End If
        
        
        'IF DESCRIPTION ALREADY ON
        NAME_PART_02 = ""
        If START_MAQ > 0 Then
            NAME_PART_02 = Mid(FILE_NAME_2, START_MAQ + START_MAQ_LEN)
            NAME_PART_02 = Mid(NAME_PART_02, 1, InStrRev(NAME_PART_02, ".") - 1)
        End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    If FS.FileExists(FILE_NAME) = True Then
        If UCase(Right(FILE_NAME, 4)) = ".JPG" Then
            Call Label2_Click
            RUNNER = True
        End If
    End If
    If RUNNER = False Then End
End If



End Sub

Private Sub FORM_LOAD()

Set FSO = CreateObject("Scripting.FileSystemObject")


If IsIDE = True Then
    Mnu_VB_ME.Enabled = False
    'MsgBox "NOT WHEN ISIDE"
    'Exit Sub
End If

Me.Caption = App.EXEName

FIRST_RUN_DONE = True

'E$ = "kujlhglkjhgf"""
'e$ = "D:\VB6\VB-NT\00_Best_VB_01\WebDates_Ace\WebDates.vbp"

'EXAMPLES
'MsgBox "-" + Command$ + "-"
'---------------------------
'#0 Send To ClipBoard Path & FileName
'---------------------------
'-"D:\VB6\VB-NT\00_Send_To\ClipBoard Path & FileName\#0 Send To ClipBoard Path & FileName.exe"-
'---------------------------
'OK
'---------------------------
'---------------------------
'#0 Send To ClipBoard Path & FileName
'---------------------------
'-"D:\VB6\VB-NT\00_Send_To\ClipBoard Path & FileName"-
'---------------------------
'OK
'---------------------------

'D:\DSC\2015 2016\2016 Sony CyberShot DSC-HX60V\DCIM2\ZZ FLICKR\20160621 -- POST OFFICE AN WALK OLD STEINE BRIGHTON -- FULL MOON\GB


For Each Control In Controls
    If UCase(Mid(Control.Name, 1, 3)) = "LAB" Then
        If Control.Top + Control.Height > xx_Height Then
            xx_Height = Control.Top + Control.Height
        End If
'        Control.Left = 50
'        Control.Width = 9120
    End If
Next

'Me.Height = xx_Height + (830) + 40
Me.Width = Label_1_FOLDER_FILE.Left + Label_1_FOLDER_FILE.Width + (300)

'LINE_CLIP = Command$
'
'If LINE_CLIP = "" Then
'
'    LINE_CLIP = Clipboard.GetText
'
'    IF DIR(LINE_CLIP,vbDirectory )
'
'End If



Call GET_PATH_OR_FILE_PATH("")



'---------------------------------------------
'VIDEO MP4
'---------------------------------------------






'LINE_CLIP = Replace(LINE_CLIP, """", "")

'Label_1_FOLDER_FILE = LINE_CLIP

'Clipboard.Clear
'Clipboard.SetText LINE_CLIP

Label2.Caption = "Create a DOC Folder Here -- And Or Also Send FILE to Doc Folder"

Timer1.Enabled = False

'Call Label2_Click

'WORKING ON SOMETHING
'SPLIT PROJECT CODE
''If IsIDE = True Then Call TEST_ROUTINE: Exit Sub

'Me.Show
'Me.WindowState = vbminnimized

Timer_Form_Load_Opperation_Do_Time_Manager.Enabled = False

End Sub


Sub GET_PATH_OR_FILE_PATH(AT_TEMP)

Set FS = CreateObject("Scripting.FileSystemObject")

'MsgBox Command$

If AT_TEMP = "" Then AT_TEMP = Command$
'AT_TEMP = "D:\DSC\2015 2016\2016 Sony CyberShot DSC-HX60V\DCIM2\ZZ FLICKR\20160621 -- POST OFFICE AN WALK OLD STEINE BRIGHTON -- FULL MOON\GB"
'AT_TEMP = "D:\DSC\2015 2016\2016 Sony CyberShot DSC-HX60V\DCIM2\ZZ FLICKR\20160621 -- ON GB"
AT_TEMP = Replace(AT_TEMP, """", "")

If FIRST_RUN_DONE = True Then
'    AT_TEMP = "D:\# MY DOCS\# 01 My Documents\Dropbox"
'    Label5.Caption = "PATH_FILE __ SOURCE __ ISIDE TEST FIRST RUN IN"
End If
FIRST_RUN_DONE = False

'PASTE TEST
'D:\DSC\2015 2016\2016 Sony CyberShot HX60V\DCIM2\00-FLIKR\A\20160621 -- 01 POST OFFICE

'CLIPBOARD DONE LATER
'If LINE_CLIP = "" Then
'    LINE_CLIP = Clipboard.GetText
'    'IF DIR(LINE_CLIP,vbDirectory )
'End If

'Me.Height = Label2.Top + Label2.Height + (830)
'Me.Width = Label_2_FOLDER.Left + Label_2_FOLDER.Width + (300)

'--------------------------
'FIND COMMANDLINE - SEND TO
'--------------------------

CMD_STRING = Replace(Command$, """", "")
i = 0
If FS.FolderExists(CMD_STRING) = True Then AT_TEMP = CMD_STRING: i = 1
If FS.FileExists(CMD_STRING) = True Then AT_TEMP = CMD_STRING: i = 1

If i = 1 Then
    Label5.Caption = "PATH_FILE __ SOURCE __ COMMAND$"
End If



On Error Resume Next
'-------------------------------------------------------------------------------------
CLIP_BOARD_STRING = Trim(Replace(Trim(Clipboard.GetText), """", ""))
On Error GoTo 0

i = 0
If FS.FolderExists(CLIP_BOARD_STRING) = True Then Label5_CLIPBOARD.Caption = CLIP_BOARD_STRING: i = 1
If FS.FileExists(CLIP_BOARD_STRING) = True Then Label5_CLIPBOARD.Caption = CLIP_BOARD_STRING: i = 1
If i = 1 Then

    'Label4.BackColor = QBColor(14)
    'Label5_CLIPBOARD.BackColor = QBColor(15)
End If


'----------------------------------------------------------------------------------------------
'WANT CLIPPER VAR IS NOT REALLY REQUIRED AS FAULT AT LOAD WAS GIVE CLIPPER WHEN EASY WORK IT IN
'----------------------------------------------------------------------------------------------

If AT_TEMP = "" Then
    On Error Resume Next
    CLIP_BOARD_STRING = Trim(Replace(Trim(Clipboard.GetText), """", ""))
    On Error GoTo 0
    i = 0
    If FS.FolderExists(CLIP_BOARD_STRING) = True Then AT_TEMP = CLIP_BOARD_STRING: i = 1
    If FS.FileExists(CLIP_BOARD_STRING) = True Then AT_TEMP = CLIP_BOARD_STRING: i = 1

    If i = 1 Then
        Label5.Caption = "PATH_FILE __ SOURCE __ ISIDE TEST FIRST RUN IN"
    End If

    DONT_WANT_CLIPBOARD_AT_START = True
    WANT_CLIPBOARD = True
    
    'Timer2.Enabled = True
Else
    DONT_WANT_CLIPBOARD_AT_START = False
    
    'Timer2.Enabled = False
End If
'-------------------------------------------------------------------------------------

'If AT_TEMP = "" Then AT_TEMP = Command$
'AT_TEMP = Replace(AT_TEMP, """", "")

i = 0
If FS.FolderExists(AT_TEMP) = True Then
    i = 1
    PATH_FILE = AT_TEMP
    SENDTO = "SEND TO RIGHT CLICK"
End If
FILE_FOUND = False
If FS.FileExists(AT_TEMP) = True Then
    i = 1
    PATH_FILE = AT_TEMP
    File = True
    FILE_NAME = AT_TEMP
    SENDTO = "SEND TO RIGHT CLICK"
    FILE_FOUND = True
End If

If i = 0 Then
    Label5.Caption = "PATH_FILE __ SOURCE __ NOT COMMAND$ AND NOT CLIPBOARD __ WAITING FOR INPUT TIMER"
    Label5.BackColor = QBColor(10)
End If

If i = 1 And FIRST_RUN_PATH_FILE_EXIST = True Then

    FIRST_RUN_PATH_FILE_EXIST = False
    FIRST_RUN_PATH_FILE_EXIST_DOUBLE = True
Else
    FIRST_RUN_PATH_FILE_EXIST_DOUBLE = 3
End If


'---------------------
'NONE SEND TO CMD LINE -- THEN
'IS CLIPBOARD ANY
'---------------------
On Error Resume Next
AT_TEMP = Replace(Trim(Clipboard.GetText), """", "")
On Error GoTo 0
If PATH_FILE = "" Then
    If FS.FolderExists(AT_TEMP) = True Then
        PATH_FILE = AT_TEMP
        SENDTO = "CLIPBOARD FIND PICK"
    End If
        
    If FS.FileExists(AT_TEMP) = True Then
        PATH_FILE = AT_TEMP
        FILE_NAME = AT_TEMP
        
        SENDTO = "CLIPBOARD FIND PICK"
        FILE_FOUND = True
    End If
End If
    
'---------
'TEST MODE -- FILE
'---------
AT_ISIDE = "C:\TEMP\bcdinfo.txt"
If IsIDE = True Then
    If PATH_FILE = "" Then
        AT_TEMP = AT_ISIDE
    End If
End If

'---------
'TEST MODE -- FOLDER
'---------
'AT_ISIDE = "C:\TEMP\"
'If IsIDE = True Then AT_TEMP = AT_ISIDE
    
Label_1_FOLDER_FILE = PATH_FILE
If FS.FolderExists(Label_1_FOLDER_FILE) = True Then
    Label_2_FOLDER.Caption = PATH_FILE
End If


If FILE_FOUND = True Then
    Frm_TIMER_PROJECT.Label1.Caption = Mid(PATH_FILE, 1, InStrRev(PATH_FILE, "\"))
Else
    Frm_TIMER_PROJECT.Label1.Caption = PATH_FILE
End If

If FILE_FOUND = True Or 1 = 2 Then
    'If InStr(PATH_FILE, "\PROJECT_TIMER.TXT") > 0 Then X1 = 1
    If InStr(PATH_FILE, "\PROJECT_TIMER") > 0 Then X1 = 1
    If InStr(PATH_FILE, "\PROJECT_TIMER_OUTPUT.TXT") > 0 Then X1 = 1
    If 1 = 2 Or X1 = 1 Then
        Load Frm_TIMER_PROJECT
        Call MNU_TIMER_PROJECT_Click
        'Frm_TIMER_PROJECT.MNU_MAIN_FORM.Visible = False
    End If
End If

AT_COMM = Command$
If 1 = 2 Or InStr(LCase(AT_COMM), "-timer_project") > 0 Then
        Load Frm_TIMER_PROJECT
        Call MNU_TIMER_PROJECT_Click
        'Frm_TIMER_PROJECT.MNU_MAIN_FORM.Visible = False
End If


If PATH_FILE <> "" Then
    'Frm_TIMER_PROJECT.Label2.Caption = "Project Timer Don't Exist -- Click On Here To Make One"
A = A
End If

If FILE_FOUND = True Then
    Label_2_FOLDER.Caption = Mid(Label_1_FOLDER_FILE.Caption, 1, InStrRev(Label_1_FOLDER_FILE.Caption, "\"))
    PATH_FILE = Label_2_FOLDER.Caption
End If

If PATH_FILE = "" Then
    I3 = "None Path/file Found Pass From Send To Or Command$ Or Clipboard" + vbCrLf
    I3 = I3 + "Do Clipboard A Path --" + vbCrLf
    I3 = I3 + "Hurry This Is A 500MSec Tax Clipboard Utility Resource Timer Clipboard Scanner"
    Label_2_FOLDER.Caption = I3
    Frm_TIMER_PROJECT.Label1.Caption = I3
    PATH_NOT_SET = True
    Form1.Timer2.Enabled = True
    
    Frm_TIMER_PROJECT.MNU_PROCESS_OUTPUT.Enabled = False
    Frm_TIMER_PROJECT.MNU_OPEN_INPUT_TEXT.Enabled = False
    Frm_TIMER_PROJECT.MNU_OPEN_INPUT_TEXT.Enabled = False
    Frm_TIMER_PROJECT.MNU_OPEN_OUTPUT_TEXT.Enabled = False
    Frm_TIMER_PROJECT.MNU_EXPLORER_FOLDER.Enabled = False
    Frm_TIMER_PROJECT.MNU_CLIP_BOARD_OUTPUT.Enabled = False
    
    'BIT OVER DOING IT CODE OVER DUPE SAFTY DEBUG
    'SORTED LATER
    'IN THE APP PATH HARDLY EVER SET FROM ONCE
    If Dir(App.Path + "\DATA_LAST_PROJECT.TXT") = "" Then Frm_TIMER_PROJECT.MNU_LAST_PROJECT.Enabled = False
    If Dir(App.Path + "\DATA_LAST_PROJECT_LOGGER.TXT") = "" Then Frm_TIMER_PROJECT.MNU_LAST_PROJECT_HISTORY.Enabled = False

End If

'MNU_STATUS.Caption = "  -- SOURCE GIVEN IS -- " + SENDTO

'If FILE = True Then
'TIMER1.Enabled = True
'End If

'--------------------------
'STILL NONE AFTER CLIPBOARD
'OR
'SEND TO
'USE REMEMBER FILE
'--------------------------

'If PATH_FILE = "" Then
'    MsgBox " NOTHING WITH VALID FOLDER OF FILE NAME GIVEN" + vbCrLf + "COMMAND$" + vbCrLf + Command$ + vbCrLf + "OR CLIPBOARD" + vbCrLf + AT_TEMP
'    End
'End If


'Label_1_FOLDER_FILE

'ChDrive w$
'ChDir w$

'If FILE = False Then
'    Shell "explorer /e, /select, " + PATH_FILE, vbNormalFocus
'    End
'Else
'    Exit Sub
'    Shell "explorer /e, /select, " + PATH_FILE, vbNormalFocus
'End If


End Sub

Private Sub Form_Resize()

Call SETUP_DIMENSIONS_INNER_FORM

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Frm_TIMER_PROJECT
Unload Me
Exit Sub

Dim Control

'SET ALL TIMERS IN ALL FORMS ENABLED=FALSE
On Error Resume Next
    For i = 0 To Forms.Count - 1
        For Each Control In Forms(i).Controls
            If InStr(UCase(Control.Name), "TIMER") > 0 Then
                'Debug.Print Control.Name
                Control.Enabled = False
            End If
        Next
    Next i
On Error GoTo 0
   
Dim Form As Form
For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form


End Sub

Private Sub MNU_EXPLORER_Click()

End Sub

Private Sub Label_2_FOLDER_FILE_Click()



End Sub


Private Sub Label4_Click()

    USE_CLIPBOARD = True
    OCBGT = ""
    Call Timer2_Timer

End Sub

Private Sub MNU_AUTO_RENAMER_Click()

i = Label_1_FOLDER_FILE.Caption
i = Mid(i, 1, InStrRev(i, "\"))

I_EXE1 = "C:\Program Files\# NO INSTALL REQUIRED\Advanced_Renamer\ARen.exe"
'I_EXE2 = Replace(UCase(I_EXE1), "C:\Program Files\", "C:\Program Files (x86)\")
I_EXE2 = I_EXE1
If InStr(UCase(I_EXE1), UCase("C:\Program Files\")) > 0 Then
    i_CUT1 = Mid(I_EXE2, 1, Len("C:\Program Files"))
    i_CUT2 = " (x86)"
    i_CUT3 = Mid(I_EXE2, Len("C:\Program Files "))
    I_EXE2 = i_CUT1 + i_CUT2 + i_CUT3
End If

If Dir(I_EXE1) = "" And Dir(I_EXE2) = "" Then
    I_INFO = ""
    I_INFO = I_INFO + "FOLDER FILE EXE NOT FOUND HERE -- " + vbCrLf + vbCrLf + I_EXE1 + vbCrLf + vbCrLf
    I_INFO = I_INFO + "FOLDER FILE EXE NOT FOUND HERE -- " + vbCrLf + vbCrLf + I_EXE2 + vbCrLf + vbCrLf
    MsgBox I_INFO, vbMsgBoxSetForeground
    Beep
    Exit Sub
End If

On Error Resume Next
PID = Shell(I_EXE1 + " " + """" + i + """", vbMaximizedFocus)
If PID > 0 Then Unload Me: Exit Sub
PID = Shell(I_EXE2 + " " + """" + i + """", vbMaximizedFocus)
If PID > 0 Then Unload Me: Exit Sub

End Sub

Private Sub MNU_CLIPBOARD_VBCRLF_VBCRLF_IN_ONE_Click()

TTXX = Clipboard.GetText

TTXX = Replace(TTXX, vbCrLf + vbCrLf, vbCrLf)

Clipboard.Clear
Sleep 10
Clipboard.SetText TTXX
Beep
End

End Sub


Private Sub MNU_EXIT_Click()
End
End Sub

Private Sub MNU_EXPLORER_FILE_Click()

i = Label_1_FOLDER_FILE.Caption

On Error Resume Next
PID = Shell("EXPLORER /SELECT, " + """" + i + """", vbMaximizedFocus)
If PID > 0 Then Unload Me: Exit Sub
Beep

End Sub

Private Sub MNU_EXPLORER_FOLDER_Click()

i = Label_1_FOLDER_FILE.Caption
i = Mid(i, 1, InStrRev(i, "\"))

On Error Resume Next
PID = Shell("EXPLORER /E," + """" + i + """", vbMaximizedFocus)
If PID > 0 Then Unload Me: Exit Sub
Beep

End Sub

Private Sub MNU_FFMPEG_VERIFY_Click()
    
    XF0 = Label_1_FOLDER_FILE.Caption
    XF1 = XF0 + ".txt"
    XF2 = Replace(XF1, ".txt", ".FFmpeg-Verify.txt")
    XF3 = Mid(XF0, 1, InStrRev(XF0, ".") - 1)
    XF3 = Mid(XF3, InStrRev(XF3, "\") + 1)
    XF4 = Mid(XF0, 1, InStrRev(XF0, "\")) + XF3 + "__FFMPEG_RUN_BATCH_" + Format(Now, "HH-MM-SS") + ".BAT"
    
    ' ---------------------------------------------------------------
    ' POINT TO PATH WHERE YOU GONNA USE __ FFMPEG.EXE
    ' COMMAND LINE VERSION
    ' LATEST VERSION WORKER GOOD
    ' HERE LOCATION
    ' ----
    ' Builds - Zeranoe FFmpeg
    ' https://ffmpeg.zeranoe.com/builds/
    ' ----
    ' ---------------------------------------------------------------
'    DH = """C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 18-ffmpeg-20150701-git-9c010ba-win32-static\ffmpeg.exe"" -v error -i """ + XF0 + """ -f null - 1>""" + XF2 + """ 2>&1"
'    DH = """C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 18-ffmpeg-20181007-0a41a8b-win64-static\bin\ffmpeg.exe"" -v error -i """ + XF0 + """ -f null - 1>""" + XF2 + """ 2>&1"
'    DH = """C:\PStart\# NOT INSTALL REQUIRED\ffmpeg-20150701-git-9c010ba-win32-static\ffmpeg.exe"" -v error -i """ + XF0 + """ -f null - 1>""" + XF2 + """ 2>&1"
    
    ' ---------------------------------------------------------------
    ' UP-TO-DATE VERSION
    ' ---------------------------------------------------------------
    PATH_FFMPEG = "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 18-ffmpeg-20181007-0a41a8b-win64-static\bin\ffmpeg.exe"
    DH = """" + PATH_FFMPEG + """ -v error -i """ + XF0 + """ -f null - 1>""" + XF2 + """ 2>&1"
    
    ' ---------------------------------------------------------------
    ' ---------------------------------------------------------------
    'TRY TO RUN SCRIPTING FROM COMMAND LINE -- WRONG DEBUG IT WORK TO DO
    'DH = "CMD /K """"C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 18-ffmpeg-20150701-git-9c010ba-win32-static\ffmpeg.exe"""" -v error -i """"" + XF0 + """""" + " -f null " + """""" + XF2 + """"""
    '-----------------------------
    'Dim WSHShell
    'Set WSHShell = CreateObject("WScript.Shell")
    'TEMP = WSHShell.ExpandEnvironmentStrings("%Temp%") + "\FFMPEG_RUN_BATCH.BAT"
    'Set WSHShell = Nothing
    ' ---------------------------------------------------------------
    ' ---------------------------------------------------------------
        
    TEMP = XF4
    
    FR1 = FreeFile
    Open TEMP For Output As #FR1
        Print #FR1, "@CD\"
        Print #FR1, DH
        Print #FR1,
        Print #FR1, "@ECHO."
        Print #FR1, "@ECHO __ "
        Print #FR1, "@ECHO __ SYSTEM AUDIO SOUND TO SIGNAL FINISH"
        Print #FR1, "@rundll32 user32.dll,MessageBeep"
        Print #FR1, "@ECHO __ "
        'Print #FR1, "@ECHO __ BATCH FILE DELETES ITSELF AFTER __ CLEANER"
        'Print #FR1, "@ECHO __ ERROR ABOUT BATCH FILE NOT FOUND ENDING IS NORMAL DELETER HAPPEN"
        Print #FR1, "@ECHO __ "
        Print #FR1, "@ECHO __ EMPTY OUTPUT RESULT FILE INDICATE NONE ERROR IN VIDEO FILE"
        Print #FR1, "@ECHO __ ERROR'S WOULD NORMALLY BE SHOWN IN GREAT DETAIL"
        Print #FR1, "@ECHO __ "
        Print #FR1, "@ECHO __ OUTPUT RESULT FILE WRITTEN IN SAME CURRENT DIRECTORY AS"
        Print #FR1, "@ECHO __ PROCESSED VIDEO FILE UNDER EXTENSION .TXT _ HERE LOCATION"
        Print #FR1, "@ECHO __ "
        Print #FR1, "@ECHO __ " + XF2
        Print #FR1, "@ECHO __ "
        Print #FR1, "@ECHO __ COMPLETER ___"
        Print #FR1, "@ECHO __ "
        
        Print #FR1, "@SET FILE_CONTAIN_DATA="
        
        '----------------------------------------------------------------
        ' Test Debug
        '----------------------------------------------------------------
        ' Print #FR1, "@ECHO TEST >>""" + XF2 +""""
    
        Print #FR1, "@ECHO OFF"
        Print #FR1, "(for /f usebackq^ eol^= %%a in (""" + XF2 + """) do break) && echo __ OutPut File For Video Has Process Result Error Data || SET FILE_CONTAIN_DATA=FALSE"
        Print #FR1, "@ECHO ON"
    
        Print #FR1, "@IF ""%FILE_CONTAIN_DATA%"" == ""FALSE"" GOTO FILE_NOT_REQUIRED_AS_EMPTY"
        Print #FR1, "@IF ""%FILE_CONTAIN_DATA%"" == """" GOTO FILE_HAS_DATA_NEXT_STEPPER"
        Print #FR1, ":FILE_NOT_REQUIRED_AS_EMPTY"
        Print #FR1, "@ECHO __ GOOD RESULT SHOWN BY PROCESS RESULT FILE BEING EMPTY"
        Print #FR1, "@ECHO __ EMPTY RESULT OUTPUT FILE TEXT NOT ANY LONGER NEEDED AND WILL BE DELETER"
        Print #FR1, "@ECHO __"
        Print #FR1, "@ECHO __ GOOD RESULT PROCESS OF VIDEO FILE HAS NONE ERROR"
        Print #FR1, "@ECHO __"
        Print #FR1, "@DEL """ + XF2 + """"
        Print #FR1, "@GOTO ENDE"
        Print #FR1, ":FILE_HAS_DATA_NEXT_STEPPER"
        
        Print #FR1, "@ECHO __ DO YOU WANT OPEN IN NOTEPAD++ __  ?(Y/N)"
        Print #FR1, "@SET INPUT="
        Print #FR1, "@SET /P INPUT=Type input: %=%"
        Print #FR1, "@IF ""%INPUT%"" == ""y"" GOTO YES"
        Print #FR1, "@IF ""%INPUT%"" == ""Y"" GOTO YES"
        Print #FR1, "@GOTO NOT"
        Print #FR1, ":YES"
        Print #FR1, "@START """" /MAX /HIGH ""C:\Program Files (x86)\Notepad++\notepad++.exe""" + " """ + XF2 + """"
        Print #FR1, ":NOT"
        Print #FR1, ":ENDE"
        
        Print #FR1, "@DEL """ + TEMP + """"
    Close FR1
        
    ' ---------------------------------------------------------------
    ' Credit Due for Source Empty File in Batch Command
    ' ---------------------------------------------------------------
    ' ----
    ' BATCH FILE CHECK FILE SIZE IS EMPTY - Google Search
    ' https://www.google.co.uk/search?q=BATCH+FILE+CHECK+FILE+SIZE+IS+EMPTY&spell=1&sa=X&ved=0ahUKEwi2rYHE5vndAhWJKMAKHTt3Cw8QBQgrKAA&biw=1536&bih=551&dpr=1.25
    ' --------
    ' windows batch: how to check if a text file is blank (the file has an empty line so size is non zero) - Stack Overflow
    ' https://stackoverflow.com/questions/33575962/windows-batch-how-to-check-if-a-text-file-is-blank-the-file-has-an-empty-line
    ' ----
    ' ---------------------------------------------------------------
        
    ' -----------------------------------------------------------------------------
    ' REMEMBER VB6 TO SET _ MICROSOFT SCRIPTING RUNTIME _ IN REFERENCES TO USE HERE
    ' -----------------------------------------------------------------------------
    ' SUPRESS COMMAND LINE ERROR OUTPUT TO CONSOLE AND SEND TO NUL INSTEAD
    ' BECUASE AN ANNOYING ERROR SHOW THAT DISPLAY THE BATCH FILE HAS BEEN DELETER
    ' WHILE IN MIDDLE OF PROCESS IT LINE BY LINE AT THE END TO CLEAN UP AND
    ' ERROR MESSENGER GETS IN THE WAY __ USE HERE + """ 2>NUL"
    ' -----------------------------------------------------------------------------
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run "CMD /K """ + TEMP + """ 2>NUL"
    Set WSHShell = Nothing

End Sub




Private Sub MNU_INFORAPID_Click()

i = Label_2_FOLDER.Caption
i = Replace(i, " ", "*")

II = " /nologo /WithSubdirectoriesYES /Dir" + i
'ii = " /nologo /WithSubdirectoriesYES /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\ /FileClipBoard-**.TXT"
'ii = " /nologo /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\ /FileClipBoard-*.TXT"

'ii = " /nologo /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\"

On Error Resume Next

PID = Shell(INFO_RAPID + " " + II, vbMaximizedFocus)
If PID > 0 Then Unload Me: Exit Sub

PID = Shell("C:\Program Files (X86)\seRapid\seRapid.exe " + II, vbMaximizedFocus)
If PID > 0 Then Unload Me: Exit Sub

'---------------------------------------------------
INFO_RAPID1 = "C:\Program Files\seRapid\seRapid.exe"
If Dir(INFO_RAPID2) = "" Then
    XX = 1
End If

INFO_RAPID2 = "C:\Program Files\seRapid\seRapid.exe"
If Dir(INFO_RAPID2) = "" Then
    
    If XX = 1 Then
        MSG_BOX_STRING = "FAILED LOAD __" + vbCrLf + vbCrLf + "HERE NOT EXISTOR" + vbCrLf + vbCrLf + INFO_RAPID1 + vbCrLf + vbCrLf + "AND ALSO" + vbCrLf + vbCrLf + "HERE NOT EXISTOR" + vbCrLf + vbCrLf + INFO_RAPID2
    Else
        MSG_BOX_STRING = "FAILED LOAD __" + vbCrLf + vbCrLf + "HERE MAYBE ERROR" + vbCrLf + vbCrLf + INFO_RAPID1 + vbCrLf + vbCrLf + "AND ALSO" + vbCrLf + vbCrLf + "HERE NOT EXISTOR" + vbCrLf + vbCrLf + INFO_RAPID2
    End If
    
    Beep
    MsgBox MSG_BOX_STRING, vbMsgBoxSetForeground
    
End If

Me.WindowState = 1

End Sub

Private Sub MNU_INFORAPID_EBAY_Click()
i = Label_2_FOLDER.Caption
i = Replace(i, " ", "*")

II = " /nologo /WithSubdirectoriesYES /Dir" + i
'ii = " /nologo /WithSubdirectoriesYES /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\ /FileClipBoard-**.TXT"
'ii = " /nologo /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\ /FileClipBoard-*.TXT"

'ii = " /nologo /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\"
'D:\# MY DOCS\# 01 My Documents\NOTEPAD DOCUMENT\EBAY 2015 TO 2017 PURCHASE HISTORY PASTED ON.txt

I1 = "/DirD:\# MY DOCS\# 01 My Documents\NOTEPAD DOCUMENT\"
I2 = "/FileEBAY 2015 TO 2017 PURCHASE HISTORY PASTED ON.txt"
I1 = Replace(I1, " ", "*")
I2 = Replace(I2, " ", "*")
II = " /nologo " + I1 + " " + I2

On Error Resume Next

PID = Shell(INFO_RAPID + " " + II, vbMaximizedFocus)
If PID > 0 Then Unload Me: Exit Sub

PID = Shell("C:\Program Files (X86)\seRapid\seRapid.exe " + II, vbMaximizedFocus)
If PID > 0 Then Unload Me: Exit Sub

'---------------------------------------------------
INFO_RAPID1 = "C:\Program Files\seRapid\seRapid.exe"
If Dir(INFO_RAPID2) = "" Then
    XX = 1
End If

INFO_RAPID2 = "C:\Program Files\seRapid\seRapid.exe"
If Dir(INFO_RAPID2) = "" Then
    
    If XX = 1 Then
        MSG_BOX_STRING = "FAILED LOAD __" + vbCrLf + vbCrLf + "HERE NOT EXISTOR" + vbCrLf + vbCrLf + INFO_RAPID1 + vbCrLf + vbCrLf + "AND ALSO" + vbCrLf + vbCrLf + "HERE NOT EXISTOR" + vbCrLf + vbCrLf + INFO_RAPID2
    Else
        MSG_BOX_STRING = "FAILED LOAD __" + vbCrLf + vbCrLf + "HERE MAYBE ERROR" + vbCrLf + vbCrLf + INFO_RAPID1 + vbCrLf + vbCrLf + "AND ALSO" + vbCrLf + vbCrLf + "HERE NOT EXISTOR" + vbCrLf + vbCrLf + INFO_RAPID2
    End If
    
    Beep
    MsgBox MSG_BOX_STRING, vbMsgBoxSetForeground
    
End If

Me.WindowState = 1
End Sub

Private Sub MNU_JUMP_TO_FOLDER_ON_NETWORK_DRIVE_Click()

'C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\Multiple_Thread Port Scanner 02 CON\NETWORK_COMPUTER_NAME.txt

End Sub

Private Sub MNU_MAKE_FOLDER_FILE_DATE_YYYY_MM_DD_Click()

'MAKE FOLDER OF FILE DATE HERE __
'MNU_MAKE_FOLDER_FILE_DATE_YYYY_MM_DD.Caption = "MAKE FOLDER OF FILE DATE HERE __ YYYY-MM-DD"
'--------------------------------------------

If FILE_NAME = "" Then Beep: MsgBox "NOT A FILE NAME PIPED": End

Set F = FS.GetFile((FILE_NAME))
ADATE1 = F.DateLastModified

'Clipboard.Clear

If Year(ADATE1) = 2016 Then at1 = "-- 2k Sixteenth"
If Year(ADATE1) = 2015 Then at1 = "-- 2k Fifteenth"
If Year(ADATE1) < 2015 Then at1 = "-- " + Format(ADATE1, "YYYY")
If Year(ADATE1) > 2016 Then at1 = "-- " + Format(ADATE1, "YYYY")

'Clipboard.SetText " " + at1 + Format(ADATE1, " MMMM DD DDDD HH:MM:SS Am/Pm")
'Clipboard.SetText Format(ADATE1, "YYYY-MM-DD")

'--------------------------------------------
If Dir(Label_2_FOLDER.Caption + "\" + Format(ADATE1, "YYYY-MM-DD"), vbDirectory) <> "" Then
    'MsgBox "FOLDER ALREADY CREATED"
Else
    On Error Resume Next
    MkDir Label_2_FOLDER.Caption + "\" + Format(ADATE1, "YYYY-MM-DD")
    If Err.Number > 0 Then
        MsgBox "CREATE FOLDER ERROR" + vbCrLf + Label_2_FOLDER.Caption + "\" + Format(ADATE1, "YYYY-MM-DD") + vbCrLf + Err.Description
    End If
End If

'FN2 = Replace(FILE_NAME, "\", "\")

Beep

End

End Sub

Private Sub MNU_MAKE_FOLDER_FILE_NVMS_DATE_YYYY_MM_DD_Click()


'MAKE FOLDER OF FILE DATE HERE __
'MNU_MAKE_FOLDER_FILE_DATE_YYYY_MM_DD.Caption = "MAKE FOLDER OF FILE DATE HERE __ YYYY-MM-DD"
'--------------------------------------------

If FILE_NAME = "" Then Beep: MsgBox "NOT A FILE NAME PIPED": End

If InStr(FILE_NAME, "NVMS") > 0 Then

    A = A

Else

    MsgBox "NOT AN NVMS FILER", vbMsgBoxSetForeground

End If

Set F = FS.GetFile((FILE_NAME))
ADATE1 = F.DateLastModified

'Clipboard.Clear

If Year(ADATE1) = 2016 Then at1 = "-- 2k Sixteenth"
If Year(ADATE1) = 2015 Then at1 = "-- 2k Fifteenth"
If Year(ADATE1) < 2015 Then at1 = "-- " + Format(ADATE1, "YYYY")
If Year(ADATE1) > 2016 Then at1 = "-- " + Format(ADATE1, "YYYY")

'Clipboard.SetText " " + at1 + Format(ADATE1, " MMMM DD DDDD HH:MM:SS Am/Pm")
'Clipboard.SetText Format(ADATE1, "YYYY-MM-DD")
'Clipboard.SetText AT1

'NVMS_OUTDOOR __________________20170126152921.avi
at1 = Mid(FILE_NAME, InStrRev(FILE_NAME, ".") - Len("20170126152921"), 8)
at1 = Mid(at1, 1, 4) + "-" + Mid(at1, 5, 2) + "-" + Mid(at1, 7, 2)


'--------------------------------------------
If Dir(Label_2_FOLDER.Caption + "\" + at1, vbDirectory) <> "" Then
    'MsgBox "FOLDER ALREADY CREATED"
Else
    On Error Resume Next
    MkDir Label_2_FOLDER.Caption + "\" + at1
    If Err.Number > 0 Then
        MsgBox "CREATE FOLDER ERROR" + vbCrLf + Label_2_FOLDER.Caption + "\" + at1 + vbCrLf + Err.Description
    End If
End If

Beep

End



End Sub

Private Sub MNU_MEDIA_PLAYER_CLASSIC_Click()

'C:\Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64_nvo.exe

'Shell "D:\VB6\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAR\#0 Send To Text List of Files and Sub Folders IRFAN.exe """ + Label_2_FOLDER.Caption + """", vbMaximizedFocus

'Shell "D:\VB6\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAR\#0 Send To Text List of Files and Sub Folders IRFAN.exe """ + "D:\VIDEO\NOT\NOT ME" + """, vbMaximizedFocus"
'Shell "D:\VB6\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAN\#0 Send To Text List of Files and Sub Folders IRFAN.exe """ + "M:\Videos\00 Vid XXX" + """, vbMaximizedFocus"
Shell "D:\VB6\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAN\#0 Send To Text List of Files and Sub Folders IRFAN.exe", vbMaximizedFocus
Beep
Unload Me

End Sub

Private Sub MNU_MEDIA_PLAYER_CLASSIC_LINE_Click()


If GetComputerName = "7-ASUS-GL522VW" Then
    Label_2_FOLDER.Caption = "D:\VIDEO\NOT\X 00 NOT ME"
End If

Shell "D:\VB6\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAN\#0 Send To Text List of Files and Sub Folders IRFAN.exe """ + Label_2_FOLDER.Caption + """", vbMaximizedFocus
Beep
Unload Me

End Sub

Private Sub MNU_NOTEPAD_EBAY_Click()
'I = Label_2_FOLDER.Caption
'I = Replace(I, " ", "*")

'ii = " /nologo /WithSubdirectoriesYES /Dir" + I
'ii = " /nologo /WithSubdirectoriesYES /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\ /FileClipBoard-**.TXT"
'ii = " /nologo /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\ /FileClipBoard-*.TXT"

'ii = " /nologo /DirD:\VB6\VB-NT\00_Best_VB_01\Fast*Clipboard\"
'D:\# MY DOCS\# 01 My Documents\NOTEPAD DOCUMENT\EBAY 2015 TO 2017 PURCHASE HISTORY PASTED ON.txt

'I1 = "/DirD:\# MY DOCS\# 01 My Documents\NOTEPAD DOCUMENT\"
'I2 = "/FileEBAY 2015 TO 2017 PURCHASE HISTORY PASTED ON.txt"
'I1 = Replace(I1, " ", "*")
'I2 = Replace(I2, " ", "*")
'II = " /nologo " + I1 + " " + I2

II = "D:\# MY DOCS\# 01 My Documents\NOTEPAD DOCUMENT\EBAY 2015 TO 2017 PURCHASE HISTORY PASTED ON.txt"


On Error Resume Next

EXE_NAME_LOADER_32 = "C:\Program Files (x86)\Notepad++\notepad++.exe"
EXE_NAME_LOADER_64 = Replace(EXE_NAME_LOADER_32, " (x86)\", "\")

If Dir(EXE_NAME_LOADER_32) <> "" Then
    PID = Shell(EXE_NAME_LOADER_32 + " """ + II + """", vbMaximizedFocus)
    If PID > 0 Then Unload Me: Exit Sub
End If

If Dir(EXE_NAME_LOADER_64) <> "" Then
    PID = Shell(EXE_NAME_LOADER_64 + " """ + II + """", vbMaximizedFocus)
    If PID > 0 Then Unload Me: Exit Sub
End If
'---------------------------------------------------

Me.WindowState = 1
End Sub

Private Sub MNU_NOTEPAD_FILE_Click()

Beep

EXE_NAME_LOADER_32 = "C:\WINDOWS\notepad.exe"
'EXE_NAME_LOADER_64 = Replace(EXE_NAME_LOADER_32, " (x86)\", "\")

If Dir(EXE_NAME_LOADER_32) <> "" Then
    PID = Shell(EXE_NAME_LOADER_32 + " """ + Label_1_FOLDER_FILE.Caption + """", vbMaximizedFocus)
    Beep
    If PID > 0 Then Unload Me: Exit Sub
End If

Me.WindowState = 1

End Sub

Private Sub MNU_NOTEPAD_PLUS_PLUS_FILE_Click()

Beep

EXE_NAME_LOADER_32 = "C:\Program Files (x86)\Notepad++\notepad++.exe"
EXE_NAME_LOADER_64 = Replace(EXE_NAME_LOADER_32, " (x86)\", "\")

If Dir(EXE_NAME_LOADER_32) <> "" Then
    PID = Shell(EXE_NAME_LOADER_32 + " """ + Label_1_FOLDER_FILE.Caption + """", vbMaximizedFocus)
    Beep
    If PID > 0 Then Unload Me: Exit Sub
End If

If Dir(EXE_NAME_LOADER_64) <> "" Then
    PID = Shell(EXE_NAME_LOADER_64 + " """ + Label_1_FOLDER_FILE.Caption + """", vbMaximizedFocus)
    Beep
    If PID > 0 Then Unload Me: Exit Sub
End If
'---------------------------------------------------

Me.WindowState = 1

End Sub

Private Sub MNU_PROCESS_LIST_CONVERT_RENAME_01_Click()

FILE_NAME_01 = "C:\TEMP\MULTI_MENU_PROCESS_LIST_CONVERT_RENAME_INPUT.TXT"
FILE_NAME_02 = "C:\TEMP\MULTI_MENU_PROCESS_LIST_CONVERT_RENAME_OUTPUT.TXT"

FR1 = FreeFile
    Open FILE_NAME_01 For Input As FR1

FR2 = FreeFile
    Open FILE_NAME_02 For Input As FR2

Do
    Line Input #FR1, STRING_PATH_01
    Line Input #FR2, STRING_PATH_02
    
    If Mid(STRING_PATH_01, 1, 2) <> "##" Then
        If Mid(STRING_PATH_01, 1, 2) <> "#;" Then
            If FSO.FileExists(STRING_PATH_01) = True Then
                If FSO.FolderExists(Mid(STRING_PATH_02, 1, InStrRev(STRING_PATH_02, "\") - 1)) = True Then
                    If STRING_PATH_01 <> STRING_PATH_02 Then
                        Debug.Print STRING_PATH_01
                        Debug.Print STRING_PATH_02
                        
                        Name STRING_PATH_01 As STRING_PATH_02
                        I_Count = I_Count + 1
                        DONE_TRUE = True
                    End If
                End If
            End If
        End If
    End If
    
Loop Until EOF(FR1)

Close #FR1
Close #FR2

If DONE_TRUE = True Then
    MsgBox "Done" + Str(I_Count) + " Changes", vbMsgBoxSetForeground
End If
If DONE_TRUE = False Then MsgBox "NONE DONE"
End Sub

Private Sub MNU_PROCESS_LIST_CONVERT_RENAME_02_Click()


'----
'VB6 OPEN A FILE FOR OUTPUT AS UNICODE - Google Search
'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB744GB744&q=VB6+OPEN+A+FILE+FOR+OUTPUT+AS+UNICODE&oq=VB6+OPEN+A+FILE+FOR+OUTPUT+AS+UNICODE&gs_l=psy-ab.3...17156.18600.0.19106.11.11.0.0.0.0.61.615.11.11.0....0...1.1.64.psy-ab..0.0.0....0.hyufwsTcO8c
'--------
'write a file in unicode format.
'http://forums.codeguru.com/showthread.php?315509-write-a-file-in-unicode-format
'----
'Tue 10 October 2017 06:45:00----------

FILE_NAME_02 = "C:\TEMP\MULTI_MENU_PROCESS_LIST_CONVERT_RENAME_OUTPUT.TXT"

FR2 = FreeFile
    Open FILE_NAME_02 For Input As FR2

Dim Test As String

Do
    Line Input #FR2, STRING_PATH_02
    
    If Mid(STRING_PATH_02, 1, 2) <> "##" Then
        If Mid(STRING_PATH_02, 1, 2) <> "#;" Then
            If FSO.FileExists(STRING_PATH_02) = True Then
                If FSO.FolderExists(Mid(STRING_PATH_02, 1, InStrRev(STRING_PATH_02, "\") - 1)) = True Then
                        
                    
                    If InStr(STRING_PATH_02, "MAQ0") > 0 Then
                        
                        Set F = FS.GetFile((STRING_PATH_02))
                        ADATE1 = F.DateLastModified
                    
                        FILE_PATH = Mid(STRING_PATH_02, 1, InStrRev(STRING_PATH_02, "\"))
                        FILE_NAME_2 = Mid(STRING_PATH_02, InStrRev(STRING_PATH_02, "\") + 1)
                        FILE_NAME_2 = Replace(FILE_NAME_2, ".MP4", "")
                        
                        STRING_PATH_03 = ""
                        MAQ = Mid(FILE_NAME_2, InStr(FILE_NAME_2, "MAQ"), 8)
                        If InStrRev(FILE_NAME_2, " _ ") > 0 Then
                            STRING_PATH_03 = Mid(FILE_NAME_2, InStrRev(FILE_NAME_2, " _ "))
                        End If
                        
                        STRING_PATH_03 = Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + STRING_PATH_03
                        STRING_PATH_03 = FILE_PATH + STRING_PATH_03 + ".MP4"
                        
                        If STRING_PATH_02 <> STRING_PATH_03 Then
                            
                            Debug.Print vbCrLf
                            Debug.Print STRING_PATH_02
                            Debug.Print STRING_PATH_03
                            Debug.Print vbCrLf
                           
                            Name STRING_PATH_02 As STRING_PATH_03
                            DONE_TRUE = True
                            I_Count = I_Count + 1
                        End If
                    End If
                End If
            End If
        End If
    End If
    
Loop Until EOF(FR2)

Close #FR1
Close #FR2

If DONE_TRUE = True Then
    MsgBox "Done" + Str(I_Count) + " Changes", vbMsgBoxSetForeground
End If
If DONE_TRUE = False Then MsgBox "NONE DONE"



End Sub

Private Sub MNU_PROCESS_LIST_CONVERT_RENAME_EDITOR_Click()

'----
'DOS COMMAND DIR DISPLAY UNICODE CAFE - Google Search
'https://www.google.co.uk/search?q=DOS+COMMAND+DIR+DISPLAY+UNICODE+CAFE&rlz=1C1CHBD_en-GBGB744GB744&oq=DOS+COMMAND+DIR+DISPLAY+UNICODE+CAFE&aqs=chrome..69i57.12414j0j7&sourceid=chrome&ie=UTF-8
'--------
'Unicode characters in Windows command line - how? - Stack Overflow
'https://stackoverflow.com/questions/388490/unicode-characters-in-windows-command-line-how
'--------
'Command Redirection, Pipes - Windows CMD - SS64.com
'https://ss64.com/nt/syntax-redirection.html
'----
'Tue 10 October 2017 06:50:02----------

FILE_NAME_01 = "C:\TEMP\MULTI_MENU_PROCESS_LIST_CONVERT_RENAME_INPUT.TXT"
FILE_NAME_02 = "C:\TEMP\MULTI_MENU_PROCESS_LIST_CONVERT_RENAME_OUTPUT.TXT"
If Dir("C:\TEMP", vbDirectory) = "" Then MkDir "C:\TEMP"

FR1 = FreeFile
    Open FILE_NAME_01 For Output As FR1
    Print #FR1, StrConv("## HERE THE GIVEN EXAMPLE BEGINNER, INPUT SCRIPT" + vbCrLf, vbUnicode);
    Print #FR1, StrConv("## DIR /B /S ""D:\DSC\2015 2017\2017 CyberShot HX60V\*.MP4"" >>""" + FILE_NAME_01 + """" + vbCrLf, vbUnicode);
    Print #FR1, StrConv("## ----" + vbCrLf, vbUnicode);
Close #FR1

FOLDER_ = "D:\DSC\2015 2017\2015 CyberShot HX60V\MP_ROOT"
FOLDER_ = "D:\VI_ DSC ME\MP_ROOT"
FD = "D:\VIDEO\NOT\ME\2017 SONY MP4\MP_ROOT"
FD = "D:\VIDEO\NOT\ME\2016 SONY MP4\# 2016 a"
FD = "D:\VIDEO\NOT\ME\2015 SONY MP4"
FD = "D:\VI_ DSC ME\2010-2018 CyberShot\2018 CyberShot HX60V\MP_ROOT\189ANV01"
FD = "D:\DSC\2009\# 2009"
FD = "D:\DSC\2008 W880i K800i\# 2008"
FD = "D:\DSC\2005 T68"
FD = "D:\DSC\2015-2018\2016 CyberShot HX60V\MP_ROOT\# TEMP"
'FD=""
'FD=""
'FD=""
'FD=""
'FD=""
'FD=""
'FD=""
'FD=""
'FD=""
'FD=""
'FD=""
'FD=""


FOLDER_ = FD + "\*.MP4"
'FOLDER_ = FD + "\*.JPG"

Shell "CMD /U /C DIR /B /S """ + FOLDER_ + """ >>""" + FILE_NAME_01 + """"


FR1 = FreeFile
    Open FILE_NAME_02 For Output As FR1
    Print #FR1, StrConv("## HERE THE GIVEN EXAMPLE BEGINNER, OUTPUT SCRIPT _ WHEN PROCESS BUTTON_01 WILL RENAME TO THIS RESULT _ WHEN PROCESS BUTTON_02 WILL Do A RENAME HERE AS GENERATOR" + vbCrLf, vbUnicode);
    Print #FR1, StrConv("## DIR /B /S ""D:\DSC\2015 2017\2017 CyberShot HX60V\*.MP4"" >>""" + FILE_NAME_02 + """" + vbCrLf, vbUnicode);
    Print #FR1, StrConv("## ----" + vbCrLf, vbUnicode);
Close #FR1

Shell "CMD /U /C DIR /B /S """ + FOLDER_ + """ >>""" + FILE_NAME_02 + """"
    

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run """" + FILE_NAME_01 + """"
WSHShell.Run """" + FILE_NAME_02 + """"
Set WSHShell = Nothing

End Sub

Private Sub MNU_REGISTER_Click()

'REGSVR32

'Label_1_FOLDER_FILE.Caption
PID = Shell("C:\WINDOWS\system32\regsvr32.exe """ + Label_1_FOLDER_FILE.Caption + """", vbMaximizedFocus)
If PID > 0 Then Beep: Unload Me

End Sub

Private Sub MNU_TAKE_OWNERSHIP_Click()

PID = Shell(BATCH_FILE_, vbMaximizedFocus)
If PID > 0 Then Beep: Unload Me

End Sub

Private Sub MNU_TREESIZE_Click()


I_EXE1 = "C:\Program Files (x86)\JAM Software\TreeSize Free\TreeSizeFree.exe"
I_EXE2 = "C:\Program Files\JAM Software\TreeSize Free\TreeSizeFree.exe"

If Dir(I_EXE1) <> "" Or Dir(I_EXE2) <> "" Then
    'Beep
    On Error Resume Next
    PID = Shell("C:\Program Files (x86)\JAM Software\TreeSize Free\TreeSizeFree.exe " + """" + Label_2_FOLDER.Caption + """", vbMaximizedFocus)
    If PID > 0 Then Beep: Unload Me: Exit Sub
    PID = Shell("C:\Program Files\JAM Software\TreeSize Free\TreeSizeFree.exe " + """" + Label_2_FOLDER.Caption + """", vbMaximizedFocus)
    If PID > 0 Then Beep: Unload Me: Exit Sub
    'C:\Program Files (x86)\JAM Software\TreeSize Free\TreeSizeFree.exe
End If


I_INFO_2 = ""
I_INFO_2 = I_INFO_2 + vbCrLf + vbCrLf + I_EXE1 + vbCrLf + vbCrLf
I_INFO_2 = I_INFO_2 + vbCrLf + vbCrLf + I_EXE2 + vbCrLf + vbCrLf


If Dir(I_EXE1) = "" And Dir(I_EXE2) = "" Then
    I_INFO_1 = "FOLDER FILE EXE NOT FOUND HERE -- " + vbCrLf + vbCrLf
        
    I_INFO = i + INFO_1 + I_INFO_2
    MsgBox I_INFO, vbMsgBoxSetForeground
    Exit Sub
End If

If PID = 0 Then
    I_INFO = "": I_INFO_1 = "": I_INFO_3 = ""
    I_INFO_1 = I_INFO_1 + "PROGRAM EXIST AN ATTEMPT EXECUTE BUT RETURN PID = 0 AND NOT EXECUTED WITH PROCESS ID NUMERIC " + vbCrLf
    I_INFO_1 = I_INFO_1 + "MUST BE AN ADMINSTRATOR PROBLEM -- "
    I_INFO_1 = I_INFO_1 + vbCrLf
    
    I_INFO_3 = I_INFO_3 + "DO YOU WANT TO GO EXPLORER THERE BUTTON -- YES -- " + vbCrLf + "STOP IS DEBUG WHEN IDE=TRUE"
    
    I_INFO = I_INFO_1 + I_INFO_2 + I_INFO_3
    If IsIDE = True Then
        I_R = vbYesNoCancel
    Else
        I_R = vbYesNo
    End If
    
    i = MsgBox(I_INFO, I_R + vbMsgBoxSetForeground)
    
    If i = vbCancel Then Stop
    
    If i = vbYes Then
    
    
        I_EXE1 = "C:\Program Files (x86)\JAM Software\TreeSize Free\TreeSizeFree.exe"
        I_EXE2 = "C:\Program Files\JAM Software\TreeSize Free\TreeSizeFree.exe"
        
        If Dir(I_EXE1) <> "" Or Dir(I_EXE2) <> "" Then
            If Dir(I_EXE1) <> "" Then
                EXE_NAME = I_EXE1
            Else
                EXE_NAME = I_EXE2
            End If
        End If
        
        Shell "EXPLORER /SELECT, """ + EXE_NAME + """", vbNormalFocus
        
        Exit Sub
    End If
    
End If






End Sub


Private Sub MNU_VB_Create_Dir_Another_Drive_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS\Create on another drive\Create Dir Another Drive.exe" + """"
Set WSHShell = Nothing

Beep
Unload Me

End Sub

Private Sub MNU_VB_SYNCRONIZER_Click()
    Dim RUN_EXE
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    RUN_EXE = "D:\VB6\VB-NT\00_Best_VB_01\10 SYNCRONIZE\SYNCRONIZER.exe"
    objShell.Run """" + RUN_EXE + """", 1, False
    Set objShell = Nothing
End Sub

Private Sub MNU_LAUNCH_BATCH_COMPILER_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "D:\VB6\VB-NT\00_Best_VB_01\Batch_Compiler_Auto\BatchCompiler.exe" + """"
Set WSHShell = Nothing
End Sub
Private Sub MNU_LAUNCH_Shell_VBasic_6_Loader_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "D:\VB6\VB-NT\00_Best_VB_01\Shell VBasic 6 Loader\Shell VBasic 6 Loader.exe" + """"
Set WSHShell = Nothing
End Sub


Private Sub MNU_HASH_MY_FILES_Click()

    XF0 = Label_1_FOLDER_FILE.Caption
    XF0 = Label_2_FOLDER.Caption
    
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    OSI = GetOsBitness
    
    If OSI = 64 Then
        HASH_EXE = "C:\PStart\Progs\0_Nirsoft_Package\NirSoft\x64\hashmyfiles.exe"
    Else
        HASH_EXE = "C:\PStart\Progs\0_Nirsoft_Package\NirSoft\hashmyfiles.exe"
    End If
    
    WSHShell.Run """" + HASH_EXE + """ /folder """ + XF0 + """"
    
    Set WSHShell = Nothing

End Sub

Private Sub MNU_WINMERGE_Click()

    XF0 = Label_1_FOLDER_FILE.Caption
        
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    XF1 = GetSettingString(HKEY_CURRENT_USER, "SOFTWARE\Thingamahoochie\WinMerge", "FirstSelection")
    
    If Trim(XF1) = "" Then
        WSHShell.Run """C:\Program Files (x86)\WinMerge\WinMergeU.exe"" /r """ + XF0 + """"
    Else
        WSHShell.Run """C:\Program Files (x86)\WinMerge\WinMergeU.exe"" /r """ + XF1 + """ """ + XF0 + """"
        Call SaveSettingString(HKEY_CURRENT_USER, "SOFTWARE\Thingamahoochie\WinMerge", "FirstSelection", "")
    End If
    Set WSHShell = Nothing
    
    ''DeleteKey "SOFTWARE\Thingamahoochie\WinMerge\FirstSelection"
    ''DeleteKey "Software\MyCompany\MyApp", rrLocalMachine

End Sub


Private Sub Timer2_Timer()

If Clipboard.GetFormat(vbCFText) = False Then Exit Sub
'-------------------------------------
'PROBLEM ERROR READ CLIPBOARD SOMETIME
'CRASH THE PROGRAM
'RETRY AGAIN IS SAFE
'-------------------------------------
On Error GoTo ENDE

If Clipboard.GetText <> OCBGT Then
On Error GoTo 0
    
    
    On Error Resume Next
    Err.Clear
    OCBGT = Clipboard.GetText
    If Err.Number > 0 Then Beep: Exit Sub
    
    
    
    If OCBGT_DOUBLE_CLIPBOARD_VAR = "" Then OCBGT_DOUBLE_CLIPBOARD_VAR = OCBGT
    
    If OCBGT <> OCBGT_DOUBLE_CLIPBOARD_VAR Then
        USE_CLIPBOARD = True
    End If

    If DONT_WANT_CLIPBOARD_AT_START = True Then
        OCBGT_DOUBLE_CLIPBOARD_VAR = OCBGT
        DONT_WANT_CLIPBOARD_AT_START = False
        USE_CLIPBOARD = False
        Exit Sub
    End If
    
    CLIP_BOARD_STRING = Trim(Replace(Trim(Clipboard.GetText), """", ""))
    
    i = 0
    If FS.FolderExists(CLIP_BOARD_STRING) = True Then Label5_CLIPBOARD.Caption = CLIP_BOARD_STRING: i = 1
    If FS.FileExists(CLIP_BOARD_STRING) = True Then Label5_CLIPBOARD.Caption = CLIP_BOARD_STRING: i = 1
    If i = 1 And USE_CLIPBOARD = False Then
        Label4.BackColor = QBColor(14)
        Label5_CLIPBOARD.BackColor = QBColor(15)
        Label4.Caption = "AVAILABLE CLIPBOARD __ PATH_FILE __ EXIST  __ REQUEST IF YOU WANT"
        If i = 1 Then
            Label5.Caption = "PATH_FILE __ SOURCE __ COMMAND$ GIVEN OR CLIPBOARD OPTION"
        End If

    End If
    
    If i = 1 And USE_CLIPBOARD = True Then
        Label4.BackColor = QBColor(10)
        Label5_CLIPBOARD.BackColor = QBColor(15)
        Label4.Caption = "AVAILABLE CLIPBOARD __ PATH_FILE __ EXIST  __ OPERATION DONE"
        If i = 1 Then
            If FIRST_RUN_PATH_FILE_EXIST_DOUBLE = True Then
                Label5.Caption = "PATH_FILE __ SOURCE __ CLIPBOARD CHANGED AT TIMER FINDER"
            Else
                If FIRST_RUN_PATH_FILE_EXIST_DOUBLE = False Then
                    Label5.Caption = "PATH_FILE __ SOURCE __ CLIPBOARD FOUND FROM BEGIN"
                End If
                If FIRST_RUN_PATH_FILE_EXIST_DOUBLE = 3 Then
                    Label5.Caption = "PATH_FILE __ SOURCE __ CLIPBOARD FOUND FROM BEGIN WHEN CHANGED AT TIMER FIND"
                End If
            End If
        End If
    
    End If
    
    If USE_CLIPBOARD = False Then Exit Sub
    
    Call GET_PATH_OR_FILE_PATH(Clipboard.GetText)
    
    Beep
       
    'If PATH_FILE <> "" Then
        Me.SetFocus ': Timer2 = False: OCBGT = ""
        'Call PROCESS_LOAD_DATA
    'End If
End If

ENDE:

End Sub


Private Sub Label_1_FOLDER_FILE_Click()
'Label_1_FOLDER_FILE.CAPTION
Timer1.Enabled = False
Beep
End Sub

Private Sub Label2_Click()

If Dir(Label_2_FOLDER.Caption + "\DOC", vbDirectory) <> "" Then
    'MsgBox "FOLDER ALREADY CREATED"
Else
    On Error Resume Next
    MkDir Label_2_FOLDER.Caption + "\DOC"
    If Err.Number > 0 Then
        MsgBox "CREATE FOLDER ERROR" + vbCrLf + Label_2_FOLDER.Caption + "\DOC" + vbCrLf + Err.Description
    End If
End If
If FS.FileExists(FILE_NAME) = True Then
    
    PATH_FROM_PATH_FILENAME = Mid(FILE_NAME, 1, InStrRev(FILE_NAME, "\"))
    PATH_FROM_PATH_FILENAME = PATH_FROM_PATH_FILENAME + "DOC\"
    
    FILENAME_STRIPED_FROM_PATH = Mid(FILE_NAME, InStrRev(FILE_NAME, "\") + 1)
    FN2 = PATH_FROM_PATH_FILENAME + FILENAME_STRIPED_FROM_PATH
    
    On Error Resume Next
    FS.MoveFile FILE_NAME, FN2
    If Err.Number > 0 Then
    
        DETAIL_VAR = "MOVE FROM --" + vbCrLf + vbCrLf + FILE_NAME
        DETAIL_VAR = "MOVE TO   --" + vbCrLf + vbCrLf + FN2
        MsgBox "WAS A ERROR MOVE FILE REQUESTED" + vbCrLf + vbCrLf + "Err.Description " + vbCrLf + Err.Description + vbCrLf + vbCrLf + DETAIL_VAR, vbMsgBoxSetForeground
    
    End If
    
End If

'FN2 = Replace(FILE_NAME, "\", "\")

End
End Sub

Sub Label_TEXT_LIST_OF_FILE_WITH_LINE_NUMBER_Click()

'AND ADD DESCRIPTON ON ANOTHER NEW LINE

'PASTE TEST
'D:\DSC\2015 2016\2016 Sony CyberShot HX60V\DCIM2\00-FLIKR\A\20160621 -- 01 POST OFFICE

Dir1.Path = Label_2_FOLDER.Caption
File1.Path = Label_2_FOLDER.Caption
File1.FileName = "*.JPG"

TT_NAUGHT = String(Len(Str(File1.ListCount)) - 1, "0")

yy$ = ""
For R1 = 0 To File1.ListCount - 1
    TT$ = File1.List(R1)
    TT$ = Mid$(TT$, InStrRev(TT$, "\") + 1)
    
    'GOT SOME EXTRA LABEL DATA
    If Len(TT$) > 51 Then
        XX22 = Mid(TT$, 1, InStrRev(TT$, " - DSC") + 10) + vbCrLf
        XX23 = XX22 + Mid(TT$, 49)
    Else
        XX23 = TT$ 'Mid(TT$, 38)
    End If
    
    yy$ = yy$ + Format(R1 + 1, TT_NAUGHT) + " of" + Str(File1.ListCount) + " -- " + XX23 + vbCrLf
Next

Clipboard.Clear
Sleep 200
Clipboard.SetText yy$

'End
'Exit Sub

End Sub


Private Sub MNU_TIMER_PROJECT_Click()

Frm_TIMER_PROJECT.Show
Me.WindowState = vbMinimized

End Sub

'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    i = CODER_VBP_FILE_NAME_2
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
    
    CODER_VBP_FILE_NAME_2 = FILE_SPEC_GO
End Sub
Public Sub MNU_VB_ME_Click()
    
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
    objShell.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set objShell = Nothing
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub
    
    
    Beep
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
    'End
End Sub
Private Sub MNU_VB_FOLDER_Click()
    Beep
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + CODER_EXE_FILE_NAME_1 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    Beep
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
    End If
End Sub
'-------------------------------------------
'-------------------------------------------
'---------------------------------------------------
'Private Type SHITEMID
'    cb As Long
'    abID As Byte
'End Type
'Private Type ITEMIDLIST
'    mkid As SHITEMID
'End Type
'Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Function GetSpecialfolder(ByVal CSIDL As Long) As String
    Dim R As Long
    Dim IDL As ITEMIDLIST
    R = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If R = NOERROR Then
        Path$ = Space$(512)
        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function
'---------------------------------------------------



Sub BRING_A_WINDOW_TO_FRONT_THAT_SELECTION_LIKE_VB()
'
'
'
'Dim ReturnHwnd As Long
'Dim I
''VB ONLY WANTS THE 1ST OF THE 2 HWND
''ReturnHwnd = FindWindowPartVB("ClipBoard Logger - Microsoft Visual Basic[")
'
''DONT NEED ABOVE USE THIS
'X1 = FindWindow("wndclass_desked_gsk", vbNullString)
''------------------------------------------------
''FIND THE WINDOW PRIZE TREASURE CHEST BUST BLOSSOM
''TRAIN SPOTTER
''------------------------------------------------
''-----------------------------------------------------------
''CHECK CLASS NAME IS VB IDE WINDOW WE WANT OR IF ANOTHER NOT
''-----------------------------------------------------------
'X2 = GetWindowTitle(X1)
'If InStr(X2, " - Microsoft Visual Basic [") > 0 Then
'If InStr(X2, "ClipBoard Logger") > 0 Then
'    MsgBox "DON'T RUN VB IDE - LOADED"
'    I = GetWindowState(X1)
'    If I = vbMinimized Then
'        SHOWVAR = SW_SHOWDEFAULT
'        ShowWindow ReturnHwnd, SHOWVAR
'    End If
'    EXIT_TRUE = True
'    Unload Me
'    Exit Sub
'End If
'End If
'
''BETTER LINE NEXT DON'T USE VB MENU IF ISIDE
'
''------------------------------------------------
''TEMP WORK AROUND
''OVER DRIVE
''OVER RIDE
''------------------------------------------------
''FINDWINDOW PART PROBLEM IN EXE AND WHEN IN WIN 7
''------------------------------------------------
''WIN 7 PROBLEM MUST USE EXTRA LAST LEFT SQUARE BRACKET OF SERACH END ABOVE
''------------------------------------------------
'If ReturnHwnd > 0 Then
'    If MsgBox("ReturnHwnd" + Str(ReturnHwnd) + " -- " + "MeHwnd" + Str(Me.hWnd) + vbCrLf + "VB Code Clipboard Logger already Running - Sure Want to Run VB IDE", vbYesNo) = vbYes Then
'        WANT_TO_RUN_ANYWAY = True
'    End If
'End If
'
'
'If ReturnHwnd > 0 Then
'
'    I = GetWindowState(ReturnHwnd1)
'
'    If I = vbMinimized Then
'
''        SHOWVAR = SW_RESTORE
''        SHOWVAR = SW_SHOW
''        SHOWVAR = SW_SHOWNA
''        SHOWVAR = SW_SHOWDEFAULT
''        SHOWVAR = SW_SHOWNORMAL
''        SHOWVAR = SW_SHOWMAXIMIZED
'        SHOWVAR = SW_SHOWDEFAULT
'
'        ShowWindow ReturnHwnd, SHOWVAR
'
'        'GUESS MAYBE
'        'SetForegroundWindow (ReturnHwnd)
'
'        DoEvents
'
'    End If
'
'    'MADE REDUNDANT CODE BY CONDICTION HERE AND ABOVE
'    If WANT_TO_RUN_ANYWAY = False Then
'        EXIT_TRUE = True
'        Unload Me
'        Exit Sub
'    End If
'
'End If
'
'If Dir("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE") <> "" Then
'    Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'    EXIT_TRUE = True
'    Unload Me
'End If
'
'If Dir("C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE") <> "" Then
'    Shell """C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'    EXIT_TRUE = True
'    Unload Me
'End If
'

End Sub

Private Sub Timer1_Timer()

Unload Me

End Sub

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************

Private Sub Timer_Form_Load_Opperation_Do_Time_Manager_Timer()

    Call MNU_TIMER_PROJECT_Click
    Timer_Form_Load_Opperation_Do.Enabled = False
    
End Sub


'---------------------------------------------------------------

'http://www.vbforums.com/showthread.php?673134-RESOLVED-Minimum-height-for-menu-bar-to-be-visible
'-------------- MENU SIZE DECLARE
'Private Type rect
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'Private Type MENUBARINFO
'  cbSize As Long
'  rcBar As rect
'  hMenu As Long
'  hwndMenu As Long
'  fBarFocused As Boolean
'  fFocused As Boolean
'End Type
'Private MenuInfo As MENUBARINFO
'Private Const OBJID_MENU As Long = &HFFFFFFFD
'Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
'Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hwnd As Long, _
'ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
'Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

'LOOK AT
'SETUP_DIMENSIONS_INNER_FORM() ---- IN FORM.RESIZE
'USE
'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'Function GetUserName() As String
'   Dim UserName As String * 255
'   Call GetUserNameA(UserName, 255)
'   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
'End Function
'
'Function GetComputerName() As String
'   Dim UserName As String * 255
'   Call GetComputerNameA(UserName, 255)
'   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
'End Function


Private Function Menu_Height()
 
    MenuInfo.cbSize = Len(MenuInfo)
    
    If GetMenuBarInfo(Me.hwnd, OBJID_MENU, 0, MenuInfo) Then
   
        With MenuInfo.rcBar
       
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            'Menu_Height = CStr(.Top) + CStr(.Bottom)
            Menu_Height = CStr(.Bottom) - CStr(.Top)
        End With
       
    End If
   
End Function
'------------------------------------------


Sub SETUP_DIMENSIONS_INNER_FORM()

'DoEvents   ' Yield for other processing.
'Line Method Example

'Line1.BorderWidth = 40
'Border Line On the Edge

'Label_1_FOLDER_FILE.Top = 0
'Label_1_FOLDER_FILE.Width = Form1.Width - 70
'
'Text1.Left = -8
'
'Text1.Top = Label_1_FOLDER_FILE.Height

'On Error Resume Next
'Mnu_Height.Caption = Menu_Height

'VER SIX IS WIN 7
'GOT THICKER BOARDERS OF FORMS WINDOW -- SCROLL BAR MISSING A BIT

'ALL WIN SETUP SEEM TO HAVE OWN SIZE

WIDTH_X = 12000

WIDTH_ADJUST = 70 ' FOR WIN XP
'If GETWinNT_Ver = "WIN XP" Then WIDTH_ADJUST = 70
'If GETWinNT_Ver = "WIN 7" Then WIDTH_ADJUST = 170
'If GETWinNT_Ver = "WIN 10" Then WIDTH_ADJUST = 250

If GetComputerName = "1-ASUS-EEE" Then WIDTH_ADJUST = 70
If GetComputerName = "1-ASUS-X5DIJ" Then WIDTH_ADJUST = 70
If GetComputerName = "3-LINDA-PC" Then WIDTH_ADJUST = 170
If GetComputerName = "4-ASUS-GL522VW" Then WIDTH_ADJUST = 250
If GetComputerName = "5-ASUS-P2520LA" Then WIDTH_ADJUST = 220 '100
If GetComputerName = "7-ASUS-GL522VW" Then WIDTH_ADJUST = 220 '100

HEIGHT_ADJUST = 70 ' FOR WIN XP
'If GETWinNT_Ver = "WIN XP" Then HEIGHT_ADJUST = 70
'If GETWinNT_Ver = "WIN 7" Then HEIGHT_ADJUST = 70
'If GETWinNT_Ver = "WIN 10" Then HEIGHT_ADJUST = 100

'Bigger adjust number smaller inner form
If GetComputerName = "1-ASUS-EEE" Then HEIGHT_ADJUST = 70
If GetComputerName = "1-ASUS-X5DIJ" Then HEIGHT_ADJUST = 70
If GetComputerName = "3-LINDA-PC" Then HEIGHT_ADJUST = 70
If GetComputerName = "4-ASUS-GL522VW" Then HEIGHT_ADJUST = 700
If GetComputerName = "5-ASUS-P2520LA" Then HEIGHT_ADJUST = 700 - 50
If GetComputerName = "7-ASUS-GL522VW" Then HEIGHT_ADJUST = 700

'HIGHER NUMBER SMALLER INNER BOX

'Text1.Width = Form1.Width - WIDTH_ADJUST + 20

'MAKE FORM TALLER OR TEXT BOX
'FORM IS PRIOITY OVER TEXT BOX

i = Label_TEXT_LIST_OF_FILE_WITH_LINE_NUMBER.Top
i = i + Label_TEXT_LIST_OF_FILE_WITH_LINE_NUMBER.Height
INNER_FORM_HEIGHT = i + HEIGHT_ADJUST
On Error Resume Next
Me.Height = (Menu_Height * Screen.TwipsPerPixelY) + INNER_FORM_HEIGHT
Me.Width = WIDTH_X + WIDTH_ADJUST
For Each Control In Controls
    If UCase(Mid(Control.Name, 1, 3)) = "LAB" Then
        Control.Left = 80
        Control.Width = Me.Width - 400
    End If
Next



'XXDD = Me.Height - (Menu_Height * Screen.TwipsPerPixelY) - 500 - INNER_FORM_HEIGHT
'If XXDD > 0 Then Text1.Height = XXDD - HEIGHT_ADJUST

'If Me.Top + Me.Left + Me.Width + Me.Height = OLTLWH Then Exit Sub
'OLTLWH = Me.Top + Me.Left + Me.Width + Me.Height

'Timer_FORM_RESIZE.Enabled = False
'Timer_FORM_RESIZE.Interval = 40
'Timer_FORM_RESIZE.Enabled = True

End Sub

Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

'-------------------------------------------------
'-------------------------------------------------
Private Sub Timer_1_MINUTE_Timer()

Timer_1_MINUTE.Interval = 60000

Call VB_PROJECT_CHECKDATE

End Sub
'-------------------------------------------------
'-------------------------------------------------
Sub VB_PROJECT_CHECKDATE()

'TOP DELCLARE -------------
'DIM XVB_DATE
'--------------------------

If Mid(App.Path, 1, 1) <> "D" Then Exit Sub

PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE'S\")

'---------------------------------------------------
'IF A NEW PROJECT NOT BEEN SYNC YET TO MIRROR FOLDER
'---------------------------------------------------
'MINOR WORK AROUND CREATE PATH
'---------------------------------------------------
If Dir(PATH_FILE_NAME2) = "" Then
    On Error Resume Next
        MkDir Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\"))
        FS.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
    If Err.Number > 0 Then Exit Sub
End If

Set F = FS.GetFile(PATH_FILE_NAME2)

VB_DATE = F.DateLastModified

'--------
'01 OF 02
'----------------------------------
'WRITE INFO ABOUT DATE CHANGE NEWER
'----------------------------------

If XVB_DATE < VB_DATE And XVB_DATE > 0 Then
    
    '-----------------------------------
    'RUN A VBS TO COPY OVER AND RELOADER
    '-----------------------------------
    
'    FR1 = FreeFile
'    Open App.Path + "\RELOAD AND COPY.VBS" For Output As #FR1


    '----------------------------------
    'Mon 10 April 2017 16:30:50--------
    '----------------------------------
    'PROJECT REFERANCE ----------------
    'wshom.ocx
    'IF HARD FIND DO BROWSER
    'IN SYSTEM32 AND SYSWOW64
    'DIR wshom.ocx /S
    '----------------------------------
    'WINDOWS SCRIPT HOST OBJECT MODEL
    'AFTER FIND ALSO HAVE TO SELECT HER
    '----------------------------------
    '----------------------------------
    'Mon 10 April 2017 15:45:11--------
    '----
    'VISUAL BASIC wshom.ocx NOT WORKING - Google Search
    'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB721GB721&q=VISUAL+BASIC+wshom.ocx+NOT+WORKING&oq=VIUAL+BASIC+wshom.ocx+NOT+WORKING&gs_l=serp.3..30i10k1.1533.3987.0.4658.12.12.0.0.0.0.127.888.11j1.12.0....0...1c.1.64.serp..0.12.886...33i160k1j33i21k1.fYCPCpVPeVk
    '--------
    'WScript reference in VB
    'http://forums.codeguru.com/showthread.php?30458-WScript-reference-in-VB
    '----
    'Set objShell = WScript.CreateObject("WScript.Shell") - ------Not WORKER
    '----------------------------------
    '----------------------------------
    
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    WSHShell.Run """" + App.Path + "\RELOAD AND COPY.VBS" + """"
    Set WSHShell = Nothing
    Unload Me
    
End If

XVB_DATE = VB_DATE

End Sub


'Option Explicit
'Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Function GetShortName(ByVal sLongFileName As String) As String
' --> (All this modules code) Obtained from : -
' ---> Microsoft's/MSDN's code
' ->  http://support.microsoft.com/kb/q175512/
' ---> The original comments were by them :

       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       'Set up buffer area for API function call return
       sShortPathName = Space(255)
       iLen = Len(sShortPathName)

       'Call the function
       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       'Strip away unwanted characters.
       GetShortName = Left(sShortPathName, lRetVal)
       If Trim(GetShortName) = "" Then GetShortName = sLongFileName
End Function


Public Function GetLongName(ByVal sShortName As String) As String

    If Mid(sShortName, 1, 2) = "\\" Then
        GetLongName = sShortName
        Exit Function
    End If

' --> (All this modules code) Obtained from : -
' ---> Microsoft's/MSDN's code
' ->  http://support.microsoft.com/default.aspx?scid=kb;EN-US;154822
' ---> The original comments were by them :

     Dim sLongName As String
     Dim sTemp As String
     Dim iSlashPos As Integer
     Dim sShortName_ENTRY As String

     sShortName_ENTRY = sShortName

     'Add \ to short name to prevent Instr from failing
     
     sShortName = sShortName & "\"

     'Start from 4 to ignore the "[Drive Letter]:\" characters
     iSlashPos = InStr(4, sShortName, "\")

     'Pull out each string between \ character for conversion
     While iSlashPos
       sTemp = Dir(Left$(sShortName, iSlashPos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
       If sTemp = "" Then
            'Error 52 - Bad File Name or Number
            GetLongName = ""
            If Trim(GetLongName) = "" Then GetLongName = sShortName_ENTRY
            'SOMETIME SHORT NAME ENTRY WORK SOMETIME NOT
            'MAYBE ERROR IF PARAM IN LINK __ TESTER
            'SOLVED ERROR IS IN PERMISSION OF FOLDER LIKE \PROGRAM FILES (X86)
            '"C:\Program Files (x86)\Process Lasso\ProcessLassoLauncher.exe"
            'SOLVED ERROR 64 BIT IN 32 BIT VB6 SHORTCUT FINDER
            'WORKED ERRRO FROM HERE AGAIN
            '"C:\Program Files (x86)\Process Lasso\ProcessLassoLauncher.exe"
            
            
            
            
             Exit Function
       End If
       sLongName = sLongName & "\" & sTemp
       iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
     Wend

     'Prefix with the drive letter
     GetLongName = Left$(sShortName, 2) & sLongName
    If Trim(GetLongName) = "" Then GetLongName = sShortName_ENTRY
    'SOMETIME SHORT NAME ENTRY WORK SOMETIME NOT
    'MAYBE ERROR IF PARAM IN LINK __ TESTER

End Function






'http://stackoverflow.com/questions/15765804/check-os-and-processor-is-32-bit-or-64-bit

'AddressWidth
'
'On a 32-bit operating system, the value is 32 and on a 64-bit operating system it is 64.
'Relevant VB6 code is:

'HEAVY TASK - SWITCH TO PROBLEM CAN HAPPEN

Public Function GetOsBitness() As String
    Dim ProcessorSet As Object
    Dim CPU As Object

    Set ProcessorSet = GetObject("Winmgmts:"). _
        ExecQuery("SELECT * FROM Win32_Processor")
    For Each CPU In ProcessorSet
        GetOsBitness = CStr(CPU.AddressWidth)
    Next
End Function

'Is processor 32-bit or 64-bit?
'
'For processor one can check DataWidth property:
'
'DataWidth
'
'On a 32-bit processor, the value is 32 and on a 64-bit processor it is 64.
'Relevant VB6 code is:

'HEAVY TASK - SWITCH TO PROBLEM CAN HAPPEN

Public Function GetCpuBitness() As String
    Dim ProcessorSet As Object
    Dim CPU As Object

    Set ProcessorSet = GetObject("Winmgmts:"). _
        ExecQuery("SELECT * FROM Win32_Processor")
    For Each CPU In ProcessorSet
        GetCpuBitness = CStr(CPU.DataWidth)
    Next
End Function


'------------------------------------------------


Public Function GetOsArchitecture()
    If IsAtLeastVista Then
        GetOsArchitecture = GetVistaOsArchitecture
    Else
        GetOsArchitecture = GetXpOsArchitecture
    End If
End Function

Private Function IsAtLeastVista() As Boolean
    IsAtLeastVista = GetOsVersion >= "6.0"
End Function

Private Function GetOsVersion() As String
    Dim OperatingSystemSet As Object
    Dim OS As Object

    Set OperatingSystemSet = GetObject("winmgmts:{impersonationLevel=impersonate}"). _
                                    InstancesOf("Win32_OperatingSystem")
    For Each OS In OperatingSystemSet
        GetOsVersion = Left$(Trim$(OS.Version), 3)
    Next
End Function

Private Function GetVistaOsArchitecture() As String
    Dim OperatingSystemSet As Object
    Dim OS As Object

    Set OperatingSystemSet = GetObject("Winmgmts:"). _
        ExecQuery("SELECT * FROM Win32_OperatingSystem")
    For Each OS In OperatingSystemSet
        GetVistaOsArchitecture = Left$(Trim$(OS.OSArchitecture), 2)
    Next
End Function

Private Function GetXpOsArchitecture() As String
    Dim ComputerSystemSet As Object
    Dim Computer As Object
    Dim SystemType As String

    Set ComputerSystemSet = GetObject("Winmgmts:"). _
        ExecQuery("SELECT * FROM Win32_ComputerSystem")
    For Each Computer In ComputerSystemSet
        SystemType = UCase$(Left$(Trim$(Computer.SystemType), 3))
    Next

    GetXpOsArchitecture = IIf(SystemType = "X86", "32", "64")
End Function


