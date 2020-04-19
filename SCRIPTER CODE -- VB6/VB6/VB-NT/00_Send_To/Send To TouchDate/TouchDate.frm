VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "FormMODIFY DATE AS CREATED DATE1"
   ClientHeight    =   7284
   ClientLeft      =   60
   ClientTop       =   648
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   ScaleHeight     =   7284
   ScaleWidth      =   15720
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1032
      Left            =   12156
      TabIndex        =   39
      Top             =   1548
      Visible         =   0   'False
      Width           =   1104
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   36
      Left            =   13620
      TabIndex        =   38
      Top             =   624
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   35
      Left            =   13068
      TabIndex        =   37
      Top             =   696
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   34
      Left            =   12996
      TabIndex        =   36
      Top             =   612
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   33
      Left            =   12780
      TabIndex        =   35
      Top             =   588
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   32
      Left            =   13392
      TabIndex        =   34
      Top             =   768
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   31
      Left            =   12960
      TabIndex        =   33
      Top             =   840
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   30
      Left            =   13152
      TabIndex        =   32
      Top             =   624
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   29
      Left            =   13068
      TabIndex        =   31
      Top             =   240
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   28
      Left            =   12792
      TabIndex        =   30
      Top             =   540
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   27
      Left            =   13548
      TabIndex        =   29
      Top             =   888
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   26
      Left            =   13116
      TabIndex        =   28
      Top             =   564
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   25
      Left            =   13380
      TabIndex        =   27
      Top             =   1104
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   0
      Left            =   13560
      TabIndex        =   26
      Top             =   444
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "PERFORM ON ALL FILES IN FOLDER OR FILE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Index           =   4
      Left            =   108
      TabIndex        =   2
      Top             =   1152
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "FILE - CREATED DATE TO MODIFY DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Index           =   8
      Left            =   120
      TabIndex        =   10
      Top             =   3468
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "BATCH - CREATED DATE TO MODIFY DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Index           =   7
      Left            =   108
      TabIndex        =   8
      Top             =   2916
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "SET ONE DATE HARDCODER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Index           =   9
      Left            =   108
      TabIndex        =   7
      Top             =   4080
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "MODIFY DATE TO CREATED DATE - NOT WORKING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   528
      Index           =   6
      Left            =   96
      TabIndex        =   1
      Top             =   2340
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "NOW DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Index           =   5
      Left            =   108
      TabIndex        =   0
      Top             =   1800
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "SET OLDER 1ST DATE TO OTHER IN FOLDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Index           =   10
      Left            =   228
      TabIndex        =   11
      Top             =   4764
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "SET ALL DATE FOLDER THE TEXTFILE HOLD DATE WITHIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   11
      Left            =   432
      TabIndex        =   12
      Top             =   5544
      Width           =   10812
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   12
      Left            =   13344
      TabIndex        =   13
      Top             =   768
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   13
      Left            =   12828
      TabIndex        =   14
      Top             =   768
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   14
      Left            =   13416
      TabIndex        =   15
      Top             =   372
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   15
      Left            =   12804
      TabIndex        =   16
      Top             =   336
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   16
      Left            =   13968
      TabIndex        =   17
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   17
      Left            =   13968
      TabIndex        =   18
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   18
      Left            =   13968
      TabIndex        =   19
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   19
      Left            =   13968
      TabIndex        =   20
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   20
      Left            =   13968
      TabIndex        =   21
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   21
      Left            =   13968
      TabIndex        =   22
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   22
      Left            =   13968
      TabIndex        =   23
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   23
      Left            =   13968
      TabIndex        =   24
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   24
      Left            =   13968
      TabIndex        =   25
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      Alignment       =   2  'Center
      BackColor       =   &H00FED3D1&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Index           =   1
      Left            =   12768
      TabIndex        =   4
      Top             =   5028
      Width           =   1632
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "FILE LABEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   528
      Index           =   3
      Left            =   108
      TabIndex        =   9
      Top             =   612
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "FOLDER LABEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   528
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   48
      Width           =   11100
   End
   Begin VB.Label Label_COLOR_GREEN 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label_COLOR_GREEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   12204
      TabIndex        =   6
      Top             =   3924
      Visible         =   0   'False
      Width           =   3072
   End
   Begin VB.Label Label_COLOR_YELLOW 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label_COLOUR_YELLOW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12192
      TabIndex        =   5
      Top             =   3576
      Visible         =   0   'False
      Width           =   3072
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB_FOLDER"
   End
   Begin VB.Menu MNU_QUICK_MODE 
      Caption         =   "QUICK MODE -- OFF -- GIVE MESSENGER"
   End
   Begin VB.Menu MNU_CREATE_TEXT_FILE_FOR_DATE 
      Caption         =   "CREATE TEXT FILE FOR DATE"
   End
   Begin VB.Menu MNU_CREATE_FOLDER_DATE_OF_FILE 
      Caption         =   "CREATE FOLDER DATE OF FILE"
   End
   Begin VB.Menu MNU_CREATE_FOLDER_DATE_MONTH 
      Caption         =   "CREATE FOLDER DATE MONTH"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FILE_NAME_MP4
Dim FILE_NAME_MP4_2
Dim MAQ
Dim START_MAQ
Dim NAME_PART_02

Dim VARCENTER

Dim FORM_ME As New Form1

Dim I_1 ' QUICK MODE RESULT
Dim I_2 ' QUICK MODE RESULT

Dim M_1()
Dim M_2()
Dim M_3()
Dim FR1

Dim FIRST_RUN
Dim a
Dim W$
Dim WORK
Dim FSO

Dim FULL_PATH_AND_FILENAME

'--------------------------------------

'Excellent Code Set the File Time To ANything you want touch date Pro Keep you file time after rotate JPg's




'Option Explicit

'THIS TOGGLES A SMALL PASSWORD FEATURE THAT I EMBEDDED
Const MPassword = False

'THIS IS THE CODE THAT MUST BE ENTERED (FINISHED BY PRESSING HASH)
'WHEN THE PASSWORD TOGGLE IS TRUE
Const MSequence = 22515

Dim MEntered As Variant

 Private Type FILETIME
     dwLowDate  As Long
     dwHighDate As Long
 End Type
 
 Private Type SYSTEMTIME
     wYear      As Integer
     wMonth     As Integer
     wDayOfWeek As Integer
     wDay       As Integer
     wHour      As Integer
     wMinute    As Integer
     wSecond    As Integer
     wMillisecs As Integer
 End Type
 
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_WRITE = &H40000000
  
Private Declare Function CreateFile Lib "kernel32" Alias _
   "CreateFileA" (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) _
   As Long

Private Declare Function LocalFileTimeToFileTime Lib _
     "kernel32" (lpLocalFileTime As FILETIME, _
      lpFileTime As FILETIME) As Long

Private Declare Function SetFileTime Lib "kernel32" _
   (ByVal hFile As Long, ByVal MullP As Long, _
    ByVal NullP2 As Long, lpLastWriteTime _
    As FILETIME) As Long

Private Declare Function SystemTimeToFileTime Lib _
    "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime _
    As FILETIME) As Long
    
Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Function SetFileDateTime(ByVal Filename As String, _
  ByVal TheDate As String) As Boolean
'************************************************
'PURPOSE:    Set File Date (and optionally time)
'            for a given file)
         
'PARAMETERS: TheDate -- Date to Set File's Modified Date/Time
'            FileName -- The File Name

'Returns:    True if successful, false otherwise
'************************************************
If Dir(Filename) = "" Then Exit Function
If Not IsDate(TheDate) Then Exit Function

Dim lFileHnd As Long
Dim lRet As Long

Dim typFileTime As FILETIME
Dim typLocalTime As FILETIME
Dim typSystemTime As SYSTEMTIME

With typSystemTime
    .wYear = Year(TheDate)
    .wMonth = Month(TheDate)
    .wDay = Day(TheDate)
    .wDayOfWeek = Weekday(TheDate) - 1
    .wHour = Hour(TheDate)
    .wMinute = Minute(TheDate)
    .wSecond = Second(TheDate)
End With

lRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)

lFileHnd = CreateFile(Filename, GENERIC_WRITE, _
    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
    OPEN_EXISTING, 0, 0)
    
lRet = SetFileTime(lFileHnd, ByVal 0&, ByVal 0&, _
         typFileTime)

CloseHandle lFileHnd
SetFileDateTime = lRet > 0

End Function


Private Sub Form_Activate()

'LABEL_SET(2).Caption = "F:\DSC\2015+SONY_MP4\2020 CyberShot HX60V __ MP4"
'Call MNU_FILEDATE_WHOLE_FOLDER_Click

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then End

i = MsgBox("", vbMsgBoxSetForeground)

End Sub

Private Sub Form_Load()
Set FSO = CreateObject("Scripting.FileSystemObject")

'If IsIDE = True Then GoTo Start2

On Error Resume Next
If Dir(App.Path + "\# DATA\", vbDirectory) = "" Then
    MkDir App.Path + "\# DATA"
End If

I_1 = "QUICK MODE -- ON -- NONE MESSENGER"
On Error Resume Next
Err.Clear
FR1 = FreeFile
If Dir(App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT") <> "" Then
    Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Input As #FR1
        Line Input #FR1, I_1
        Line Input #FR1, I_2
    Close FR1
End If
I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
I_3 = "QUICK MODE -- ON -- GIVE MESSENGER"
If I_1 = I_3 Then
    If DateValue(I_2) + TimeValue(I_2) < Now Then
        FR1 = FreeFile
        Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Output As #FR1
            Print #FR1, I_1
        Close FR1
    End If
End If
MNU_QUICK_MODE.Caption = I_1

W$ = "D:\VB6\VB-NT\00_Best_VB_01\Auto Start Menu\Auto Start Menu.exe"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\Auto Start Menu\Auto Start Menu.exe"

W$ = "E:\01 FAVS\Ministers\Images"

W$ = "M:\0 00 Art\Google Downloader\00 MICHELLE"
W$ = "U:\# MY DOCS\# 01 My Documents\02_GOOGLE_EARTH\#Google Earth\My Places ASUS BIG\files"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2009 FaceBook Photo's\#\Part 1\DSC02338.JPG"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 FaceBook Photo's\#\MEADOWFIELD HOSPITAL-2010-12-10 21-41-00 - SONY DSC-HX5V - DSC00321.JPG"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 FaceBook Photo's\#\MEADOWFIELD HOSPITAL-2010-12-10 21-41-14 - SONY DSC-HX5V - DSC00322.JPG"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 FaceBook Photo's\#\MEADOWFIELD HOSPITAL-2010-12-10 21-41-26 - SONY DSC-HX5V - DSC00323.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 21 13.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 20 38.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 19 43.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 17 52.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 14 11.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 10 51.JPG"
W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 03 54 49.JPG"

W$ = "M:\DSC\2015\10350614"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\Facebook Holiday Inn Hotel"
W$ = "M:\DSC\2005-2007\2005 NEXT"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-04-19 Kings Hotel"
W$ = "M:\DSC\2010\# TEST"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2008-11-29 Best Western Brighton Stay in On the Run"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2008-12-08 Best Western"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-04-09 Best Western\test"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-04-16 Holiday Inn"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-04-19 Kings Hotel\test"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-10-27 Meadowfield Hospital Maple Ward Asylum\test"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2009 FaceBook London"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Rutland Gardens"
W$ = "D:\Videos\00 Vid XXX\ME\2011 GALAXY SAMSUNG GT-P1000 MP4"
W$ = "D:\Videos\00 Vid XXX\ME"
W$ = "D:\Videos\1999 Accidents\"
W$ = "D:\0 00 ART LOGGERS - WEBCAM\VIDEO\"
W$ = "D:\DSC\# Docus Proofs Texts\# BHT\NOTICE BOARD"

W$ = "D:\DSC\2018 Double Screen Cam\DCIM\2018-10-21"
W$ = "D:\VI_ DSC ME\2010+NOKIA\" '"2015 NOKIA E72 _ 008 AUG _ MP4 _ x001 _ Home Front Room.MP4"
' CARE WHEN MODIFY DATE FROM MOD TO CREATED _
' EXPLORER WITH A DISPLAY OF JPG TREAT AS PICTURE AND NOT THE REAL FILE SYSTEM MOD-DATE SELECTION FOR COLOUMN HEADER
W$ = ""

If Command$ <> "" Or W$ <> "" Then
    If W$ = "" Then W$ = Command$
End If

If Command$ = "" And W$ = "" Then
    If Clipboard.GetFormat(vbCFText) Then
        W1$ = Clipboard.GetText(vbCFText) ' Get Clipboard text.
        If Mid(W1$, 1, 1) = """" Then
            W1$ = Mid(W1$, 2): W1$ = Mid(W1$, 1, Len(W1$) - 1)
        End If
        If FSO.FolderExists(W1$) = True Then
            W$ = W1$
        End If
        If FSO.FileExists(W1$) = True Then
            W$ = W1$
        End If
    End If
End If


If Mid(W$, 1, 1) = """" Then
    W$ = Mid(W$, 2): W$ = Mid(W$, 1, Len(W$) - 1)
End If
FULL_PATH_AND_FILENAME = W$

If FSO.FileExists(W$) Then
    CONVET_FOLDER = True
    W$ = Mid(W$, 1, InStrRev(W$, "\"))
End If

App.Title = "#0 Send To Modify Date With Menu"
Me.Caption = App.Title

If FSO.FolderExists(W$) = False Then
    If CONVET_FOLDER = True Then
        MsgBox "Not Proper Folder Given" + vbCrLf + vbCrLf + "Command$ = " + vbCrLf + Command$ + vbCrLf + vbCrLf + "Convert File to Folder = " + vbCrLf + vbCrLf + W$, vbMsgBoxSetForeground
        End
    Else
'        MsgBox "Not Proper Folder Given" + vbCrLf + vbCrLf + "Command$ = " + vbCrLf + Command$, vbMsgBoxSetForeground
'        End
    End If
End If

If W$ <> "" Then
    If Mid$(W$, Len(W$), 1) <> "\" Then
        W$ = W$ + "\"
    End If
End If

If Mid(W$, 1, 1) = """" Then
    W$ = Mid(W$, 2): W$ = Mid(W$, 1, Len(W$) - 1)
End If

'FOLDER LABEL
LABEL_SET(2).Caption = W$
If W$ = "" Then LABEL_SET(2).Caption = "NOT FOLDER GIVEN"

'FILE LABEL
LABEL_SET(3).Caption = FULL_PATH_AND_FILENAME  ' -- FULL_PATH_AND_FILENAME
If FULL_PATH_AND_FILENAME = "" Then LABEL_SET(3).Caption = "NOT FILE GIVEN"

LABEL_SET(1).BackColor = Label_COLOR_YELLOW.BackColor


End Sub

Private Sub Form_Resize()

Dim r

ReDim M_1(LABEL_SET.Count)
ReDim M_2(LABEL_SET.Count)
ReDim M_3(LABEL_SET.Count)
i = 0

i = i + 1: M_1(i) = "GO"
i = i + 1: M_1(i) = "Folder Label"
i = i + 1: M_1(i) = "File Label"
'i = i + 1: M_1(i) = "PERFORM ON ALL FILES IN FOLDER OR FILE"
i = i + 1: M_1(i) = "NOW_DATE"
i = i + 1: M_1(i) = "MODIFY DATE TO CREATED DATE - NOT WORKING"
i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "BATCH - CREATED DATE TO MODIFY DATE"
i = i + 1: M_1(i) = "FILE  - CREATED DATE TO MODIFY DATE"
i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "DATE CONVERTOR__ MMM D, YYYY H MM AM __ TO YYYY-MM-DD--HH-MM-DD FOR SCREENCASTIFY"
i = i + 0: M_3(i) = "DATE_CONVERTOR___MMM_D__YYYY_H_MM_AM____TO_YYYY_MM_DD__HH_MM_DD_FOR_SCREENCASTIFY"

'i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "ONE_FOLDER AS OTHER SAME ___ HARDCODER __ BATCH"
i = i + 0: M_3(i) = "ONE_FOLDER_AS_OTHER_SAME_____HARDCODER____BATCH"
'i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "MAKE_FOLDER YYYY-MM-DD OF FILE AND MOVE THERE ___ BATCH IT"
i = i + 0: M_3(i) = "MAKE_FOLDER_YYYY_MM_DD_OF_FILE_AND_MOVE_THERE_____BATCH_IT"
'i = i + 1: M_1(i) = "----"

i = i + 1: M_1(i) = "RENAME -- YYYY_MM_DD MMM_DDD HH_MM_SS__MA_.MP4 -- BATCH"
i = i + 0: M_3(i) = "RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA__MP4____BATCH"
i = i + 1: M_1(i) = "----"

i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- CH00_YYYY_MM_DD HH_MM_SS.MP4 -- HIKVISION -- SINGLE"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_CH00_YYYY_MM_DD_HH_MM_SS_MP4_HIKVISION_SINGLE"
i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- YYYY_MM_DD .MP4 .WMV -- WITH_ONE_FOLDER"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_YYYY_MM_DD__WITH_ONE_FOLDER"
i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- YYYY_MM_DD MMM_DDD HH_MM_SS__MA_.MP4 -- SINGLE"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA_SINGLE"
i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- YYYY_MM_DD MMM_DDD HH_MM_SS__MA_.MP4 -- BATCH"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA"
i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- YYYY MM DD HH-MM-SS - DDD NOKIA__.MP4 -- BATCH"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS_DDD_NOKIA_AH"

i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- YYYY MM DD HH-MM-SS_.JPG -- CAMERA -- SINGLE"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS__JPG_SINGLE"

i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "SET_ONE_DATE_HARDCODER"
i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "SET_MOST_RECENT_DATE_TO_OTHER_IN_FOLDER"
i = i + 1: M_1(i) = "SET_OLDER_DATE_TO_OTHER_IN_FOLDER"
i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH_--_\#_TEXT_DATER*.TXT"
i = i + 0: M_3(i) = "SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH"

For r = 0 To i
    M_2(r) = Replace(M_1(r), "_", " ")
Next

For r = 0 To LABEL_SET.Count - 1
    If M_2(r) <> "Folder Label" Then
        If M_2(r) <> "File Label" Then
            If LABEL_SET(r).Caption <> M_2(r) Then
                LABEL_SET(r).Caption = M_2(r)
            End If
        End If
    End If
    If LABEL_SET(r).Caption = "" Then
        LABEL_SET(r).Visible = False
    End If
Next

On Error Resume Next
Me.Width = Screen.Width - 1000 ' 24000
For r = 0 To LABEL_SET.Count - 1
    LABEL_SET(r).FontSize = 12
    LABEL_SET(r).fontname = "ARIAL"
Next

MIN_HL = 4000
For r = 0 To LABEL_SET.Count - 1
    LABEL_SET(r).Refresh
    LABEL_SET(r).AutoSize = True
    LABEL_SET(r).Refresh
    If LABEL_SET(r).Caption <> "" Then
    If LABEL_SET(r).Height < MIN_HL Then
        MIN_HL = LABEL_SET(r).Height
    End If
    End If
    LABEL_SET(r).AutoSize = False
Next
LABEL_SET(2).FontSize = LABEL_SET(4).FontSize - 2
LABEL_SET(3).FontSize = LABEL_SET(2).FontSize

HL = MIN_HL
LENGHT_LABEL = 380
LENGHT_LABEL_AND_WIDTH = Me.Width - LENGHT_LABEL + 100
STEP_H = 100
For r = 2 To LABEL_SET.Count - 1
    If LABEL_SET(r).Caption = "" Then
        LABEL_SET(r).Visible = False
    End If
    
    If LABEL_SET(r).Visible = True Then
        LABEL_SET(r).Left = 100
        LABEL_SET(r).Width = LENGHT_LABEL_AND_WIDTH
        LABEL_SET(r).Height = HL
        LABEL_SET(r).Top = STEP_H
        STEP_H = STEP_H + 40 + HL
    End If
Next

r = 1
LABEL_SET(r).Left = 100
LABEL_SET(r).Width = LENGHT_LABEL_AND_WIDTH
LABEL_SET(r).Height = HL
LABEL_SET(r).Top = STEP_H
STEP_H = STEP_H + 40 + HL

On Error Resume Next
Me.Height = STEP_H + 900 ' + MENUBAR CODE REQUIRE --- REQUIRE HIGH NUMBER IF WINDOWS 10 NOT GRAPHIC DPI CLUE
Me.Top = 0

If Me.WindowState = 0 Then
    If VARCENTER = True Then Exit Sub
    Me.Top = Screen.Height / 2 - Me.Height / 2 '+400
    Me.Left = Screen.Width / 2 - Me.Width / 2 '+400
    VARCENTER = True
    On Error GoTo 0
End If


End Sub


Private Sub LABEL_SET_Click(index As Integer)

' LABEL_SET(3).Caption

DISPLAY_NAMER = M_1(index)
If M_3(index) <> "" Then
    DISPLAY_NAMER = M_3(index)
End If

If DISPLAY_NAMER <> "GO" Then
    WORK = DISPLAY_NAMER
End If


Select Case DISPLAY_NAMER

Case 2
    ' FOLDER BUTTON
    ' LABEL_SET(2)
Case 3
    ' FILE BUTTON
    ' LABEL_SET(3)

Case "PERFORM ON ALL FILES IN FOLDER OR FILE"
    ' Call Label3_Click
    
Case "NOW_DATE"
    'WORK = "MOD_TO_NOW_DATE"

Case "MODIFY DATE TO CREATED DATE - NOT WORKING"
    Call Label2_Click
    
Case "BATCH - CREATED DATE TO MODIFY DATE"
    Call Label9_Click
    
Case "FILE  - CREATED DATE TO MODIFY DATE"
    Call Label11_Click
    
Case "SET_ONE_DATE_HARDCODER"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
    
Case "SET_MOST_RECENT_DATE_TO_OTHER_IN_FOLDER"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_OLDER_DATE_TO_OTHER_IN_FOLDER"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "ONE_FOLDER_AS_OTHER_SAME_____HARDCODER____BATCH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "MAKE_FOLDER_YYYY_MM_DD_OF_FILE_AND_MOVE_THERE_____BATCH_IT"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "DATE_CONVERTOR___MMM_D__YYYY_H_MM_AM____TO_YYYY_MM_DD__HH_MM_DD_FOR_SCREENCASTIFY"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA__MP4____BATCH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_DATE_OF_FILENAME_CH00_YYYY_MM_DD_HH_MM_SS_MP4_HIKVISION_SINGLE"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_DATE_OF_FILENAME_YYYY_MM_DD__WITH_ONE_FOLDER"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA_SINGLE"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS__JPG_SINGLE"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS_DDD_NOKIA_AH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
  
Case "SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
    
Case "GO"
    If LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor Then
        LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor + 30
        LABEL_SET(1).Refresh
        DoEvents
        Call Label_GO_AH_Click
    End If
End Select

End Sub




Sub SET_MOST_RECENT_DATE_TO_OTHER_IN_FOLDER()
    
    ScanPath.txtPath.Text = LABEL_SET(2).Caption
    ' ScanPath.txtPath.Text = "D:\UTILS\2011 GALAXY SAMSUNG GT-P1000 - Copy"
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        
        Set F = FSO.GetFile(A11 + B11)
        DT2 = F.DateCreated
        DT1 = F.datelastmodified
        
        If DT4 = 0 Then DT4 = DT1
        If DT1 > DT4 Then DT4 = DT1
        
        Set F = Nothing
        
    Next
    
    If DT4 = 0 Then
        MsgBox "NAUGHT DATE FOUND IN FILE GATHER -- EXIT"
        End
    End If
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        
        TT = SetFileDateTime(A11 + B11, DT4)
        
        XC = XC + 1
        
    Next
    
    MsgBox "Done = " + vbCrLf + str(XC)
    End

End Sub

Sub SET_OLDER_DATE_TO_OTHER_IN_FOLDER()
    
    ScanPath.txtPath.Text = LABEL_SET(2).Caption
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    'D:\VIDEO\NOT\X 01 ME\2017 SONY MP4\DOC\2017 05 31
    
    For RR = ScanPath.ListView1.ListItems.Count To 1 Step -1
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        If InStr(UCase(B11), ".TXT") > 0 Then
            ScanPath.ListView1.ListItems.Remove (RR)
        End If
    Next
    
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        
        FLAG_OPER = True
        If UCase(B11) = UCase("thumbs.db") Then FLAG_OPER = False
        If UCase(B11) = UCase("DESKTOP.INI") Then FLAG_OPER = False
        
        Set F = FSO.GetFile(A11 + B11)
        DT2 = F.DateCreated
        DT1 = F.datelastmodified
        If FLAG_OPER = True Then
            If DT4 = 0 Then DT4 = DT1
            If DT1 < DT4 Then DT4 = DT1
        End If
        Set F = Nothing
        
    Next
    
    If DT4 = 0 Then
        MsgBox "NAUGHT DATE FOUND IN FILE GATHER -- EXIT"
        End
    End If
    
    DT4 = DT4 + TimeSerial(0, 1, 0)
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        
        TT = SetFileDateTime(A11 + B11, DT4)
        XC = XC + 1
        
        If ScanPath.ListView1.ListItems.Count > 2 Then
        FI = UCase(B11)
        FI = Mid(FI, InStrRev(FI, ".") + 1) + " "
        If InStr("MPG MPEG MP4 JPG AVI ", FI) > 0 Then
            DT4 = DT4 + TimeSerial(0, 1, 0)
        End If
        End If
        
        
    Next
    On Error Resume Next
    FR1 = FreeFile
    Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Input As #FR1
        Line Input #FR1, I_1
    Close FR1
    MNU_QUICK_MODE.Caption = I_1
    I_1 = "QUICK MODE -- ON -- NONE MESSENGER"
    I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC)
    End If
    
    End
End Sub

Sub SET_DATE_OF_FILENAME_YYYY_MM_DD__WITH_ONE_FOLDER()

    ScanPath.chkSubFolders = vbUnchecked
    ScanPath.cboMask.Text = "*.*"
    
    ' SCAN DO ON TEXTPATH CHANGE
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    
    If IsIDE Then
        '2009-02-17-4896_S03_.WMV
        ScanPath.txtPath.Text = "D:\VIDEO\NOT\X 00 NOT ME\00 Vid XXX\BUNKER.COM\2009-06 JUN"
    End If
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    If ScanPath.ListView1.ListItems.Count > 500 Then
        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
        End
    End If
    
    If ScanPath.ListView1.ListItems.Count > 0 Then
        MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    End If
    
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If InStr("WMV MP4 TXT", EXT_STR) And InStr(A11, "_gsdata_") = 0 Then
            DATE_FILENAME_D_1 = Mid(B11, 1, 10)
            DATE_FILENAME_T_2 = "10:00:00" ' Mid(B11, 20, 8) -- CLOCK TIMEZONE IF MIDNIGHT DAY OFFSET OUT KEEP IN
            i2 = DATE_FILENAME_D_1
            i4 = DATE_FILENAME_T_2
            DATE_FILENAME_D_1 = Mid(i2, 1, 4) + "/" + Mid(i2, 6, 2) + "/" + Mid(i2, 9, 2)
            DATE_FILENAME_T_2 = Mid(i4, 1, 2) + ":" + Mid(i4, 4, 2) + ":" + Mid(i4, 7, 2)
            
            DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
            
            DT4 = DATE_FILENAME_SUCCESS
            If IsDate(DT4) = False Then
                ' FALSE DATE
                ' ----------
                MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
                End
            End If
            
            If DT4 = 0 Then
                MsgBox "NAUGHT DATE FINDER -- EXIT"
                End
            End If
        
            TT = SetFileDateTime(A11 + B11, DT4)
            
            EXT_ORG = Mid(B11, Len(B11) - 2)
            EXT_STR = UCase(Mid(B11, Len(B11) - 2))
            If EXT_ORG <> EXT_STR Then
                'RENAME EXTENSION CAPITAL
                FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
                Name A11 + B11 As A11 + FILE_RENAME1
            End If
            
            XC = XC + 1
            MM_1 = MM_1 + B11 + vbCrLf
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End
End Sub

Sub RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA__MP4____BATCH()

Call MNU_FILEDATE_WHOLE_FOLDER_Click

End Sub



Sub SET_DATE_OF_FILENAME_CH00_YYYY_MM_DD_HH_MM_SS_MP4_HIKVISION_SINGLE()
    ' ---------------------------------------------------------------------
    ' ANOTHER SUB ROUTINE
    ' LEECH AND MODIFY FROM BATCH CODE SET
    ' Fri 28-Feb-2020 07:18:00
    ' ---------------------------------------------------------------------
    ' LABEL_SET(3).Caption ' -- FULL_PATH_AND_FILENAME
    ' ---------------------------------------------------------------------
    A11 = LABEL_SET(3).Caption
    D11 = A11
    If A11 <> "NOT FILE GIVEN" Then
        A11 = Mid(D11, 1, InStrRev(A11, "\"))
        B11 = Mid(D11, InStrRev(A11, "\") + 1)
    End If
    If LABEL_SET(2).Caption = "NOT FOLDER GIVEN" Then
'        If IsIDE Then
'            D11 = "F:\DSC--2018+CCSE_HIKVISION_SCREENCASTIFY\2020-01-27--EDDIE\CH01_2020 01 27--16 47 36_BLAGGER_CCTV-1.mp4"
'            A11 = Mid(D11, 1, InStrRev(D11, "\"))
'            B11 = Mid(D11, InStrRev(D11, "\") + 1)
'        End If
    End If
    If Len(D11) < 5 Or D11 = "NOT FILE GIVEN" Then
        MsgBox "FULL PATH AND FILE NOT FIND -- EXIT" + vbCrLf + vbCrLf + D11
        End
    End If
    
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    EXT_STR = UCase(Mid(B11, Len(B11) - 2))
    If InStr("MP4 TXT", EXT_STR) Then
        DATE_FILENAME_D_1 = Mid(B11, 6, 10)
        DATE_FILENAME_T_2 = Mid(B11, 18, 8)
        If IsDate(DATE_FILENAME_D_1) = True Then
            Mid(DATE_FILENAME_D_1, 5, 1) = "/"
            Mid(DATE_FILENAME_D_1, 8, 1) = "/"
            Mid(DATE_FILENAME_T_2, 3, 1) = ":"
            Mid(DATE_FILENAME_T_2, 6, 1) = ":"
        End If
        If IsDate(DATE_FILENAME_D_1) = False Then
            ' CH01_20200225--142741
            ' IN YYYYMMDD
            ' OUT DD-MM-YYYY
            DATE_FILENAME_D_1 = Mid(B11, 12, 2) + "/" + Mid(B11, 10, 2) + "/" + Mid(B11, 6, 4)
            DATE_FILENAME_T_2 = Mid(B11, 16, 2) + ":" + Mid(B11, 18, 2) + ":" + Mid(B11, 20, 2)
        End If
        
        DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
        If IsDate(DATE_FILENAME_SUCCESS) = False Then
        a = a
        End If
        
        
        DT4 = DATE_FILENAME_SUCCESS
        If IsDate(DT4) = False Then
            ' FALSE DATE
            ' ----------
            MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
            End
        End If
        
        If DT4 = 0 Then
            MsgBox "NAUGHT DATE FINDER -- EXIT"
            End
        End If
    
        TT = SetFileDateTime(A11 + B11, DT4)
        
        EXT_ORG = Mid(B11, Len(B11) - 2)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If EXT_ORG <> EXT_STR Then
            'RENAME EXTENSION CAPITAL
            FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
            Name A11 + B11 As A11 + FILE_RENAME1
        End If
        
        XC = XC + 1
        MM_1 = MM_1 + B11 + vbCrLf
    End If
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End

End Sub

Sub SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA_SINGLE()
    ' ---------------------------------------------------------------------
    ' ANOTHER SUB ROUTINE
    ' LEECH AND MODIFY FROM BATCH CODE SET
    ' Wed 29-Jan-2020 06:38:17
    ' ---------------------------------------------------------------------
    ' LABEL_SET(3).Caption ' -- FULL_PATH_AND_FILENAME
    ' ---------------------------------------------------------------------
    
    A11 = LABEL_SET(3).Caption
    ' A11 = "\\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY\2020 CyberShot HX60V\MP4\2020_01_20 Jan_Mon 12_25_38__MAH01294_HOVE GRAND AVENUE DEMOLITION__.MP4"
    D11 = A11
    If A11 <> "NOT FILE GIVEN" Then
        A11 = Mid(D11, 1, InStrRev(A11, "\"))
        B11 = Mid(D11, InStrRev(A11, "\") + 1)
    End If
    
    If Len(D11) < 5 Or D11 = "NOT FILE GIVEN" Then
        MsgBox "FULL PATH AND FILE NOT GOOD -- EXIT" + vbCrLf + vbCrLf + D11
        End
    End If
    
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    EXT_STR = UCase(Mid(B11, Len(B11) - 2))
    If InStr("MP4 TXT", EXT_STR) Then
        DATE_FILENAME_D_1 = Mid(B11, 1, 10)
        DATE_FILENAME_T_2 = Mid(B11, 20, 8)
        i2 = DATE_FILENAME_D_1
        i4 = DATE_FILENAME_T_2
        DATE_FILENAME_D_1 = Mid(i2, 1, 4) + "/" + Mid(i2, 6, 2) + "/" + Mid(i2, 9, 2)
        DATE_FILENAME_T_2 = Mid(i4, 1, 2) + ":" + Mid(i4, 4, 2) + ":" + Mid(i4, 7, 2)
        
        DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
        
        DT4 = DATE_FILENAME_SUCCESS
        If IsDate(DT4) = False Then
            ' FALSE DATE
            ' ----------
            MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
            End
        End If
        
        If DT4 = 0 Then
            MsgBox "NAUGHT DATE FINDER -- EXIT"
            End
        End If
    
        TT = SetFileDateTime(A11 + B11, DT4)
        
        EXT_ORG = Mid(B11, Len(B11) - 2)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If EXT_ORG <> EXT_STR Then
            'RENAME EXTENSION CAPITAL
            FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
            Name A11 + B11 As A11 + FILE_RENAME1
        End If
        
        XC = XC + 1
        MM_1 = MM_1 + B11 + vbCrLf
    End If
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End

End Sub

Sub SET_QUICK_MODE_RESULT()

    ' SET QUICK MODE RESULT
    ' ---------------------
    On Error Resume Next
    FR1 = FreeFile
    Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Input As #FR1
        Line Input #FR1, I_1
    Close FR1
    MNU_QUICK_MODE.Caption = I_1
    I_1 = "QUICK MODE -- ON -- NONE MESSENGER"
    I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"

End Sub

Sub DATE_CONVERTOR___MMM_D__YYYY_H_MM_AM____TO_YYYY_MM_DD__HH_MM_DD_FOR_SCREENCASTIFY()

    ScanPath.chkSubFolders = vbUnchecked
    ScanPath.cboMask.Text = "*.MP4"
    
    ' SCAN DO ON TEXTPATH CHANGE
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    If LABEL_SET(2).Caption = "NOT FOLDER GIVEN" Then
        If IsIDE Then
            ScanPath.txtPath.Text = "F:\DSC--SCREENCASTIFY__GD_CLOUD"
        End If
    End If
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
'    If ScanPath.ListView1.ListItems.Count > 500 Then
'        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
'        End
'    End If
    
    'If ScanPath.ListView1.ListItems.Count > 0 Then
        MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    'End If
    
    
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If InStr("MP4 TXT", EXT_STR) And InStr(A11, "_gsdata_") = 0 Then
            
            Set F = FSO.GetFile(A11 + B11)
            DT1 = 0
            DT1 = F.datelastmodified
            DT1 = 0
            'Aug 5, 2019 9 57 Am-4
            DATE_STRING = ""
            If InStrRev(B11, "m-") > 0 Then
                THREE_BACK_SPACE = InStrRev(B11, " ", Len(B11))
                For r = 1 To 2
                    THREE_BACK_SPACE = InStrRev(B11, " ", THREE_BACK_SPACE - 1)
                Next
                DATE_STRING_1 = Mid(B11, 1, THREE_BACK_SPACE)
                DATE_STRING_2 = Mid(B11, THREE_BACK_SPACE, InStrRev(B11, "m-") - THREE_BACK_SPACE + 1)
                DATE_STRING_2 = Trim(DATE_STRING_2)
            End If
            DATE_STRING_2 = Replace(DATE_STRING_2, " ", ":", , 1)
            If DATE_STRING_1 <> "" Then
            If IsDate(DATE_STRING_1) Then
                DT1 = DateValue(DATE_STRING_1)
            End If
            End If
            
            If DATE_STRING_2 <> "" Then
            If IsDate(DATE_STRING_2) Then
                DT1 = DT1 + TimeValue(DATE_STRING_2)
            End If
            End If
            
            DT1_STRING = Format(DT1, "YYYY-MM-DD--HH-MM-SS")
            WHAT_AT_AFTER_MARK_DASH = Mid(B11, InStrRev(B11, "-"))
            DT1_STRING = DT1_STRING + WHAT_AT_AFTER_MARK_DASH
            
            If DT1 > 0 Then
                
                On Error Resume Next
                Err.Clear
                Name A11 + "\" + B11 As A11 + "\" + DT1_STRING
                If Err.Number > 0 Then
                
                    DETAIL_VAR = "RENAME FROM --" + vbCrLf + vbCrLf + A11 + "\" + B11
                    DETAIL_VAR = "MOVE TO   --" + vbCrLf + vbCrLf + A11 + "\" + DT1_STRING
                    MsgBox "WAS A ERROR RENAME FILE REQUEST" + vbCrLf + vbCrLf + "Err.Description = " + Err.Description + vbCrLf + vbCrLf + DETAIL_VAR, vbMsgBoxSetForeground
                
                End If
                TT = SetFileDateTime(A11 + "\" + DT1_STRING, DT1)
                
                XC = XC + 1
                MM_1 = MM_1 + B11 + vbCrLf
            End If
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End


End Sub

Sub ONE_FOLDER_AS_OTHER_SAME_____HARDCODER____BATCH()

    Set FSO = CreateObject("Scripting.FileSystemObject")

    DR_1 = "C:\DD\ABBYWINTERS.COM\"
    DR_2 = "C:\DD\"
    ScanPath.chkSubFolders = vbUnchecked
    ScanPath.cboMask.Text = "*.MP4;*.WMV"
    
    ' SCAN DO ON TEXTPATH CHANGE
    ScanPath.txtPath.Text = DR_1
    
    'If ScanPath.ListView1.ListItems.Count > 0 Then
        'MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    'End If
    
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        B11_NOT_EXT = Mid(B11, 1, Len(B11) - 4)
        ' WANT PROCESS WITH -- MP4 WMV -- AND THEN PUT THEM HERE
        ' ----------------------------------------------------------
        If InStr("MP4 WMV", EXT_STR) And InStr(A11, "_gsdata_") = 0 Then
            
            Set F = FSO.GetFile(A11 + B11)
            DT1 = F.datelastmodified

            On Error Resume Next
            Err.Clear
            
            OUT_FILE_MP4 = DR_2 + B11_NOT_EXT + ".MP4"
            OUT_FILE_WMV = DR_2 + B11_NOT_EXT + ".WMV"
            If Dir(OUT_FILE_MP4) <> "" Then
                TT = SetFileDateTime(OUT_FILE_MP4, DT1)
                Name OUT_FILE_MP4 As OUT_FILE_MP4
                XC = XC + 1
                MM_1 = MM_1 + B11 + vbCrLf
            End If
            If Err.Number > 0 Then MsgBox "EEROR WITH " + vbCrLf + vbCrLf + OUT_FILE_MP4
            If Dir(OUT_FILE_WMV) <> "" Then
                TT = SetFileDateTime(OUT_FILE_WMV, DT1)
                Name OUT_FILE_WMV As OUT_FILE_WMV
                XC = XC + 1
                MM_1 = MM_1 + B11 + vbCrLf
            End If
            If Err.Number > 0 Then MsgBox "EEROR WITH " + vbCrLf + vbCrLf + OUT_FILE_WMV
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    ' If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    ' End If
    
    End


End Sub

Sub MAKE_FOLDER_YYYY_MM_DD_OF_FILE_AND_MOVE_THERE_____BATCH_IT()

    ScanPath.chkSubFolders = vbUnchecked
    ScanPath.cboMask.Text = "*.MP4;*.MP3;*.WAV"
    
    ' SCAN DO ON TEXTPATH CHANGE
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    If IsIDE Then
        ' ScanPath.txtPath.Text = "F:\DSC--2018+CCTV_HIKVISION\2018+CCTV_HIKVISION_DS-7204H\2020_CCTV_HIKVISION_DS-7204H"
        ScanPath.txtPath.Text = "D:\0 00 MOBILE-2\RETEKESS\RETEKESS_RECORD 2020\MRECORD"
    End If
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
'    If ScanPath.ListView1.ListItems.Count > 500 Then
'        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
'        End
'    End If
    
    'If ScanPath.ListView1.ListItems.Count > 0 Then
        MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    'End If
    
    
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        ' WANT PROCESS WITH -- MP4 TXT MP3 -- AND THEN PUT THEM HERE
        ' ----------------------------------------------------------
        If InStr("MP4 TXT MP3", EXT_STR) And InStr(A11, "_gsdata_") = 0 Then
            
            Set F = FSO.GetFile(A11 + B11)
            DT1 = F.datelastmodified

            OUT_FOLDER__ = A11 + "\" + Format(DT1, "YYYY-MM-DD")
            OUT_FOLDER_AND_FILENAME = A11 + "\" + Format(DT1, "YYYY-MM-DD") + "\" + B11
'            If InStr(LABEL_SET(3).Caption, "REC") > 0 And InStr(LABEL_SET(3).Caption, ".WAV") > 0 Then
'                OUT_FOLDER_AND_FILENAME = LABEL_SET(2).Caption + "\RECORD " + Format(DT1, "YYYY-MM-DD")
'            End If
            'CREATE
            If Dir(OUT_FOLDER__, vbDirectory) = "" Then
                CreateFolderTree OUT_FOLDER__
            End If
            FILE_NAME = A11 + B11
            On Error Resume Next
            Err.Clear
            FSO.MoveFile FILE_NAME, OUT_FOLDER_AND_FILENAME
            If Err.Number > 0 Then
            
                DETAIL_VAR = "MOVE FROM --" + vbCrLf + vbCrLf + FILE_NAME
                DETAIL_VAR = "MOVE TO   --" + vbCrLf + vbCrLf + OUT_FOLDER_AND_FILENAME
                MsgBox "WAS A ERROR MOVE FILE REQUESTED" + vbCrLf + vbCrLf + "Err.Description = " + Err.Description + vbCrLf + vbCrLf + DETAIL_VAR, vbMsgBoxSetForeground
            
            End If
            
            XC = XC + 1
            MM_1 = MM_1 + B11 + vbCrLf
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End


End Sub


Sub SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA()

    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.MP4"
    
    ' SCAN DO ON TEXTPATH CHANGE
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    If IsIDE Then
        ScanPath.txtPath.Text = "D:\DD"
    End If
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    If ScanPath.ListView1.ListItems.Count > 500 Then
        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
        End
    End If
    
    If ScanPath.ListView1.ListItems.Count > 0 Then
        MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    End If
    
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If InStr("MP4 TXT", EXT_STR) And InStr(A11, "_gsdata_") = 0 Then
            DATE_FILENAME_D_1 = Mid(B11, 1, 10)
            DATE_FILENAME_T_2 = Mid(B11, 20, 8)
            i2 = DATE_FILENAME_D_1
            i4 = DATE_FILENAME_T_2
            DATE_FILENAME_D_1 = Mid(i2, 1, 4) + "/" + Mid(i2, 6, 2) + "/" + Mid(i2, 9, 2)
            DATE_FILENAME_T_2 = Mid(i4, 1, 2) + ":" + Mid(i4, 4, 2) + ":" + Mid(i4, 7, 2)
            
            DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
            
            DT4 = DATE_FILENAME_SUCCESS
            If IsDate(DT4) = False Then
                ' FALSE DATE
                ' ----------
                MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
                End
            End If
            
            If DT4 = 0 Then
                MsgBox "NAUGHT DATE FINDER -- EXIT"
                End
            End If
        
            TT = SetFileDateTime(A11 + B11, DT4)
            
            EXT_ORG = Mid(B11, Len(B11) - 2)
            EXT_STR = UCase(Mid(B11, Len(B11) - 2))
            If EXT_ORG <> EXT_STR Then
                'RENAME EXTENSION CAPITAL
                FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
                    Name A11 + B11 As A11 + FILE_RENAME1
            End If
            
            XC = XC + 1
            MM_1 = MM_1 + B11 + vbCrLf
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End
End Sub



Sub SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS__JPG_SINGLE()
    ' ---------------------------------------------------------------------
    ' ANOTHER SUB ROUTINE
    ' LEECH AND MODIFY FROM BATCH CODE SET
    ' Mon 30-Mar-2020 22:51:11
    ' ---------------------------------------------------------------------
    ' LABEL_SET(3).Caption ' -- FULL_PATH_AND_FILENAME
    ' ---------------------------------------------------------------------
    A11 = LABEL_SET(3).Caption
    ' A11 = "D:\DSC\2015+SONY\2020 CyberShot HX60V\JPG\2020 03 30\2020-03-30 13-05-50 - SONY DSC-HX60V - DSC02044 - EDITOR.JPG"
    D11 = A11
    If A11 <> "NOT FILE GIVEN" Then
        A11 = Mid(D11, 1, InStrRev(A11, "\"))
        B11 = Mid(D11, InStrRev(A11, "\") + 1)
    End If
    
    If Len(D11) < 5 Or D11 = "NOT FILE GIVEN" Then
        MsgBox "FULL PATH AND FILE NOT GOOD -- EXIT" + vbCrLf + vbCrLf + D11
        End
    End If
    
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    EXT_STR = UCase(Mid(B11, Len(B11) - 2))
    If InStr("JPG TXT", EXT_STR) Then
        DATE_FILENAME_D_1 = Mid(B11, 1, 10)
        DATE_FILENAME_T_2 = Mid(B11, 12, 8)
        i2 = DATE_FILENAME_D_1
        i4 = DATE_FILENAME_T_2
        DATE_FILENAME_D_1 = Mid(i2, 1, 4) + "/" + Mid(i2, 6, 2) + "/" + Mid(i2, 9, 2)
        DATE_FILENAME_T_2 = Mid(i4, 1, 2) + ":" + Mid(i4, 4, 2) + ":" + Mid(i4, 7, 2)
        
        DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
        
        DT4 = DATE_FILENAME_SUCCESS
        If IsDate(DT4) = False Then
            ' FALSE DATE
            ' ----------
            MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
            End
        End If
        
        If DT4 = 0 Then
            MsgBox "NAUGHT DATE FINDER -- EXIT"
            End
        End If
    
        TT = SetFileDateTime(A11 + B11, DT4)
        
        EXT_ORG = Mid(B11, Len(B11) - 2)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If EXT_ORG <> EXT_STR Then
            'RENAME EXTENSION CAPITAL
            FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
            Name A11 + B11 As A11 + FILE_RENAME1
        End If
        
        XC = XC + 1
        MM_1 = MM_1 + B11 + vbCrLf
    End If
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End

End Sub


Sub SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS_DDD_NOKIA_AH()

    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    
    ScanPath.txtPath.Text = LABEL_SET(2).Caption
'   ScanPath.txtPath.Text = "D:\VI_ DSC ME\2010+NOKIA\2017 NOKIA E72 _ 007 JULY_ MP4____x003 _ Mill View Room"
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        
        DATE_FILENAME_D_1 = Mid(B11, 1, 10)
        DATE_FILENAME_T_2 = Mid(B11, 12, 8)
        i2 = DATE_FILENAME_D_1
        i4 = DATE_FILENAME_T_2
        DATE_FILENAME_D_1 = Mid(i2, 1, 4) + "/" + Mid(i2, 6, 2) + "/" + Mid(i2, 9, 2)
        DATE_FILENAME_T_2 = Mid(i4, 1, 2) + ":" + Mid(i4, 4, 2) + ":" + Mid(i4, 7, 2)
        
        DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
        
        DT4 = DATE_FILENAME_SUCCESS
        If IsDate(DT4) = False Then
            ' FALSE DATE
            ' ----------
            MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
            End
        End If
        
        If DT4 = 0 Then
            MsgBox "NAUGHT DATE FINDER -- EXIT"
            End
        End If
    
        TT = SetFileDateTime(A11 + B11, DT4)
        XC = XC + 1
        MM_1 = MM_1 + B11 + vbCrLf
    Next
    
    ' SET QUICK MODE RESULT
    ' ---------------------
    On Error Resume Next
    FR1 = FreeFile
    Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Input As #FR1
        Line Input #FR1, I_1
    Close FR1
    MNU_QUICK_MODE.Caption = I_1
    I_1 = "QUICK MODE -- ON -- NONE MESSENGER"
    I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End
End Sub
    
Sub SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH()
    
    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    
    ScanPath.txtPath.Text = LABEL_SET(2).Caption
'    ScanPath.txtPath.Text = "D:\VIDEO\NOT\X 00 NOT ME\00 Vid XXX\00 ROOT_02_(DATE-ALPHABETICAL)\1984 a11 -- NURSE -- 1984"
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        
        Set F = FSO.GetFile(A11 + B11)
        DT2 = F.DateCreated
        DT1 = F.datelastmodified
        
        If DT4 = 0 Then DT4 = DT1
        If DT1 < DT4 Then DT4 = DT1
        
        Set F = Nothing
    Next
    
    XX = Dir(ScanPath.txtPath.Text + "\# TEXT_DATER*.TXT")
    If XX = "" Then
        XX = Dir(ScanPath.txtPath.Text + "\#_TEXT_DATER*.TXT")
    End If
    If XX = "" Then
        XX = Dir(ScanPath.txtPath.Text + "\# TEXT DATER*.TXT")
    End If
    If XX = "" Then
        MsgBox "NONE (DATE WITHIN TEXT FILE) FOUND HERE" + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\# TEXT_DATER*.txt" + vbCrLf + vbCrLf + "-- EXIT"
        End
    End If
        
    FR1 = FreeFile
    Open ScanPath.txtPath.Text + "\" + XX For Input As #FR1
        Line Input #FR1, TIME_STRING
    Close FR1
    
    TIME_STRING = Replace(TIME_STRING, "\", "/")
    TIME_STRING = Replace(TIME_STRING, "_", ":")
    
    If InStr(ScanPath.txtPath.Text, "XXX") > 0 Then
        If Val(TimeValue(TIME_STRING)) = 0 Then TIME_STRING = Format(DateValue(TIME_STRING), "YYYY/MM/DD") + " 10:00:00"
    End If
    DT4 = DateValue(TIME_STRING) + TimeValue(TIME_STRING)
    If IsDate(DT4) = False Then
        ' FALSE DATE
        ' ----------
        MsgBox "DATE FOUND WITHIN TEXT FILE FOUND HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
        End
    End If
    
    If DT4 = 0 Then
        MsgBox "NAUGHT DATE FOUND OVERALL IN FILE GATHER -- EXIT"
        End
    End If
    
    If TimeValue(TIME_STRING) = 0 Then
        ' ADD A BIT ON DATE FOR DAYLIGHT SAVING AND NTFS AT EXACT MIDNIGHT
        ' CLOUD SYSTEM REQUIRING A BIT MORE ALSO OR DAY BEFORE
        ' ----------------------------------------------------------------
        DT4 = DT4 + TimeSerial(11, 0, 0)
        ' ----------------------------------------------------------------
    End If
    MM_1 = Format(DT4, "DD-MM-YYYY  HH:MM:SS  DDDD")
    MM_1 = MM_1 + vbCrLf
    MM_1 = MM_1 + ScanPath.ListView1.ListItems.Item(1).SubItems(1) + vbCrLf
    MM_1 = MM_1 + vbCrLf
    COUNT_FILE = COUNT_FILE + 1
    If COUNT_FILE > 1 Then
        DT4 = DT4 + TimeSerial(0, 1, 0)
    End If
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        TT = SetFileDateTime(A11 + B11, DT4)
        XC = XC + 1
        MM_1 = MM_1 + B11 + vbCrLf
        FI = UCase(B11)
        FI = Mid(FI, InStrRev(FI, ".") + 1) + " "
        If InStr(" MPG MPEG MP4 JPG AVI ", FI) > 0 Then
            DT4 = DT4 + TimeSerial(0, 1, 0)
        End If
    Next
    
    On Error Resume Next
    FR1 = FreeFile
    Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Input As #FR1
        Line Input #FR1, I_1
    Close FR1
    MNU_QUICK_MODE.Caption = I_1
    I_1 = "QUICK MODE -- ON -- NONE MESSENGER"
    I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End
End Sub


Sub BATCH_MODIFIED_TO_CREATED_TIME()
    
On Error Resume Next
XX = FSO.FolderExists(W$)

Dim DateSet As Date

If XX = False Then
'time2$ = "2011-11-01 10:00:00"
'DateSet = DateValue(time2$) + TimeValue(time2$)
'DateSet = Now
'
'tt = SetFileDateTime(W$, DateSet)

'tt = LastModifiedToCurrent(W$)
End
End If

' If DIRW$ <> "" Then W$ = DIRW$

'MsgBox str(XX) + "--" + W$
'End

Start2:

Dim TPath$

ScanPath.chkSubFolders = vbChecked
ScanPath.cboMask.Text = "*.*"

'ScanPath.txtPath.Text = "E:\01 FAVS\#00 Palm\#000 New\2009-08 Aug\Unet\Mine\Word Example Usage"
'ScanPath.txtPath.Text = "E:\01 FAVS\Ministers\Images"
ScanPath.txtPath.Text = W$

If Len(W$) < 5 Then End

'ScanPath.Show
Dim DT1 As Date
'Dim Ds1 As Date
Dim DS2 As Date
'Dim DateSet As Date
For we = 1 To ScanPath.ListView1.ListItems.Count
    A11 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B11 = ScanPath.ListView1.ListItems.Item(we)
    
    Set F = FSO.GetFile(A11 + B11)
    DT3 = F.DateCreated
    DT1 = F.datelastmodified
    
    DateSet = DT3
    
    Set F = Nothing
    
    TT = SetFileDateTime(A11 + B11, DateSet)
    XC = XC + 1
    
Next

MsgBox "Done = " + vbCrLf + str(XC) '+dd$
End


End Sub

Sub BATCH_CREATED_TO_MODIFIED_TIME()
    
On Error Resume Next
XX = FSO.FolderExists(W$)

Dim DateSet As Date

'If XX = False Then
''time2$ = "2011-11-01 10:00:00"
''DateSet = DateValue(time2$) + TimeValue(time2$)
''DateSet = Now
''tt = SetFileDateTime(W$, DateSet)
''tt = LastModifiedToCurrent(W$)
'End
'End If

Start2:

Dim TPath$

ScanPath.chkSubFolders = vbChecked
ScanPath.cboMask.Text = "*.*"

'ScanPath.txtPath.Text = "E:\01 FAVS\#00 Palm\#000 New\2009-08 Aug\Unet\Mine\Word Example Usage"
'ScanPath.txtPath.Text = "E:\01 FAVS\Ministers\Images"
ScanPath.txtPath.Text = W$

If Len(W$) < 5 Then End

'ScanPath.Show
Dim DT1 As Date
'Dim Ds1 As Date
Dim DS2 As Date
I_MSG = ""
For we = 1 To ScanPath.ListView1.ListItems.Count
    A11 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B11 = ScanPath.ListView1.ListItems.Item(we)
    
    Set F = FSO.GetFile(A11 + B11)
    DT3 = F.DateCreated
    DT1 = F.datelastmodified
    Set F = Nothing
    
    TT = SetFileDateTime(A11 + B11, DT3)
    XC = XC + 1
    
    I_MSG = I_MSG + B11 + vbCrLf
    
Next

MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + I_MSG '+dd$
End

End Sub


Sub FILE_CREATED_TO_MODIFIED_TIME()
        
    'On Error Resume Next
    XX = FSO.FileExists(FULL_PATH_AND_FILENAME)
    
    Dim DateSet As Date
    
    Dim DT1 As Date
    Dim DT3 As Date
    I_MSG = ""
        
    Set F = FSO.GetFile(FULL_PATH_AND_FILENAME)
    DT3 = F.DateCreated
    DT1 = F.datelastmodified
    Set F = Nothing
    
    TT = SetFileDateTime(FULL_PATH_AND_FILENAME, DT3)
    XC = XC + 1
    
    I_MSG = I_MSG + FULL_PATH_AND_FILENAME + vbCrLf
    
    
    MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + I_MSG '+dd$
    End
    
End Sub



Sub CODE_RUN()

Exit Sub

If Mid(W$, 1, 1) = """" Then
    W$ = Mid(W$, 2): W$ = Mid(W$, 1, Len(W$) - 1)
End If

If Dir$(W$) <> "" Then
    DIRW$ = Mid$(W$, 1, InStrRev(W$, "\") - 1)
End If
    
If Mid$(W$, Len(W$), 1) <> "\" Then
    W$ = W$ + "\"
End If
    
On Error Resume Next
XX = FSO.FolderExists(W$)

Dim DateSet As Date

If DIRW$ <> "" Then W$ = DIRW$

'MsgBox str(XX) + "--" + W$
'End

Start2:

Dim TPath$

ScanPath.chkSubFolders = vbChecked
ScanPath.cboMask.Text = "*.avi"

'ScanPath.txtPath.Text = "E:\01 FAVS\#00 Palm\#000 New\2009-08 Aug\Unet\Mine\Word Example Usage"
'ScanPath.txtPath.Text = "E:\01 FAVS\Ministers\Images"
ScanPath.txtPath.Text = W$

If Len(W$) < 5 Then End

'ScanPath.Show
Dim DT1 As Date
'Dim Ds1 As Date
Dim DS2 As Date
'Dim DateSet As Date
For we = 1 To ScanPath.ListView1.ListItems.Count
    A11 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B11 = ScanPath.ListView1.ListItems.Item(we)
    
    Set F = FSO.GetFile(A11 + B11)
    DT1 = F.datelastmodified
    
    DateSet = DT1 - TimeSerial(1, 0, 0)
    
    Set F = Nothing
    
    'Ds1 = Mid(B11, 1, 20)
'    Ds1 = Replace(Ds1, "+", ":") + ":00"
'    If InStr(B11, "2008 011(Nov) 29 Sat 08-04-10") > 0 Then
    'If Year(DT1) = 2009 Then
        '2008 012(Dec) 04 Thu 18-22-52 - W880I - DSC00331.JPG
      '  Ds1 = DateValue("04-Dec-2008") + TimeValue("18:22:52")
        
'        xB11 = B11
'        xB11 = Replace(xB11, "MILL VIEW HOSPITAL-", "")
'
        If A11 + B11 = "D:\0 00 ART LOGGERS - WEBCAM\VIDEO\Web_Video_Capture- 2015-10-03 14-28-59 Sat -- Microsoft WDM Image Capture (Win32)2.avi" Then
'            Ds1 = DateValue(Mid(xB11, 9, 2) + "-" + Mid(xB11, 6, 2) + "-" + Mid(xB11, 1, 4))
'            Ds2 = TimeValue(Replace(Mid(xB11, 12, 8), "-", ":"))
        '    DateSet = DateValue(Ds1) + TimeValue(Ds1)
'            DateSet = Ds1 + Ds2
    '        DateSet = GetExif(A11 + B11)
            DateSet = DateValue("2015-10-03") + TimeValue("14:28:59")
                  
            '
            TT = SetFileDateTime(A11 + B11, DateSet)
            XC = XC + 1
            'tt = LastModifiedToCurrent(A11 + B11)
        End If
'        Stop
        'End
    'End If
    
Next

MsgBox "Done = " + vbCrLf + str(XC) '+dd$
End


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


Function GetExif(File$)
    Dim objExif As New ExifReader
    Dim txtExifInfo As String
    
    objExif.Load File$
    txtExifInfo = objExif.Tag(DateTimeOriginal)
    GetExif = txtExifInfo
End Function


'Function GetProperty(strFile, n)
'Dim objShell
'Dim objFolder
'Dim objFolderItem
'Dim i
'Dim strPath
'Dim strName
'Dim intPos
'
'On Error GoTo ErrHandler
'
'intPos = InStrRev(strFile, "")
'strPath = Left(strFile, intPos)
'strName = Mid(strFile, intPos + 1)
'Set objShell = CreateObject("Shell.Application")
'Set objFolder = objShell.Namespace(strPath)
'Set objFolderItem = objFolder.ParseName(strName)
'If Not objFolderItem Is Nothing Then
'GetProperty = objFolder.GetDetailsOf(objFolderItem, n)
'End If
'
'ExitHandler:
'Set objFolderItem = Nothing
'Set objFolder = Nothing
'Set objShell = Nothing
'Exit Function
'
'ErrHandler:
'MsgBox Err.Description, vbExclamation
'Resume ExitHandler
'End Function

Private Sub Label11_Click()

If Label1.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label2.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label9.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label11.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub

Label11.BackColor = Label_COLOR_YELLOW.BackColor
Label_GO_AH.BackColor = Label_COLOR_GREEN.BackColor

WORK = "FILE_CREATED_TO_MODIFIED_TIME"


End Sub

Private Sub Label2_Click()

If Label1.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label2.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label9.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label2.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub

Label2.BackColor = Label_COLOR_YELLOW.BackColor
Label_GO_AH.BackColor = Label_COLOR_GREEN.BackColor

WORK = "MOD_TO_CREATED_DATE"

' NEXT -- Label_GO_AH_Click

End Sub

Private Sub Label9_Click()


If Label11.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label9.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label1.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label2.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub

Label9.BackColor = Label_COLOR_YELLOW.BackColor
Label_GO_AH.BackColor = Label_COLOR_GREEN.BackColor

WORK = "CREATED_TO_MOD_DATE"

' NEXT -- Label_GO_AH_Click
' NEXT -- Call BATCH_CREATED_TO_MODIFIED_TIME

End Sub

Sub SET_BATCH_DATE_CAMERA_VIDEO_FILENAME_TO_DATE_FILE()

    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_10_31 Oct_Wed 11_55_50__MAQ01836 (2).mp4"

    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    ScanPath.txtPath.Text = "\\7-asus-gl522vw\7_ASUS_GL522VW_02_D_DRIVE\VI_ DSC ME 01\2010+SONY\2017 CyberShot HX60V_#\New folder"
    'ScanPath.txtPath.Text = LABEL_SET(2).Caption
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        a = B11
        
        XX = InStr(a, "\201_")
        DATEVAR = Replace(Mid(a, XX + 1, 10) + " " + Mid(a, XX + 20, 8), "_", "-")
        
        For r = 1 To 10
            If Mid(DATEVAR, r, 1) = "-" Then Mid(DATEVAR, r, 1) = "/"
        Next
        For r = 10 To Len(DATEVAR)
            If Mid(DATEVAR, r, 1) = "-" Then Mid(DATEVAR, r, 1) = ":"
        Next
        DateSet = DateValue(DATEVAR) + TimeValue(DATEVAR)
        
    '    Ds1 = DateValue(Mid(xB11, 9, 2) + "-" + Mid(xB11, 6, 2) + "-" + Mid(xB11, 1, 4))
    '    Ds2 = TimeValue(Replace(Mid(xB11, 12, 8), "-", ":"))
    '    DateSet = DateValue(Ds1) + TimeValue(Ds1)
    '    DateSet = Ds1 + Ds2
    '    DateSet = GetExif(A11 + B11)
                
        TT = SetFileDateTime(a, DateSet)
        XC = XC + 1
    Next
    
    MsgBox "Done = " + str(XC)
    End

End Sub


Private Sub Label_GO_AH_Click()

If WORK = "NOW_DATE" Then
    ' -------------------------------
    a = LABEL_SET(3).Caption           ' GET FILE
    DateSet = Now
    TT = SetFileDateTime(a, DateSet)
    MsgBox "Done " + vbCrLf + vbCrLf + a + vbCrLf + vbCrLf + DATEVAR, vbMsgBoxSetForeground
    End

    ' Call CODE_RUN
    Exit Sub
End If

If WORK = "MOD_TO_CREATED_DATE" Then
    Call BATCH_MODIFIED_TO_CREATED_TIME
    Exit Sub
End If

If WORK = "CREATED_TO_MOD_DATE" Then
    Call BATCH_CREATED_TO_MODIFIED_TIME
    Exit Sub
End If

If WORK = "FILE_CREATED_TO_MODIFIED_TIME" = True Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_MOST_RECENT_DATE_TO_OTHER_IN_FOLDER" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_OLDER_DATE_TO_OTHER_IN_FOLDER" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "DATE_CONVERTOR___MMM_D__YYYY_H_MM_AM____TO_YYYY_MM_DD__HH_MM_DD_FOR_SCREENCASTIFY" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "MAKE_FOLDER_YYYY_MM_DD_OF_FILE_AND_MOVE_THERE_____BATCH_IT" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "ONE_FOLDER_AS_OTHER_SAME_____HARDCODER____BATCH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS_DDD_NOKIA_AH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS__JPG_SINGLE" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA__MP4____BATCH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If


If WORK = "SET_DATE_OF_FILENAME_CH00_YYYY_MM_DD_HH_MM_SS_MP4_HIKVISION_SINGLE" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If
If WORK = "SET_DATE_OF_FILENAME_YYYY_MM_DD__WITH_ONE_FOLDER" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If
If WORK = "SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA_SINGLE" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_ONE_DATE_HARDCODER" = True Then
    
    Call SET_BATCH_DATE_CAMERA_VIDEO_FILENAME_TO_DATE_FILE
    End
    
'   SIMPLE
'   -------------------------------
    ' a = "D:\UTILS\2011 GALAXY SAMSUNG GT-P1000 - Copy\2012 07 GALAXY SAMSUNG GT-P1000_ VIDEO.MP4"
    a = LABEL_SET(3).Caption
    DATEVAR = "2012/07/01 18:00:00"
    DATEVAR = "2017/04/04 23:35:44"
    DateSet = DateValue(DATEVAR) + TimeValue(DATEVAR)
    TT = SetFileDateTime(a, DateSet)
    MsgBox "Done " + vbCrLf + vbCrLf + a + vbCrLf + vbCrLf + DATEVAR, vbMsgBoxSetForeground
    End

End If


If WORK = "SET_ONE_DATE" = True Then
        
    a = "D:\DSC\# Docus Proofs Texts\# BHT\NOTICE BOARD\2009-09-22 21-23-03 - Sony Ericsson K800i - DSC03044.JPG"
    a = "D:\DSC\# Docus Proofs Texts\# BHT\SMS PROOFS\2009-09-22 23-09-55 - Sony Ericsson K800i - DSC03047.JPG"
    a = "D:\DSC\# Docus Proofs Texts\# BHT\SMS PROOFS\2009-09-22 23-10-12 - Sony Ericsson K800i - DSC03048.JPG"
    a = "D:\VI_ DSC ME\2010+NOKIA\2015 NOKIA E72 _ 008 AUG _ MP4 _ x001 _ Home Front Room\2015_08_07 AUG_FRI 15_13_38__MAQ06717 - MY HOME.MP4"
    
    
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_07_24 Jul_Wed 22_48_08__MAQ07083_A Piezoelectric Speaker (AKA A Piezo Bender)_ELECTRONIC TECHNO RECORDER.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_02_07 Feb_Thu 14_42_48__MAQ03670_SEAGULLS AT BEACH SWARM.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_02_27 Feb_Wed 08_04_38__MAQ04106 VIDEO OF DEVICE THAT DOESN'T WORK GOT FROM EBAY ALARM AT POWER FAILURE ALERT.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_02_27 Feb_Wed 08_04_38__MAQ04106 VIDEO OF DEVICE THAT DOESN'T WORK GOT FROM EBAY ALARM AT POWER FAILURE ALERT.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_08_01 Aug_Thu 20_19_40__MAQ07292_VIEW TOP SOUTHWICK HILL SOUTH DOWNS.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_07_28 Jul_Sun 17_00_39__MAQ07182_UP THE DOWNS SUOTHWICK HILL.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_07_12 Jul_Fri 16_04_05__MAQ06893_UP THE SOUTH DOWN_.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_07_06 Jul_Sat 09_58_34__MAQ06634_UP THE SOUTH DOWN_ VIEW ON CHANCTONBURY RING _ LOT OF ANIMAL_ CATTLE & OTHER.txt"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_06_27 Jun_Thu 16_11_26__MAQ06357_A Walk Up The Downs _ Was Giver Show Sound Of Pylon Wire Thrash By Heavy Wind And Sunny A Type Of Undescribable Audio A View of Chantonbury Ring.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_02_27 Feb_Wed 08_04_38__MAQ04106 VIDEO OF DEVICE THAT DOESN'T WORK GOT FROM EBAY ALARM AT POWER FAILURE ALERT.MP4"
    XX = InStr(a, "\2019_")
    DATEVAR = Replace(Mid(a, XX + 1, 10) + " " + Mid(a, XX + 20, 8), "_", "-")
    
    
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_02_27 Feb_Tue 15_27_03__MAQ08676 _ Down at the Docks, Shoreham Port The Digger Weighs How Much Hardcore In The Spade_2.mp4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_05_28 May_Mon 15_25_36__MAQ00290 _ STARLING NEW BORN.MP4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_06_17 Jun_Sun 13_21_11__MAQ00586.mp4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_07_13 Jul_Fri 16_05_40__MAQ01488 _ Rotated_Garden_Water.mp4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_07_19 Jul_Thu 10_58_21__MAQ01563 _ Logitech 533 Speaker Standby Mode Removed Circuit Delay Control Gear.mp4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_08_20 Aug_Mon 19_02_51__MAQ02416 _ Helicopter Along the Seafront.mp4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_08_20 Aug_Mon 16_39_01__MAQ02389 _ SQUIRREL IN THE GARDEN.mp4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_10_27 Oct_Sat 10_18_46__MAQ01732_SEAGULL AND THE BIRD FEAST FAT TRAY WITH CHERRIES ON.MP4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_10_27 Oct_Sat 10_18_46__MAQ01732_SEAGULL AND THE BIRD FEAST FAT TRAY WITH CHERRIES ON.MP4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_10_27 Oct_Sat 10_18_46__MAQ01732_SEAGULL AND THE BIRD FEAST FAT TRAY WITH CHERRIES ON.MP4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_10_31 Oct_Wed 11_55_50__MAQ01836 (2).mp4"
    
    XX = InStr(a, "\2018_")
    DATEVAR = Replace(Mid(a, XX + 1, 10) + " " + Mid(a, XX + 20, 8), "_", "-")
    
    'SET THE FILE ABOVE
    'SET DATE BELOW

'    Set F = FSO.GetFile(a)
'    DT1 = F.datelastmodified
'    Set F = Nothing
    
    'DateSet = DT1 - TimeSerial(1, 0, 0)

'    DATEVAR = "2009-09-22 21-23-03"
'    DATEVAR = "2009-09-22 23-09-55"
'    DATEVAR = "2009-09-22 23-10-12"
'    DATEVAR = "2015-08-07 15-13-38"
    'DATEVAR = ""
    'DATEVAR = ""
    
    
    For r = 1 To 10
    If Mid(DATEVAR, r, 1) = "-" Then Mid(DATEVAR, r, 1) = "/"
    Next
    For r = 10 To Len(DATEVAR)
    If Mid(DATEVAR, r, 1) = "-" Then Mid(DATEVAR, r, 1) = ":"
    Next
    DateSet = DateValue(DATEVAR) + TimeValue(DATEVAR)
    
'    Ds1 = DateValue(Mid(xB11, 9, 2) + "-" + Mid(xB11, 6, 2) + "-" + Mid(xB11, 1, 4))
'    Ds2 = TimeValue(Replace(Mid(xB11, 12, 8), "-", ":"))
'    DateSet = DateValue(Ds1) + TimeValue(Ds1)
'    DateSet = Ds1 + Ds2
'    DateSet = GetExif(A11 + B11)
            
    TT = SetFileDateTime(a, DateSet)
    XC = XC + 1
    
    MsgBox "Done = " + str(XC)
    End
    
End If




End Sub


Private Sub SET_ONE_DATE_HARDCODER()
    
    
    WORK = "SET_ONE_DATE_HARDCODER"
    
    Call Label_GO_AH_Click

End Sub


Private Sub MNU_CREATE_FOLDER_DATE_MONTH_Click()

    Set F = FSO.GetFile(LABEL_SET(3).Caption)
    DT1 = F.datelastmodified

    OUT_FOLDER = LABEL_SET(2).Caption + "\" + Format(DT1, "YYYY-MM MMM")
    
    MkDir OUT_FOLDER

    End


End Sub

Private Sub MNU_CREATE_FOLDER_DATE_OF_FILE_Click()

    Set F = FSO.GetFile(LABEL_SET(3).Caption)
    DT1 = F.datelastmodified

    OUT_FOLDER = LABEL_SET(2).Caption + "\" + Format(DT1, "YYYY-MM-DD")
    If InStr(LABEL_SET(3).Caption, "REC") > 0 And InStr(LABEL_SET(3).Caption, ".WAV") > 0 Then
        OUT_FOLDER = LABEL_SET(2).Caption + "\RECORD " + Format(DT1, "YYYY-MM-DD")
    End If
    
    
    CreateFolderTree OUT_FOLDER

    End


End Sub

Private Sub MNU_CREATE_TEXT_FILE_FOR_DATE_Click()
    
    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    ScanPath.txtPath.Text = LABEL_SET(2).Caption
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        Exit For
    Next
    
    XX = ScanPath.txtPath.Text + "\# TEXT_DATER.TXT"
    FR1 = FreeFile
    Open XX For Output As #FR1
        Print #FR1, A11
        Print #FR1, B11
    Close FR1
    
    Me.WindowState = vbMinimized
    Shell "C:\Program Files (x86)\Notepad++\notepad++.exe """ + XX + """", vbMaximizedFocus

    End

End Sub

Private Sub MNU_EXIT_Click()
End
End Sub

Private Sub MNU_QUICK_MODE_Click()

'QUICK MODE -- ON -- GIVE MESSENGER

I_1 = "QUICK MODE -- ON -- NONE MESSENGER"
I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
If InStr(MNU_QUICK_MODE.Caption, I_1) Then
    MNU_QUICK_MODE.Caption = I_2
    Else
    MNU_QUICK_MODE.Caption = I_1
End If


Dim VAR_DATE As Date

On Error Resume Next
Err.Clear
FR1 = FreeFile
Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Output As #FR1
    Print #FR1, MNU_QUICK_MODE.Caption
    VAR_DATE = Now + TimeSerial(24, 0, 0)
    Print #FR1, VAR_DATE
Close FR1

If Err.Number = 0 Then
    I_1 = "QUICK MODE -- ON -- NONE MESSENGER"
    I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
    If InStr(MNU_QUICK_MODE.Caption, I_1) Then
        MsgBox "ENJOY GLOBAL FOR ALL INSTANCE RUNNER" + vbCrLf + vbCrLf + "REMEMEBER FOR 1 DAY DEFAULT OFF", vbMsgBoxSetForeground
    End If
Else
    MsgBox "ERROR -- FILE NOT WRITEN -- TRY AGAIN" + vbCrLf + vbCrLf + App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT", vbMsgBoxSetForeground
End If

End Sub



Private Sub MNU_VB_FOLDER_Click()
Shell "EXPLORER /SELECT, " + App.Path + "\" + App.EXEName + ".VBP", vbMaximizedFocus
End Sub

Private Sub MNU_VB_LAUNCH_FAV_SET_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 31-AUTORUN SET FAV VB & AUTOHOTKEY.ahk" + """"
Set WSHShell = Nothing
End Sub

Private Sub MNU_VB_ME_Click()
    
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
    On Error Resume Next
    Dim IGETTER
    IGETTER = GetSpecialfolder(38)
    
    If IGETTER = "" Then
        MsgBox "ERROR LOAD ADMIN"
        MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__<" + GetSpecialfolder(38) + ">__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
        'MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__< GetSpecialfolder(38) >__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
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
    
    ' Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
'    MsgBox """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """"
    
    Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    
    Me.WindowState = vbMinimized 'to down and en-able exit
    
    EXIT_TRUE = True
    Unload Me
    'End

End Sub


Public Function GetSpecialFolder_Show_Script_Debug(CSIDL As Long) As String

Dim r As Long
On Error Resume Next
For r = 0 To 120
    If Trim(GetSpecialfolder(r)) <> "" Then
        'Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
        'AAX = GetSpecialfolder(R)
    End If
Next
End Function



Public Function GetSpecialfolder(CSIDL As Long) As String
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




' --------------------------------------------------------------------
' ALL IN ONE FUNCTION LESS API LESS DEMAND JOB WORKER
' --------------------------------------------------------------------
' CreateFolderTree CreatePATH Create_PATH
' --------------------------------------------------------------------
Private Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
'    If Mid(sPath, 1, 2) = "\\" Then
'        nPos = InStr(nPos + 1, sPath, "\")
'        nPos = InStr(nPos + 1, sPath, "\")
'        nPos = InStr(nPos + 1, sPath, "\")
'    End If
    
    While nPos > 0
        If nPos - 1 > 3 Then
            If Dir(Left$(sPath, nPos - 1), vbDirectory) = "" Then
                MkDir Left$(sPath, nPos - 1)
            End If
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    If Dir(sPath, vbDirectory) = "" Then MkDir sPath
    
    CreateFolderTree = True
    If Dir(sPath, vbDirectory) = "" Then
        CreateFolderTree = False
    End If

    Exit Function

CreateFolderTreeError:
    Exit Function
End Function

Private Function CreateFolderTree_AND_NETWORK(ByVal sPath As String) As Boolean
    Dim nPos As Integer
    Dim strarray
    Dim r
    Dim LINE_PATH_BUILDER

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)
    
    If Mid(sPath, 1, 2) = "\\" Then
        LINE_PATH_BUILDER = "\\"
    End If
    If Mid(sPath, 2, 2) = ":\" Then
        LINE_PATH_BUILDER = "" 'Mid(sPath, 1, 3)
    End If
    
    On Error Resume Next
    
    strarray = Split(sPath, "\")
    For r = 0 To UBound(strarray)
        If strarray(r) <> "" Then
            LINE_PATH_BUILDER = LINE_PATH_BUILDER + strarray(r)
                        
            FSO.CreateFolder LINE_PATH_BUILDER
            ' -----------------------------------
            ' THIS HAS ERROR IS NOT ABLE MAKE PATH
            ' USE FSO
            ' -----------------------------------
            MkDir LINE_PATH_BUILDER
            
            LINE_PATH_BUILDER = LINE_PATH_BUILDER + "\"
            
        End If
    Next
    
    
    
    
    nPos = InStr(sPath, "\")
'    If Mid(sPath, 1, 2) = "\\" Then
'        nPos = InStr(nPos + 1, sPath, "\")
'        nPos = InStr(nPos + 1, sPath, "\")
'        nPos = InStr(nPos + 1, sPath, "\")
'    End If
    
'    While nPos > 0
'        If nPos - 1 > 3 Then
'            If Dir(Left$(sPath, nPos - 1), vbDirectory) = "" Then
'                MkDir Left$(sPath, nPos - 1)
'            End If
'        End If
'        nPos = InStr(nPos + 1, sPath, "\")
'    Wend
'    If Dir(sPath, vbDirectory) = "" Then MkDir sPath
    
'    CreateFolderTree_AND_NETWORK = True
'    If Dir(sPath, vbDirectory) = "" Then
'        CreateFolderTree_AND_NETWORK = False
'    End If

    Exit Function

CreateFolderTreeError:
    Exit Function
End Function





' --------------------------------------------------------------------
' WORK BUT EASIER ALL IN ONE FUNCTION NOT WITH API THOUGH
' MORE TRANSPORTBALE TO OTHER CODE MAYBE IN ONE FUNCTION LESS DELCLARE
' --------------------------------------------------------------------
Private Function CreateFolderTree_API(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        If nPos - 1 > 3 Then ' NOT CHECK ROOT
            If Not FolderExists(Left$(sPath, nPos - 1)) Then
                MkDir Left$(sPath, nPos - 1)
            End If
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    If Not FolderExists(sPath) Then MkDir sPath
    
    If Not FolderExists(sPath) Then
        CreateFolderTree_API = False
        Exit Function
    End If
    CreateFolderTree_API = True
    Exit Function

CreateFolderTreeError:
    Exit Function
End Function

Private Function FolderExists(sFolder As String) As Boolean
    '##############################################################################################
    'Returns True if the specified folder exists
    '##############################################################################################
    
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFolder, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FolderExists = True
        Else
            FolderExists = False
        End If
    End If
End Function


'LABEL_SET(2).Caption
Private Sub MNU_FILEDATE_WHOLE_FOLDER_Click()
    '---------------------------------------------
    'VIDEO MP4
    '---------------------------------------------

    ' If FILE_NAME_MP4 = "" Then Beep: MsgBox "NOT A FILE NAME PIPED": End
    
    File1.Path = LABEL_SET(2).Caption
    
    ' SOMETIME THE PASS BY CONTEXT MENU ISN'T LONG NAME BUT SHORTNAME
        
    R_C_COUNTER_MAX = 0
    For R_C_2 = 1 To 3
        
        A_M = ""
        
        R_C_COUNTER = 0
            
        For R_C = 0 To File1.ListCount - 1
    
            FILE_NAME_MP4 = File1.Path + "\" + File1.List(R_C)
        
            If Mid(FILE_NAME_MP4, 1, 2) <> "\\" Then
                FILE_NAME_MP4 = GetLongName(FILE_NAME_MP4)
            End If
            
            Set F = FSO.GetFile((FILE_NAME_MP4))
            ADATE1 = F.datelastmodified
            
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
            'If InStr(FILE_NAME_MP4, "D:\DSC\") > 0 Then
            'Clipboard.SetText " " + at1 + Format(ADATE1, " MMMM DD DDDD HH-MM-SS Am/Pm")
            FILE_NAME_MP4_2 = Mid(FILE_NAME_MP4, InStrRev(FILE_NAME_MP4, "\") + 1)
            FILE_NAME_MP4_2_EXT = Mid(FILE_NAME_MP4, InStrRev(FILE_NAME_MP4, "."))
            
            MAQ = ""
            START_MAQ = 0
            
            Call MNU_FILEDATE_CLIPBOARD_ROUTINE_CHECKER
            
            If START_MAQ = 1 Then
                If File1.List(R_C) <> Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02 + FILE_NAME_MP4_2_EXT Then
                    If MAQ <> "" Then
                    
                        R_C_COUNTER = R_C_COUNTER + 1
                        If R_C_2 = 1 Then
                            R_C_COUNTER_MAX = R_C_COUNTER_MAX + 1
                        
                        End If
                        
                        A_M = A_M + "RENAMER --" + str(R_C_COUNTER) + " OF" + str(R_C_COUNTER_MAX) + " OF" + str(File1.ListCount) + vbCrLf + File1.Path + "\" + vbCrLf + "\" + File1.List(R_C) + vbCrLf + "\" + Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02 + FILE_NAME_MP4_2_EXT
                        A_M = A_M + vbCrLf
                        A_M = A_M + vbCrLf
                        
                        
                        If R_C_2 = 3 Then
                        
                            If R_C_COUNTER < 10 Then
                            
                                A_M_2 = "RENAMER -- VERIFY 1ST 10  AND THEN AUTO" + vbCrLf + Trim(str(R_C + 1)) + " OF" + str(File1.ListCount) + vbCrLf + File1.Path + "\" + vbCrLf + "\" + File1.List(R_C) + vbCrLf + "\" + Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02 + FILE_NAME_MP4_2_EXT
                                
                                X_M = MsgBox(A_M_2, vbYesNo)
                                If X_M = vbNo Then
                                    R_C_COUNTER = R_C_COUNTER - 1
                                    Exit For
                                End If
                            End If
                            
                            Name File1.Path + "\" + File1.List(R_C) As File1.Path + "\" + Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02 + FILE_NAME_MP4_2_EXT
                        
                        End If
                    End If
                End If
            End If
        Next
            
        If R_C_2 = 2 Then
            If A_M <> "" Then
                R_C_COUNTER = 0
                X_M = MsgBox("RENAME CONVERSION AS FOLLOW PROCESS YES / NOT" + vbCrLf + vbCrLf + A_M, vbYesNo)
                If X_M = vbNo Then Exit For
            End If
        End If
    Next
        
    If R_C_COUNTER > 0 Then
    
    MsgBox "DONE " + Trim(str(R_C_COUNTER)) + " CHANGER"
    
    End If
    
    End
End Sub



Private Sub MNU_FILEDATE_CLIPBOARD_Click()
    ' MNU_FILEDATE_CLIPBOARD
    ' [ Sunday 09:46:40 Am_18 November 2018 ]

    '---------------------------------------------
    'VIDEO MP4
    '---------------------------------------------

    If FILE_NAME_MP4 = "" Then Beep: MsgBox "NOT A FILE NAME PIPED": End
    
    Beep
    
    Me.Hide

'    On Error Resume Next
'    Me.Visible = True
'    DoEvents
'    Me.WindowState = vbMinimized
'    DoEvents
'    MsgBox FILE_NAME_MP4, vbMsgBoxSetForeground
'    On Error GoTo 0

'    SOMETIME THE PASS BY CONTEXT MENU ISN;T LONG NAME BUT SHORTNAME

    If Mid(FILE_NAME_MP4, 1, 2) <> "\\" Then
        FILE_NAME_MP4 = GetLongName(FILE_NAME_MP4)
    End If
    
    Set F = fs.GetFile((FILE_NAME_MP4))
    ADATE1 = F.datelastmodified
    
'    Clipboard.Clear
'    Sleep 200
    'If Year(ADATE1) = 2016 Then at1 = "-- 2k Sixteenth"
    'If Year(ADATE1) = 2015 Then at1 = "-- 2k Fifteenth"
    'If Year(ADATE1) <= 2015 Then at1 = "-- " + Format(ADATE1, "YYYY")
    'If Year(ADATE1) > 2016 Then at1 = "-- " + Format(ADATE1, "YYYY")
    
    'at1 = "-- " + Format(ADATE1, "YYYY")
    
    
    '------------------------
    ' LONG DATE FOR YOUTUBING
    '------------------------
    'If InStr(FILE_NAME_MP4, "D:\DSC\") > 0 Then
    'Clipboard.SetText " " + at1 + Format(ADATE1, " MMMM DD DDDD HH-MM-SS Am/Pm")
    FILE_NAME_MP4_2 = Mid(FILE_NAME_MP4, InStrRev(FILE_NAME_MP4, "\") + 1)
    'MsgBox FILE_NAME_MP4_2
    MAQ = ""
    START_MAQ = 0
    

    Call MNU_FILEDATE_CLIPBOARD_ROUTINE_CHECKER
    
    Clipboard.Clear
    Sleep 200

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
    If InStr(FILE_NAME_MP4_2, TEE) > 0 Then
        START_MAQ_LEN = 8
        MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
        START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
    End If
    
    TEE = "MAH"
    If InStr(FILE_NAME_MP4_2, TEE) > 0 Then
        START_MAQ_LEN = 8
        MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
        START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
    End If
    
    TEE = "M4V"
    If InStr(FILE_NAME_MP4_2, TEE) > 0 Then
        START_MAQ_LEN = 8
        MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
        START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
    End If
    
    TEE = "M4H"
    If InStr(FILE_NAME_MP4_2, TEE) > 0 Then
        START_MAQ_LEN = 8
        MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
        START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
    End If
    
    TEE = "DSCF"
    If InStr(GetLongName(FILE_NAME_MP4), TEE) > 0 Then
        If InStr(GetLongName(FILE_NAME_MP4), ".MOV") > 0 Then
            START_MAQ_LEN = 8
            MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
            START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
        End If
    End If

    'DOUBLE SCREEN CAMERA
    TEE = "DSCF"
    If InStr(GetLongName(FILE_NAME_MP4), TEE) > 0 Then
        If InStr(GetLongName(FILE_NAME_MP4), ".AVI") > 0 Then
            START_MAQ_LEN = 8
            MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
            START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
        End If
    End If
    
    TEE = "image"
    If InStr(GetLongName(FILE_NAME_MP4), TEE) > 0 Then
        If InStr(GetLongName(FILE_NAME_MP4), ".AVI") > 0 Then
            START_MAQ_LEN = 7
            MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
            START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
        End If
    End If
    
    TEE = "Photo"
    If InStr(GetLongName(FILE_NAME_MP4), TEE) > 0 Then
        If InStr(GetLongName(FILE_NAME_MP4), ".JPG") > 0 Then
            START_MAQ_LEN = 7
            MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
            START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
        End If
    End If
    
    
    'IF DESCRIPTION ALREADY ON
    NAME_PART_02 = ""
    If START_MAQ > 0 Then
        NAME_PART_02 = Mid(FILE_NAME_MP4_2, START_MAQ + START_MAQ_LEN)
        NAME_PART_02 = Mid(NAME_PART_02, 1, InStrRev(NAME_PART_02, ".") - 1)
    End If
End Sub


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


