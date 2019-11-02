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
      Left            =   132
      TabIndex        =   11
      Top             =   4692
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
      Left            =   288
      TabIndex        =   12
      Top             =   5364
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
      Left            =   14352
      TabIndex        =   13
      Top             =   1704
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
      Left            =   13872
      TabIndex        =   14
      Top             =   1008
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
      Left            =   13968
      TabIndex        =   15
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
      Index           =   15
      Left            =   13968
      TabIndex        =   16
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M()
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

'FOLDER LABEL
LABEL_SET(2).Caption = W$
If W$ = "" Then LABEL_SET(2).Caption = "NOT FOLDER GIVEN"

'FILE LABEL
LABEL_SET(3).Caption = FULL_PATH_AND_FILENAME
If FULL_PATH_AND_FILENAME = "" Then LABEL_SET(3).Caption = "NOT FILE GIVEN"

LABEL_SET(1).BackColor = Label_COLOR_YELLOW.BackColor

End Sub

Private Sub Form_Resize()

Dim r

ReDim M(LABEL_SET.Count)
ReDim M_2(LABEL_SET.Count)
ReDim M_3(LABEL_SET.Count)
i = 0

i = i + 1: M(i) = "GO"
i = i + 1: M(i) = "Folder Label"
i = i + 1: M(i) = "File Label"
'i = i + 1: M(i) = "PERFORM ON ALL FILES IN FOLDER OR FILE"
i = i + 1: M(i) = "NOW_DATE"
i = i + 1: M(i) = "MODIFY DATE TO CREATED DATE - NOT WORKING"
i = i + 1: M(i) = "----"
i = i + 1: M(i) = "BATCH - CREATED DATE TO MODIFY DATE"
i = i + 1: M(i) = "FILE  - CREATED DATE TO MODIFY DATE"
i = i + 1: M(i) = "----"
i = i + 1: M(i) = "SET_ONE_DATE_HARDCODER"
i = i + 1: M(i) = "----"
i = i + 1: M(i) = "SET_MOST_RECENT_DATE_TO_OTHER_IN_FOLDER"
i = i + 1: M(i) = "SET_OLDER_DATE_TO_OTHER_IN_FOLDER"
i = i + 1: M(i) = "----"
i = i + 1: M(i) = "SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH_--_\#_TEXT_DATER*.TXT"
i = i + 0: M_3(i) = "SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH"

For r = 0 To i
    M_2(r) = Replace(M(r), "_", " ")
Next

For r = 1 To LABEL_SET.Count
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

'For r = 1 To LABEL_SET.Count
'    LABEL_SET(r).Caption = Replace(LABEL_SET(r).Caption, "_", " ")
'Next

LABEL_SET(2).FontSize = 12
LABEL_SET(3).FontSize = LABEL_SET(2).FontSize

On Error Resume Next
Me.Width = 17000
On Error GoTo 0

' TOP LABEL
' HEIGHT LABEL
HL = LABEL_SET(2).Height
HL = 500

LENGHT_LABEL = 380

STEP_H = 100
For r = 2 To LABEL_SET.Count
    If LABEL_SET(r).Caption = "" Then
        LABEL_SET(r).Visible = False
    End If
    
    If LABEL_SET(r).Visible = True Then
        LABEL_SET(r).Left = 100
        LABEL_SET(r).Width = Me.Width - LENGHT_LABEL
        LABEL_SET(r).Height = HL
        LABEL_SET(r).Top = STEP_H
        STEP_H = STEP_H + 40 + HL
    End If
Next

r = 1
LABEL_SET(r).Left = 100
LABEL_SET(r).Width = Me.Width - LENGHT_LABEL
LABEL_SET(r).Height = HL
LABEL_SET(r).Top = STEP_H
STEP_H = STEP_H + 40 + HL

On Error Resume Next
Me.Height = STEP_H + 1000
Me.Top = 0

End Sub


Private Sub LABEL_SET_Click(index As Integer)

DISPLAY_NAMER = M(index)
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
    LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

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
    
    For rr = 1 To ScanPath.ListView1.ListItems.Count
        A1$ = ScanPath.ListView1.ListItems.Item(rr).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(rr)
        
        Set F = FSO.GetFile(A1$ + B1$)
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
    
    For rr = 1 To ScanPath.ListView1.ListItems.Count
        A1$ = ScanPath.ListView1.ListItems.Item(rr).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(rr)
        
        TT = SetFileDateTime(A1$ + B1$, DT4)
        
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
    
    For rr = 1 To ScanPath.ListView1.ListItems.Count
        A1$ = ScanPath.ListView1.ListItems.Item(rr).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(rr)
        
        FLAG_OPER = True
        If UCase(B1$) = UCase("thumbs.db") Then FLAG_OPER = False
        If UCase(B1$) = UCase("DESKTOP.INI") Then FLAG_OPER = False
        
        Set F = FSO.GetFile(A1$ + B1$)
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
    For rr = 1 To ScanPath.ListView1.ListItems.Count
        A1$ = ScanPath.ListView1.ListItems.Item(rr).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(rr)
        
        TT = SetFileDateTime(A1$ + B1$, DT4)
        XC = XC + 1
        
        If ScanPath.ListView1.ListItems.Count > 2 Then
        FI = UCase(B1$)
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
    
    For rr = 1 To ScanPath.ListView1.ListItems.Count
        A1$ = ScanPath.ListView1.ListItems.Item(rr).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(rr)
        
        Set F = FSO.GetFile(A1$ + B1$)
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
    DT4 = DT4 + TimeSerial(0, 1, 0)
    For rr = 1 To ScanPath.ListView1.ListItems.Count
        A1$ = ScanPath.ListView1.ListItems.Item(rr).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(rr)
        TT = SetFileDateTime(A1$ + B1$, DT4)
        XC = XC + 1
        MM_1 = MM_1 + B1$ + vbCrLf
        FI = UCase(B1$)
        FI = Mid(FI, InStrRev(FI, ".") + 1) + " "
        If InStr("MPG MPEG MP4 JPG AVI ", FI) > 0 Then
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
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    Set F = FSO.GetFile(A1$ + B1$)
    DT3 = F.DateCreated
    DT1 = F.datelastmodified
    
    DateSet = DT3
    
    Set F = Nothing
    
    TT = SetFileDateTime(A1$ + B1$, DateSet)
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
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    Set F = FSO.GetFile(A1$ + B1$)
    DT3 = F.DateCreated
    DT1 = F.datelastmodified
    Set F = Nothing
    
    TT = SetFileDateTime(A1$ + B1$, DT3)
    XC = XC + 1
    
    I_MSG = I_MSG + B1$ + vbCrLf
    
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
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    Set F = FSO.GetFile(A1$ + B1$)
    DT1 = F.datelastmodified
    
    DateSet = DT1 - TimeSerial(1, 0, 0)
    
    Set F = Nothing
    
    'Ds1 = Mid(B1$, 1, 20)
'    Ds1 = Replace(Ds1, "+", ":") + ":00"
'    If InStr(B1$, "2008 011(Nov) 29 Sat 08-04-10") > 0 Then
    'If Year(DT1) = 2009 Then
        '2008 012(Dec) 04 Thu 18-22-52 - W880I - DSC00331.JPG
      '  Ds1 = DateValue("04-Dec-2008") + TimeValue("18:22:52")
        
'        xb1$ = B1$
'        xb1$ = Replace(xb1$, "MILL VIEW HOSPITAL-", "")
'
        If A1$ + B1$ = "D:\0 00 ART LOGGERS - WEBCAM\VIDEO\Web_Video_Capture- 2015-10-03 14-28-59 Sat -- Microsoft WDM Image Capture (Win32)2.avi" Then
'            Ds1 = DateValue(Mid(xb1$, 9, 2) + "-" + Mid(xb1$, 6, 2) + "-" + Mid(xb1$, 1, 4))
'            Ds2 = TimeValue(Replace(Mid(xb1$, 12, 8), "-", ":"))
        '    DateSet = DateValue(Ds1) + TimeValue(Ds1)
'            DateSet = Ds1 + Ds2
    '        DateSet = GetExif(a1$ + B1$)
            DateSet = DateValue("2015-10-03") + TimeValue("14:28:59")
                  
            '
            TT = SetFileDateTime(A1$ + B1$, DateSet)
            XC = XC + 1
            'tt = LastModifiedToCurrent(a1$ + B1$)
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
    
    For rr = 1 To ScanPath.ListView1.ListItems.Count
        A1$ = ScanPath.ListView1.ListItems.Item(rr).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(rr)
        a = B1$
        
        XX = InStr(a, "\201_")
        DATEVAR = Replace(Mid(a, XX + 1, 10) + " " + Mid(a, XX + 20, 8), "_", "-")
        
        For r = 1 To 10
            If Mid(DATEVAR, r, 1) = "-" Then Mid(DATEVAR, r, 1) = "/"
        Next
        For r = 10 To Len(DATEVAR)
            If Mid(DATEVAR, r, 1) = "-" Then Mid(DATEVAR, r, 1) = ":"
        Next
        DateSet = DateValue(DATEVAR) + TimeValue(DATEVAR)
        
    '    Ds1 = DateValue(Mid(xb1$, 9, 2) + "-" + Mid(xb1$, 6, 2) + "-" + Mid(xb1$, 1, 4))
    '    Ds2 = TimeValue(Replace(Mid(xb1$, 12, 8), "-", ":"))
    '    DateSet = DateValue(Ds1) + TimeValue(Ds1)
    '    DateSet = Ds1 + Ds2
    '    DateSet = GetExif(a1$ + B1$)
                
        TT = SetFileDateTime(a, DateSet)
        XC = XC + 1
    Next
    
    MsgBox "Done = " + str(XC)
    End

End Sub


Private Sub Label_GO_AH_Click()

If WORK = "NOW_DATE" = True Then
    ' -------------------------------
    a = LABEL_SET(3).Caption           ' GET FILE
    DateSet = Now
    TT = SetFileDateTime(a, DateSet)
    MsgBox "Done " + vbCrLf + vbCrLf + a + vbCrLf + vbCrLf + DATEVAR, vbMsgBoxSetForeground
    End

    ' Call CODE_RUN
    Exit Sub
End If

If WORK = "MOD_TO_CREATED_DATE" = True Then
    Call BATCH_MODIFIED_TO_CREATED_TIME
    Exit Sub
End If

If WORK = "CREATED_TO_MOD_DATE" = True Then
    Call BATCH_CREATED_TO_MODIFIED_TIME
    Exit Sub
End If

If WORK = "FILE_CREATED_TO_MODIFIED_TIME" = True Then
    Call FILE_CREATED_TO_MODIFIED_TIME
    Exit Sub
End If


If WORK = "SET_MOST_RECENT_DATE_TO_OTHER_IN_FOLDER" Then
    Call SET_MOST_RECENT_DATE_TO_OTHER_IN_FOLDER
    Exit Sub
End If

If WORK = "SET_OLDER_DATE_TO_OTHER_IN_FOLDER" Then
    Call SET_OLDER_DATE_TO_OTHER_IN_FOLDER
    Exit Sub
End If

If WORK = "SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH" Then
    Call SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH
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
    
'    Ds1 = DateValue(Mid(xb1$, 9, 2) + "-" + Mid(xb1$, 6, 2) + "-" + Mid(xb1$, 1, 4))
'    Ds2 = TimeValue(Replace(Mid(xb1$, 12, 8), "-", ":"))
'    DateSet = DateValue(Ds1) + TimeValue(Ds1)
'    DateSet = Ds1 + Ds2
'    DateSet = GetExif(a1$ + B1$)
            
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


Private Sub MNU_CREATE_FOLDER_DATE_OF_FILE_Click()

    Set F = FSO.GetFile(LABEL_SET(3).Caption)
    DT1 = F.datelastmodified

    OUT_FOLDER = LABEL_SET(2).Caption + "\" + Format(DT1, "YYYY-MM-DD")
    If InStr(LABEL_SET(3).Caption, "REC") > 0 And InStr(LABEL_SET(3).Caption, ".WAV") > 0 Then
        OUT_FOLDER = LABEL_SET(2).Caption + "\RECORD " + Format(DT1, "YYYY-MM-DD")
    End If
    
    
    MkDir OUT_FOLDER

    End


End Sub

Private Sub MNU_CREATE_TEXT_FILE_FOR_DATE_Click()
    
    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    ScanPath.txtPath.Text = LABEL_SET(2).Caption
    
    For rr = 1 To ScanPath.ListView1.ListItems.Count
        A1$ = ScanPath.ListView1.ListItems.Item(rr).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(rr)
        Exit For
    Next
    
    XX = ScanPath.txtPath.Text + "\# TEXT_DATER.TXT"
    FR1 = FreeFile
    Open XX For Output As #FR1
        Print #FR1, A1$
        Print #FR1, B1$
    Close FR1
    
    Me.WindowState = vbMinimized
    Shell "C:\Program Files (x86)\Notepad++\notepad++.exe """ + XX + """", vbMaximizedFocus

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




