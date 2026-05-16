VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ScanPath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ScanPath 2.0 - Anything -- Demo of Class Object"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   9780
      Top             =   345
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "REMOVE LINE WITHOUT ERROR TO VIEW"
      Default         =   -1  'True
      Height          =   1005
      Left            =   7095
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2040
      Width           =   2040
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "TIDY TEMP FOLDER"
      Height          =   1005
      Left            =   9150
      Picture         =   "ScanPathDemo.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2025
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "WHEN YOUR READY"
      Height          =   825
      Left            =   8220
      Picture         =   "ScanPathDemo.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3195
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton CmdScanDir 
      Caption         =   "ScanDir"
      Height          =   825
      Left            =   6585
      Picture         =   "ScanPathDemo.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3390
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   585
      Top             =   4545
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Height          =   825
      Left            =   4230
      Picture         =   "ScanPathDemo.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3390
      Width           =   1245
   End
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   2115
      Left            =   2805
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      Text            =   "ScanPathDemo.frx":34BC
      Top             =   4065
      Visible         =   0   'False
      Width           =   7245
   End
   Begin VB.ComboBox cboMask 
      Height          =   315
      Left            =   540
      TabIndex        =   1
      Top             =   1110
      Width           =   6105
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   6090
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   750
      Width           =   555
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   555
      TabIndex        =   0
      Text            =   "E:\My Music Zen"
      Top             =   735
      Width           =   5505
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "Search Sub-Folders"
      Height          =   195
      Left            =   1530
      TabIndex        =   8
      Top             =   2070
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkPatternMatching 
      Caption         =   "Pattern Matching"
      Height          =   195
      Left            =   1530
      TabIndex        =   7
      Top             =   1830
      Width           =   1815
   End
   Begin VB.CheckBox chkNormal 
      Caption         =   "Normal"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   1830
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   2070
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   2310
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   2535
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   5325
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2010
      Width           =   1545
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   0
      Left            =   2565
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3480
      Width           =   1365
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3090
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   1215
      TabIndex        =   16
      Top             =   3480
      Width           =   1305
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   1
      Left            =   2565
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3840
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   1215
      TabIndex        =   18
      Top             =   3840
      Width           =   1305
   End
   Begin VB.CheckBox chkRefreshListView 
      Caption         =   "Refresh ListView (slower)"
      Height          =   195
      Left            =   1530
      TabIndex        =   10
      Top             =   2610
      Value           =   1  'Checked
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   10950
      TabIndex        =   22
      Top             =   0
      Width           =   11010
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single File solution to quickly add file processing to any Utility Project."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1260
         TabIndex        =   25
         Top             =   390
         Width           =   6015
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A High Performance API file/folder scanner Class."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1260
         TabIndex        =   24
         Top             =   60
         Width           =   6825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   23
         Top             =   60
         Width           =   1035
      End
   End
   Begin VB.CheckBox chkCopyMemory 
      Caption         =   "Use CopyMemory (display Date && Size)"
      Height          =   195
      Left            =   1530
      TabIndex        =   11
      Top             =   2850
      Width           =   3120
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3060
      Left            =   450
      TabIndex        =   21
      Top             =   4275
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   5398
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CheckBox chkSort 
      Caption         =   "Sort Files(A-Z) while Scanning"
      Height          =   195
      Left            =   1530
      TabIndex        =   9
      Top             =   2310
      Width           =   2535
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   825
      Left            =   5565
      Picture         =   "ScanPathDemo.frx":35F4
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3390
      Visible         =   0   'False
      Width           =   945
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   0
      Left            =   5325
      TabIndex        =   13
      Top             =   2400
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   21954561
      CurrentDate     =   37299
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   1
      Left            =   5325
      TabIndex        =   14
      Top             =   2760
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   21954561
      CurrentDate     =   37296
   End
   Begin VB.Label lblCount3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6900
      TabIndex        =   46
      Top             =   1425
      Width           =   3855
   End
   Begin VB.Label lblCount2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6900
      TabIndex        =   45
      Top             =   1065
      Width           =   3855
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Date/Size Filter:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4650
      TabIndex        =   38
      Top             =   1590
      Width           =   1410
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Options:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1530
      TabIndex        =   30
      Top             =   1590
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mask"
      Height          =   195
      Left            =   90
      TabIndex        =   28
      Top             =   1140
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      Height          =   195
      Left            =   90
      TabIndex        =   26
      Top             =   810
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Attributes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   29
      Top             =   1590
      Width           =   885
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      Height          =   195
      Index           =   0
      Left            =   555
      TabIndex        =   35
      Top             =   3540
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
      Height          =   195
      Left            =   555
      TabIndex        =   33
      Top             =   3120
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   195
      Index           =   1
      Left            =   555
      TabIndex        =   37
      Top             =   3900
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   4650
      TabIndex        =   34
      Top             =   2460
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "File Date"
      Height          =   195
      Left            =   4650
      TabIndex        =   32
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   4650
      TabIndex        =   36
      Top             =   2760
      Width           =   195
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6900
      TabIndex        =   31
      Top             =   705
      Width           =   3855
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------
'Del_Emptys
'SORT_THE_TEMP_FOLDER_OUT
'DELETE ELITE
'-------------------------
Dim DONT_RESIZE_WHILE_SETUP
Dim DONT_RESIZE_RUN_ONCE_OR_NORM

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function FindWindowDLL _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long


Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

'Function GetUserName() As String
'   Dim UserName As String * 255
'   Call GetUserNameA(UserName, 255)
'   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
'End Function
'Function GetComputerName() As String
'   Dim UserName As String * 255
'   Call GetComputerNameA(UserName, 255)
'   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
'End Function


'Jeepers
'Dim Fs22
Dim FS
'Option Explicit

'##############################################################################################
'Purpose:   This project is a Demo for my cScanPath Class

'           This Class scans a specified path and returns the files it finds.
'           It has fairly comprehensive Filters. You can select files by:
'           Attributes (Normal, Hidden, Read Only, System etc)
'           File Size (>, <, Range)
'           File Date (From, To, Range)
'           File Extensions (multiple allowed i.e. *.txt;*.dat;*.tmp)

'           You can optionally scan sub-folders

'           To keep the demo simple I have only used attributes & Extensions
'           for Filter. For full example of this Class see WipeIt3 submission
       
'Author:    Richard Mewett ©2005
'##############################################################################################

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'We must declare with WithEvents to process the files returned
Private WithEvents SP As cScanPath
Attribute SP.VB_VarHelpID = -1

Public Function StripNulls(OriginalStr As String) As String
        ' This removes the extra Nulls so String comparisons will work
        If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
        End If
        StripNulls = OriginalStr
End Function

Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim L As Long
   Dim S As String
   L = GetWindowTextLength(hWnd)
   S = Space(L + 1)
   GetWindowText hWnd, S, L + 1
   GetWindowTitle = Left$(S, L)

End Function


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


Sub USER_DIR()

'Dim Fs22
'Dim FS
'Set Fs22 = CreateObject("Scripting.FileSystemObject")
Set FS = CreateObject("Scripting.FileSystemObject")


'txtPath.Text = "M:\01 Installations\01 DVD Nfo Txt MP3 JPG"
''txtPath.Text = "M:\01 Installations\01 DVD Nfo Txt Log Diz"
'txtPath = "M:\01 Installations\01 DVD MP3 JPG GIF PDF"
'txtPath = "M:\00 Archive Work"
''txtPath = "M:\01 Archives"
'
'txtPath = "M:\01 Installations Icy"
'cmdScanDir_Click
'
'txtPath = "M:\01 Installations Lacie"
'cmdScanDir_Click
'
'txtPath = "M:\01 Installations WD200"
'cmdScanDir_Click
'
'
'
'
'
'
'T:\0 00 ART LOGGERS\# FEED STATION\My FeedStation Podcasts DATE ORDER\## ARCHIVE

ScanPath.ListView1.ListItems.Clear

If FS.FolderExists(Clipboard.GetText) = True Then
    tdk = Clipboard.GetText
Else
    tdk = "D:\0 00 MOBILE\MOBILE"
'    tdk = "D:\0 00 MOBILE\MOBILE\MOBILE NOKIA - COMMON"
    tdk = "T:\0 00 ART LOGGERS"
End If

If tdk <> "" And iside = False Then
    If FS.FolderExists(tdk) = True Then
    
        ScanPath.txtPath = tdk
        'Make Sure Backslash at End of Scan Path
        If Mid(ScanPath.txtPath, Len(ScanPath.txtPath)) <> "\" Then ScanPath.txtPath = ScanPath.txtPath + "\"
        
        'Dont Do Root's
        If Len(ScanPath.txtPath) < 4 Then End
        'Scan Folder Must Exist
        If FS.FolderExists(ScanPath.txtPath) = False Then End
'        Call cmdScanDir_Click
    End If
End If

End Sub


Sub SORT_THE_TEMP_FOLDER_OUT()

' -------------------------------------------------------
' STAGE ONE OF TWO 01 OF 02
' DELETE FILES

Command3.Enabled = False

Set FS = CreateObject("Scripting.FileSystemObject")

'Me.Show
'DoEvents

'1. DEL ALL *.TMP IF OLDER THAN A 24 HOUR
'2. DEL ALL NAUGHT IF OLDER THAN A 24 HOUR
'3. NOT IS ERROR READ FILE



txtPath = "C:\TEMP\-" + GetComputerName
cboMask = "*.*"

Call cmdScan_Click

Dim ADATE1 As Date

For WE = ListView1.ListItems.Count To 1 Step -1
    
    lblCount = Str(WE) + " OF" + Str(ListView1.ListItems.Count)
    DoEvents
    
    ListView1.ListItems.Item(WE).EnsureVisible
    DoEvents
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    c1$ = A1$ + B1$
    
    Err.Clear
    On Local Error Resume Next
    Set F = FS.GetFile(A1$ + B1$)
    ADATE1 = F.datelastmodified
    ASIZE = F.Size
    
    GO = 0
    If ADATE1 < Now - TimeSerial(24, 0, 0) Then GO = 1
    If GO = 1 And InStr(UCase(B1$), ".TMP") > 0 Then GO = 2
    If GO = 1 And ASIZE = 0 Then GO = 2
    
    If Err.Number <> 0 Then GO = 0
    
    If GO = 2 Then
    
        ' -------------------------------------------------------
        ' STAGE ONE OF TWO 01 OF 02
        ' DELETE FILES
        Err.Clear
        VAR_ERROR = False
        On Error Resume Next
            Kill A1$ + B1$
            If Error.Number > 0 Then VAR_ERROR = True
        On Error GoTo 0
        
        
        'If Err.Number > 0 Then MsgBox "Error When Delete" + vbCrLf + c1$ + "\*.*", vbMsgBoxSetForeground: Exit Sub
        
        'RmDir c1$
        If VAR_ERROR > 0 Then
'                MsgBox "Cant Remove Folder " + vbCrLf + c1$
            errhuge = errhuge + 1
            datavar = "Error Delete File"
            ListView1.ListItems.Item(WE).SubItems(2) = datavar
        Else
            datavar = "Deleted"
            huge = huge + 1
            ListView1.ListItems.Item(WE).SubItems(2) = datavar
        
        End If
        
        ListView1.ListItems.Item(WE).EnsureVisible
        DoEvents
        On Error GoTo 0
    End If

Next

AD1 = Str(huge) + " OF" + Str(ListView1.ListItems.Count) + " *.TMP AND NAUGHT FILES Deleted" + vbCrLf + Str(errhuge) + " OF *.TMP AND NAUGHT ERROR Not Deleted"

TOTALFILE = Str(ListView1.ListItems.Count)

'ListView1.ListItems.Clear



' -------------------------------------------------------
' STAGE ONE OF TWO 01 OF 02
' DELETE FILES
'MsgBox AD1


' -------------------------------------------------------
' STAGE TWO OF TWO 02 OF 02
' DELETE EMPTY FOLDERS

'txtPath = "C:\TEMP\"

Call cmdScanDir_Click


'JUST IN CASE SORTED OPTION NOT SELECTED
'SORT ON FILE-PATH
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 0
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False
'SORT ON PATH-PATH
ScanPath.ListView1.SortOrder = lvwAscending  'lvwAscending
ScanPath.ListView1.SortKey = 1
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False

    
    
    
For WE = ListView1.ListItems.Count To 1 Step -1
    
    lblCount = Str(WE) + " OF" + Str(ListView1.ListItems.Count)

    DoEvents
    
    ListView1.ListItems.Item(WE).EnsureVisible
    DoEvents
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    c1$ = A1$ + B1$
    
    On Error Resume Next
    Err.Clear
    X22 = Dir(c1$, vbDirectory) <> ""
        
    'Set F = FS.GetFolder(c1$)
    If Err.Number > 0 Then
        'MsgBox "ERROR GETFOLDER WITH THIS FILE" + vbCrLf + Err.Description + vbCrLf + A1$ + B1$, vbMsgBoxSetForegroundL ': Exit Sub
        'TOO MANY LOCKED FOLDER TO DISPLAY EROR MESSAGE
        'ERR.DESCRIPTION
        
        errhuge2 = errhuge2 + 1
        
        datavar = "ERROR Read FOLDER"
        ListView1.ListItems.Item(WE).SubItems(2) = datavar
    
    End If
        
'    If Err.Number = 0 And Dir(c1$, vbDirectory) <> "" Then
    
    If FS.FileExists(c1$) = False And FS.FolderExists(c1$) = False Then
        errhuge2 = errhuge2 + 1
        datavar = "ERROR Read FOLDER OR FILE MAYBE LOCKED"
        ListView1.ListItems.Item(WE).SubItems(2) = datavar
    
        'MsgBox "Error Read Folder Files" + vbCrLf + Err.Description + vbCrLf + c1$ + "\*.*", vbMsgBoxSetForeground ': Exit Sub
    End If
    
    If Err.Number = 0 And FS.FileExists(c1$) = False And FS.FolderExists(c1$) = True Then
        Err.Clear
        
        X1 = Dir(c1$ + "\*", vbDirectory)
        X2 = Dir
        x3 = Dir
        
        'NEED CHECK X3 AND LINE FO THEM IS FILE SIZE NAUGHT AND THEN MAYBE DELETE AS WELL
        'Can be Processor Heavy if Folders Contain Lot Files
        'If f.Size > 0 Then Stop
        'If Err.Number > 0 Then MsgBox "Error Get Size" + vbCrLf + c1$, vbMsgBoxSetForeground: Exit Sub
        'ERR.DESCRIPTION
        
        If Dir(c1$ + "\*.*") = "" And x3 = "" Then
            If Err.Number > 0 Then
                errhuge2 = errhuge2 + 1
                datavar = "ERROR Read FOLDER"
                ListView1.ListItems.Item(WE).SubItems(2) = datavar
            
                MsgBox "Error Read Folder Files" + vbCrLf + Err.Description + vbCrLf + c1$ + "\*.*", vbMsgBoxSetForeground ': Exit Sub
            
            End If
            
            If Err.Number = 0 Then
                
                RmDir c1$
                
                If Err.Number > 0 Then
                    errhuge = errhuge + 1
                    datavar = "ERROR DELETE FOLDER"
                    ListView1.ListItems.Item(WE).SubItems(2) = datavar
                Else
                    datavar = "FOLDER Deleted"
                    huge = huge + 1
                    ListView1.ListItems.Item(WE).SubItems(2) = datavar
                
                End If
                ListView1.ListItems.Item(WE).EnsureVisible
                DoEvents
                On Error GoTo 0
            End If
        End If
    End If
    
    DoEvents
Next

TOTALFOLDER = Str(ListView1.ListItems.Count) - TOTALFILE


' -------------------------------------------------------
' STAGE ONE OF TWO 01 OF 02
' DELETE FILES
'If Command$ = "" Then MsgBox AD1
'MsgBox AD1

' -------------------------------------------------------
' STAGE TWO OF TWO 02 OF 02
' DELETE EMPTY FOLDERS
AD2 = Str(huge) + " OF" + Str(ListView1.ListItems.Count) + " Folder Deleted" + vbCrLf + Str(errhuge) + " OF Folders ERROR Not Deleted" + vbCrLf + Str(errhuge2) + " OF ERROR READ FOLDER Not Deleted"
'If Command$ = "" Then MsgBox AD2
TOTALENDE = ListView1.ListItems.Count

Command3.Enabled = True
I = MsgBox(AD1 + vbCrLf + vbCrLf + AD2 + vbCrLf + vbCrLf + "DO YOU WANT TO REMOVE LINE WITHOUT ERROR TO VIEW BETTER" + vbCrLf + "CAN DO YOURSELF WITH GREEN BUTTON", vbYesNo)


'If Command$ = "" Then
'    I = MsgBox("DO YOU WANT TO REMOVE LINE WITHOUT ERROR TO VIEW BETTER" + vbCrLf + "CAN DO YOURSELF WITH GREEN BUTTON", vbYesNo)
If I = vbYes Then
    Command3.Enabled = False
    Call Command3_Click

End If
'Else
    'Call Command3_Click


'End If

lblCount = Str(TOTALENDE) + " TOTAL AND --" + Str(ListView1.ListItems.Count) + " WITH ERROR TIDY VIEW"
lblCount2 = Str(TOTALFILE) + " TOTAL FILE AFTER TIDY"
lblCount3 = Str(TOTALFOLDER) + " TOTAL FOLDER AFTER TIDY"

'Me.WindowState = vbNormal

Timer2.Enabled = False
Command2.BackColor = Command3.BackColor
Command3.Enabled = False


End Sub

Sub Del_Emptys()
Me.Show
DoEvents


'SORT_THE_TEMP_FOLDER_OUT

'C:\TEMP\-1-ASUS-X5DIJ


'USER_DIR


Call cmdScanDir_Click


'JUST IN CASE SORTED OPTION NOT SELECTED
'SORT ON FILE-PATH
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 0
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False
'SORT ON PATH-PATH
ScanPath.ListView1.SortOrder = lvwAscending  'lvwAscending
ScanPath.ListView1.SortKey = 1
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False



'If iside = True Or 1 = 1 Then
''    ScanPath.txtPath = "E:\01 Start Menu\"
'    ScanPath.txtPath = "U:\0 00 ART LOGGERS\# TEMP INTERNET\Temp Inet Files JPGs\"
'    w$ = "F:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call cmdScanDir_Click
'Else
'    Call cmdScanDir_Click
'End If

    
    
    
    
    
    
    
Dim WFD As WIN32_FIND_DATA
Dim lResult As Long

For WE = ListView1.ListItems.Count To 1 Step -1
    
    lblCount = Str(WE)
    DoEvents
    
    ListView1.ListItems.Item(WE).EnsureVisible
    DoEvents
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    c1$ = A1$ + B1$
    
    On Error Resume Next
    Set F = FS.GetFolder(c1$)
    If Err.Number > 0 Then
        MsgBox "ERROR GETFOLDER WITH THIS FILE" + vbCrLf + Err.Description + vbCrLf + A1$ + B1$, vbMsgBoxSetForegroundL ': Exit Sub
    End If
    'On Error GoTo 0
    
    'Err.Clear
'    lblCount = Trim(Str(we))
    'Can be Processor Heavy if Folders Contain Lot Files
    'If f.Size > 0 Then Stop
    'If f.Size = 0 Then
        'If Err.Number > 0 Then MsgBox "Error Get Size" + vbCrLf + c1$, vbMsgBoxSetForeground: Exit Sub
        
'        If dir(c1$ + "\*.*") <> "" Then Stop
        
    If Err.Number = 0 Then
        Err.Clear
        X1 = Dir(c1$ + "\*", vbDirectory)
        X2 = Dir
        x3 = Dir
        
        'NEED CHECK X3 AND LINE FO THEM IS FILE SIZE NAUGHT AND THEN MAYBE DELETE AS WELL
        
        If Dir(c1$ + "\*.*") = "" And x3 = "" Then
            If Err.Number > 0 Then MsgBox "Error Dir Folder for Files" + vbCrLf + Err.Description + vbCrLf + c1$ + "\*.*", vbMsgBoxSetForeground ': Exit Sub
            
            If Err.Number = 0 Then
    '            Err.Clear
                'FS.Deletefolder c1$', True 'True = Readonly deleted
                RmDir c1$
                
                'If Err.Number > 0 Then MsgBox "Error When Delete" + vbCrLf + c1$ + "\*.*", vbMsgBoxSetForeground: Exit Sub
                
                'RmDir c1$
                If Err.Number > 0 Then
    '                MsgBox "Cant Remove Folder " + vbCrLf + c1$
                    errhuge = errhuge + 1
                    datavar = "Can't Remove"
                    ListView1.ListItems.Item(WE).SubItems(2) = datavar
                Else
                    datavar = "Deleted"
                    huge = huge + 1
                    ListView1.ListItems.Item(WE).SubItems(2) = datavar
                
                End If
                ListView1.ListItems.Item(WE).EnsureVisible
                DoEvents
                On Error GoTo 0
            'Else
            '    huge2 = huge2 + 1
            End If
        End If
    End If
    
    'End If
    DoEvents
Next

MsgBox Str(huge) + " of" + Str(ListView1.ListItems.Count) + " Folders Deleted" + vbCrLf + Str(errhuge) + " Folders Not Deleted"
'End
Exit Sub

End




On Local Error GoTo jeep
Drived2$ = Mid$(txtPath.Text, 1, 3)
MkDir Drived2$ + "Temp\Anything"
v1$ = Mid$(txtPath.Text, InStrRev(txtPath.Text, "\"))

MkDir Drived2$ + "Temp\Anything" + v1$
On Local Error GoTo 0

List1.AddItem "Stage 1 of 2 : Make Dir's And Move Files to Temp"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For WE = 1 To ListView1.ListItems.Count
    
    lblCount = Str(WE)
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
    c1$ = Drived2$ + "Temp\Anything" + v1$ + "\" + Mid$(A1$, ets1 + 2)

    etd1 = 13
    ets2 = 1
    Do
        'jc1$ = Mid$(c1$, etd1 + ets2-1)
        ets2 = InStr(etd1 + ets2, c1$, "\")
        etd1 = 1
        If ets2 = 0 Then Exit Do

        d2$ = Mid$(c1$, 1, ets2)

        If InStr(f1$, d2$) = 0 Then
            On Local Error GoTo jeep
            MkDir d2$
            If errs2 <> 75 And errs2 > 0 Then
                MsgBox ("Error Making Temp Directory")
                Stop
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    f1$ = d2$

    errs2 = 0
    On Local Error GoTo jeep
    Fs22.moveFile A1$ + B1$, c1$ + B1$
    On Local Error GoTo 0

    If errs2 <> 0 Then
        MsgBox ("Error Moving File to Temp")
        End
    End If

    List1.AddItem Format$(WE, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'Fs22.DeleteFolder comrad$ + "\" + WFD$

Next

Err.Clear
On Local Error Resume Next
Set Fs22 = CreateObject("Scripting.FileSystemObject")
Fs22.Deletefolder txtPath.Text, True
If Err.Number > 0 Then Call HardDel
On Local Error GoTo 0


List1.AddItem "----------------------------------------"
List1.AddItem "Stage 2 of 2 : Move Files Back From Temp"
List1.AddItem "----------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

v2$ = Mid$(txtPath.Text, 1, InStrRev(txtPath.Text, "\"))

Err.Clear
On Local Error Resume Next
Set Fs22 = CreateObject("Scripting.FileSystemObject")
Fs22.Movefolder Drived2$ + "Temp\Anything\*", v2$
If Err.Number > 0 Then Call HardMove
On Local Error GoTo 0

List1.AddItem "Stage 2 of 2 : Complete --------------"
List1.AddItem "--------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

'Timer1.Enabled = True

Exit Sub
jeep:
errs2 = Err.Number
errs3$ = Err.Description
Resume Next

End Sub

Private Sub cboSize_Click()
    Select Case cboSize.ListIndex
    Case 0
        lblSize(0).Visible = False
        lblSize(1).Visible = False
        txtSize(0).Visible = False
        txtSize(1).Visible = False
        cboSizeType(0).Visible = False
        cboSizeType(1).Visible = False
    Case 1, 2
        lblSize(0).Caption = "Size"
        lblSize(0).Visible = True
        lblSize(1).Visible = False
        txtSize(0).Visible = True
        txtSize(1).Visible = False
        cboSizeType(0).Visible = True
        cboSizeType(1).Visible = False
    Case Else
        lblSize(0).Caption = "Min"
        lblSize(1).Caption = "Max"
        lblSize(1).Visible = True
        lblSize(0).Visible = True
        txtSize(0).Visible = True
        txtSize(1).Visible = True
        cboSizeType(0).Visible = True
        cboSizeType(1).Visible = True
    End Select
End Sub


Private Sub chkCopyMemory_Click()
    chkSort.Enabled = (chkCopyMemory.Value = vbUnchecked)
    
    With ListView1
        If chkCopyMemory.Value Then
            .ColumnHeaders("PATH").Width = 5000
            .ColumnHeaders("SIZE").Width = 1250
            .ColumnHeaders("DATE").Width = 1750
        Else
            .ColumnHeaders("PATH").Width = 8000
            .ColumnHeaders("SIZE").Width = 0
            .ColumnHeaders("DATE").Width = 0
        End If
    End With
End Sub


Private Sub chkSort_Click()
    chkCopyMemory.Enabled = (chkSort.Value = vbUnchecked)
End Sub

Private Sub Cmd_scanDir_Click()

End Sub

Private Sub cmdBrowse_Click()
    txtPath.Text = GetFolder(Me.hWnd, "Scan Path:", txtPath.Text)
'    fg1 = FreeFile
'    Open App.Path + "\Scan Path.txt" For Output As #fg1
'    Print #fg1, txtPath.Text
'    Close #fg1

'    fg1 = FreeFile
'    Open App.Path + "\Scan Path.txt" For Input As #fg1
'    Line Input #fg1, ae$
'    Close #fg1
'    txtPath.Text = ae$



End Sub

Private Sub cmdHelp_Click()
    txtHelp.Visible = Not txtHelp.Visible
End Sub

Public Sub cmdScan_Click()
    Dim lCount As Long
    Dim lSize(1) As Long
    
    If cmdScan.Caption = "Scan" Then
        '####################################################################################################
        'Process our Selection Criteria
        For lCount = 0 To 1
            lSize(lCount) = Val(txtSize(lCount).Text)
        
            Select Case cboSizeType(lCount).ListIndex
            Case 1: lSize(lCount) = lSize(lCount) * 1024
            Case 2: lSize(lCount) = (lSize(lCount) * 1024) * 1024
            End Select
        Next lCount
            
        If txtSize(1).Visible Then
            If lSize(0) = 0 Then
                MsgBox "Maximium Size value must exceed 0", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            End If
        End If
        
        '####################################################################################################
        'Reset Form
        cmdScan.Caption = "Stop"
        lblCount.Caption = "-"
 
        Screen.MousePointer = vbHourglass
        ListView1.ListItems.Clear
        
        If chkRefreshListView.Value = vbUnchecked Then
            LockWindowUpdate ListView1.hWnd
        End If
        
        '####################################################################################################
        'Create our Scan Object
        Set SP = New cScanPath
        
        With SP
            'Attributes
            .Archive = chkArchive.Value
            .Compressed = True
            .Hidden = chkHidden.Value
            .Normal = chkNormal.Value
            .ReadOnly = chkReadOnly.Value
            .System = chkSystem.Value
            
            'Date/Size (NOTE: you don't have to use these!)
            Select Case cboDate.ListIndex
            Case 0
                .DateType = Modified
            Case 1
                .DateType = Created
            Case 2
                .DateType = LastAccessed
            End Select
            
            If IsDate(DTPicker1(0).Value) Then
                .FromDate = DTPicker1(0).Value
            End If
            If IsDate(DTPicker1(1).Value) Then
                .ToDate = DTPicker1(1).Value
            End If
            
            Select Case cboSize.ListIndex
            Case 1
                .MaximumSize = lSize(0)
            Case 2
                .MinimumSize = lSize(0)
            Case 3
                .MinimumSize = lSize(0)
                .MaximumSize = lSize(1)
            End Select
            
            'Mask
            .Filter = cboMask.Text
            
            'Go - that was easy wasn't it!
            .StartScan txtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
        End With
        
       With ListView1
            If .ListItems.Count > 0 Then
                .SelectedItem = .ListItems(1)
            End If
            
            lblCount.Caption = .ListItems.Count
        End With
        
        LockWindowUpdate 0
        
        cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
    End If






Me.Refresh

End Sub

Public Sub cmdScanDir_Click()
    Dim lCount As Long
    Dim lSize(1) As Long
    
    If CmdScanDir.Caption = "ScanDir" Then
        '####################################################################################################
        'Process our Selection Criteria
        For lCount = 0 To 1
            lSize(lCount) = Val(txtSize(lCount).Text)
        
            Select Case cboSizeType(lCount).ListIndex
            Case 1: lSize(lCount) = lSize(lCount) * 1024
            Case 2: lSize(lCount) = (lSize(lCount) * 1024) * 1024
            End Select
        Next lCount
            
        If txtSize(1).Visible Then
            If lSize(0) = 0 Then
                MsgBox "Maximium Size value must exceed 0", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            End If
        End If
        
        '####################################################################################################
        'Reset Form
        CmdScanDir.Caption = "Stop"
        lblCount.Caption = "-"
 
        Screen.MousePointer = vbHourglass
        'ListView1.ListItems.Clear
        
        If chkRefreshListView.Value = vbUnchecked Then
            LockWindowUpdate ListView1.hWnd
        End If
        
        '####################################################################################################
        'Create our Scan Object
        Set SP = New cScanPath
        
        With SP
            'Attributes
            .Archive = chkArchive.Value
            .Compressed = True
            .Hidden = chkHidden.Value
            .Normal = chkNormal.Value
            .ReadOnly = chkReadOnly.Value
            .System = chkSystem.Value
            
            'Date/Size (NOTE: you don't have to use these!)
            Select Case cboDate.ListIndex
            Case 0
                .DateType = Modified
            Case 1
                .DateType = Created
            Case 2
                .DateType = LastAccessed
            End Select
            
            If IsDate(DTPicker1(0).Value) Then
                .FromDate = DTPicker1(0).Value
            End If
            If IsDate(DTPicker1(1).Value) Then
                .ToDate = DTPicker1(1).Value
            End If
            
            Select Case cboSize.ListIndex
            Case 1
                .MaximumSize = lSize(0)
            Case 2
                .MinimumSize = lSize(0)
            Case 3
                .MinimumSize = lSize(0)
                .MaximumSize = lSize(1)
            End Select
            
            'Mask
            .Filter = cboMask.Text
            
            'Go - that was easy wasn't it!
            .StartScanDir txtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
        End With
        
       With ListView1
            If .ListItems.Count > 0 Then
                .SelectedItem = .ListItems(1)
            End If
            
            lblCount.Caption = .ListItems.Count
        End With
        
        LockWindowUpdate 0
        
        CmdScanDir.Caption = "ScanDir"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
    End If



For WE = ListView1.ListItems.Count To 2 Step -1
    A1$ = Me.ListView1.ListItems.Item(WE)
    a2$ = Me.ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = Me.ListView1.ListItems.Item(WE - 1)
    b2$ = Me.ListView1.ListItems.Item(WE - 1).SubItems(1)
    If A1$ + B1$ = a2$ + b2$ Then
        Me.ListView1.ListItems.Remove (WE)
    End If
    
    
    'b1$ = frmMain.ListView1.ListItems.Item(we)

Next

lblCount.Caption = ListView1.ListItems.Count
    
Me.Refresh

End Sub




Private Sub Command1_Click()


Call Del_Emptys


End Sub

Private Sub Command2_Click()

Call SORT_THE_TEMP_FOLDER_OUT

End Sub

Private Sub Command3_Click()

'REMOVE EMPLY LINE ERROR TO VIEW WHAT IS ERROR

For WE = ListView1.ListItems.Count To 1 Step -1
    
    lblCount = Str(WE)
    DoEvents
    
    ListView1.ListItems.Item(WE).EnsureVisible
    DoEvents
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    c1$ = A1$ + B1$
    
    If ListView1.ListItems.Item(WE).SubItems(2) = "" Then
        
        ListView1.ListItems.Remove (WE)
    
    End If


Next


End Sub

Private Sub Form_Load()
        
    'Me.Caption = "ScanPath 2.0 - Anything -- Demo of Class Object -- Remove Empty Folders"
    'Me.Caption = "ScanPath 2.0 - Anything -- Demo of Class Object -- Tidy Temp Folder"
    Me.Caption = "Tidy Temp Folder -- ScanPath 2.0 - Anything -- Demo of Class Object"
    
    Dim lCount As Long
    
    With cboMask
        .AddItem "*.*"
        .AddItem "*.dll;*.exe;*.ocx"
        .AddItem "*.doc;*.mdb;*.xls"
        .AddItem "*.bmp;*.gif;*.jpg;*.tif"
        .AddItem "*.bas;*.cls;*.ctl;*.frm;*.vbp"
        .ListIndex = 0
    End With
    
    DTPicker1(0).Value = DateSerial(Year(Now), Month(Now) - 3, Day(Now))
    DTPicker1(1).Value = Now
    
    DTPicker1(0).Value = Empty
    DTPicker1(1).Value = Empty
    
    With cboDate
        .AddItem "Modified"
        .AddItem "Created"
        .AddItem "Last Accessed"
        .ListIndex = 0
    End With
    
    With cboSize
        .AddItem "Any Size"
        .AddItem "Less Than"
        .AddItem "Greater Than"
        .AddItem "Between"
        .ListIndex = 0
    End With
        
    For lCount = 0 To 1
        With cboSizeType(lCount)
            .AddItem "Bytes"
            .AddItem "Kilobytes"
            .AddItem "Megabytes"
            .ListIndex = 1
        End With
    Next lCount
    
    With ListView1
        .ColumnHeaders.Add , "FILE", "File", 3000, lvwColumnLeft
        .ColumnHeaders.Add , "PATH", "Path", 8000, lvwColumnLeft
        .ColumnHeaders.Add , "STATUS", "Status", 4000, lvwColumnRight
        .ColumnHeaders.Add , "DATE", "Date", 0, lvwColumnLeft
        
        .View = lvwReport
    End With

'    fg1 = FreeFile
'    Open App.Path + "\Scan Path.txt" For Input As #fg1
'    Line Input #fg1, ae$
'    Close #fg1
'    txtPath.Text = ae$



'Del_Emptys

'Call USER_DIR

'Call SORT_THE_TEMP_FOLDER_OUT

DONT_RESIZE_RUN_ONCE_OR_NORM = True

Command3.Enabled = False

Exit Sub

'DONT_RESIZE_WHILE_SETUP = True
''-----------------------------------------
'Call SETUP_DIMENSIONS_NORM
''Call SETUP_DIMENSIONS_INNER_FORM
''-----------------------------------------
''RESIZE_AT_LOAD = False
'DONT_RESIZE_WHILE_SETUP = False


'Me.Enabled = False
'Me.WindowState = vbNormal
'Me.Show
''DoEvents
'Me.Hide
'
'Me.WindowState = vbMinimized 'vbNormal
'Me.Show
'Me.Enabled = True



End Sub

Private Sub Form_Resize()

If DONT_RESIZE_WHILE_SETUP = True Then Exit Sub


'MAYBE SOMETIMES WINDOW IS MIN WHILE TAKE A CHANGE
'AT CLIPBOARD
'THINK NOT
'If Me.WindowState = vbMinimized Then RRR = Now - TimeSerial(0, 0, 3)

'HHH = Now + TimeSerial(0, 2, 0) ' vbMinimized WINDOW

'If RESIZE_AT_LOAD = True And Me.WindowState <> vbMinimized Then
If Me.WindowState <> vbMinimized And DONT_RESIZE_RUN_ONCE_OR_NORM = True Then
'    Me.WindowState = vbNormal
    DONT_RESIZE_WHILE_SETUP = True
    '-----------------------------------------
    Call SETUP_DIMENSIONS_NORM
    'Call SETUP_DIMENSIONS_INNER_FORM
    ListView1.Left = 0
    ListView1.Width = Me.Width - 200
    ListView1.Height = 3000
    
    
    '-----------------------------------------
    'RESIZE_AT_LOAD = False
    DONT_RESIZE_WHILE_SETUP = False
    DONT_RESIZE_RUN_ONCE_OR_NORM = False

End If

'If Me.WindowState = vbMinimized Then
'    MNU_SCANPATH_COUNTER.Visible = False
'    Timer_API_OKAY_COLOUR.Enabled = False
'Else
'    Timer_API_OKAY_COLOUR.Enabled = True
'End If


'If Me.WindowState <> vbMinimized Then
'    I = SetForegroundWindow(Me.hwnd)
'End If


'Call SETUP_DIMENSIONS_INNER_FORM

Exit Sub

End Sub


Sub SETUP_DIMENSIONS_CENTER()

   
    Dim RECTT1 As RECT
    Dim RECTT2 As RECT
    Dim RECTT4 As RECT
    
    On Error Resume Next
    
    If Me.WindowState <> vbMinimized Then
        
        DONT_RESIZE_WHILE_SETUP = True
        
        Test1 = FindWindow("Shell_TrayWnd", vbNullString)
        T1 = GetWindowRect(Test1, RECTT1) ' BOTTOM BAR
        Test2 = FindWindow("MOM Class", vbNullString)
        T1 = GetWindowRect(Test2, RECTT2) ' TOP BAR
        'IRON BAR - DUMB BELLS
        TEST4 = FindWindowPart_BASEBAR("BaseBar")
        T1 = GetWindowRect(TEST4, RECTT4) ' LEFT BAR
        A = RECTT4.Top
        Y1 = (RECTT1.Bottom - RECTT1.Top) * Screen.TwipsPerPixelY
        Y2 = (RECTT2.Bottom - RECTT2.Top) * Screen.TwipsPerPixelY
        
        If TEST4 > 0 Then
            If RECTT4.Right = RECTT4.Left Then
                X4 = (RECTT4.Right) * Screen.TwipsPerPixelX
            Else
                X4 = (RECTT4.Right - RECTT4.Left) * Screen.TwipsPerPixelX
            End If
        End If
        
        SCREEN_WIDTH_SPACE = Screen.Width - X4
        SCREEN_HEIGHT_SPACE = Screen.Height - Y1 - Y2
        
        'Form1.Width = SCREEN_WIDTH_SPACE - 800
        'Form1.Height = SCREEN_HEIGHT_SPACE - 500
        
        'Form1.Left = (Screen.Width - Me.Width) / 2
        Form1.Left = X4 + (SCREEN_WIDTH_SPACE - Me.Width) / 2
        Form1.Top = Y2 + (SCREEN_HEIGHT_SPACE - Me.Height) / 2
        
        'DoEvents
        
        'Call SETUP_DIMENSIONS_INNER_FORM
    
    End If



End Sub






Function FindWindowPart_BASEBAR(Test) As Long

'Testing
'Test$
'TEST$ 123

FindWindowPart_BASEBAR = False

Dim test_hwnd As Long
Dim cText As String
Dim RECT8 As RECT


test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
Do While test_hwnd <> 0
        

        
        cText = Space$(255)
        G = GetClassName(test_hwnd, cText, 255)
        cText = StripNulls(cText)
        If cText = Test Then
            hwndtest = GetWindowRect(test_hwnd, RECT8)
            If RECT8.Top < 35 And RECT8.Bottom > 700 And RECT8.Left = 0 Then
            
                FindWindowPart_BASEBAR = test_hwnd: Exit Function

            End If
        End If
        
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop


End Function



Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_Click()

End Sub

Private Sub SP_DirMatch(Directory As String, Path As String)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################
End Sub

Private Sub SP_DirMatchxx(Directory As String, Path As String)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################



Dim LV As ListItem
    Dim uWIN32 As WIN32_FIND_DATA
    
    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
    On Local Error GoTo GetFileError
    
    With ListView1
        Set LV = .ListItems.Add(, , Directory)
        LV.SubItems(1) = Path
        
        
        
        
        
        
        
        
        
        If chkCopyMemory.Value Then
            'VB does not allows UDT's to be passed directly from a Class but we can access the structure
            'using pointers. We use CopyMemory to place the WIN32_FIND_DATA variable from the Class into
            'the uWIN32 declared in this Sub.
            
            'This is primarily a speed trick - we could retrieve the date & time information "ourselves"
            'using the standard VB functions or an API call but since we already have it...
            CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)
            LV.SubItems(2) = uWIN32.nFileSizeLow
            LV.SubItems(3) = FormatFileDate(uWIN32.ftLastAccessTime)
        End If
        
        'Refresh Count & Scroll the list every 10th file
        'If (LV.Index Mod 10) = 0 Then
            If chkRefreshListView.Value Then
                .SelectedItem = LV
                .SelectedItem.EnsureVisible
                'DoEvents
            End If
            
            lblCount.Caption = LV.Index
            DoEvents
        'End If
    End With
    Exit Sub

GetFileError:
    MsgBox Err.Description, vbCritical


End Sub

Private Sub SP_FileMatch(Filename As String, Path As String)
    Dim LV As ListItem
    Dim uWIN32 As WIN32_FIND_DATA
    
    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
    On Local Error GoTo GetFileError
    
    With ListView1
        Set LV = .ListItems.Add(, , Filename)
        LV.SubItems(1) = Path
        
        
        
        
        
        
        
        
        
        If chkCopyMemory.Value Then
            'VB does not allows UDT's to be passed directly from a Class but we can access the structure
            'using pointers. We use CopyMemory to place the WIN32_FIND_DATA variable from the Class into
            'the uWIN32 declared in this Sub.
            
            'This is primarily a speed trick - we could retrieve the date & time information "ourselves"
            'using the standard VB functions or an API call but since we already have it...
            CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)
            LV.SubItems(2) = uWIN32.nFileSizeLow
            LV.SubItems(3) = FormatFileDate(uWIN32.ftLastAccessTime)
        End If
        
        'Refresh Count & Scroll the list every 10th file
        If (LV.Index Mod 10) = 0 Then
            If chkRefreshListView.Value Then
                .SelectedItem = LV
                .SelectedItem.EnsureVisible
            End If
            
            lblCount.Caption = LV.Index
        End If
    End With
    Exit Sub

GetFileError:
    MsgBox Err.Description, vbCritical
End Sub



Private Function FormatFileDate(CT As FILETIME) As String
    Const SHORT_DATE = "Short Date"
    
    Dim ST As SYSTEMTIME
    Dim ds As Single
       
    If FileTimeToSystemTime(CT, ST) Then
          ds = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
          
          FormatFileDate = Format$(ds, SHORT_DATE)
    End If
End Function

Private Sub Timer1_Timer()
End
End Sub



Sub HardDel()

Dim Fs22

Set Fs22 = CreateObject("Scripting.FileSystemObject")

Drived2$ = Mid$(txtPath.Text, 1, 3)
v1$ = Mid$(txtPath.Text, InStrRev(txtPath.Text, "\"))


List1.AddItem "Stage Opp : Del Original Dir"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For WE = 1 To ListView1.ListItems.Count
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
    c1$ = Drived2$ + "My Music" + v1$ + "\" + Mid$(A1$, ets1 + 2)

    etd1 = 13
    ets2 = 1
    Do
        'jc1$ = Mid$(c1$, etd1 + ets2-1)
        ets2 = InStr(etd1 + ets2, c1$, "\")
        etd1 = 1
        If ets2 = 0 Then Exit Do

        d2$ = Mid$(c1$, 1, ets2 - 1)

        If InStr(f1$, d2$) = 0 Then
            Err.Clear
            On Local Error Resume Next
            'MkDir d2$
            Fs22.Deletefolder d2$, True
            If Err.Number <> 75 And Err.Number > 0 Then
                'MsgBox ("Error Making Temp Directory")
                'Stop
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    f1$ = d2$

    errs2 = 0
    On Local Error Resume Next
'    Fs22.DeleteFile a1$ + b1$, True
    On Local Error GoTo 0

    'If errs2 <> 0 Then
    '    MsgBox ("Error Moving File to Temp")
    '    End
    'End If

    List1.AddItem Format$(WE, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'Fs22.DeleteFolder comrad$ + "\" + WFD$

Next

A = A
jeep:
End Sub


Sub HardMove()
Dim Fs22

Set Fs22 = CreateObject("Scripting.FileSystemObject")

Drived2$ = Mid$(txtPath.Text, 1, 3)
v1$ = Mid$(txtPath.Text, InStrRev(txtPath.Text, "\"))


List1.AddItem "Stage Opp : Del Original Dir"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For WE = 1 To ListView1.ListItems.Count
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
    c1$ = Drived2$ + "My Music" + v1$ + "\" + Mid$(A1$, ets1 + 2)
    c2$ = Drived2$ + "Temp\Anything" + v1$ + "\" + Mid$(A1$, ets1 + 2)

    etd1 = 13
    ets2 = 1
    Do
        'jc1$ = Mid$(c1$, etd1 + ets2-1)
        ets2 = InStr(etd1 + ets2, c1$, "\")
        etd1 = 1
        If ets2 = 0 Then Exit Do

        d2$ = Mid$(c1$, 1, ets2 - 1)

        If InStr(f1$, d2$) = 0 Then
            Err.Clear
            On Local Error Resume Next
            MkDir d2$
   '         Fs22.deletefolder d2$, True
            If Err.Number <> 75 And Err.Number > 0 Then
                'MsgBox ("Error Making Temp Directory")
                'Stop
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    f1$ = d2$
    Err.Clear
    errs2 = 0
    On Local Error Resume Next
    Set Fs22 = CreateObject("Scripting.FileSystemObject")
    Fs22.moveFile c2$ + B1$, A1$ + B1$
    'err.number
    'err.description
    On Local Error GoTo 0

    'If errs2 <> 0 Then
    '    MsgBox ("Error Moving File to Temp")
    '    End
    'End If

    List1.AddItem Format$(WE, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'Fs22.DeleteFolder comrad$ + "\" + WFD$

Next
    
    v1$ = Mid$(txtPath.Text, 1, InStrRev(txtPath.Text, "\") - 1)
    Err.Clear
    On Local Error Resume Next
    Set Fs22 = CreateObject("Scripting.FileSystemObject")
    Fs22.copyfolder Drived2$ + "Temp\Anything", v1$, True
    
    Set Fs22 = CreateObject("Scripting.FileSystemObject")
    Fs22.Deletefolder Drived2$ + "Temp\Anything", True


jeep:
End Sub

Private Sub Timer2_Timer()
If Command2.BackColor <> Command3.BackColor Then
    Command2.BackColor = Command3.BackColor
    Else
    Command2.BackColor = cmdHelp.BackColor
End If
    
End Sub


Sub SETUP_DIMENSIONS_NORM()

   
    Dim RECTT1 As RECT
    Dim RECTT2 As RECT
    Dim RECTT4 As RECT
    
    On Error Resume Next
    
    If Me.WindowState <> vbMinimized Then
        
        DONT_RESIZE_WHILE_SETUP = True
        
        Test1 = FindWindow("Shell_TrayWnd", vbNullString)
        T1 = GetWindowRect(Test1, RECTT1) ' BOTTOM BAR
        Test2 = FindWindow("MOM Class", vbNullString)
        T1 = GetWindowRect(Test2, RECTT2) ' TOP BAR
        'IRON BAR - DUMB BELLS
        TEST4 = FindWindowPart_BASEBAR("BaseBar")
        T1 = GetWindowRect(TEST4, RECTT4) ' LEFT BAR
        A = RECTT4.Top
        
        
        Y1 = (RECTT1.Bottom - RECTT1.Top) * Screen.TwipsPerPixelY
        Y2 = (RECTT2.Bottom - RECTT2.Top) * Screen.TwipsPerPixelY
        
        ' IF BAR NOT AT TOP BUT ON SIDE AND THEN USUAL HEIHGT
        If RECTT2.Bottom <= 40 Then Y2 = 480
        
        If TEST4 > 0 Then
            If RECTT4.Right = RECTT4.Left Then
                X4 = (RECTT4.Right) * Screen.TwipsPerPixelX
            Else
                X4 = (RECTT4.Right - RECTT4.Left) * Screen.TwipsPerPixelX
            End If
        End If
        
        SCREEN_WIDTH_SPACE = Screen.Width - X4
        SCREEN_HEIGHT_SPACE = Screen.Height - Y1 - Y2
        
        Me.Width = SCREEN_WIDTH_SPACE - 800
        Me.Height = SCREEN_HEIGHT_SPACE - 500
        
        'Form1.Left = (Screen.Width - Me.Width) / 2
        Me.Left = X4 + (SCREEN_WIDTH_SPACE - Me.Width) / 2
        Me.Top = Y2 + (SCREEN_HEIGHT_SPACE - Me.Height) / 2
    
'        DoEvents
        
'        Call SETUP_DIMENSIONS_INNER_FORM
    
    End If



End Sub



