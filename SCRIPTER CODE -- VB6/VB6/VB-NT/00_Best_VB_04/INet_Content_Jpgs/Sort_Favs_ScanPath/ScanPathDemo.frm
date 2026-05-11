VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ScanPath 
   Caption         =   "ScanPath 2.0 - Anything -- "
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12975
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   12975
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT TIME DO A FULL SCAN"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   1080
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GOTO RESULT FOLDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   120
      Width           =   810
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Height          =   825
      Left            =   11430
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1920
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4140
      Top             =   2115
   End
   Begin VB.ComboBox cboMask 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   495
      TabIndex        =   1
      Text            =   "*.jpg"
      Top             =   405
      Width           =   6105
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   6045
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   45
      Width           =   555
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   510
      TabIndex        =   0
      Text            =   "E:\My Music Zen"
      Top             =   30
      Width           =   5505
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "Search Sub-Folders"
      Height          =   195
      Left            =   1530
      TabIndex        =   8
      Top             =   2310
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkPatternMatching 
      Caption         =   "Pattern Matching"
      Height          =   195
      Left            =   1530
      TabIndex        =   7
      Top             =   2070
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkNormal 
      Caption         =   "Normal"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   2070
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   2310
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   2550
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   2775
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   5940
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2250
      Width           =   1545
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   0
      Left            =   9930
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2640
      Width           =   1365
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   8580
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2250
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   8580
      TabIndex        =   16
      Top             =   2640
      Width           =   1305
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   1
      Left            =   9930
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3000
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   8580
      TabIndex        =   18
      Top             =   3000
      Width           =   1305
   End
   Begin VB.CheckBox chkRefreshListView 
      Caption         =   "Refresh ListView (slower)"
      Height          =   195
      Left            =   1530
      TabIndex        =   10
      Top             =   2850
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9660
      ScaleHeight     =   255
      ScaleWidth      =   1665
      TabIndex        =   22
      Top             =   1860
      Visible         =   0   'False
      Width           =   1725
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
         Top             =   300
         Width           =   6015
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A High Performance API file/folder scanner Class."
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
         Left            =   2655
         TabIndex        =   24
         Top             =   240
         Width           =   4260
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
         Left            =   60
         TabIndex        =   23
         Top             =   30
         Width           =   1035
      End
   End
   Begin VB.CheckBox chkCopyMemory 
      Caption         =   "Use CopyMemory (display Date && Size)"
      Height          =   195
      Left            =   1530
      TabIndex        =   11
      Top             =   3090
      Width           =   3465
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3510
      Left            =   75
      TabIndex        =   21
      Top             =   3435
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   6191
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
      Top             =   2550
      Width           =   2535
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   810
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   0
      Left            =   5940
      TabIndex        =   13
      Top             =   2640
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   16515073
      CurrentDate     =   37299
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   1
      Left            =   5940
      TabIndex        =   14
      Top             =   3000
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   16515073
      CurrentDate     =   37296
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label13"
      Height          =   225
      Left            =   11640
      TabIndex        =   46
      Top             =   720
      Width           =   405
   End
   Begin VB.Label lblCount5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OOOOOOOOOO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   6780
      TabIndex        =   45
      Top             =   1455
      Width           =   2295
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   11640
      TabIndex        =   43
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Working"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   495
      TabIndex        =   42
      Top             =   765
      Width           =   6120
   End
   Begin VB.Label lblCount2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OOOOOOOOOO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   6780
      TabIndex        =   41
      Top             =   390
      Width           =   2295
   End
   Begin VB.Label lblCount4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OOOOOOOOOO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   6780
      TabIndex        =   40
      Top             =   1110
      Width           =   2295
   End
   Begin VB.Label lblCount3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OOOOOOOOOO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   6780
      TabIndex        =   39
      Top             =   750
      Width           =   2295
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
      Left            =   5130
      TabIndex        =   38
      Top             =   1830
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
      Top             =   1830
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mask"
      Height          =   195
      Left            =   45
      TabIndex        =   28
      Top             =   435
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      Height          =   195
      Left            =   45
      TabIndex        =   26
      Top             =   105
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
      Top             =   1830
      Width           =   885
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      Height          =   195
      Index           =   0
      Left            =   7770
      TabIndex        =   35
      Top             =   2700
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
      Height          =   195
      Left            =   7770
      TabIndex        =   33
      Top             =   2280
      Width           =   585
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   195
      Index           =   1
      Left            =   7770
      TabIndex        =   37
      Top             =   3060
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   5130
      TabIndex        =   34
      Top             =   2700
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "File Date"
      Height          =   195
      Left            =   5130
      TabIndex        =   32
      Top             =   2280
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   5130
      TabIndex        =   36
      Top             =   3000
      Width           =   195
   End
   Begin VB.Label lblCount1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OOOOOOOOOO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   6780
      TabIndex        =   31
      Top             =   30
      Width           =   2295
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sub INetContent()



Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Public OutPutPath2
Public Dog, OutPutPath, Xdate As Date, Tdate As Date, Zdate$, Ydate As Date, Path

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

'Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long


'Private Type FILETIME
'    dwLowDateTime     As Long
'    dwHighDateTime    As Long
'End Type

'Private Type WIN32_FIND_DATA
'    dwFileAttributes  As Long
'    ftCreationTime    As FILETIME
'    ftLastAccessTime  As FILETIME
'    ftLastWriteTime   As FILETIME
'    nFileSizeHigh     As Long
'    nFileSizeLow      As Long
'    dwReserved0       As Long
'    dwReserved1       As Long
'    cFileName         As String * MAX_PATH
'    cAlternate        As String * 14
'End Type

'Private mWFD As WIN32_FIND_DATA

'Put This in a Bas Mobule
'Public m_CRC As clsCRC

'Public WxHex$

Public A1$, B1$, G1$, FF$

'Put in Sub Of Code
'Set m_CRC = New clsCRC
'm_CRC.Algorithm = CRC32
'WxHex$ = Hex(m_CRC.CalculateString(Form3.Text1.Text))
'WxHex$ = Hex(m_CRC.CalculateFile(Filename))


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


Private Sub FORM_LOAD()

If App.PrevInstance = True Then End

ScanPath.Caption = ScanPath.Caption + App.EXEName
    
    
Set FS = CreateObject("SCRIPTING.FILESYSTEMOBJECT")


'--------------------------------------------
'OVERRIDE M DRIVE DONT EXIST LAPTOP ONLY RUN --------------------------------------
'ONE WHEN ISIDE
'If IsIDE = True Then GoTo JUMP
'--------------------------------------------
If FS.DRIVEEXISTS("M") = False Then End
Dim D
Set D = FS.GETDRIVE("M:\")
If D.ISREADY = False Then End
    
'EXIT SUB
    
FILESPEC = App.Path + "\" + App.EXEName + ".VBP"
If IsIDE = False And Dir$(FILESPEC) <> "" Then
    'TT = SHELL("C:\PROGRAM FILES\MICROSOFT VISUAL STUDIO\VB98\VB6.EXE /RUNEXIT """ + FILESPEC + """", VBMINIMIZEDNOFOCUS)
    'CALL KINGMOD.PROCESS_LOW(TT)
    'END
End If
    
JUMP:
  
  
ScanPath.Height = ListView1.Top + ListView1.Height + 450
ScanPath.Width = ListView1.Width + 210
    
    Dim lCount As Long
    
    With cboMask
        .AddItem "*.*"
        .AddItem "*.DLL;*.EXE;*.OCX"
        .AddItem "*.DOC;*.MDB;*.XLS"
        .AddItem "*.BMP;*.GIF;*.JPG;*.TIF"
        .AddItem "*.BAS;*.CLS;*.CTL;*.FRM;*.VBP"
        .ListIndex = 0
    End With
    
    DTPicker1(0).Value = DateSerial(Year(Now), Month(Now) - 3, Day(Now))
    DTPicker1(1).Value = Now
    
    DTPicker1(0).Value = Empty
    DTPicker1(1).Value = Empty
    
    With cboDate
        .AddItem "MODIFIED"
        .AddItem "CREATED"
        .AddItem "LAST ACCESSED"
        .ListIndex = 0
    End With
    
    With cboSize
        .AddItem "ANY SIZE"
        .AddItem "LESS THAN"
        .AddItem "GREATER THAN"
        .AddItem "BETWEEN"
        .ListIndex = 0
    End With
        
    For lCount = 0 To 1
        With cboSizeType(lCount)
            .AddItem "BYTES"
            .AddItem "KILOBYTES"
            .AddItem "MEGABYTES"
            .ListIndex = 0 ' SET IT TO BYTES HERE
        End With
    Next lCount
    
    With ListView1
        .ColumnHeaders.Add , "FILE", "FILE", 3000, lvwColumnLeft
        .ColumnHeaders.Add , "PATH", "PATH", 8000, lvwColumnLeft
        .ColumnHeaders.Add , "SIZE", "SIZE", 2000, lvwColumnLeft
        .ColumnHeaders.Add , "DATE", "DATE", 0, lvwColumnLeft
        
        .View = lvwReport
    End With


'SCANPATH.SHOW


'CALL BANGERS

Call INetContent

End Sub



Sub INetContent()
Label17 = "Process INet Content STAGE 1 - JPGs" + vbCrLf + "STATGE 2 - PNG GIF ICO" + vbCrLf + "STAGE 3 - INDEX.DAT"


If App.PrevInstance = True Then End

ScanPath.lblCount1 = ""
ScanPath.lblCount2 = ""
ScanPath.lblCount3 = ""
ScanPath.lblCount4 = ""
ScanPath.lblCount5 = ""

If IsIDE = False Then
    'Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" /runexit """ + App.Path + "\" + App.EXEName + ".vbp"" /c iiii", vbMinimizedNoFocus
    'End
End If


ChDrive "C:"
Set FS = CreateObject("Scripting.FileSystemObject")
    
If Command$ = "" Then
    Me.Show
    Me.WindowState = 0
End If

If IsIDE = True Then
    Me.WindowState = 0
End If

If Command$ <> "" Then
    Me.Show
    Me.WindowState = 1
End If

'Me.Show



DoEvents
'Exit Sub



'If Command$ = "" Then ScanPath.Show



OutPutPath = "D:\0 00 Art Loggers\# TEMP INTERNET\Temp Inet Files JPGs\"
OutPutPath2 = "D:\0 00 Art Loggers\# TEMP INTERNET\Temp Inet Files Gif Ico Png\"
OutPutPathDAT = "D:\0 00 Art Loggers\# TEMP INTERNET\Temp Inet Files DAT\"

On Error Resume Next
MkDir "D:\0 00 Art Loggers\# TEMP INTERNET\Temp Inet Files Gif Ico Png"
MkDir "D:\0 00 Art Loggers\# TEMP INTERNET\Temp Inet Files JPGs"
MkDir "D:\0 00 Art Loggers"
MkDir "D:\0 00 Art"

MkDir "D:\0 00 Art Loggers\# TEMP INTERNET\Temp Inet Files DAT"

'If GetComputerName = "55-88-HAPPY" Then
'    If FS.FolderExists(OutPutPath) = False Then End
'End If



MkDir OutPutPath2 + "\#Gif"
MkDir OutPutPath2 + "\#Png"
MkDir OutPutPath2 + "\#Ico"
'MkDir OutPutPathDAT + "\#DAT"

'On Error GoTo 0


'If Xdate = 0 Then Xdate = Now - 10
'Xdate = Now - 10


On Error GoTo 0



Path = "C:\Documents and Settings\" + GetUserName + "\Local Settings\Temporary Internet Files\"
'Path = "C:\Documents and Settings\MATT 01\Local Settings\Temporary Internet Files\"
If FS.FolderExists(Path) = False Then MsgBox "FOLDER NOT EXISTS"
TWOFOLDERS = "-02"

'Path = "\\55-88-happy\ASUS C (C)\Documents and Settings\Matt2\Local Settings\Temporary Internet Files\"
'If FS.FolderExists(Path) = False Then TWOFOLDERS = "-01"
'
TWOFOLDERS = "-01"
On Error Resume Next

Open App.Path + "\Xdate-" + GetComputerName + "--" + GetUserName + "-" + TWOFOLDERS + ".txt" For Input As #1
    Line Input #1, Zdate$
Close #1
'Zdate$ = ""



Ydate = DateValue(Zdate$) + TimeValue(Zdate$)
Ydate = Ydate - TimeSerial(0, 5, 0)

cboDate.ListIndex = 0
DTPicker1(0) = Ydate

cboMask.Text = "*.jpg"
'cboSize.ListIndex = 1 'Less than
cboSize.ListIndex = 2 'Bigger Than
'cboSizeType(lCount).ListIndex = 0 'Byte
'cboSizeType(lCount).ListIndex = 2 'Meg
cboSizeType(lCount).ListIndex = 1 'K
txtSize(0) = 3

DoEvents

'Path = "M:\My FeedStation Podcasts\"
'txtPath.Text = Path
'Call cmdScan_Click

Path = "C:\Documents and Settings\" + GetUserName + "\Local Settings\Temporary Internet Files\"
'Path = "C:\Documents and Settings\MATT 01\Local Settings\Temporary Internet Files\"
txtPath.Text = Path
Call cmdScan_Click

'Path = "\\55-88-happy\ASUS C (C)\Documents and Settings\Matt2\Local Settings\Temporary Internet Files\"
'txtPath.Text = Path
'Call cmdScan_Click






For we = ListView1.ListItems.Count To 1 Step -1
    
    lblCount2 = ListView1.ListItems.Count
    DoEvents
    
    G1$ = ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)
    
    On Error Resume Next
    Set F = FS.getfile(A1$ + G1$)
    Tdate = F.DateLastModified
    If Tdate > Xdate Then Xdate = Tdate
    
    remo = 0

'    If Tdate <= Ydate Then
'        ListView1.ListItems.Remove (we)
'        remo = 1
'    End If
    If remo = 0 Then
        hole = "Inet " + Format$(Now, "YYYY-MM-DD") + "\"
        GeText = LCase(Mid$(B1$, Len(B1$) - 2, 3)) + "\"
        If GeText = "jpg\" Then GeText = ""
        secext = OutPutPath + hole + GeText + B1$
        
        
        If FS.FileExists(secext) = True Then
            ListView1.ListItems.Remove (we)
        End If
        secext = OutPutPath + hole + GeText + Mid$(B1$, 1, Len(B1$) - 4) + ".zip"
        If FS.FileExists(secext) = True Then
            ListView1.ListItems.Remove (we)
        End If
        secext = OutPutPath + hole + GeText + Mid$(B1$, 1, Len(B1$) - 4) + ".rar"
        If FS.FileExists(secext) = True Then
            ListView1.ListItems.Remove (we)
        End If

    End If
Next

Open App.Path + "\Xdate-" + GetComputerName + "--" + GetUserName + "-" + TWOFOLDERS + ".txt" For Output Lock Write As #1
    Print #1, Xdate
Close #1
que = 0
' Xdate = 0
hole = "INet " + Format$(Now, "YYYY-MM-DD") + "\"
On Error Resume Next
If que = 0 Then MkDir OutPutPath + hole: que = 1
On Error GoTo 0

For we = ListView1.ListItems.Count To 1 Step -1
    
    lblCount3 = we
    DoEvents
    
    G1$ = ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)
    
    On Error Resume Next
    Set F = FS.getfile(A1$ + G1$)
'    Tdate = F.DateLastModified
    
    
    GeText = LCase(Mid$(B1$, Len(B1$) - 2, 3)) + "\"
    If GeText = "jpg\" Then GeText = ""
    
        
    halo = 0
    hatt = LCase(B1$)
    If InStr(B1$, ".flv-0[1].jpg") Then halo = 1
    If InStr(B1$, ".mp4-3[1].jpg") Then halo = 1
    If InStr(B1$, ".mpg-3[1].jpg") Then halo = 1
    If InStr(B1$, ".mmv-3[1].jpg") Then halo = 1
    If InStr(B1$, ".avi-3[1].jpg") Then halo = 1
    If InStr(B1$, "_bay-web.large[1].jpg") Then halo = 1
    If InStr(B1$, "10495_1[1].jpg") Then halo = 1 ' May no not been a run of these
    
    
    '.flv-0[1].jpg
    quein = Right$(LCase(B1$), 13)
    If Mid(B1$, 1, 1) = "." And Mid(B1$, 5, 1) = "-" And InStr(hatt, "squirt") Then halo = 1
    
    '-10497_1[1].jpg
    quein = Right$(LCase(B1$), 15)
    If Mid(B1$, 1, 1) = "-" And Mid(B1$, 5, 1) = "_" And InStr(hatt, "anal") Then halo = 1
    'fucking
    
    'same as above if 5 digit number on end
    quein = Mid$(LCase(B1$), 15, 5)
    If IsNumeric(quein) > 0 Then halo = 1
    
    
    tg = LCase(B1$)
    Do
    t = 0
    tx = "[1].jpg"
    If t = 0 Then t = InStr(tg, tx): If t > 0 Then tg = Mid$(tg, 1, t - 2) + Mid$(tg, t + Len(tx))
    tx = "%"
    If t = 0 Then t = InStr(tg, tx): If t > 0 Then tg = Mid$(tg, 1, t - 1) + Mid$(tg, t + Len(tx))
    tx = "-"
    If t = 0 Then t = InStr(tg, tx): If t > 0 Then tg = Mid$(tg, 1, t - 1) + Mid$(tg, t + Len(tx))
    tx = "_"
    If t = 0 Then t = InStr(tg, tx): If t > 0 Then tg = Mid$(tg, 1, t - 1) + Mid$(tg, t + Len(tx))
    tx = "."
    If t = 0 Then t = InStr(tg, tx): If t > 0 Then tg = Mid$(tg, 1, t - 1) + Mid$(tg, t + Len(tx))
    Loop Until t = 0
    
    tw = IsNumeric(tg)
    On Error Resume Next
    If tw = False Then tw = Str("&h" + tg)
    On Error GoTo 0
    If tw = True Then tw = 1
    
    If Mid(B1$, 7, 1) = "." And Mid(B1$, 9, 1) = "." And Mid$(B1$, 13, 1) = "." And tw > 0 Then halo = 1
    If Mid(B1$, 6, 1) = "." And Mid(B1$, 8, 1) = "." And Mid$(B1$, 12, 1) = "." And tw > 0 Then halo = 1

    
    
    'If Val(Mid(B1$, 1, 6)) > 0 Then halo= 1
    
        
    '#0XxXRay
    
    halo2 = ""
    If halo = 1 Then
        halo2 = "#0XxXRay\"
    End If
        
    On Error Resume Next
    MkDir OutPutPath + hole + GeText + halo2
    
    MkDir OutPutPath + hole + GeText
    If halo = 1 Then MkDir OutPutPath + hole + GeText + halo2
    On Error GoTo 0
    
    
    On Error Resume Next
    F.Copy OutPutPath + hole + GeText + halo2
    On Error GoTo 0
'    Set F = Fs.getfile(OutPutPath + hole + GeText)
'    pdate = F.DateLastModified
'    If pdate <> Tdate Then
'        MsgBox "Dates of File Copy are not in sync Inet Content"
'    End If

Next

'-------------------------------------
'DTPicker1(0).Value = vbNullString
ListView1.ListItems.Clear






cboSize.ListIndex = 0 ' Not Used
cboMask.Text = "*.png;*.ico;*.gif;*.dat"

    
   Path = "C:\Documents and Settings\" + GetUserName + "\Local Settings\Temporary Internet Files\"
'   Path = "C:\Documents and Settings\MATT 01\Local Settings\Temporary Internet Files\"
    If FS.FolderExists(Path) = False Then MsgBox "FOLDER NOT EXISTS"
    If FS.FolderExists(Path) = True Then
        txtPath.Text = Path
        Call cmdScan_Click
    End If
    
    'Path = "\\55-88-happy\ASUS C (C)\Documents and Settings\Matt2\Local Settings\Temporary Internet Files\"
    'If FS.FolderExists(Path) = False Then MsgBox "FOLDER NOT EXISTS"
    If FS.FolderExists(Path) = True Then
        txtPath.Text = Path
        Call cmdScan_Click
    End If


For we = ListView1.ListItems.Count To 1 Step -1
    
    lblCount4 = we
    DoEvents
    
    G1$ = ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)
    Set F = FS.getfile(A1$ + G1$)
    Tdate = F.DateLastModified
    If Tdate > Xdate Then Xdate = Tdate
    remo = 0
    If Tdate <= Ydate Then
        ListView1.ListItems.Remove (we)
        remo = 1
    End If
    If remo = 0 Then
        GeText = "#" + LCase(Mid$(B1$, Len(B1$) - 2, 3)) + "\"
        If FS.FileExists(OutPutPath2 + GeText + B1$) = True Then
        ListView1.ListItems.Remove (we)
        End If
    End If
Next

For we = ListView1.ListItems.Count To 1 Step -1
    
    lblCount5 = we
    DoEvents
    
    G1$ = ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)
    
    On Error Resume Next
    Set F = FS.getfile(A1$ + G1$)

    'test
    'Tdate = F.DateLastModified
    
    If InStr(UCase(B1$), ".DAT") > 0 Then
        GeText = "#" + LCase(Mid$(B1$, Len(B1$) - 2, 3)) + "\"
        
        XDATE5 = Format$(F.DateLastModified, "YYYY-MM-DD HH-MM-SS")
        F.Copy OutPutPathDAT + "\" + XDATE5 + " -- " + GetComputerName + " -- " + GetUserName + " -- " + B1$
'        F.Copy OutPutPathDAT + "\" + XDATE5 + " -- " + GetComputerName + " -- " + "MATT 01" + " -- " + B1$
    Else
        GeText = "#" + LCase(Mid$(B1$, Len(B1$) - 2, 3)) + "\"
        F.Copy OutPutPath2 + "\" + GeText
    End If
    On Error GoTo 0

'    Set F = Fs.getfile(OutPutPath + GeText)
'    pdate = F.DateLastModified

'    If pdate <> Tdate Then
'        MsgBox "Dates of File Copy are not in sync Inet Content"
'    End If
Next

'ListView1.ListItems.Clear
'txtPath.Text = "C:\Documents and Settings\Matt4\Local Settings\Temporary Internet Files\Content.IE5\"
'cboSize.ListIndex = 0 ' Not Used
'cboMask.Text = "*.png;*.ico;*.gif"
'Call cmdScan_Click
'Call cmdScanDir_Click

'Next rx


Open App.Path + "\Xdate-" + GetComputerName + "-" + GetUserName + "-" + TWOFOLDERS + ".txt" For Output Lock Write As #1
Print #1, Format$(Xdate, "DD-MM-YYYY HH:MM:SS")
Close #1



If Command$ <> "" Then q = " " + Command$

TT$ = "D:\VB6\VB-NT\00_Best_VB_04\INet_Content_Jpgs_CRC\INet_Content_Jpgs_CRC.exe"
If Dir$(TT$) = "" Then End

'If IsIDE = False And GetComputerName <> "55-88-HAPPY" Then
If GetComputerName <> "55-88-HAPPY" Then
'    Shell TT$ + q, vbNormalNoFocus
End If

ScanPath.ListView1.BackColor = ScanPath.Label13.BackColor
lblCount5 = " ---- DONE ----"
If IsIDE = False Then End
If ScanPath.Visible = False Then End
If ScanPath.WindowState <> 0 Then End
If Command$ <> "" Then End
If ScanPath.WindowState = 0 Then Exit Sub

End Sub

Private Sub FORM_RESIZE()
'CENTER SCREEN
On Error Resume Next
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2

End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
End
End Sub


Private Sub cboSize_Click()
    'cboSize.ListIndex = 2
    'txtSize(0).Text = 1
    
    'cboSizeType(0).Visible = True
    'cboSizeType(0).

    
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


Private Sub cboSizeType_Change(Index As Integer)
'cboSizeType = 0
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
            .ColumnHeaders("SIZE").Width = 1000
            .ColumnHeaders("DATE").Width = 0
        End If
    End With
End Sub


Private Sub chkSort_Click()
    chkCopyMemory.Enabled = (chkSort.Value = vbUnchecked)
End Sub

Private Sub cmdBrowse_Click()
    txtPath.Text = GetFolder(Me.hWnd, "Scan Path:", txtPath.Text)
    fg1 = FreeFile
    Open "\Temp\Scan Path.txt" For Output As #fg1
    Print #fg1, txtPath.Text
    Close #fg1

'    fg1 = FreeFile
'    Open App.Path + "\Scan Path.txt" For Input As #fg1
'    Line Input #fg1, ae$
'    Close #fg1
'    txtPath.Text = ae$



End Sub

Private Sub cmdHelp_Click()
    'txtHelp.Visible = Not txtHelp.Visible
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
        
        lblCount1.Caption = "-"
 
        Screen.MousePointer = vbHourglass
'        ListView1.ListItems.Clear
        
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
            
            lblCount1.Caption = .ListItems.Count
        End With
        
        LockWindowUpdate 0
        
        'cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
        'End
    End If

'frmMain.Refresh

'End
'If Command$ <> "" Then End
cmdScan.Caption = "Scan"
End Sub
Public Sub cmdScanDir_Click()
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
        lblCount2.Caption = "-"
 
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
            .StartScanDir txtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
        End With
        
       With ListView1
            If .ListItems.Count > 0 Then
                .SelectedItem = .ListItems(1)
            End If
            
            lblCount1.Caption = .ListItems.Count
        End With
        
        LockWindowUpdate 0
        
        cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
    End If


GoTo Skip

For we = ListView1.ListItems.Count To 2 Step -1
    A1$ = frmMain.ListView1.ListItems.Item(we)
    a2$ = frmMain.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = frmMain.ListView1.ListItems.Item(we - 1)
    b2$ = frmMain.ListView1.ListItems.Item(we - 1).SubItems(1)
    If A1$ + B1$ = a2$ + b2$ Then frmMain.ListView1.ListItems.Remove (we)
    
    
    
    'b1$ = frmMain.ListView1.ListItems.Item(we)

Next

Skip:

lblCount1.Caption = ListView1.ListItems.Count
    

End Sub




Private Sub Command1_Click()
Shell "EXPLORER.EXE /select, D:\0 00 Art Loggers\# TEMP INTERNET\Temp Inet Files JPGs\", vbMinimizedNoFocus
End Sub


Private Sub List1_Click()

Dim fs22
Set fs22 = CreateObject("Scripting.FileSystemObject")


tpath3$ = "M:\04 Music ---\Del\"

B1$ = Mid$(List1.List(List1.ListIndex), 14)
C1$ = B1$ + "--Todel"

'tpath1$ = "M:\04 Music ---\"

C1$ = Mid$(B1$, InStrRev(B1$, "\") + 1)
List1.RemoveItem (List1.ListIndex)
'Name b1$ As c1$

'fs22.copyFile b1$, tpath3$ + c1$
If Dir$(tpath3$ + C1$) <> "" Then Kill tpath3$ + C1$
fs22.moveFile B1$, tpath3$ + C1$



End Sub


Private Sub Label18_Click()
Dog = True

If cmdScan.Caption = "Stop" Then cmdScan_Click
Label17 = "Stopped before Complete"
End Sub

Private Sub lblCount1_Change()
    'Main.Command1.Caption = ScanPath.lblCount1

End Sub

Private Sub SP_DirMatch(Directory As String, DDirectory As String, Path As String)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################
End Sub

Private Sub SP_DirMatchxx(Directory As String, DDirectory As String, Path As String)
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
        LV.SubItems(2) = DDirectory
        
        
        
        
        
        
        
        
        
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
            
            lblCount2.Caption = LV.Index
        End If
    End With
    Exit Sub

GetFileError:
    MsgBox Err.Description, vbCritical


End Sub


Private Sub SP_FileMatch(FileName As String, DFilename As String, Path As String)
    Dim LV As ListItem
    Dim uWIN32 As WIN32_FIND_DATA
    
    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
    On Local Error GoTo GetFileError
    
    With ListView1
        Set LV = .ListItems.Add(, , FileName)
        LV.SubItems(1) = Path
        'mod by matt l
        LV.SubItems(2) = DFilename
        
        
        
        
        
        
        
        
        
        
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
            
            lblCount1.Caption = LV.Index
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
'End
End Sub














Sub Jeepers()


Dim fs22

Set fs22 = CreateObject("Scripting.FileSystemObject")

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

For we = 1 To ListView1.ListItems.Count
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
    C1$ = Drived2$ + "Temp\Anything" + v1$ + "\" + Mid$(A1$, ets1 + 2)

    etd1 = 13
    ets2 = 1
    Do
        'jc1$ = Mid$(c1$, etd1 + ets2-1)
        ets2 = InStr(etd1 + ets2, C1$, "\")
        etd1 = 1
        If ets2 = 0 Then Exit Do

        d2$ = Mid$(C1$, 1, ets2)

        If InStr(f1$, d2$) = 0 Then
            On Local Error GoTo jeep
            MkDir d2$
            If errs2 <> 75 And errs2 > 0 Then
                MsgBox ("Error Making Temp Directory")
                Me.WindowState = 0
                Stop
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    f1$ = d2$

    'Dim Fs22
    'Set Fs22 = CreateObject("Scripting.FileSystemObject")
    errs2 = 0
    On Local Error GoTo jeep
    fs22.moveFile A1$ + B1$, C1$ + B1$
    On Local Error GoTo 0

    If errs2 <> 0 Then
        MsgBox ("Error Moving File to Temp")
        End
    End If

    List1.AddItem Format$(we, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'Fs22.DeleteFolder comrad$ + "\" + WFD$

Next

Err.Clear
On Local Error Resume Next
Set fs22 = CreateObject("Scripting.FileSystemObject")
fs22.deletefolder txtPath.Text, True
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
Set fs22 = CreateObject("Scripting.FileSystemObject")
fs22.Movefolder Drived2$ + "Temp\Anything\*", v2$
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



Sub HardDel()
Dim fs22

Set fs22 = CreateObject("Scripting.FileSystemObject")

Drived2$ = Mid$(txtPath.Text, 1, 3)
v1$ = Mid$(txtPath.Text, InStrRev(txtPath.Text, "\"))


List1.AddItem "Stage Opp : Del Original Dir"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For we = 1 To ListView1.ListItems.Count
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
    C1$ = Drived2$ + "My Music" + v1$ + "\" + Mid$(A1$, ets1 + 2)

    etd1 = 13
    ets2 = 1
    Do
        'jc1$ = Mid$(c1$, etd1 + ets2-1)
        ets2 = InStr(etd1 + ets2, C1$, "\")
        etd1 = 1
        If ets2 = 0 Then Exit Do

        d2$ = Mid$(C1$, 1, ets2 - 1)

        If InStr(f1$, d2$) = 0 Then
            Err.Clear
            On Local Error Resume Next
            'MkDir d2$
            fs22.deletefolder d2$, True
            If Err.Number <> 75 And Err.Number > 0 Then
                'MsgBox ("Error Making Temp Directory")
                '
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

    List1.AddItem Format$(we, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'Fs22.DeleteFolder comrad$ + "\" + WFD$

Next

a = a
jeep:
End Sub


Sub HardMove()
Dim fs22

Set fs22 = CreateObject("Scripting.FileSystemObject")

Drived2$ = Mid$(txtPath.Text, 1, 3)
v1$ = Mid$(txtPath.Text, InStrRev(txtPath.Text, "\"))


List1.AddItem "Stage Opp : Del Original Dir"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For we = 1 To ListView1.ListItems.Count
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
    C1$ = Drived2$ + "My Music" + v1$ + "\" + Mid$(A1$, ets1 + 2)
    c2$ = Drived2$ + "Temp\Anything" + v1$ + "\" + Mid$(A1$, ets1 + 2)

    etd1 = 13
    ets2 = 1
    Do
        'jc1$ = Mid$(c1$, etd1 + ets2-1)
        ets2 = InStr(etd1 + ets2, C1$, "\")
        etd1 = 1
        If ets2 = 0 Then Exit Do

        d2$ = Mid$(C1$, 1, ets2 - 1)

        If InStr(f1$, d2$) = 0 Then
            Err.Clear
            On Local Error Resume Next
            MkDir d2$
   '         Fs22.deletefolder d2$, True
            If Err.Number <> 75 And Err.Number > 0 Then
                'MsgBox ("Error Making Temp Directory")
                'T
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    f1$ = d2$
    Err.Clear
    errs2 = 0
    On Local Error Resume Next
    Set fs22 = CreateObject("Scripting.FileSystemObject")
    fs22.moveFile c2$ + B1$, A1$ + B1$
    'err.number
    'err.description
    On Local Error GoTo 0

    'If errs2 <> 0 Then
    '    MsgBox ("Error Moving File to Temp")
    '    End
    'End If

    List1.AddItem Format$(we, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'Fs22.DeleteFolder comrad$ + "\" + WFD$

Next
    
    v1$ = Mid$(txtPath.Text, 1, InStrRev(txtPath.Text, "\") - 1)
    Err.Clear
    On Local Error Resume Next
    Set fs22 = CreateObject("Scripting.FileSystemObject")
    fs22.copyfolder Drived2$ + "Temp\Anything", v1$, True
    
    Set fs22 = CreateObject("Scripting.FileSystemObject")
    fs22.deletefolder Drived2$ + "Temp\Anything", True


jeep:
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


