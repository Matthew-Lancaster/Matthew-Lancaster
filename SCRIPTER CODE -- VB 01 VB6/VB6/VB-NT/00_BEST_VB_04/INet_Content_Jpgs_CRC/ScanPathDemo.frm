VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ScanPath 
   Caption         =   "ScanPath 2.0 - Sort Anything -"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14055
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   14055
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox Chk_SortOnDate 
      Caption         =   "Sort On Date"
      Height          =   195
      Left            =   3225
      TabIndex        =   47
      Top             =   2640
      Width           =   1260
   End
   Begin VB.CommandButton CmdScanDir 
      Caption         =   "Scan Dir"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   825
      Left            =   12255
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   795
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Enabled         =   0   'False
      Height          =   825
      Left            =   12480
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3000
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   585
      Top             =   5355
   End
   Begin VB.ComboBox cboMask 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   480
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
      Width           =   525
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Text            =   "E:\My Music Zen"
      Top             =   45
      Width           =   5520
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "Search Sub-Folders"
      Height          =   195
      Left            =   1530
      TabIndex        =   8
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkPatternMatching 
      Caption         =   "Pattern Matching"
      Height          =   195
      Left            =   1530
      TabIndex        =   7
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkNormal 
      Caption         =   "Normal"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   3345
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   3570
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   5385
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2820
      Width           =   1545
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   0
      Left            =   10980
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3210
      Width           =   1365
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   9630
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2820
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   9630
      TabIndex        =   16
      Top             =   3210
      Width           =   1305
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   1
      Left            =   10980
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3570
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   9630
      TabIndex        =   18
      Top             =   3570
      Width           =   1305
   End
   Begin VB.CheckBox chkRefreshListView 
      Caption         =   "Refresh ListView (slower)"
      Height          =   195
      Left            =   1530
      TabIndex        =   10
      Top             =   3420
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   11400
      ScaleHeight     =   255
      ScaleWidth      =   1665
      TabIndex        =   22
      Top             =   2640
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
      Top             =   3660
      Width           =   3060
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3510
      Left            =   75
      TabIndex        =   21
      Top             =   4005
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   6191
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
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
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Enabled         =   0   'False
      Height          =   825
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   780
      Visible         =   0   'False
      Width           =   810
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   0
      Left            =   5385
      TabIndex        =   13
      Top             =   3210
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   16449537
      CurrentDate     =   37299
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   1
      Left            =   5385
      TabIndex        =   14
      Top             =   3570
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   16449537
      CurrentDate     =   37296
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   2
      Left            =   7020
      TabIndex        =   48
      Top             =   3210
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16449538
      CurrentDate     =   37299
   End
   Begin VB.Label lblCount8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100.00%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   4215
      TabIndex        =   52
      Top             =   900
      Width           =   2370
   End
   Begin VB.Label lblCount7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6675
      TabIndex        =   51
      Top             =   1650
      Width           =   4935
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6675
      TabIndex        =   50
      Top             =   105
      Width           =   4935
   End
   Begin VB.Label lblCount6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6675
      TabIndex        =   49
      Top             =   2250
      Width           =   4935
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   13365
      TabIndex        =   44
      Top             =   1710
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Working"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   13215
      TabIndex        =   43
      Top             =   2010
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblCount2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6675
      TabIndex        =   42
      Top             =   735
      Width           =   4935
   End
   Begin VB.Label lblCount5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6675
      TabIndex        =   41
      Top             =   1950
      Width           =   4935
   End
   Begin VB.Label lblCount4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6675
      TabIndex        =   40
      Top             =   1350
      Width           =   4935
   End
   Begin VB.Label lblCount3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6675
      TabIndex        =   39
      Top             =   1050
      Width           =   4935
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
      Left            =   4725
      TabIndex        =   38
      Top             =   2595
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
      Top             =   2400
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
      Top             =   2400
      Width           =   885
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      Height          =   195
      Index           =   0
      Left            =   9045
      TabIndex        =   35
      Top             =   3270
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
      Height          =   195
      Left            =   9030
      TabIndex        =   33
      Top             =   2850
      Width           =   585
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   195
      Index           =   1
      Left            =   9045
      TabIndex        =   37
      Top             =   3630
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   4725
      TabIndex        =   34
      Top             =   3270
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "File Date"
      Height          =   195
      Left            =   4725
      TabIndex        =   32
      Top             =   2850
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   4740
      TabIndex        =   36
      Top             =   3570
      Width           =   195
   End
   Begin VB.Label lblCount1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6675
      TabIndex        =   31
      Top             =   420
      Width           =   4935
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' LOOK FOR call INETPix


'-------------Is This Sub Get on It at start      at       Bangers

Public Dog, OutPutPath, Xdate As Date, Tdate As Date, Zdate$, Ydate As Date

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




Private Sub Form_Activate()
ScanPath.ListView1.SetFocus

End Sub

Private Sub Form_GotFocus()
ScanPath.ListView1.SetFocus

End Sub

Private Sub Form_Load()
    
If App.PrevInstance = True Then End



Set FS = CreateObject("Scripting.FileSystemObject")
    
 '--------------------------------------------
'OVERRIDE M DRIVE DONT EXIST LAPTOP ONLY RUN --------------------------------------
'ONE WHEN ISIDE
If IsIDE = True Then GoTo JUMP
    
If FS.DRIVEEXISTS("m") = False Then End
Dim D
Set D = FS.GETDRIVE("M:\")
If D.ISREADY = False Then End

JUMP:
    
ScanPath.Caption = ScanPath.Caption + App.EXEName
'scanpath.Caption = frmMain.Caption + App.EXEName
  
ScanPath.Height = ListView1.Top + ListView1.Height + 450
ScanPath.Width = ListView1.Width + 210
    
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
            .ListIndex = 0 ' set it to bytes here
        End With
    Next lCount
    
    With ListView1
        .ColumnHeaders.Add , "FILE", "File", 3000, lvwColumnLeft
        .ColumnHeaders.Add , "PATH", "Path", 8000, lvwColumnLeft
        .ColumnHeaders.Add , "SIZE", "Size", 2000, lvwColumnLeft
        .ColumnHeaders.Add , "DATE", "Date", 0, lvwColumnLeft
        
        .View = lvwReport
    End With





Call INETPix

End Sub


Private Sub Form_Resize()
On Error Resume Next
ScanPath.ListView1.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
Reset
End
Exit Sub

Dim D
On Local Error Resume Next
Set dc = FS.Drives
For Each D In dc
    dr = D.DriveLetter
    n = D.VolumeName
    If InStr(n, "RAM") > 0 Then Exit For
Next
On Local Error GoTo 0
ramdrive = dr
F = FreeFile
Open ramdrive + ":\Temp\JpegCRCDoneFlag.txt" For Output As #F
Close #F
 
Reset
End
End Sub




Public Sub INETPix()
'--------------------
'THE DATE IS TAKEN AND ANY UNDER THE VALUE WILL BE CHECKED AGAINST CRC UNTIL .............
'--------------------
'THIST HAS AN ERROR OF BAD CRASH BUT SPEED IS UP IS GOOD
'--------------------
'THE CRC ARE TAKEN AND STORED AND IF THE DATE IS MISSING OR SET BACK
'BEYOND LAST CRC UNWANTED DELETES ARE MADE
'--------------------
'ALSO GOOD AS MORE PORTABLE THE OLD CRC CHECKED FILES DONT HAVE TO EXIST
'SO CAN STORE ON D DRIVE
'--------------------

'Open "D:\TEMP\123.txt" For Output As #1
'Close #1

'Call ShellFileDelete("D:\TEMP\123.txt", True)

'End
'




If Command$ = "" Then
    ScanPath.Show
    DoEvents
    ScanPath.ListView1.SetFocus
Else
    ScanPath.Show
    'DoEvents
    'ScanPath.ListView1.SetFocus
    ScanPath.WindowState = 1


End If

Call ScanPath.chkCopyMemory_Click
'Put in Sub Of Code

Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32




Dim Echo
Echo = "\Temp Inet Files JPGs\"
'Echo = Mid$(ScanPath.txtPath.Text, InStrRev(ScanPath.txtPath.Text, "\", Len(ScanPath.txtPath.Text) - 1))

CRCStore = App.Path + "\VBDataNoArchive" + Echo + "DataCRC.txt"
CrcVars = App.Path + "\VBDataNoArchive" + Echo + "DataCount & Date.txt"

Dim pos&
Dim src As String
Dim des As String

'****************
' Fast Version
'****************
'Allocates a big string, then copies the smaller strings into it.
'this means that VB does not need to perform the expensive
'dynamic realocation of string memory!

On Error Resume Next
f1 = FreeFile
Open CrcVars For Input As #f1
Line Input #f1, o1
Line Input #f1, o2
Close #f1
CountD = Val(o1)
DateD = DateValue(o2) + TimeValue(o2)
'DateD = 0
If DateD = 0 Then MsgBox "Error No Date": Unload ScanPath: Exit Sub
If DateD = 0 Then
    DateD = DateSerial(1601, 1, 1) 'Min Date Picker Value - dont over do it
    Open CRCStore For Output As #1
    Close #1
End If
On Error GoTo 0
If CountD = 0 Then MsgBox "Error ":  Unload ScanPath: Exit Sub
DelCount1 = CountD
ScanPath.lblCount3 = Str(DelCount1) + " Total Dupes Deleted - NOV 2010"
ScanPath.lblCount4 = Str(DelCount2) + " Dupes Deleted This Time"
ScanPath.lblCount5 = Format(DateD, "DDD DD-MM-YY HH:MM:SSa/p") + " -- Last Run"



On Error Resume Next
    f1 = FreeFile
    Open CRCStore For Input As #f1
        des = Input(LOF(f1), f1)
    Close #f1
On Error GoTo 0
pos& = LenB(des)

des = des & Space$(10000)
    
ScanPath.lblCount8 = "0%"
    
'DateD = DateSerial(1601, 1, 1) 'Min Date Picker Value - dont over do it

ScanPath.cboDate.ListIndex = 0
ScanPath.DTPicker1(0) = DateD
ScanPath.DTPicker1(2) = DateD
ScanPath.cboMask.Text = "*.jpg;*.ico;*.gif;*.png;*.dat"


DRVAR = "D"
If FS.DRIVEEXISTS("M") = False Then
    DRVAR = "D"
Else
    Dim D
    Set D = FS.GETDRIVE("M:")
    If D.ISREADY = False Then DRVAR = "D"
End If


'ScanPath.txtPath.Text = DRVAR + ":\0 00 Art\My FeedStation Podcasts\## TODAY"
'Call ScanPath.cmdScan_Click



ScanPath.txtPath.Text = DRVAR + ":\D:\0 00 Art Loggers"
Call ScanPath.cmdScan_Click
'If InStr(LCase(Command$), "full") > 0 Then
'    ScanPath.txtPath.Text = "M:\0 00 Art Loggers\Temp Inet Files Gif Ico Png\"
'    Call ScanPath.cmdScan_Click
'End If

'TRYING TO MAKE A SCAN OF JUST FOLDERS WITH DATE MATCH BEYOND BECOZ BEEN TO BIG
'daysback = DateDiff("d", DateD - 1, Now)
'For R = daysback To 1 Step -1
'    Td = "INet " + Format$(Now - R, "YYYY-MM-DD")
'    ScanPath.txtPath.Text = DRVAR + ":\0 00 Art Loggers\Temp Inet Files JPGs\" + Td
'
'Next
'    Call ScanPath.cmdScan_Click
'EXTRA FOLDERS
'ScanPath.txtPath.Text = "M:\0 00 Art\Google Downloader"
'Call ScanPath.cmdScan_Click

'EXTRA FOLDERS






'Whole Scan if you Like
'ScanPath.txtPath.Text = "M\0 00 Art Loggers\Temp Inet Files JPGs\"
'Call ScanPath.cmdScan_Click

'Exit Sub


'If any rescanning is done you have to delete that part of the CRC list or all of it for Complete Scan
'OR DO IT WITH A FULL SCANNER THEN CONTINUE USING THE CRC SCRIPT

ScanPath.lblCount1 = Str(ScanPath.ListView1.ListItems.Count) + " Total After Scan"

On Error Resume Next
MkDir App.Path + "\VBDataNoArchive"
MkDir App.Path + "\VBDataNoArchive" + Echo
On Error GoTo 0

f1 = FreeFile
Open CRCStore + ".bak" For Append As #f1
ScanPath.lblCount2 = Str(we) + "  Processed"
we2 = ScanPath.ListView1.ListItems.Count
For we = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)

    w = 0
    If FS.FileExists(A1$ + G1$) = True Then

        Set F = FS.getfile(A1$ + G1$)
        ADate1 = F.datelastmodified
        
        'If (we Mod 10) = 0 Then
            ScanPath.lblCount2 = Str(we) + " Processed"
            ScanPath.lblCount8 = Format((we / we2) * 100, "0.00") + "%"
            ScanPath.lblCount7 = Format(ADate1, "DD-MMM-YY") + " File Now"
            
        
        'End If


        If ADate1 > DateDx Then
            DateDx = ADate1
            'is this correct datedx was dated dated not right surely
            ScanPath.lblCount6 = Format(DateDx, "DDD DD-MM-YY HH:MM:SSa/p") + " -- This Run"
        End If
        
        TSize = F.Size
        WxHex$ = Space$(10)
        LSet WxHex$ = Hex(m_CRC.CalculateFile(A1$ + G1$))
        If InStr(des, WxHex$ + Str(TSize)) = 0 Then
            Print #f1, WxHex$ + Str(TSize)
            w = 1
        Else
            DelCount1 = DelCount1 + 1
            DelCount2 = DelCount2 + 1
            ScanPath.lblCount3 = Str(DelCount1) + " Total Dupes Deleted"
            ScanPath.lblCount4 = Str(DelCount2) + " Dupes Deleted"
            
            ScanPath.ListView1.ListItems.Item(we).EnsureVisible
            ScanPath.ListView1.ListItems.Item(we).Selected = True
            DoEvents
            
            'To Recycle
            'Call ShellFileDelete(A1$ + B1$, True)
            w = DeleteFile(A1$ + B1$)
            w = FS.FileExists(A1$ + G1$)
            If w = True Then
                'Call ShellFileDelete(A1$ + G1$, True)
                w = DeleteFile(A1$ + G1$)
            End If
            w = FS.FileExists(A1$ + G1$)
            
            'TEMP OUT
            'If w = True Then MsgBox "Delete Failed For this File " + vbCrLf + a14 + "\" + B1$ + vbCrLf + G1$ + vbCrLf + "Will try again on next Run - you wish"
        
        End If
        
        If w = 1 Then
            src = " " + WxHex$ + Str(TSize)
            If pos& + LenB(src) > LenB(des) Then
                des = des & Space$(50000)
            End If
            CopyMemory ByVal StrPtr(des) + pos&, ByVal StrPtr(src), LenB(src)
            pos& = pos& + LenB(src)
        End If
    End If

Next

'des = Left$(des, pos& \ 2) 'trim back to size

Close fr1

If ScanPath.lblCount6 = "-" Then
    ScanPath.lblCount6 = "This Run Nothing Done"
Else
    FS.CopyFile CRCStore + ".bak", CRCStore

    f1 = FreeFile
    Open CrcVars For Output As #f1
        Print #f1, DelCount1
        Print #f1, Format(DateDx, "DD-MM-YYYY HH:MM:SS")
    Close #f1
End If


'If DateDX > 0 then something must have been set or else no need write file markers nothing been done
FS.CopyFile CRCStore + ".bak", CRCStore
'DateDx = 0 when deletes have been made


If DateDx > 0 Then
    f1 = FreeFile
    Open CrcVars For Output As #f1
        Print #f1, DelCount1
        Print #f1, Format(DateDx, "DD-MM-YYYY HH:MM:SS")
    Close #f1
End If


If Command$ <> "" Then Unload ScanPath: Exit Sub
If ScanPath.Visible = False Then Unload ScanPath

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

Public Sub chkCopyMemory_Click()
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
    Open App.Path + "\Scan Path.txt" For Output As #fg1
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
        
        'lblCount.Caption = "-"
 
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
            
            'we always use direct scan as is better to sort after job then during thats all we need
            If Chk_SortOnDate.Value = vbChecked Then
                ListView1.SortOrder = lvwAscending
                ListView1.SortKey = 3
                ListView1.Sorted = True
                ListView1.Sorted = False
            Else
                ListView1.SortOrder = lvwAscending
                ListView1.SortKey = 0
                ListView1.Sorted = True
                ListView1.SortKey = 1
                ListView1.Sorted = True
                ListView1.Sorted = False
            End If
        
        End With
        
       With ListView1
            If .ListItems.Count > 0 Then
                .SelectedItem = .ListItems(1)
            End If
            
            'lblCount.Caption = .ListItems.Count
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

If Chk_SortOnDate = vbChecked Then

End If

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
        
            
            'Correct Yes 1st one column sort keeps sort then another
            ListView1.SortOrder = lvwAscending
            ListView1.SortKey = 0
            ListView1.Sorted = True
            ListView1.SortKey = 1
            ListView1.Sorted = True
            ListView1.Sorted = False
        
            If Chk_SortOnDate.Value = vbChecked Then
                ListView1.SortOrder = lvwAscending
                ListView1.SortKey = 3
                ListView1.Sorted = True
                ListView1.Sorted = False
            End If
        
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

lblCount.Caption = ListView1.ListItems.Count
    

End Sub




Private Sub Label13_Click()

End Sub

Private Sub Label18_Click()
Dog = True

If cmdScan.Caption = "Stop" Then cmdScan_Click
Label17 = "Stopped before Complete"
End Sub

Private Sub lblCount_Change()
    'Main.Command1.Caption = ScanPath.lblCount

End Sub

Private Sub lblCount7_Click()
'lblCount7
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
'            CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)
'            LV.SubItems(2) = uWIN32.nFileSizeLow
'            LV.SubItems(3) = FormatFileDate(uWIN32.ftLastAccessTime)
        
            Set F = FS.getfile(Path + DFilename)
            LV.SubItems(2) = F.Size
            LV.SubItems(3) = Format(F.datelastmodified, "YYYY-MM-DD HH:MM:SS")


        
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


Private Sub SP_FileMatch(Filename As String, DFilename As String, Path As String)
    Dim LV As ListItem
    Dim uWIN32 As WIN32_FIND_DATA
    
    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
    On Local Error GoTo GetFileError
    
    With ListView1
        Set LV = .ListItems.Add(, , Filename)
        LV.SubItems(1) = Path
        'mod by matt l
        LV.SubItems(2) = DFilename
        
        
        
        
        
        
        
        
        
        
        If chkCopyMemory.Value Then
            'VB does not allows UDT's to be passed directly from a Class but we can access the structure
            'using pointers. We use CopyMemory to place the WIN32_FIND_DATA variable from the Class into
            'the uWIN32 declared in this Sub.
            
            'This is primarily a speed trick - we could retrieve the date & time information "ourselves"
            'using the standard VB functions or an API call but since we already have it...
'            CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)
'            LV.SubItems(2) = uWIN32.nFileSizeLow
'            LV.SubItems(3) = FormatFileDate(uWIN32.ftLastWriteTime)
        
            Set F = FS.getfile(Path + DFilename)
            LV.SubItems(2) = F.Size
            LV.SubItems(3) = Format(F.datelastmodified, "YYYY-MM-DD HH:MM:SS")

        
        End If
        
        'Refresh Count & Scroll the list every 10th file
        If (LV.Index Mod 10) = 0 Then
            If chkRefreshListView.Value Then
                .SelectedItem = LV
                .SelectedItem.EnsureVisible
            End If
            
            'lblCount.Caption = LV.Index + " Total Scanned"
        End If
    End With
    Exit Sub

GetFileError:
    MsgBox Err.Description, vbCritical
    'Resume
End Sub



Private Function FormatFileDate(CT As FILETIME) As String
    Const SHORT_DATE = "Short Date"
    
    Dim ST As SYSTEMTIME
    Dim ds1 As Single
    Dim ds2 As Single
    'ct.dwHighDateTime
    If FileTimeToSystemTime(CT, ST) Then
          ds1 = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
          ds2 = TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
          
          FormatFileDate = Format$(ds1 + ds2, "YYYY-MM-DD HH-MM-SS")
    End If
End Function

Private Sub Timer1_Timer()
'End
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


