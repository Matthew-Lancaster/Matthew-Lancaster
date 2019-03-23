VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ScanPath 
   Caption         =   "ScanPath 2.0 - Anything -- "
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   12975
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   12975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdScanDir 
      Caption         =   "Scan DIR"
      Default         =   -1  'True
      Height          =   880
      Left            =   11328
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3488
      Width           =   810
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   585
      Top             =   5190
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Height          =   825
      Left            =   11312
      Picture         =   "ScanPathDemo.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5312
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   2115
      Left            =   2055
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Text            =   "ScanPathDemo.frx":1A5E
      Top             =   4575
      Visible         =   0   'False
      Width           =   7245
   End
   Begin VB.ComboBox cboMask 
      Height          =   336
      Left            =   495
      TabIndex        =   1
      Text            =   "*.jpg"
      Top             =   930
      Width           =   4928
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   660
      Left            =   8820
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   30
      Width           =   600
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   510
      TabIndex        =   0
      Text            =   "C:\Documents and Settings\Matt4\Local Settings\Temporary Internet Files\Content.IE5\"
      Top             =   30
      Width           =   8265
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "Search Sub-Folders"
      Height          =   195
      Left            =   1530
      TabIndex        =   8
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkPatternMatching 
      Caption         =   "Pattern Matching"
      Height          =   195
      Left            =   1530
      TabIndex        =   7
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkNormal 
      Caption         =   "Normal"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   3225
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   3450
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   5865
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2700
      Width           =   1545
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   0
      Left            =   9930
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3090
      Width           =   1365
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   8580
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2700
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   8580
      TabIndex        =   16
      Top             =   3090
      Width           =   1305
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   1
      Left            =   9930
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3450
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   8580
      TabIndex        =   18
      Top             =   3450
      Width           =   1305
   End
   Begin VB.CheckBox chkRefreshListView 
      Caption         =   "Refresh ListView (slower)"
      Height          =   195
      Left            =   1530
      TabIndex        =   10
      Top             =   3424
      Width           =   2205
   End
   Begin VB.CheckBox chkCopyMemory 
      Caption         =   "Use CopyMemory (display Date && Size)"
      Height          =   195
      Left            =   1530
      TabIndex        =   11
      Top             =   3632
      Width           =   3465
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3510
      Left            =   75
      TabIndex        =   21
      Top             =   3840
      Width           =   11265
      _ExtentX        =   19844
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
      Height          =   368
      Left            =   1530
      TabIndex        =   9
      Top             =   3008
      Width           =   2535
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   825
      Left            =   11340
      Picture         =   "ScanPathDemo.frx":1B96
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2624
      Width           =   810
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   0
      Left            =   5865
      TabIndex        =   13
      Top             =   3090
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   57737217
      CurrentDate     =   37299
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   1
      Left            =   5865
      TabIndex        =   14
      Top             =   3450
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   57737217
      CurrentDate     =   37296
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "DONT CHAIN RUN NX PROGRAM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3540
      TabIndex        =   47
      Top             =   1290
      Width           =   5910
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      Caption         =   "Label6"
      Height          =   208
      Left            =   5616
      TabIndex        =   46
      Top             =   896
      Visible         =   0   'False
      Width           =   256
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dont Do Net Drives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6015
      TabIndex        =   45
      Top             =   870
      Width           =   3420
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   11340
      TabIndex        =   43
      Top             =   4545
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Working"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   512
      Left            =   16
      TabIndex        =   42
      Top             =   1744
      Width           =   9344
   End
   Begin VB.Label lblCount2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   432
      Left            =   9488
      TabIndex        =   41
      Top             =   480
      Width           =   6736
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   416
      Left            =   9488
      TabIndex        =   40
      Top             =   2176
      Width           =   6736
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   416
      Left            =   9488
      TabIndex        =   39
      Top             =   1760
      Width           =   6736
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   416
      Left            =   9488
      TabIndex        =   38
      Top             =   1328
      Width           =   6736
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   416
      Left            =   9488
      TabIndex        =   37
      Top             =   912
      Width           =   6736
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
      Left            =   5055
      TabIndex        =   34
      Top             =   2280
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
      TabIndex        =   26
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mask"
      Height          =   195
      Left            =   45
      TabIndex        =   24
      Top             =   885
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      Height          =   195
      Left            =   45
      TabIndex        =   22
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
      TabIndex        =   25
      Top             =   2280
      Width           =   885
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      Height          =   195
      Index           =   0
      Left            =   7770
      TabIndex        =   31
      Top             =   3150
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
      Height          =   195
      Left            =   7770
      TabIndex        =   29
      Top             =   2730
      Width           =   585
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   195
      Index           =   1
      Left            =   7770
      TabIndex        =   33
      Top             =   3510
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   5055
      TabIndex        =   30
      Top             =   3150
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "File Date"
      Height          =   195
      Left            =   5055
      TabIndex        =   28
      Top             =   2730
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   5055
      TabIndex        =   32
      Top             =   3450
      Width           =   195
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   416
      Left            =   9488
      TabIndex        =   27
      Top             =   64
      Width           =   6736
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------Is This Sub Get on It at start      at       Bangers
Dim TTO()
Public Dog, RR
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Type WIN32_FIND_DATA
    dwFileAttributes  As Long
    ftCreationTime    As FILETIME
    ftLastAccessTime  As FILETIME
    ftLastWriteTime   As FILETIME
    nFileSizeHigh     As Long
    nFileSizeLow      As Long
    dwReserved0       As Long
    dwReserved1       As Long
    cFileName         As String * MAX_PATH
    cAlternate        As String * 14
End Type
Private mWFD As WIN32_FIND_DATA

'Put This in a Bas Mobule
Public m_CRC As clsCRC

Public WxHex$

Public A1$, B1$, FF$

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







Sub Bangers()

ReDim TTO(50)


TTO(1) = "*.mp3"
TTO(2) = "*.mid"
TTO(3) = "*.midi"
TTO(4) = "*.wav"
TTO(5) = "*.wma"
TTO(6) = "*.mpg"
TTO(7) = "*.mpeg"
TTO(8) = "*.avi"
TTO(9) = "*.wmv"
TTO(10) = "*.rm"
TTO(11) = "*.asf"
TTO(12) = "*.qt"
TTO(13) = "*.rmi"
TTO(14) = "*.flv"

ReDim Preserve TTO(14)


Label17 = "Process Banger Favs Quick"
'Call Banger_Favs
Label17 = "Working"

Set fs = CreateObject("Scripting.FileSystemObject")

cboMask.Text = "*.mp3"
chkSubFolder = vbChecked

Dim ra(30)

DA = 0
DA = DA + 1: ra(DA) = "D:\0 00 Music ---"
DA = DA + 1: ra(DA) = "M:\0 00 Music ---"
DA = DA + 1: ra(DA) = "U:\0 00 Music ---"
'DA = DA + 1: ra(DA) = "M:\04 Music ---\04 My Music"
'DA = DA + 1: ra(DA) = "M:\04 Music ---\05 Whole Lot"
'DA = DA + 1: ra(DA) = "M:\00 HTTRACK\Techno BeatPlexity.com MusicGateWay\http___www.beatplexity.com_musicgateway\www.beatplexity.com\musicgateway"
'DA = DA + 1: ra(DA) = "M:\00 HTTRACK\Techno PoMac.Net Swarm\Http___pomac.netswarm.net_music.html_music\pomac.netswarm.net\music"


'DA = DA + 1: ra(DA) = "M:\Videos"


'DA = DA + 1: ra(DA) = "M:\Videos\00 Vid Films"
'DA = DA + 1: ra(DA) = "M:\Videos\00 Vid XXX"
'DA = DA + 1: ra(DA) = "M:\04 Music ---\07 Midi Mod"
'DA = DA + 1: ra(DA) = "M:\04 Music ---\08 Tbooks All"
'DA = DA + 1: ra(DA) = "M:\Videos\00 Vid Films\MP3's"


'DA = DA + 1: ra(DA) = "\\55-88-happy\D\04 Music ---"


For ret = 1 To DA
    DoEvents
    If Right$(ra(ret), 1) <> "\" Then ra(ret) = ra(ret) + "\"
    If fs.FolderExists(ra(ret)) = False Then
        MsgBox "FOLDER " + ra(ret) + " DONT EXIST"
    End If
Next


For ret = 1 To DA
If IsIDE = True Then Me.SetFocus

lblCount2 = "-"
Label13 = "-"
Label14 = "-"
Label15 = "-"
Label16 = "-"

If IsIDE = True Then Me.SetFocus
'If IsIDE = True Then Label5.BackColor = Label6.BackColor

If Mid(ra(ret), 1, 2) = "\\" And Label5.BackColor = Label6.BackColor Then
    ret = ret + 1
    If ret > DA Then Exit For
End If


Label17 = "Working --" + Str(ret) + " OF" + Str(DA)

DoEvents
If Right$(ra(ret), 1) <> "\" Then ra(ret) = ra(ret) + "\"

txtPath.Text = ra(ret)



'FIRST FIND ALL ZERO SIZE MP3 FILES
cboSize.ListIndex = 1 'Less than
txtSize(0) = 1

If Dog = True Then Exit Sub

Label17 = Str(ret) + " OF" + Str(DA) + " SCAN ZERO SIZES"

Call cmdScan_Click

If Dog = True Then Exit Sub
DoEvents
If Dog = True Then Exit Sub

'Exit Sub



DX = ListView1.ListItems.Count
Label16 = Str(ListView1.ListItems.Count) + " Items In List"
For WE = ListView1.ListItems.Count To 1 Step -1
    
    Label13 = Str(WE) + " OF" + Str(DX) + " DO ZERO SIZES"

    DoEvents
    
    G1$ = ListView1.ListItems.Item(WE).SubItems(2)
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    Err.Clear
    If G1 = "" Then G1 = B1
    On Error Resume Next
    w = DeleteFile(A1$ + G1$)
    
    'Debug.Print A1 + G1
    If Err.Number > 0 Then MsgBox "Error Cant Delete File you wanted to Know Been Bug With This"
    On Error GoTo 0

Next

'Exit Sub

'RESET THIS TO NORM
cboSize.ListIndex = 0 ' DONT USE SIZE FILTER
txtSize(0) = 0

lblCount2 = "-"
Label13 = "-"
Label14 = "-"
Label15 = "-"
Label16 = "-"

Label17 = Str(ret) + " OF" + Str(DA) + " SCAN DIR'S"

'DIR SCAN
Call cmdScanDir_Click
'Exit Sub



DX = ListView1.ListItems.Count
Label16 = Str(ListView1.ListItems.Count) + " Items In List"

lblCount2 = "-"
Label13 = "-"
Label14 = "-"
Label15 = "-"
Label16 = "-"


'Exit Sub

For WE = ListView1.ListItems.Count To 1 Step -1
  
    Label13 = Str(WE) + " OF" + Str(DX) + " Remove Zero DIR's in List"
    DoEvents
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    A1$ = A1$ + B1$ + "\"
  
    
    
    
    'Err.Clear
    'On Error Resume Next
    Set F = fs.GetFolder(A1$)
    
       
    
    
    'RR = Get_All_Directory_Files_With_Wildcard_1ST_MP3_BIGGER_ZERO_SIZE(A1$, True, "*.MP3")
    
    RR = False
    
    'If (InStr(A1$, " - 20") > 0 And InStr(LCase(A1$), "banging") > 0) = False Then
    '    For RST = 1 To UBound(TTO)
    '        DoEvents
    '        If RR = False Then RR = Get_All_Directory_Files_With_Wildcard_1ST_MP3_BIGGER_ZERO_SIZE(A1$, True, TTO(RST))
    '        If RR = True Then Exit For
    '    Next
    'End If
    
    If F.Size > 0 And Err.Number = 0 Then RR = True
    
    If InStr(A1$, " - 20") > 0 And InStr(LCase(A1$), "banging") > 0 And RR = False Then
        RR = True
        'If fs.FileExists(A1$ + "\*.*") = True Or fs.FolderExists(A1$ + "\*") = True Then
        '    RR = False
        'End If
    
    End If
    
    
    
    If RR = False Then
        
        egg = egg + 1
        Label14 = Str(egg) + " Deleted In List"
        
        ListView1.ListItems.Remove (WE)
            
        Label16 = Str(ListView1.ListItems.Count) + " Items In List"
        End If
    
Next
    
On Error Resume Next

A1$ = txtPath.Text

If fs.FolderExists(A1$) = True Then

Call Tagg2

DX = ListView1.ListItems.Count
Label16 = Str(ListView1.ListItems.Count) + " Items In List"
For WE = 1 To ListView1.ListItems.Count
  
    Label13 = Str(WE) + " OF" + Str(DX) + " TaggBang and Tagg"
    DoEvents
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    
    A1$ = A1$ + B1$ + "\"
    
    If InStr(A1$, " - 20") > 0 And InStr(LCase(A1$), "banging") > 0 Then
        Call TaggBang
    Else
        Call Tagg2
    End If

Next

End If 'folderexists

Next

'End


Label17 = " *********** Finished ** DONE ********** "
Label17.BackColor = QBColor(2)
Label17.ForeColor = RGB(255, 255, 255)

ListView1.BackColor = QBColor(14) '6 or 14


'Exit Sub


If Me.WindowState = 1 Then aw = " Quite"




If Label7.BackColor <> Label6.BackColor Then
    If IsIDE = False Then
        Shell "D:\VB6\VB-NT\00 MP3 Sound\MP3 Winamp PlayList Generator ID3v2\MP3 Winamp PlayList Generator ID3V2.exe" + aw, vbNormalFocus
        End
    End If

    'aw = " Quite"
    If IsIDE = True Then
        Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit ""D:\VB6\VB-NT\00 MP3 Sound\MP3 Winamp PlayList Generator ID3v2\MP3 Winamp PlayList Generator ID3V2.vbp""", vbMinimizedNoFocus
        End
    End If
End If


End Sub



Sub Tagg2()
 
FF$ = Mid$(A1$, Len(txtPath.Text) + 1)

If FF$ = "" Then
    FF$ = Mid$(A1$, 4)
End If

FF$ = Left$(FF$, Len(FF$) - 1)

Do
    ttgx = InStr(FF$, "\")
    If ttgx > 0 Then
        FF$ = Mid$(FF$, 1, ttgx - 1) + " -- " + Mid$(FF$, ttgx + 1)
    End If
Loop Until ttgx = 0

Call WriteTagg

End Sub

Sub TaggBang()
 
    
ty$ = ""
If InStr(A1$, "My Music\") Then ty$ = "My Music"
If InStr(A1$, "My Music Zen\") Then ty$ = "My Music Zen"
If ty$ = "My Music" And InStr(A1$, "Favs") Then ty$ = "My Music - Favs"
If ty$ = "My Music Zen" And InStr(A1$, "Favs") Then ty$ = "My Music Zen - Favs"
If ty$ = "My Music" And InStr(A1$, " Bang Sets") Then ty$ = "My Music - Sets"
If ty$ = "My Music Zen" And InStr(A1$, " Bang Sets") Then ty$ = "My Music Zen - Sets"
If ty$ = "My Music" And InStr(A1$, " Bang Fav Sets") Then ty$ = "My Music - Fav Sets"
If ty$ = "My Music Zen" And InStr(A1$, " Bang Fav Sets") Then ty$ = "My Music Zen - Fav Sets"
'03 Bang Sets
    
    
gg$ = Mid$(A1$, InStrRev(A1$, "\", Len(A1$) - 1) + 1)
tx$ = Left$(gg$, Len(gg$) - 1)

FF$ = "---------------[[(({{ " + tx$ + " - " + ty$ + " }}))]]---------------.mp3"


On Error Resume Next
Do
fg1 = FreeFile
Err.Clear
Open A1$ + FF$ For Output As fg1
Close fg1
If Err.Number > 0 Then FF$ = Left$(FF$, Len(FF$) - 5) + ".mp3"
Loop Until Err.Number = 0


'Call WriteTagg



End Sub



Sub WriteTagg()

On Local Error Resume Next
'Kill A1$ + "0#-" + FF$ + "-^.mp3"
'Kill A1$ + "01#-" + FF$ + "-^.mp3"
'Err.Clear
easy = Dir$(A1$ + "\*.txt") <> ""
 
 dtest = FF$
 
 '
Do
    DoEvents
    Err.Clear

    'Debug.Print A1$
    fg1 = FreeFile
    Open A1$ + "0#01--" + FF$ + "-^.mp3" For Output As fg1
    Close fg1
    
    'ENCRYPTION PROBLEM ACCESS
    If Err.Description = "Path/File access error" Then
        Exit Sub
    End If
    
    If Err.Number > 0 Then FF$ = Left$(FF$, Len(FF$) - 1)
Loop Until Err.Number = 0 Or FF$ = ""

'TEST TRAP UNUSUAL ERRORS
If FF$ = "" Then Debug.Print dtest: Stop: Exit Sub
    

Do
    Err.Clear
    ek$ = String$(Len(FF$) + 20, "-")
    fg1 = FreeFile
    Open A1$ + "0#00--" + ek$ + "-^.mp3" For Output As fg1
    Close fg1
    If easy = True Then
        fg1 = FreeFile
        Open A1$ + "0#02----- Text File In Dir -----^.mp3" For Output As fg1
        Close fg1
    
    End If
    fg1 = FreeFile
    Open A1$ + "0#02--" + ek$ + "-^.mp3" For Output As fg1
    Close fg1
    FF$ = Left$(FF$, Len(FF$) - 1)
Loop Until Err.Number = 0
    
    
On Local Error GoTo 0
    


End Sub


Sub AutoPix()

Dim fs22
Set fs22 = CreateObject("Scripting.FileSystemObject")
'Fs22.moveFile a1$ + b1$, c1$ + b1$
'    Fs22.moveFile a1$ + b1$, tp1$ + b1$


'Put in Sub Of Code
Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32
'WxHex$ = Hex(m_CRC.CalculateString(Form3.Text1.Text))
'WxHex$ = Hex(m_CRC.CalculateFile(Filename))

'tp1$ = "G:\Art\AutoPix Dont Want CRC - Done\"

'txtPath.Text = "G:\Art\AutoPix Dont Want CRC\"
'If Mid$(txtPath.Text, Len(txtPath.Text), 1) <> "\" Then
'    txtPath.Text = txtPath.Text + "\"
'End If

cboMask.Text = "*.mp3"

Call cmdScan_Click

For WE = 1 To ListView1.ListItems.Count
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
'    WxHex$ = Space$(10)
'    LSet WxHex$ = Hex(m_CRC.CalculateFile(a1$ + b1$))
'    Print #f1, WxHex$
Err.Clear
On Local Error Resume Next
'    tp2$ = tp1$ + Mid$(a1$, Len(txtPath.Text) + 1)
'    Fs22.moveFile a1$ + b1$, tp1$ + b1$

If Err.Number > 0 Then
'    tp2$ = tp1$ + Mid$(a1$, Len(txtPath.Text) + 1)
'    ff$ = Err.Description
'    Err.Clear
'    MkDir (tp2$)
    'If Err.Number > 0 Then Stop
'    Err.Clear
    
'    Fs22.moveFile a1$ + b1$, tp2$ + Left$(b1$, Len(b1$) - 4) + "_" + Trim(Str(we)) + ".jpg"
'    If Err.Number > 0 Then Stop

End If
Next
Close #f1

f1 = FreeFile
Open txtPath.Text + "CRC File.txt" For Input As #f1
RR = Input$(LOF(f1), f1)
Close #f1

For r3 = 1 To 1

If r3 = 2 Then txtPath.Text = "G:\Art\AutoPix\AutoPix 000-000\"
If r3 = 1 Then txtPath.Text = "G:\Art\AutoPix\"
'If r3 = 1 Then txtPath.Text = "G:\Art\AutoPix\AutoPix zz 000-000\"

If Mid$(txtPath.Text, Len(txtPath.Text), 1) <> "\" Then
    txtPath.Text = txtPath.Text + "\"
End If
cboMask.Text = "*.jpg"

Call cmdScan_Click





'MsgBox "Ready to Go"

'Open App.Path + "\CRC Dupe.txt" For Output As #1


For WE = 1 To ListView1.ListItems.Count
'For we = 1 To 5000
    
    
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)


    ii = ii + 1
'    List1.AddItem Str$(ii) + " -- " + a1$ + b1$
 '   List1.ListIndex = List1.ListCount - 1
    
    WxHex$ = Space$(8)

    LSet WxHex$ = Hex(m_CRC.CalculateFile(A1$ + B1$))
    
    ListView1.ListItems.Item(WE).SubItems(1) = WxHex$ + " - - " + Format$(WE, "0000000") + " - - " + A1$
    ListView1.SelectedItem = ListView1.ListItems(WE)
    ListView1.SelectedItem.EnsureVisible
    Label13.Caption = Str(WE)
    Label14.Caption = Str(WE)

 '   Print #1, WxHex$, " - - "; a1$, b1$
DoEvents



Next
    
    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True        'Use default sorting to sort the
    ListView1.SortKey = 1

For WE = 2 To ListView1.ListItems.Count
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    
    a12$ = ListView1.ListItems.Item(WE - 1).SubItems(1)
    b12$ = ListView1.ListItems.Item(WE - 1)




If Mid$(A1$, 1, 3) <> "G:\" Then
    
    G1$ = Mid$(A1$, 1, 8)
    g2 = Val(Mid$(A1$, 14, 7))
    g3$ = Mid$(A1$, 26)
    
    h1$ = Mid$(a12$, 1, 8)
    h2 = Val(Mid$(a12$, 14, 7))
    h3$ = Mid$(a12$, 26)

rt = 0
If InStr(RR, G1$) Then
    If rt = 0 Then MsgBox "1st Del CRC Dont Want"
    rt = 1
    
    Kill g3$ + B1$
    aga = aga + 1
    Label14.Caption = Str(aga)
End If


'Shell "C:\Program Files\IceView\IceView.exe " + h3$ + b12$, vbNormalFocus
'Shell "C:\Program Files\IceView\IceView.exe " + g3$ + b1$, vbNormalFocus
    
    If G1$ = h1$ Then
        aga = aga + 1
        Label14.Caption = Str(aga)
        If g2 > h2 Then xxcrc = 1
        If g2 < h2 Then xxcrc = 2
        If xxcrc = 1 Then Kill g3$ + B1$
        If xxcrc = 2 Then Kill h3$ + b12$
    End If
    
    End If
    
    ListView1.SelectedItem = ListView1.ListItems(WE)
    ListView1.SelectedItem.EnsureVisible
    Label13.Caption = Str(WE)

Next



Next




MsgBox "End"

Exit Sub

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

Private Sub cmdScan_Click()
    Dim lCount As Long
    Dim lSize(1) As Long
    
    ListView1.Sorted = False        'Use default sorting to sort the
    
    
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
 
        'Screen.MousePointer = vbHourglass
        ListView1.ListItems.Clear
        
        'If chkRefreshListView.Value = vbUnchecked Then
        '    LockWindowUpdate ListView1.hWnd
        'End If
        
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
            
            'lblCount.Caption = .ListItems.Count
        End With
        
        'LockWindowUpdate 0
        
        'cmdScan.Caption = "Scan"
        'Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
        'End
    End If

'frmMain.Refresh

'End
'If Command$ <> "" Then End
cmdScan.Caption = "Scan"

    
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 1
    ListView1.Sorted = True        'Use default sorting to sort the
    ListView1.Refresh
    DoEvents
    ListView1.Sorted = False

    'SORT ON PATH LAST
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 2
    ListView1.Sorted = True        'Use default sorting to sort the
    ListView1.Refresh
    DoEvents
    ListView1.Sorted = False



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
            .StartScanDir txtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
        End With
        
       With ListView1
            If .ListItems.Count > 0 Then
                .SelectedItem = .ListItems(1)
            End If
            
            'lblCount.Caption = .ListItems.Count
        End With
        
        LockWindowUpdate 0
        
        cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
    End If



'lblCount.Caption = ListView1.ListItems.Count
    

End Sub




Private Sub Form_Load()
    
ScanPath.Caption = ScanPath.Caption + App.EXEName
'scanpath.Caption = frmMain.Caption + App.EXEName
  
'ScanPath.Height = ListView1.Top + ListView1.Height + 450
'ScanPath.Width = ListView1.Width + 210
    
'Code to Auto Size Form based on controls used Including Menu Bar Sketchy Style
'Make sure form set to scale Twips
'put in your form load

ScanPath.Show
'DoEvents
If Command$ <> "" Then Me.WindowState = 1
    
    
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


Call Bangers


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



Private Sub Form_Resize()

x = 1
y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next
'On Error GoTo 0

Me.Width = (x + 120)
If mnu = 1 Then
    pluso = 720: pluso = 1100 'Sometimes Different
Else
    pluso = 450
End If

Me.Height = (y + pluso)
Me.Refresh
DoEvents

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_Click()

Exit Sub

Dim fs22
Set fs22 = CreateObject("Scripting.FileSystemObject")


tpath3$ = "M:\04 Music ---\Del\"

B1$ = Mid$(List1.List(List1.ListIndex), 14)
c1$ = B1$ + "--Todel"

'tpath1$ = "M:\04 Music ---\"

c1$ = Mid$(B1$, InStrRev(B1$, "\") + 1)
List1.RemoveItem (List1.ListIndex)
'Name b1$ As c1$

'fs22.copyFile b1$, tpath3$ + c1$
If Dir$(tpath3$ + c1$) <> "" Then Kill tpath3$ + c1$
fs22.MoveFile B1$, tpath3$ + c1$



End Sub

Private Sub Label18_Click()
Dog = True

If cmdScan.Caption = "Stop" Then cmdScan_Click
Label17 = "Stopped before Complete"
End Sub

Private Sub Label5_Click()
Label5.BackColor = Label6.BackColor

End Sub

Private Sub Label7_Click()
Label7.BackColor = Label6.BackColor

End Sub

Private Sub MNU_VB_Click()
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
End
End Sub

Private Sub SP_DirMatch(Directory As String, Path As String)
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
            DoEvents
            If chkRefreshListView.Value Then
                .SelectedItem = LV
                .SelectedItem.EnsureVisible
            End If
            
            lblCount.Caption = Str(LV.Index) + " Folder's"
        End If
    End With

lblCount.Caption = Str(LV.Index) + " Folder's"
    
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
            
            lblCount.Caption = Str(LV.Index) + " File's"
        End If
    End With

lblCount.Caption = Str(LV.Index) + " File's"
    
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

For WE = 1 To ListView1.ListItems.Count
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

    'Dim Fs22
    'Set Fs22 = CreateObject("Scripting.FileSystemObject")
    errs2 = 0
    On Local Error GoTo jeep
    fs22.MoveFile A1$ + B1$, c1$ + B1$
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
Set fs22 = CreateObject("Scripting.FileSystemObject")
fs22.DeleteFolder txtPath.Text, True
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
fs22.MoveFolder Drived2$ + "Temp\Anything\*", v2$
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
            fs22.DeleteFolder d2$, True
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

    List1.AddItem Format$(WE, "000 ") + "Move > " + A1$ + B1$
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
    fs22.MoveFile c2$ + B1$, A1$ + B1$
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
    Set fs22 = CreateObject("Scripting.FileSystemObject")
    fs22.CopyFolder Drived2$ + "Temp\Anything", v1$, True
    
    Set fs22 = CreateObject("Scripting.FileSystemObject")
    fs22.DeleteFolder Drived2$ + "Temp\Anything", True


jeep:
End Sub

