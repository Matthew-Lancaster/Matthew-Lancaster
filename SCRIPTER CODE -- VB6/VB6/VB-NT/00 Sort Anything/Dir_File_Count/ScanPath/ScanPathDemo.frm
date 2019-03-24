VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ScanPath 
   AutoRedraw      =   -1  'True
   Caption         =   "ScanPath 2.0"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12630
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   12630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Com1 
      Caption         =   "ReScan"
      Height          =   870
      Left            =   9945
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   720
      Width           =   705
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Default         =   -1  'True
      Height          =   885
      Left            =   10695
      Picture         =   "ScanPathDemo.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   705
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   585
      Top             =   4815
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   30
      TabIndex        =   41
      Top             =   6210
      Width           =   11295
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Height          =   390
      Left            =   11850
      Picture         =   "ScanPathDemo.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   855
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   2115
      Left            =   11475
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      Text            =   "ScanPathDemo.frx":1EA0
      Top             =   3810
      Visible         =   0   'False
      Width           =   7245
   End
   Begin VB.ComboBox cboMask 
      Height          =   315
      Left            =   540
      TabIndex        =   1
      Text            =   "*.MP3"
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
      Left            =   540
      TabIndex        =   0
      Text            =   "E:\My Music Zen"
      Top             =   750
      Width           =   5505
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "Search Sub-Folders"
      Height          =   195
      Left            =   1530
      TabIndex        =   8
      Top             =   2340
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkPatternMatching 
      Caption         =   "Pattern Matching"
      Height          =   195
      Left            =   1530
      TabIndex        =   7
      Top             =   2100
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkNormal 
      Caption         =   "Normal"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   2100
      Width           =   1365
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   2340
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   2580
      Width           =   1365
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   2805
      Width           =   1365
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   3030
      Width           =   1365
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   5940
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2280
      Width           =   1545
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   0
      Left            =   9930
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2670
      Width           =   1365
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   8580
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2280
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   8580
      TabIndex        =   16
      Top             =   2670
      Width           =   1305
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   1
      Left            =   9930
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3030
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   8580
      TabIndex        =   18
      Top             =   3030
      Width           =   1305
   End
   Begin VB.CheckBox chkRefreshListView 
      Caption         =   "Refresh ListView (slower)"
      Height          =   195
      Left            =   1530
      TabIndex        =   10
      Top             =   2880
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   12570
      TabIndex        =   22
      Top             =   0
      Width           =   12630
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
         Left            =   1260
         TabIndex        =   24
         Top             =   60
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
      Top             =   3120
      Value           =   1  'Checked
      Width           =   3465
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2685
      Left            =   30
      TabIndex        =   21
      Top             =   3480
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4736
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CheckBox chkSort 
      Caption         =   "Sort Files(A-Z) while Scanning"
      Height          =   195
      Left            =   1530
      TabIndex        =   9
      Top             =   2580
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   735
      Left            =   11775
      Picture         =   "ScanPathDemo.frx":1FD8
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1410
      Visible         =   0   'False
      Width           =   675
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   0
      Left            =   5940
      TabIndex        =   13
      Top             =   2670
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   20709377
      CurrentDate     =   39358
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   1
      Left            =   5925
      TabIndex        =   14
      Top             =   3030
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   20709377
      CurrentDate     =   39358
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "After Filters"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8505
      TabIndex        =   50
      Top             =   1005
      Width           =   1410
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Processing Now"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8505
      TabIndex        =   49
      Top             =   1305
      Width           =   1410
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Scanned"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8505
      TabIndex        =   48
      Top             =   705
      Width           =   1410
   End
   Begin VB.Label LblCount2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6705
      TabIndex        =   47
      Top             =   705
      Width           =   1770
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6705
      TabIndex        =   46
      Top             =   1305
      Width           =   1770
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modified in Last 1 day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5130
      TabIndex        =   45
      Top             =   1710
      Width           =   2100
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   240
      Left            =   11655
      TabIndex        =   43
      Top             =   2430
      Width           =   870
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
      Top             =   2010
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
      Top             =   1860
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
      Top             =   1860
      Width           =   885
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      Height          =   195
      Index           =   0
      Left            =   7770
      TabIndex        =   35
      Top             =   2730
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
      Height          =   195
      Left            =   7770
      TabIndex        =   33
      Top             =   2310
      Width           =   585
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   195
      Index           =   1
      Left            =   7770
      TabIndex        =   37
      Top             =   3090
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   5130
      TabIndex        =   34
      Top             =   2730
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "File Date"
      Height          =   195
      Left            =   5130
      TabIndex        =   32
      Top             =   2310
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   5130
      TabIndex        =   36
      Top             =   3030
      Width           =   195
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6705
      TabIndex        =   31
      Top             =   1005
      Width           =   1770
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Jeepers


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


Sub Jeepers()

List1.Clear

ScanPath.Show

Dim outInfo As MPEGInfo

'ScanPath.List1.AddItem "Converting MP3 Tags..."
    ScanPath.List1.Refresh
    DoEvents

'On Local Error Resume Next
'f1 = FreeFile
'Open App.Path + "CRC Such.txt" For Input As #f1
'wdm$ = Input(LOF(f1), f1)
'Close #f1
'On Local Error GoTo 0


Dim InTag As ID3Tag
Dim DummyInTag As ID3Tag

Dim f
Dim Fs22
Set Fs22 = CreateObject("Scripting.FileSystemObject")
    
Dim ra(20)
'ra(4) = ""
ra(1) = "E:\04 Music ---\04 My Music\01 Banging Tunes\"
'ra(2) = "E:\04 Music ---\04 My Music\01 Banging Tunes\"


'ra(1) = "E:\04 Music ---\04 My Music\02 Chart Songs\00 IDTag-Able\"
ra(2) = "E:\04 Music ---\04 My Music\Audio Recordings\"
ra(3) = "E:\04 Music ---\04 My Music\Radio\"
ra(4) = "E:\04 Music ---\04 My Music\Sterns - Fantazia\"
ra(5) = "E:\04 Music ---\04 My Music\02 Chart Songs\00 IDTag-Able\"
ra(6) = "E:\04 Music ---\02 My Music Zen Removed\"
ra(7) = "E:\04 Music ---\03 My Music Zen\"
ra(8) = "E:\04 Music ---\04 My Music Talking Books\"
ra(9) = "E:\04 Music ---\04 My Music\Comedy Mp3's\"
'ra(4) = ""
'ra(4) = ""


'ra(10) = "F:\04 Music ---\04 My Music\01 Banging Tunes\"
'ra(9) = "E:\04 Music ---\04 My Music\02 Chart Songs\00 IDTag-Able\"
'ra(11) = "F:\04 Music ---\04 My Music\Audio Recordings\"
'ra(12) = "F:\04 Music ---\04 My Music\Radio\"
'ra(13) = "F:\04 Music ---\04 My Music\Sterns - Fantazia\"
'ra(14) = "F:\04 Music ---\04 My Music\02 Chart Songs\00 IDTag-Able\"
'ra(15) = "F:\04 Music ---\02 My Music Zen Removed\"
'ra(16) = "F:\04 Music ---\03 My Music Zen\"
'ra(17) = "F:\04 Music ---\04 My Music Talking Books\"
'ra(18) = "E:\04 Music ---\04 My Music\Comedy Mp3's\"


'ra(2) = "E:\04 Music ---\04 My Music\Audio Recordings\"

'For rtu = 1 To 18
'ra(1) = "E:\04 Music ---\04 My Music\01 Banging Tunes\"
'ra(2) = "F:\04 Music ---\04 My Music\01 Banging Tunes\"

ra(1) = "C:\00 Zen\"
ra(1) = "F:\04 Music ---\04 My Music Talking Books\_Children\"
ra(1) = "F:\04 Music ---\03 My Music Zen\03 Small'S\"
'ra(2) = "F:\04 Music ---\03 My Music Zen\01 Banging Tunes\"
'ra(1) = "F:\04 Music ---\03 My Music Zen\"
ra(1) = "F:\04 Music ---\04 My Music\"
ra(1) = "F:\04 Music ---\04 My Music Talking Books\"
ra(1) = "F:\04 Music ---\03 My Music Zen\04 Talking\04 My Music Talking Books\"
ra(1) = "F:\04 Music ---\04 My Music\01 Banging Tunes\00 Bang Favs\01 Fil Devious"
ra(1) = "F:\04 Music ---\03 My Music Zen\"
ra(1) = "F:\04 Music ---\03 My Music Zen\03 Small'S\"
ra(1) = "E:\04 Music ---\03 My Music Zen\"
ra(2) = "E:\04 Music ---\04 My Music\"
ra(3) = "E:\04 Music ---\04 My Music\Comedy Mp3'S\Bill Hick'S Mp3\"
ra(1) = "E:\04 Music ---\04 My Music\01 Banging Tunes\"
ra(1) = "E:\04 Music ---\03 My Music Zen\"
ra(1) = "E:\04 Music ---\"
ra(1) = "S:\My Music Talking Books"
ra(2) = "S:\Tbooks\"


For rtu = 1 To 2

If Right$(ra(rtu), 1) <> "\" Then ra(rtu) = ra(rtu) + "\"

List1.AddItem ra(rtu)

ttn$ = ttn$ + ra(rtu) + vbCrLf
'If rtu = 3 Then ScanPath.chkNormal.Value = vbChecked
    
    
ScanPath.txtPath = ra(rtu)
ScanPath.cboMask.Text = "*.mp3"

t0 = 10
Label14 = "Modified in Last " + Str(t0) + " days"
'DTPicker1(0).Value = Now - t0
'DTPicker1(1).Value = Now
'Chk Box ticked when date set
    
ScanPath.chkArchive.Value = vbUnchecked

ScanPath.chkArchive.Value = vbChecked

ScanPath.chkNormal.Value = vbChecked

ScanPath.chkHidden.Value = vbChecked
ScanPath.chkReadOnly.Value = vbChecked
ScanPath.chkSystem.Value = vbChecked

ScanPath.txtSize(0) = "0"


Call ScanPath.cmdScan_Click

List1.AddItem lblCount
ttn$ = ttn$ + lblCount + vbCrLf

      
    
tt = InStr(4, ScanPath.txtPath, "\") + 1
a3$ = Mid$(ScanPath.txtPath, tt)


If chkArchive = vbChecked Then
    For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    
        a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(we)
    
        On Local Error Resume Next
        Set f = Fs22.Getfile(a1$ + B1$)
        On Local Error GoTo 0
    
        If (f.Attributes And FILE_ATTRIBUTE_ARCHIVE) = False Then
'            ScanPath.ListView1.ListItems.Remove (we)
        End If
    
    Next
End If

'c1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1) + ScanPath.ListView1.ListItems.Item(we)
'e1 = Len(a1$) + Len(B1$)
 

'For we = ScanPath.ListView1.ListItems.Count To 1 Step -1

 '   a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
  '  B1$ = ScanPath.ListView1.ListItems.Item(we)
   ' On Local Error Resume Next

'Next

On Local Error GoTo 0

For we = 1 To ScanPath.ListView1.ListItems.Count
    
    we3 = we3 + 1
    
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    'If InStr(a1$, "Mos Clubber's Guide To Ibiza Summer 2001") > 0 Then Stop
    'If B1$ = "Mos Clubber'S Guide To Ibiza Summer 2001 Cd2-05 Stringer.mp3" Then Stop
    'If a1$ = "E:\04 Music ---\04 My Music\02 Chart Songs\00 IDTag-Able\Eddies CDs\Min Of Sound Annual 2003\CD 01\" Then Stop
    'If b1$ = "Min Of Sound Annual 2003 Cd1 - 05 The Theme.mp3" Then Stop
    
    On Local Error Resume Next
    Set f = Fs22.Getfile(a1$ + B1$)
    On Local Error GoTo 0
    
    xcount = xcount + 1
    
    yy$ = Mid$(a1$, Len(ra(rtu)) + 1)
    yy$ = Mid$(yy$, 1, InStr(yy$, "\"))
    If yy$ <> d1$ Then
        If d1$ <> "" Then
            List1.AddItem "File Count " + Format$(xcount) + " - For " + d1$
            ttn$ = ttn$ + "File Count " + Format$(xcount) + " - For " + d1$ + vbCrLf
        End If
        xcount = 0
        d1$ = yy$
    End If
    
    
Next
Next
    
Open App.Path + "\Test.txt" For Output As #1
Print #1, ttn$
Close #1
Shell "notepad " + App.Path + "\Test.txt", vbNormalFocus

End Sub


Public Sub Jeepersx2()
    
List1.Clear

ScanPath.Show

Dim outInfo As MPEGInfo

ScanPath.List1.AddItem "Converting MP3 Tags..."
    ScanPath.List1.Refresh
    DoEvents

'On Local Error Resume Next
'f1 = FreeFile
'Open App.Path + "CRC Such.txt" For Input As #f1
'wdm$ = Input(LOF(f1), f1)
'Close #f1
'On Local Error GoTo 0


Dim InTag As ID3Tag
Dim DummyInTag As ID3Tag

Dim f
Dim Fs22
Set Fs22 = CreateObject("Scripting.FileSystemObject")
    
Dim ra(20)
'ra(4) = ""
ra(1) = "E:\04 Music ---\04 My Music\01 Banging Tunes\"
'ra(2) = "E:\04 Music ---\04 My Music\01 Banging Tunes\"


'ra(1) = "E:\04 Music ---\04 My Music\02 Chart Songs\00 IDTag-Able\"
ra(2) = "E:\04 Music ---\04 My Music\Audio Recordings\"
ra(3) = "E:\04 Music ---\04 My Music\Radio\"
ra(4) = "E:\04 Music ---\04 My Music\Sterns - Fantazia\"
ra(5) = "E:\04 Music ---\04 My Music\02 Chart Songs\00 IDTag-Able\"
ra(6) = "E:\04 Music ---\02 My Music Zen Removed\"
ra(7) = "E:\04 Music ---\03 My Music Zen\"
ra(8) = "E:\04 Music ---\04 My Music Talking Books\"
ra(9) = "E:\04 Music ---\04 My Music\Comedy Mp3's\"
'ra(4) = ""
'ra(4) = ""


'ra(10) = "F:\04 Music ---\04 My Music\01 Banging Tunes\"
'ra(9) = "E:\04 Music ---\04 My Music\02 Chart Songs\00 IDTag-Able\"
'ra(11) = "F:\04 Music ---\04 My Music\Audio Recordings\"
'ra(12) = "F:\04 Music ---\04 My Music\Radio\"
'ra(13) = "F:\04 Music ---\04 My Music\Sterns - Fantazia\"
'ra(14) = "F:\04 Music ---\04 My Music\02 Chart Songs\00 IDTag-Able\"
'ra(15) = "F:\04 Music ---\02 My Music Zen Removed\"
'ra(16) = "F:\04 Music ---\03 My Music Zen\"
'ra(17) = "F:\04 Music ---\04 My Music Talking Books\"
'ra(18) = "E:\04 Music ---\04 My Music\Comedy Mp3's\"


'ra(2) = "E:\04 Music ---\04 My Music\Audio Recordings\"

'For rtu = 1 To 18
'ra(1) = "E:\04 Music ---\04 My Music\01 Banging Tunes\"
'ra(2) = "F:\04 Music ---\04 My Music\01 Banging Tunes\"

ra(1) = "C:\00 Zen\"
ra(1) = "F:\04 Music ---\04 My Music Talking Books\_Children\"
ra(1) = "F:\04 Music ---\03 My Music Zen\03 Small'S\"
'ra(2) = "F:\04 Music ---\03 My Music Zen\01 Banging Tunes\"
'ra(1) = "F:\04 Music ---\03 My Music Zen\"
ra(1) = "F:\04 Music ---\04 My Music\"
ra(1) = "F:\04 Music ---\04 My Music Talking Books\"
ra(1) = "F:\04 Music ---\03 My Music Zen\04 Talking\04 My Music Talking Books\"
ra(1) = "F:\04 Music ---\04 My Music\01 Banging Tunes\00 Bang Favs\01 Fil Devious"
ra(1) = "F:\04 Music ---\03 My Music Zen\"
ra(1) = "F:\04 Music ---\03 My Music Zen\03 Small'S\"
ra(1) = "E:\04 Music ---\03 My Music Zen\"
ra(2) = "E:\04 Music ---\04 My Music\"
ra(3) = "E:\04 Music ---\04 My Music\Comedy Mp3'S\Bill Hick'S Mp3\"
ra(1) = "E:\04 Music ---\04 My Music\01 Banging Tunes\"
ra(1) = "E:\04 Music ---\03 My Music Zen\"
ra(1) = "E:\04 Music ---\"
'ra(1) = "S:\My Music Talking Books"
ra(1) = "S:\Tbooks\"


For rtu = 1 To 1

If Right$(ra(rtu), 1) <> "\" Then ra(rtu) = ra(rtu) + "\"


'If rtu = 3 Then ScanPath.chkNormal.Value = vbChecked
    
    
ScanPath.txtPath = ra(rtu)
ScanPath.cboMask.Text = "*.mp3"

t0 = 10
Label14 = "Modified in Last " + Str(t0) + " days"
'DTPicker1(0).Value = Now - t0
'DTPicker1(1).Value = Now
'Chk Box ticked when date set
    
ScanPath.chkArchive.Value = vbUnchecked

ScanPath.chkArchive.Value = vbChecked

'ScanPath.chkNormal.Value = vbChecked

'ScanPath.chkHidden.Value = vbChecked
'ScanPath.chkReadOnly.Value = vbChecked



Call ScanPath.cmdScan_Click

      
    
tt = InStr(4, ScanPath.txtPath, "\") + 1
a3$ = Mid$(ScanPath.txtPath, tt)


If chkArchive = vbChecked Then
    For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    
        a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(we)
    
        On Local Error Resume Next
        Set f = Fs22.Getfile(a1$ + B1$)
        On Local Error GoTo 0
    
        If (f.Attributes And FILE_ATTRIBUTE_ARCHIVE) = False Then
'            ScanPath.ListView1.ListItems.Remove (we)
        End If
    
    Next
End If

'c1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1) + ScanPath.ListView1.ListItems.Item(we)
'e1 = Len(a1$) + Len(B1$)
 

'For we = ScanPath.ListView1.ListItems.Count To 1 Step -1

 '   a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
  '  B1$ = ScanPath.ListView1.ListItems.Item(we)
   ' On Local Error Resume Next

'Next

On Local Error GoTo 0

For we = 1 To ScanPath.ListView1.ListItems.Count
    
    we3 = we3 + 1
    
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    'If InStr(a1$, "Mos Clubber's Guide To Ibiza Summer 2001") > 0 Then Stop
    'If B1$ = "Mos Clubber'S Guide To Ibiza Summer 2001 Cd2-05 Stringer.mp3" Then Stop
    'If a1$ = "E:\04 Music ---\04 My Music\02 Chart Songs\00 IDTag-Able\Eddies CDs\Min Of Sound Annual 2003\CD 01\" Then Stop
    'If b1$ = "Min Of Sound Annual 2003 Cd1 - 05 The Theme.mp3" Then Stop
    
    On Local Error Resume Next
    Set f = Fs22.Getfile(a1$ + B1$)
    On Local Error GoTo 0
    
    
    
    
    Comment$ = ""
    If ReadID3v1(a1$ + B1$, InTag) = True Then
        Comment$ = "ID3v1 Previous Info " + Format$(Now, "DD-MM-YYYY HH:MM:SS") + vbCrLf
        If InTag.title <> "" Then Comment$ = Comment$ + "Title-----" + InTag.title + vbCrLf
        If InTag.Artist <> "" Then Comment$ = Comment$ + "Artist----" + InTag.Artist + vbCrLf
        If InTag.Album <> "" Then Comment$ = Comment$ + "Album-----" + InTag.Album + vbCrLf
        If InTag.SongYear <> "" Then Comment$ = Comment$ + "SongYear--" + InTag.SongYear + vbCrLf
        If InTag.Genre <> "" Then Comment$ = Comment$ + "Genre-----" + InTag.Genre + vbCrLf
        If InTag.Comment <> "" Then Comment$ = Comment$ + "Comment----" + InTag.Comment + vbCrLf
        If InTag.TrackNr <> "" Then Comment$ = Comment$ + "TrackNr----" + InTag.TrackNr + vbCrLf
        
        
        wa = DeleteID3v1(a1$ + B1$)
    
        Comment$ = "Matty Roid - " + Format$(Now, "DD-MM-YYYY HH:MM:SS")
    
    End If
        
    kiss2 = 0
    wa = ReadID3v2(a1$ + B1$, InTag)
    If wa = False Then kiss2 = 1 ': Stop
    
    tagger$ = InTag.Comment
    
    tea = 0
    If InStr(InTag.Comment, "Roids Rim") = 0 Or InStr(InTag.Comment, "Orig") = 0 Then
        Comment$ = "Roids Rim " + vbCrLf
        Comment$ = Comment$ + "Orig: " + Format$(Now, "DDD DD-MM-YYYY HH:MM:SSp") + vbCrLf
        Comment$ = Comment$ + "Modi: " + Format$(Now, "DDD DD-MM-YYYY HH:MM:SSp") + vbCrLf
        Comment$ = Comment$ + "Size: " + Str(f.Size) + vbCrLf
        tea = 1
    End If
        
    tea2 = 0
    If InStr(InTag.Comment, "Roids Rim") > 0 And tea = 0 Then
        If InStr(InTag.Comment, "Orig") = 0 Then
            Comment$ = "Roids Rim " + vbCrLf
            rt = InStrRev(InTag.Comment, "Orig")
            rt2 = InStr(rt + 1, InTag.Comment, vbCrLf) - rt
            Comment$ = Comment$ + Mid$(InTag.Comment, rt, rt2) + vbCrLf
            Comment$ = Comment$ + "Modi: " + Format$(Now, "DDD DD-MM-YYYY HH:MM:SSp") + vbCrLf
            Comment$ = Comment$ + "Size: " + Str(f.Size) + vbCrLf
            tea2 = 1
            tea = 1
        End If
    End If
    
    If InStr(InTag.Comment, "Roids Rim") > 0 And tea = 0 Then
        Comment$ = "Roids Rim " + vbCrLf
            Comment$ = "Roids Rim " + vbCrLf
            rt = InStrRev(InTag.Comment, "Orig")
            rt2 = InStr(rt + 1, InTag.Comment, vbCrLf) - rt
            Comment$ = Comment$ + Mid$(InTag.Comment, rt, rt2) + vbCrLf
            Comment$ = Comment$ + "Modi: " + Format$(Now, "DDD DD-MM-YYYY HH:MM:SSp") + vbCrLf
            Comment$ = Comment$ + "Size: " + Str(f.Size) + vbCrLf
        
        tea = 1
    End If
    
        
    'cc$ = outInfo.Length
        
    kiss = 0
    If InTag.title <> Mid$(B1$, 1, Len(B1$) - 4) Then
    InTag.title = Mid$(B1$, 1, Len(B1$) - 4)
    kiss = 1
    End If

    mdlID3.ReadMPEGInfo a1$ + B1$, outInfo
        lblMPEG = outInfo.MPEGVersion
        lblLen = "Length: " & outInfo.Length \ 60 & ":" & Format$(outInfo.Length Mod 60, "00")
        lblkBit = "Bitrate: " & outInfo.Bitrate & " kbit / sec, " & IIf(outInfo.HasVBR, "variable bit rate", "constant bit rate")
        lblFreq = "Frequency: " & outInfo.Frequency & " hz, " & LCase$(outInfo.ChannelMode)
        lblCopy = "Copyrighted: " & IIf(outInfo.IsCopyrighted, "yes", "no")
        lblOrig = "Original: " & IIf(outInfo.IsOriginal, "yes", "no")
        lblCRC = "Uses checksums: " & IIf(outInfo.HasCRC, "yes", "no")
        lblEmphasis = "Uses emphasis: " & IIf(outInfo.HasEmphasis, "yes", "no")
        lblv1Vers = "ID3 v1 tag: " & IIf(outInfo.ID3v1Version = -1, "none", "version 1." & outInfo.ID3v1Version)
        lblv2Vers = "ID3 v2 tag: " & IIf(outInfo.ID3v2Version = -1, "none", "version 2." & outInfo.ID3v2Version)
    

    
    'txt1$ = InTag.Album
    'If txt1$ <> a3$ Then
    '    kiss2 = 1
    'End If
    'txt1$ = InTag.Artist
'    If txt1$ <> a3$ Then kiss2 = 1
    
    'timemin$ = Trim(Str$(outInfo.Length \ 60))
     timemin$ = Trim(Str$(outInfo.Length \ 60))
    If outInfo.Length \ 60 < 100 Then
        timemin$ = " " + Trim(Str$(outInfo.Length \ 60))
        'dont work so weel dir list in explor dont show start space
    End If
    If outInfo.Length \ 60 < 10 Then
        timemin$ = "  " + Trim(Str$(outInfo.Length \ 60))
        'dont work so weel dir list in explor dont show start space
    End If
    
    leng$ = timemin$ & ":" & Format$(outInfo.Length Mod 60, "00")
    'kbps$ = """;""" + Trim(Str$(outInfo.Bitrate)) + " Kbps"
    
    If outInfo.Bitrate < 100 Then
        kbps$ = " " + Trim(Str$(outInfo.Bitrate)) + "K"
    End If
    If outInfo.Bitrate < 10 Then
        kbps$ = "  " + Trim(Str$(outInfo.Bitrate)) + "K"
    End If
    If outInfo.Bitrate >= 100 Then
        kbps$ = Trim(Str$(outInfo.Bitrate)) + "K"
    End If
    'kbps
    
    
    leng$ = leng$ + " " + kbps$
    c1$ = """" + B1$ + a1$

    txt1$ = InTag.Artist
    If txt1$ <> leng$ Then
        kiss = 1
    End If
    
    If outInfo.Bitrate = 0 Then
        kiss = 1
    End If
    
    'If ReadID3v2(a1$ + B1$, DummyInTag) = False Then kiss2 = 1

    If leng$ <> InTag.Artist Or tea2 = 1 Then
    kiss = 1
    
    End If
        
        
        
    'If Skiproid = True Then c1$ = " --- Skip Comment" Else c1$ = ""
    
    'ScanPath.List1.AddItem Format$(we, "000 ") + a1$ + B1$ '+ c1$
    DoEvents
    
'    If InTag.Comment <> Comment$ Then InTag.Comment = Comment$
    'Comment$ = InTag.Comment
    
    If InTag.Artist <> leng$ Then
        InTag.Artist = leng$
        kiss = 1
    End If
    
    ''If Comment$ <> "" Then Stop
    'If Trim(Comment$) = "" Then Comment$ = "Roids Rim " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p")
    'If Mid$(Comment$, 1, 9) = "Roids Rim" Then Comment$ = "Roids Rim " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf + Mid$(Comment$, 36)
    ''If Mid$(Comment$, 1, 9) = "Roids Rim" Then Comment$ = len("Roids Rim " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p")) + vbCrLf) + Mid$(Comment$, 40)
    

    
    'If InStr(a1$, "2004") = 0 And InStr(a1$, "2005") = 0 And InStr(a1$, "2006") = 0 Then
    'a3$ = Mid$(a1$, InStrRev(a1$, "\", Len(a1$) - 2))
    'Else
    'a4 = InStr(41, a1$, "\")
    'a3$ = Mid$(a1$, 30, 10) + " " + Mid$(a1$, 40, a4 - 40)
    ''a3$ = Left$(a3$, Len(a3$) - 1)
    'End If
    
    
    a3$ = Mid$(a1$, tt)
    w4 = InStr(a3$, "\")
    w4 = InStr(w4 + 1, a3$, "\")
'    If w4 > 0 Then
'        a4$ = Mid$(a3$, 1, w4)
'    End If
    
    
    
    t1 = InStrRev(a3$, "\")
    t2 = InStrRev(a3$, "\", t1 - 1)
    t3 = InStrRev(a3$, "\", t2 - 1)
'    t4 = InStrRev(a3$, "\", t3 - 1)

    If InStr(a3$, "\200") = 0 Then
        If t3 = 0 Then a3$ = Mid$(a3$, t2)
        If t3 > 0 Then a3$ = Mid$(a3$, t3)
    Else
        a3$ = Mid$(a3$, InStr(a3$, "\200") + 1)
    End If
    'If t3 > 0 Then a3$ = Mid$(a3$, t4)
    
'    a3$ = Left$(a3$, Len(a3$) - 1)
'    If InStr(a3$, Mid$(ra(rtu), 17)) > 0 Then
'    a3$ = Mid$(a3$, Len(ra(rtu)))
'    End If
    
    g1$ = a1$
    For r = 1 To Len(g1$)
        If Mid$(g1$, r, 1) = " " Or _
        (Mid$(g1$, r, 1) = "\" And r < Len(g1$)) Or _
        (Mid$(g1$, r, 1) = "-" And r < Len(g1$)) Then
            r = r + 1
            Mid$(g1$, r, 1) = UCase$(Mid$(g1$, r, 1))
        End If
    Next
    
    On Local Error Resume Next
    
    ttg$ = Format$(we3, "000 ") + " --- " + InTag.Artist + " --- " + InTag.Album + " --- " + InTag.title
    ttg$ = Format$(we3, "000 ") + " --- " + InTag.Artist + " --- " + InTag.Album + " --- " + InTag.title
    
    If a1$ <> g1$ Then
    Name a1$ As g1$
    ttg$ = ttg$ + " -- Renamed"
    End If
    On Local Error GoTo 0
'    If InTag.title = "Anx2005Web" Then Stop
    
    'If Mid$(a3$, 1, 5) = "\2005" Then Stop
    If InTag.Album <> a3$ Then InTag.Album = a3$
    
    
    ate7$ = "E:\04 Music ---\04 My Music\01 Banging Tunes\Above 128\"
    ate7$ = "C:\Above 128 OUT\"
    'MkDir ate7$
    If InTag.TrackNr <> "" Then InTag.TrackNr = ""
    
    
    
    If InStr(tagger$, "Size:") > 0 Then ut = Val(Mid$(tagger$, InStr(tagger$, "Size:") + 5))
    
 '   If ut = f.Size Then kiss = 0: kiss2 = 0
    
    If kiss = 1 Or kiss2 = 1 Then
    
        wa = WriteID3v2(a1$ + B1$, InTag)
        ttg$ = ttg$ + "  ----  Write Id3v2 Tag"
        Change = Change + 1
        Label15 = "Changed " + Str(Change)
        'ScanPath.List1.AddItem ttg$, 0
        'ScanPath.List1.AddItem "Converted -- " + ttg$, 1
    
    Else
        ttg$ = ttg$ + "  ----  No Change..."
        'ScanPath.ListView1.ListItems.Remove (we)
        'we = we - 1
        
    End If
        Label15 = "Changed " + Str(Change)
    
    
    ScanPath.List1.AddItem ttg$
    ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1
    ScanPath.List1.Refresh
    
    
    Set f = Fs22.Getfile(a1$ + B1$)
    
    If f.Attributes And FILE_ATTRIBUTE_ARCHIVE Then
        f.Attributes = f.Attributes - FILE_ATTRIBUTE_ARCHIVE
    End If
    If f.Attributes And FILE_ATTRIBUTE_HIDDEN Then
        f.Attributes = f.Attributes - FILE_ATTRIBUTE_HIDDEN
    End If
    If f.Attributes And FILE_ATTRIBUTE_READONLY Then
        f.Attributes = f.Attributes - FILE_ATTRIBUTE_READONLY
    End If
'    If f.Attributes And FILE_ATTRIBUTE_SYSTEM Then
'        f.Attributes = f.Attributes - FILE_ATTRIBUTE_SYSTEM
'    End If
'    Public Const FILE_ATTRIBUTE_HIDDEN = &H2
'Public Const FILE_ATTRIBUTE_NORMAL = &H80
'Public Const FILE_ATTRIBUTE_READONLY = &H1
'Public Const FILE_ATTRIBUTE_SYSTEM = &H4
'Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

    
    If outInfo.Bitrate > 128 Then
    'dupe declaration
    'Dim Fs22
    'Set Fs22 = CreateObject("Scripting.FileSystemObject")


        'write out to one folder
    
        'put back if want 128's
    '   wa = WriteID3v2(ate7$ + b1$, InTag)
'       no
    ate7$ = "C:\Above 128 OUT\"
        
    'write back into selected folders
    'Fs22.CopyFile ate7$ + b1$, a1$ + b1$
    'wa = WriteID3v2(a1$ + b1$, InTag)
    End If
    DoEvents

If Power = True Then Exit For

Next
If Power = True Then Exit For

Next
    
If Power = True Then ScanPath.List1.AddItem "Aborted......."
ScanPath.List1.AddItem "Completed......."

If Changed = 0 Then
    ScanPath.List1.AddItem "No Changes......", 0
    ScanPath.List1.AddItem "--------------------------", 1
Else
    ScanPath.List1.AddItem "Yes Changes......", 0
    ScanPath.List1.AddItem "---------------------------", 1
End If

ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1
ScanPath.List1.ListIndex = 0
ScanPath.List1.Refresh

'End

End Sub







Public Sub cboSize_Click()
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

Private Sub cmdBrowse_Click()
'    txtPath.Text = GetFolder(Me.hWnd, "Scan Path:", txtPath.Text)

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
    txtHelp.Visible = Not txtHelp.Visible
End Sub

Private Sub cmdQuit_Click()
Power = True
End
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

'Call Jeepers

End Sub





Private Sub Com1_Click()
Call Jeepers
End Sub

Private Sub Form_Load()
    
    
ScanPath.Height = List1.Top + List1.Height + 400
ScanPath.Width = List1.Width + 150
    
' frmMain.Caption = frmMain.Caption + App.EXEName
ScanPath.Caption = ScanPath.Caption + " - " + App.EXEName
   
    
    Dim lCount As Long
    
    With cboMask
        .AddItem "*.MP3"
        .AddItem "*.*"
        .AddItem "*.dll;*.exe;*.ocx"
        .AddItem "*.doc;*.mdb;*.xls"
        .AddItem "*.bmp;*.gif;*.jpg;*.tif"
        .AddItem "*.bas;*.cls;*.ctl;*.frm;*.vbp"
        .ListIndex = 0
    End With
    
    'DTPicker1(0).Value = DateSerial(Year(Now), Month(Now) - 3, Day(Now))
    'DTPicker1(1).Value = Now
    
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
        .ColumnHeaders.Add , "SIZE", "Size", 0, lvwColumnRight
        .ColumnHeaders.Add , "DATE", "Date", 0, lvwColumnLeft
        
        .View = lvwReport
    End With

Power = False
cboSizeType(0).ListIndex = 0
cboSize.ListIndex = 2
Call cboSize_Click
txtSize(0).Text = "1"
ScanPath.Show
'Call StartUp

Exit Sub

fg1 = FreeFile
Open App.Path + "\Scan Path.txt" For Input As #fg1
Line Input #fg1, ae$
Close #fg1
txtPath.Text = ae$
'txtPath.Text = App.Path


End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub SP_DirMatch(Directory As String, Path As String)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################
End Sub

Private Sub SP_FileMatch(fileName As String, Path As String)
    Dim LV As ListItem
    Dim uWIN32 As WIN32_FIND_DATA
 
    'Dim gg As String * 500

    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
'   On Local Error GoTo GetFileError
    
    With ListView1
        Set LV = .ListItems.Add(, , fileName)
        LV.SubItems(1) = Path
        
        
        
        GoTo jumpk
    Dim f
    Dim Fs22
    Set Fs22 = CreateObject("Scripting.FileSystemObject")
       
        GoTo jumpk
        jj$ = ""
        dd$ = ""
        For r = 1 To Len(fileName)
        jj$ = jj$ + Hex(Asc(Mid$(fileName, r, 1))) + Mid$(fileName, r, 1) + "-"
        f = 0
        If InStr("0123456789", Mid$(fileName, r, 1)) = 0 Then
        f = 1
        End If
        If f = 0 Then dd$ = dd$ + Mid$(fileName, r, 1)
        Next
        
jumk:


On Local Error Resume Next
        
        'LSet gg = (Path + fileName)
        GG$ = Path + fileName
        hh$ = GetShortName(GG$)
        hh$ = Path + fileName
        xx = InStrRev(hh$, "\")
        vv$ = Mid$(hh$, 1, xx)
        GG$ = Path
        
    Set f = Fs22.Getfile(hh$)
    If f.Attributes And 2 Then
        f.Attributes = f.Attributes - 2
'        Stop
    End If

    Set f = Fs22.GetFolder(GG$)
    If f.Attributes And 2 Then
        f.Attributes = f.Attributes - 2
        Stop
    End If
    
    
    
    'readonly = 1
    Set f = Fs22.GetFolder(GG$)
    If f.Attributes And 1 Then
        f.Attributes = f.Attributes - 1
        Stop
    End If
    
    Set f = Fs22.Getfile(hh$)
    If f.Attributes And 1 Then
        f.Attributes = f.Attributes - 1
        Stop
    End If
        
        
        
        
jumpk:
        
        
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


Sub StartUp()
'ra(3) = "E:\04 Music ---\04 My Music\"
'Load ScanPath
ScanPath.txtPath = "E:\04 Music ---\04 My Music\"
'ScanPath.txtPath = "E:\04 Music ---\03 My Music Zen\"
'ScanPath.txtPath = "E:\My Music\Talking Books"
ScanPath.cboMask.Text = "*.mp3"

Call ScanPath.cmdScan_Click

List1.AddItem "Finished"

End Sub



