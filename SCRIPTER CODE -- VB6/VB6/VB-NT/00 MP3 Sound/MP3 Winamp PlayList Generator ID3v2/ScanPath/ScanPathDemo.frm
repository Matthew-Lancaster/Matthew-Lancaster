VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ScanPath 
   AutoRedraw      =   -1  'True
   Caption         =   "ScanPath 2.0"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12945
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   12945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3255
      Top             =   2835
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PAUSE"
      Height          =   735
      Left            =   11415
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3105
      Width           =   645
   End
   Begin VB.CommandButton Com1 
      Caption         =   "ReScan"
      Height          =   870
      Left            =   11328
      Picture         =   "ScanPathDemo.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4920
      Width           =   768
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Default         =   -1  'True
      Height          =   912
      Left            =   11456
      Picture         =   "ScanPathDemo.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3885
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   585
      Top             =   6360
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Height          =   555
      Left            =   11536
      Picture         =   "ScanPathDemo.frx":1EA0
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5835
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   2115
      Left            =   12045
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   38
      Text            =   "ScanPathDemo.frx":276A
      Top             =   2715
      Visible         =   0   'False
      Width           =   3900
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
      Top             =   3885
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkPatternMatching 
      Caption         =   "Pattern Matching"
      Height          =   195
      Left            =   1530
      TabIndex        =   7
      Top             =   3645
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkNormal 
      Caption         =   "Normal"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   3645
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   3885
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   4125
      Width           =   1365
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   4350
      Width           =   1365
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   4575
      Width           =   1365
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   5940
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3825
      Width           =   1545
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   0
      Left            =   9930
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4215
      Width           =   1365
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   8580
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3825
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   8580
      TabIndex        =   16
      Top             =   4215
      Width           =   1305
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   1
      Left            =   9930
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4575
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   8580
      TabIndex        =   18
      Top             =   4575
      Width           =   1305
   End
   Begin VB.CheckBox chkRefreshListView 
      Caption         =   "Refresh ListView (slower)"
      Height          =   195
      Left            =   1530
      TabIndex        =   10
      Top             =   4425
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   12030
      TabIndex        =   22
      Top             =   0
      Width           =   12090
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
         Height          =   192
         Left            =   112
         TabIndex        =   23
         Top             =   64
         Width           =   1040
      End
   End
   Begin VB.CheckBox chkCopyMemory 
      Caption         =   "Use CopyMemory (display Date && Size)"
      Height          =   195
      Left            =   1530
      TabIndex        =   11
      Top             =   4665
      Width           =   3465
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2955
      Left            =   30
      TabIndex        =   21
      Top             =   5025
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5186
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
      Top             =   4125
      Width           =   2944
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   735
      Left            =   10860
      Picture         =   "ScanPathDemo.frx":28A2
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3105
      Visible         =   0   'False
      Width           =   510
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   0
      Left            =   5940
      TabIndex        =   13
      Top             =   4215
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   20578305
      CurrentDate     =   39358
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   1
      Left            =   5925
      TabIndex        =   14
      Top             =   4575
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   20578305
      CurrentDate     =   39358
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   8400
      TabIndex        =   50
      Top             =   2280
      Width           =   3690
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PERCENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   6660
      TabIndex        =   49
      Top             =   2280
      Width           =   1770
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EDIT GEN FOR NX RUN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   7290
      TabIndex        =   48
      Top             =   3150
      Width           =   3525
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   0
      TabIndex        =   47
      Top             =   1440
      Width           =   6645
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Processing Now"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   6660
      TabIndex        =   46
      Top             =   1515
      Width           =   1770
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Scanned"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   6660
      TabIndex        =   45
      Top             =   750
      Width           =   1770
   End
   Begin VB.Label LblCount2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   8415
      TabIndex        =   44
      Top             =   750
      Width           =   3675
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   8400
      TabIndex        =   43
      Top             =   1515
      Width           =   3690
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
      TabIndex        =   42
      Top             =   3255
      Width           =   2100
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
      TabIndex        =   37
      Top             =   3525
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
      Top             =   3375
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
      Top             =   3375
      Width           =   885
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      Height          =   195
      Index           =   0
      Left            =   7770
      TabIndex        =   34
      Top             =   4275
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
      Height          =   195
      Left            =   7770
      TabIndex        =   32
      Top             =   3855
      Width           =   585
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   195
      Index           =   1
      Left            =   7770
      TabIndex        =   36
      Top             =   4635
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   5130
      TabIndex        =   33
      Top             =   4275
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "File Date"
      Height          =   195
      Left            =   5130
      TabIndex        =   31
      Top             =   3855
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   5130
      TabIndex        =   35
      Top             =   4575
      Width           =   195
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


Dim OX1, OY1, MOUSE


Const DIMSIZE = 50
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        x As Long
        Y As Long
End Type

Dim P As POINTAPI


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



Public Sub Jeepers()
    
Power = False
    
'List1.Clear

ScanPath.Show


DoEvents

If Command$ <> "" Then Me.WindowState = 1

DoEvents

'ScanPath.WindowState = 0

Dim outInfo As MPEGInfo

'ScanPath.List1.AddItem "Converting MP3 Tags..."
    'ScanPath.List1.Refresh
    DoEvents

'On Local Error Resume Next
'f1 = FreeFile
'Open App.Path + "CRC Such.txt" For Input As #f1
'wdm$ = Input(LOF(f1), f1)
'Close #f1
'On Local Error GoTo 0


Dim InTag As ID3Tag
Dim DummyInTag As ID3Tag

    
Dim RA(DIMSIZE)
Dim RS(DIMSIZE)
Dim FOLDER_NAME(DIMSIZE)
Dim PLAYLIST_NAME(DIMSIZE)
Dim PLAYLIST_NAME2(DIMSIZE)
Dim RY1(DIMSIZE)
Dim RY2(DIMSIZE)
Dim RY3(DIMSIZE)
Dim RY4(DIMSIZE)


TA = 0
TA = TA + 1

RA(TA) = "M:\04 Music ---\07 Midi Mod"
FOLDER_NAME(TA) = "Winamp Gold 05 MIDI"
PLAYLIST_NAME(TA) = FOLDER_NAME(TA) + ".m3u"


TA = TA + 1
RA(TA) = "M:\Videos"
FOLDER_NAME(TA) = "Winamp Gold 07 Video 0"
PLAYLIST_NAME(TA) = FOLDER_NAME(TA) + ".m3u"


TA = TA + 1
RA(TA) = "M:\Videos\00 Vid XXX"
FOLDER_NAME(TA) = "Winamp Gold 07 Video X"
PLAYLIST_NAME(TA) = FOLDER_NAME(TA) + ".m3u"


TA = TA + 1
RA(TA) = "M:\04 Music ---\08 TBooks All"
FOLDER_NAME(TA) = "Winamp Gold 02 TBook"
PLAYLIST_NAME(TA) = FOLDER_NAME(TA) + ".m3u"


TA = TA + 1
RA(TA) = "M:\04 Music ---\05 Whole Lot"
FOLDER_NAME(TA) = "Winamp Gold 03 Whole Single's"
PLAYLIST_NAME(TA) = FOLDER_NAME(TA) + ".m3u"


TA = TA + 1
RA(TA) = "M:\04 Music ---\04 My Music"
FOLDER_NAME(TA) = "Winamp Gold 05 My Music 01"
PLAYLIST_NAME(TA) = FOLDER_NAME(TA) + ".m3u"
'PLAYLIST_NAME2(TA) = FOLDER_NAME(TA) + " - NEG-ONE.m3u"


'GoTo JUMPHERE


'TA = TA + 1
'RA(TA) = "M:\04 Music ---\04 My Music"
'FOLDER_NAME(TA) = "Winamp Gold 05 My Music 01"
'PLAYLIST_NAME(TA) = FOLDER_NAME(TA) + "- NEG-ONE.m3u"



'###### ---- ASUS MUSIC
'TA = TA + 1
'ra(TA) = "\\55-88-happy\D\04 Music ---\04 My Music"
'RY1(TA) = "\\55-88-happy\C\Program Files\00 Winamp's\ASUS Winamp 05 MY MUSIC 01\WinAmp ASUS.m3u"
'RY2(TA) = "\\55-88-happy\C\Program Files\00 Winamp's\ASUS Winamp 05 MY MUSIC 02\WinAmp ASUS.m3u"

'TA = TA + 1
'RA(TA) = "\\55-88-happy\D\04 Music ---\04 My Music"
'FOLDER_NAME(TA) = "Winamp Gold 05 My Music 01"
'PLAYLIST_NAME(TA) = "zz ASUS - " + FOLDER_NAME(TA) + ".m3u"

'TA = TA + 1
'ra(TA) = "\\55-88-happy\D\04 Music ---\06 Midi Mod"
'RY1(TA) = "\\55-88-happy\C\Program Files\00 Winamp's\ASUS Winamp 05 MIDI\WinAmp ASUS.m3u"



Dim RR
For RR = 1 To TA
    RY1(RR) = "C:\Program Files\00 WinAmp's\" + FOLDER_NAME(RR) + "\" + PLAYLIST_NAME(RR)
    
    If Mid(RA(RR), 1, 2) = "\\" Then
        'RY1(RR) = "\\55-88-happy\C\Program Files\00 WinAmp's\" + FOLDER_NAME(RR) + "\" + PLAYLIST_NAME(RR)
        RY1(RR) = "\\55-88-happy\C\TEMP2\00 WinAmp's\" + FOLDER_NAME(RR) + "\" + PLAYLIST_NAME(RR)
    End If
    
    If FS.FileExists(RY1(RR)) Then
        FS.COPYFILE RY1(RR), "M:\04 Music ---\WinAmp PlayLists\" + PLAYLIST_NAME(RR)
    End If
    
    If Mid(RA(RR), 1, 2) <> "\\" Then
        FS.COPYFILE RY1(RR), GetSpecialfolder(26) + "\" + FOLDER_NAME(RR) + "\Winamp.m3u"
        FS.COPYFILE RY1(RR), GetSpecialfolder(26) + "\" + FOLDER_NAME(RR) + "\Winamp.m3u8"
    End If

Next



'GoTo jump
On Error Resume Next
Open App.Path + "\DirSizeFlags.txt" For Input As #1
Do
Line Input #1, kl$
ty = ty + 1
RS(ty) = Val(Mid(kl$, 1, 15))
Loop Until EOF(1)
Close #1


'NEEDS QUICK WAY TO SEE DIR CHANGED IF THERE IS REAL ONE
Err.Clear
Open App.Path + "\DirSizeFlags.txt" For Output As #1
For RTU = 1 To TA
    If Right$(RA(RTU), 1) <> "\" Then RA(RTU) = RA(RTU) + "\"
    
    Debug.Print RA(RTU)
    Label13 = "WORKING - GET TREE SIZE" + Str(RTU) + " OF" + Str(TA)
    ScanPath.txtPath = RA(RTU)
    Me.Refresh
    
    TTG = 0
    DoEvents
    Set F = FS.GetFolder(RA(RTU))
    TTG = F.Size
    
    Print #1, Format(TTG, "000000000000000") + " " + RA(RTU)
    If TTG = RS(RTU) Then RA(RTU) = ""
Next
Close #1

'If Err.Number > 0 `Then MsgBox "Error Open App.Path + ""\DirSizeFlags.txt"" For Output As #1"

On Error GoTo 0

For RTU = 1 To TA
'If RTU > 1 Then Exit Sub

Do
    If RA(RTU) = "" Then RTU = RTU + 1
Loop Until RTU = TA Or RA(RTU) <> ""
If RA(RTU) = "" Then Exit For
'rtu = rtu + 1
'If rtu = 3 Then ScanPath.chkNormal.Value = vbChecked
    
ScanPath.txtPath = RA(RTU)
ScanPath.cboMask.Text = "*.mp3;*.avi;*.wav;*.wma;*.mpeg;*.mpg;*.wmv;*.rm;*.asf;*.qt;*.rmi;*.mid;*.mod;*.midi;*.flv"

t0 = 10
Label14 = "Modified in Last " + Str(t0) + " days"
'DTPicker1(0).Value = Now - t0
'DTPicker1(1).Value = Now
'Chk Box ticked when date set
    
cboSizeType(0).ListIndex = 0
cboSize.ListIndex = 0

Call cboSize_Click

txtSize(0).Text = "0"

'ScanPath.chkArchive.Value = vbUnchecked

ScanPath.chkArchive.Value = vbChecked
ScanPath.chkNormal.Value = vbChecked
ScanPath.chkSystem.Value = vbChecked
ScanPath.chkHidden.Value = vbChecked
ScanPath.chkReadOnly.Value = vbChecked


ScanPath.ListView1.ListItems.Clear

Label13 = "WORKING - SCAN PATH - " + Str(RTU) + " OF" + Str(TA)

 XXh = 0


Call ScanPath.cmdScan_Click


Dim TF1, TF2
TF1 = RA(RTU) = "M:\Videos"
TF2 = RA(RTU) = "M:\04 Music ---\04 My Music\"
'EXTRA ONES FOR THIS
If TF1 Or TF2 Then
    
    Dim CONDU(50)
    
    'CT = CT + 1: CONDU(CT) = "M:\04 Music ---\08 Tbooks All\Tbook 01 Children Ladys\"
    
    'M:\04 Music ---\08 Tbooks All\Lord Of The Rings J.R.R Tolkien
    
    'M:\04 Music ---\07 Midi Mod\MIDI\06 MIDI ENCARTA MIDI
    'M:\04 Music ---\07 Midi Mod\MIDI\07 Rmi
    'M:\04 Music ---\07 Midi Mod\MIDI\04 CHART MIDI
    'M:\04 Music ---\07 Midi Mod\MIDI\03 FILM MIDI\01 Film Tracks
    'M:\04 Music ---\07 Midi Mod\MIDI\02 MIDI
    'M:\04 Music ---\07 Midi Mod\MIDI\01 Best
    'M:\04 Music ---\07 Midi Mod\MIDI\00 Bigger Than Size Copy
    
    
    CT = 0
    'CT = CT + 1: CONDU(CT) = "M:\04 Music ---\08 Tbooks All\Lord Of The Rings J.R.R Tolkien\"
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\08 Tbooks All\"
    
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\05 Whole Lot\"
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\05 HTTrack\"
    
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\07 Midi Mod\MIDI\00 Bigger Than Size Copy"
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\07 Midi Mod\MIDI\01 Best\"
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\07 Midi Mod\MIDI\06 MIDI ENCARTA MIDI"
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\07 Midi Mod\MIDI\07 Rmi"
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\07 Midi Mod\MIDI\04 CHART MIDI"
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\07 Midi Mod\MIDI\03 FILM MIDI\01 Film Tracks"
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\07 Midi Mod\MIDI\02 MIDI"
    
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\07 Midi Mod\MODS"
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\07 Midi Mod\WAVE"
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\01 SOUND EFFECT"
    
    CT = CT + 1: CONDU(CT) = "M:\04 Music ---\00 - CD-CLASSICAL"
    
    'CT = CT + 1: CONDU(CT) = "M:\04 Music ---\08 Tbooks All\Tbook 03 Non Fictions\"
    'CT = CT + 1: CONDU(CT) = "M:\04 Music ---\08 Tbooks All\Tbook 02 Non Fictions Ladys\"
    'CT = CT + 1: CONDU(CT) = "M:\04 Music ---\08 Tbooks All\Tbook 03 Poetry All Male\"
    'CT = CT + 1: CONDU(CT) = "M:\04 Music ---\08 Tbooks All\Tbook 04 Children\"
    'CT = CT + 1: CONDU(CT) = "M:\04 Music ---\08 Tbooks All\TBook 10 Children 2010\"
    'CT = CT + 1: CONDU(CT) = "M:\04 Music ---\08 Tbooks All\TBook 10 Tom\"
    
    
    CT = CT + 1: CONDU(CT) = "M:\Videos\00 Vid Films\MP3's\"
    CT = CT + 1: CONDU(CT) = "M:\Videos\00 Vid Films"
    CT = CT + 1: CONDU(CT) = "M:\Videos\00 Vid XXX"
    CT = CT + 1: CONDU(CT) = "M:\Videos\03 VIDEO\"
    CT = CT + 1: CONDU(CT) = "M:\Videos\01 Beavis & Butthead\"
    CT = CT + 1: CONDU(CT) = "M:\Videos\04 MOBILE\"
    CT = CT + 1: CONDU(CT) = "M:\Videos\05 BILL HICKS - TRIP A TRON -  TAE KWON DO"
    CT = CT + 1: CONDU(CT) = "M:\Videos\05 Red Planet"
    CT = CT + 1: CONDU(CT) = "M:\Videos\05 Videos 01 Simpsons"
    CT = CT + 1: CONDU(CT) = "M:\Videos\05 WebCam"
    CT = CT + 1: CONDU(CT) = "M:\Videos\07 Photo Slides"
    
   
    
    For RST = 1 To CT
        DoEvents
        ScanPath.txtPath = CONDU(RST)
        Call ScanPath.cmdScan_Click
    Next
    
    
    'REMOVE NOT WANTED
    
    SCPL = ScanPath.ListView1.ListItems.Count
    For WE = SCPL To 1 Step -1
        
        Label13 = "WORKING - REMOVE FROM LIST SELECTED - " + Str(WE) + " OF" + Str(SCPL)
        
        A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
        B1 = ScanPath.ListView1.ListItems.Item(WE)
        
        ReDim CONDUC(20)
        
        CT = 0
        CT = CT + 1: CONDU(CT) = "01 Matts Vids XX orignals"
        'CT = CT + 1: CONDU(CT) = "M:\04 Music ---\04 My Music\06 Talking\04 Talking Books"
        CT = CT + 1: CONDU(CT) = "Stephen R. Covey - The 7 Habits Of Highly Effective Families"
        'CT = CT + 1: CONDU(CT) = "JOINED Books"
        CT = CT + 1: CONDU(CT) = "Fictions\Art Of War"
        'CT = CT + 1: CONDU(CT) = "\LADY"
        CT = CT + 1: CONDU(CT) = "\MALE"
        CT = CT + 1: CONDU(CT) = "MALE"
        CT = CT + 1: CONDU(CT) = "01 Matts Vids XX orignals"
        CT = CT + 1: CONDU(CT) = "01 SX"
        If TF2 = True Then CT = CT + 1: CONDU(CT) = "01 XXX Mp3"
                    
                
        For RST = 1 To CT
            If InStr(LCase(A1), LCase(CONDU(RST))) > 0 Then
                ScanPath.ListView1.ListItems.Remove (WE)
            End If
        Next
    
    Next
    ScanPath.txtPath = RA(RTU)
End If
      
      
      
'Exit Sub

      
    
TT = InStr(4, ScanPath.txtPath, "\") + 1
a3$ = Mid$(ScanPath.txtPath, TT)



Label13 = "WORKING - SCAN PATH - " + Str(RTU) + " OF" + Str(TA)


'TOTAL SCANNED
'LblCount2 = XXh

'AFTER FILTERS
lblCount = Trim(Str(ScanPath.ListView1.ListItems.Count))


On Local Error GoTo 0

'Reset
On Error Resume Next
    fr1 = FreeFile
    Open RY1(RTU) + ".tmp" For Output As #fr1
    Print #fr1, "#EXTM3U"
    Change = 0

    'If Err.Number > 0 Then MsgBox "Error - Open RY1(rtu) + "".tmp"" For Output As #fr1"
On Error GoTo 0


For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    If XPAUSE = True Then
        Do
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            Sleep 500
        Loop Until XPAUSE = False
    End If
    
    we3 = we3 + 1
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    'If InStr(A1, "Mos Clubber's Guide To Ibiza Summer 2001") > 0 Then Stop
    'If B1 = "Mos Clubber'S Guide To Ibiza Summer 2001 Cd2-05 Stringer.mp3" Then Stop
    'If A1 = "E:\04 Music ---\04 My Music\02 Chart Songs\00 IDTag-Able\Eddies CDs\Min Of Sound Annual 2003\CD 01\" Then Stop
    'If B1 = "Min Of Sound Annual 2003 Cd1 - 05 The Theme.mp3" Then Stop
    
    On Local Error Resume Next
    Err.Clear
    Set F = FS.Getfile(A1 + B1)
    'On Local Error GoTo 0
    
    
    xyh2 = InStrRev(B1, ".")
    title3$ = Mid$(B1, 1, xyh2 - 1)
    qq = False
    If LCase(Mid$(B1, xyh2)) = ".mp3" Then qq = True
    
    
    Comment$ = ""
    If qq = True Then
    If ReadID3v1(A1 + B1, InTag) = True Then
        Comment$ = "ID3v1 Previous Info " + Format$(Now, "DD-MM-YYYY HH:MM:SS") + vbCrLf
        If InTag.title <> "" Then Comment$ = Comment$ + "Title-----" + InTag.title + vbCrLf
        If InTag.Artist <> "" Then Comment$ = Comment$ + "Artist----" + InTag.Artist + vbCrLf
        If InTag.Album <> "" Then Comment$ = Comment$ + "Album-----" + InTag.Album + vbCrLf
        If InTag.SongYear <> "" Then Comment$ = Comment$ + "SongYear--" + InTag.SongYear + vbCrLf
        If InTag.Genre <> "" Then Comment$ = Comment$ + "Genre-----" + InTag.Genre + vbCrLf
        If InTag.Comment <> "" Then Comment$ = Comment$ + "Comment----" + InTag.Comment + vbCrLf
        If InTag.TrackNr <> "" Then Comment$ = Comment$ + "TrackNr----" + InTag.TrackNr + vbCrLf
        Comment$ = "Matty Roid - " + Format$(Now, "DD-MM-YYYY HH:MM:SS")
    End If
    End If
        
    If qq = True Then
    kiss2 = 0
    wa = ReadID3v2(A1 + B1, InTag)
    If wa = False Then kiss2 = 1 ': Stop
    
    tagger$ = InTag.Comment
        
    mdlID3.ReadMPEGInfo A1 + B1, outInfo
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
    
    End If
    
    timemin$ = Trim(Str$(outInfo.Length \ 60))
    If outInfo.Length \ 60 < 100 Then
        timemin$ = " " + Trim(Str$(outInfo.Length \ 60))
        'dont work so well dir list in explor dont show start space
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
    c1$ = """" + B1 + A1

    txt1$ = InTag.Artist
    If txt1$ <> leng$ Then
        kiss = 1
    End If
    
    If outInfo.Bitrate = 0 Then
        kiss = 1
    End If
    
    'If ReadID3v2(A1 + B1, DummyInTag) = False Then kiss2 = 1

    If leng$ <> InTag.Artist Or tea2 = 1 Then
    kiss = 1
    
    End If
        
        
    If InTag.Artist <> leng$ And outInfo.Length > 0 Then
        InTag.Artist = leng$
        kiss = 1
    End If
    
    TT = InStr(4, ScanPath.txtPath, "\")

    a3$ = Mid$(A1, TT)
    w4 = InStr(2, a3$, "\")
    w4 = InStr(w4 + 1, a3$, "\")
    
    'On Local Error GoTo 0
    
    'On Local Error Resume Next
    
'    If InStr(A1, "New 07") > 0 Then Stop
    
        
    yy1$ = "#EXTINF:"
    
    xyh2 = InStrRev(B1, ".")
    title3$ = Mid$(B1, 1, xyh2 - 1)
    yy1$ = ""
    If LCase(Mid$(B1, xyh2)) = ".mp3" Then
        If F.Size = 0 Then outInfo.Length = 0: InTag.title = "": InTag.Artist = ""
        yy1$ = "#EXTINF:"
        yy1$ = yy1$ + Trim(Str(outInfo.Length)) + ","
        If InTag.title <> "" Then title3$ = InTag.title
        Do
            xs = InStr(title3$, "_")
            If xs > 0 Then
                Mid$(title3$, xs, 1) = " "
            End If
        Loop Until xs = 0
        If InTag.Artist = "" Then yy1$ = yy1$ + title3$ + vbCrLf
        If InTag.Artist <> "" Then yy1$ = yy1$ + InTag.Artist + " - " + title3$ + vbCrLf
    End If
    
    
    
    yy1$ = yy1$ + A1 + B1

    Print #fr1, yy1$
    
        
    Change = Change + 1
    'Label15 = "Count" + Str(Change)
    Label15 = Trim(Str(Change))
    
    'ak54lmo abby car number plate
    
    
    'ScanPath.List1.AddItem ttg$
    'ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1
    'ScanPath.List1.Refresh
    
    
    DoEvents

    If Power = True Then Exit For

Next

If Power = True Then Exit For

Close fr1

FS.COPYFILE RY1(RTU) + ".tmp", RY1(RTU)
Kill RY1(RTU) + ".tmp"

Next
    
For RR = 1 To TA
    'ALREADY DONE
    'RY1(RR) = "C:\Program Files\00 WinAmp's\" + FOLDER_NAME(RR) + "\" + PLAYLIST_NAME(RR)
    'RY4(RR) = "\\55-88-happy\c\Program Files\00 Winamp's\ASUS " + FOLDER_NAME(RR) + "\" + PLAYLIST_NAME(RR)
    
    
    FS.COPYFILE RY1(RR), "M:\04 Music ---\WinAmp PlayLists\" + PLAYLIST_NAME(RR)
    FS.COPYFILE RY1(RR), GetSpecialfolder(26) + "\" + FOLDER_NAME(RR) + "\Winamp.m3u"
    FS.COPYFILE RY1(RR), GetSpecialfolder(26) + "\" + FOLDER_NAME(RR) + "\Winamp.m3u8"

Next
    
Dim Line1, LINE2
JUMPHERE:
For RR = 1 To TA
    
    TXF1 = "M:\04 Music ---\WinAmp PlayLists\" + PLAYLIST_NAME(RR)
    TXF2 = "Winamp Gold 05 My Music 01.m3u"
    TXF3 = "M:\04 Music ---\WinAmp PlayLists\Winamp Gold 05 My Music 01-ONE NO.m3u"
    
    If Dir(TXF1) = TXF2 Then
        Open TXF1 For Input As #1
        Open TXF3 For Output As #2
        Line Input #1, HEADERLINE
        Print #2, HEADERLINE
        
        Do
                
            DoEvents
            
            Line Input #1, LINE2
            If Mid(LINE2, 1, 1) = "#" Then Line1 = LINE2
                
            'If InStr(LINE2, "Videos\00 Vid XXX\") > 0 Then Stop
            
            If InStr(LINE2, "Videos\00 Vid XXX\") = 0 And Line1 <> LINE2 Then
                'LINES DONT HAVE A HEADER # WHEN THEY NOT MP3
                'If LINE1 = "" Then Stop
                If Line1 <> "" Then Print #2, Line1: Line1 = ""
                Print #2, LINE2
            End If
        
        Loop Until EOF(1)
        
        Close #1, #2
        
    
    End If

Next
    
    
    
    
    
    

ScanPath.WindowState = 0
ScanPath.ListView1.BackColor = vbYellow


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
        'lblCount.Caption = "-"
 
        'Screen.MousePointer = vbHourglass
        'ListView1.ListItems.Clear
        
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
        
        cmdScan.Caption = "Scan"
        'Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
    End If

'Call Jeepers
    
    
    
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





Private Sub Com1_Click()

PROG_START = Now + TimeSerial(0, 0, 10)

Call Jeepers


Label13 = "** DONE **"
Label13.BackColor = vbGreen

MsgBox "** DONE ** PLAYLIST GENERATE **", vbMsgBoxSetForeground

End Sub

Private Sub Command1_Click()
XPAUSE = Not XPAUSE
If XPAUSE = True Then Command1.BackColor = &HFF&
If XPAUSE = 0 Then Command1.BackColor = &H808080


End Sub

Private Sub Form_Load()
'Jeepers
Set FS = CreateObject("Scripting.FileSystemObject")
    
PROG_START = Now + TimeSerial(0, 0, 10)
    
Dim r As Long
On Error Resume Next
'For R = 0 To 120
    'Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
'Next
'End
    
    
Label13 = "WORKING"
    
'ScanPath.Height = ListView1.Top + ListView1.Height + 400
'ScanPath.Width = ListView1.Width + 150
    
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
        .ColumnHeaders.Add , "PATH", "Path", 5000, lvwColumnLeft
        .ColumnHeaders.Add , "SIZE", "Size", 0, lvwColumnRight
        .ColumnHeaders.Add , "DATE", "Date", 0, lvwColumnLeft
        
        .View = lvwReport
    End With


Call Jeepers

Label13 = "** DONE **"
Label13.BackColor = vbGreen
MsgBox "** DONE ** PLAYLIST GENERATE **", vbMsgBoxSetForeground

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
Y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.Top > Y Then Y = Control.Height + Control.Top
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

Me.Height = (Y + pluso)
Me.Refresh
DoEvents

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label15_Change()
Label20.Caption = Format((Val(Label15) / Val(LblCount2)) * 100, "0.00") + "%"
End Sub

Private Sub Label18_Click()
    Shell "NOTEPAD " + App.Path + "\DirSizeFlags.txt", vbMaximizedFocus
End Sub

Private Sub LblCount2_Change()

'LblCount2.Refresh

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
    Dim F
    Dim FS
    Set FS = CreateObject("Scripting.FileSystemObject")
       
        GoTo jumpk
        jj$ = ""
        dd$ = ""
        For r = 1 To Len(fileName)
        jj$ = jj$ + Hex(Asc(Mid$(fileName, r, 1))) + Mid$(fileName, r, 1) + "-"
        F = 0
        If InStr("0123456789", Mid$(fileName, r, 1)) = 0 Then
        F = 1
        End If
        If F = 0 Then dd$ = dd$ + Mid$(fileName, r, 1)
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
        
    Set F = FS.Getfile(hh$)
    If F.Attributes And 2 Then
        F.Attributes = F.Attributes - 2
'        Stop
    End If

    Set F = FS.GetFolder(GG$)
    If F.Attributes And 2 Then
        F.Attributes = F.Attributes - 2
        Stop
    End If
    
    
    
    'readonly = 1
    Set F = FS.GetFolder(GG$)
    If F.Attributes And 1 Then
        F.Attributes = F.Attributes - 1
        Stop
    End If
    
    Set F = FS.Getfile(hh$)
    If F.Attributes And 1 Then
        F.Attributes = F.Attributes - 1
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
        If (LV.Index Mod 10) = 0 And TESTNOT = 100 Then
            If chkRefreshListView.Value Then
                .SelectedItem = LV
                .SelectedItem.EnsureVisible
            End If
            
            lblCount.Caption = Trim(Str(LV.Index))
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



Private Sub Timer2_Timer()

Dim x1, y1

FindCursor x1, y1


If OX1 <> x1 Or OY1 <> y1 Then
    MOUSE = Now + TimeSerial(0, 0, 5)
End If

If MOUSE > Now Then
    XPAUSE = True

Else
    XPAUSE = False
End If

If XPAUSE = True Then Command1.BackColor = &HFF&
If XPAUSE = 0 Then Command1.BackColor = &H808080

OX1 = x1
OY1 = y1


End Sub

Public Sub FindCursor(x, Y)

'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'FindCursor Tx, Ty
'Private Type POINTAPI
'        x As Long
'        Y As Long
'End Type

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
x = P.x ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
Y = P.Y '/ GetSystemMetrics(1) * Screen.Height

End Sub


