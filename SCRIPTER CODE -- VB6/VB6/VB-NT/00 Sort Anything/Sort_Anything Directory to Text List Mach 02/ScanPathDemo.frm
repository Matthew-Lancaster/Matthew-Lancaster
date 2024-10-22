VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmMain 
   Caption         =   "ScanPath 2.0 - Anything -- Demo of Class Object"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13950
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   13950
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   75
      TabIndex        =   43
      Top             =   6225
      Width           =   13815
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
      Left            =   9300
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3675
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   2115
      Left            =   1275
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      Text            =   "ScanPathDemo.frx":1194
      Top             =   3675
      Visible         =   0   'False
      Width           =   7245
   End
   Begin VB.ComboBox cboMask 
      Height          =   315
      Left            =   540
      TabIndex        =   1
      Text            =   "*.jpg"
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
      Value           =   1  'Checked
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
      Left            =   5940
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2010
      Width           =   1545
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   0
      Left            =   9930
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2400
      Width           =   1365
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   8580
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2010
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   8580
      TabIndex        =   16
      Top             =   2400
      Width           =   1305
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   1
      Left            =   9930
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2760
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   8580
      TabIndex        =   18
      Top             =   2760
      Width           =   1305
   End
   Begin VB.CheckBox chkRefreshListView 
      Caption         =   "Refresh ListView (slower)"
      Height          =   195
      Left            =   1530
      TabIndex        =   10
      Top             =   2610
      Width           =   2205
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   13890
      TabIndex        =   22
      Top             =   0
      Width           =   13950
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
      Top             =   2850
      Width           =   3465
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2970
      Left            =   45
      TabIndex        =   21
      Top             =   3210
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   5239
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
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Default         =   -1  'True
      Height          =   825
      Left            =   9450
      Picture         =   "ScanPathDemo.frx":12CC
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   720
      Width           =   1875
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   0
      Left            =   5940
      TabIndex        =   13
      Top             =   2400
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   22216705
      CurrentDate     =   37299
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   1
      Left            =   5940
      TabIndex        =   14
      Top             =   2760
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   22216705
      CurrentDate     =   37296
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7425
      TabIndex        =   42
      Top             =   1425
      Width           =   1875
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7440
      TabIndex        =   41
      Top             =   1080
      Width           =   1875
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
      Left            =   7770
      TabIndex        =   35
      Top             =   2460
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
      Height          =   195
      Left            =   7770
      TabIndex        =   33
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   195
      Index           =   1
      Left            =   7770
      TabIndex        =   37
      Top             =   2820
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   5130
      TabIndex        =   34
      Top             =   2460
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "File Date"
      Height          =   195
      Left            =   5130
      TabIndex        =   32
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   5130
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
      Height          =   285
      Left            =   7425
      TabIndex        =   31
      Top             =   705
      Width           =   1875
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MACH 02 CHANGED TO INCLUDE SUBFOLDERS AND FULL LIST OF
'FILENAME


'This is Directory PROGRAM -- it will look this format
'----------------------------------------------------------------------
'Directory to Text File VB6
'----------------------------------------------------------------------
'Dir = D:\# MY DOCS\CallerID\Text-9\Andy\
'----------------------------------------------------------------------
'File Count = 162
'----------------------------------------------------------------------
'#00-Andy
'1 ANDY11 Prison
'1 ANDY13
'1 ANDY14
'1 Andy15 Diary
'1 Andy16 Pub
'2Andy Double Words List Sent
'Andy SMS InBox
'Andy SMS OutBox
'Andy 1
'Andy 2
'Andy 03 3 Famous Stories



'-------------Is This Folder Get on It at start      at       Bangers

'Put This in a Bas Mobule
Public m_CRC As clsCRC
Public WxHex$

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
       
'Author:    Richard Mewett �2005
'##############################################################################################

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'We must declare with WithEvents to process the files returned
Private WithEvents SP As cScanPath
Attribute SP.VB_VarHelpID = -1







Sub Bangers()

cboMask.Text = "*.asf;*.mpg;*.rm;*.avi;*.mp3;*.txt;*.jpg;*.gif;*.txt"
cboMask.Text = "*.nws;*.txt"
cboMask.Text = "*.*"

tpath1$ = "M:\04 Music ---\03 My Music Zen\03 Small'S\Smalls Tv Theme'S\Tv Themes All\"
tpath1$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Alt.Sz_Best\Temp Al.SZBest\01_All_Sz\"
tpath1$ = "M:\Videos\Beavis & Butthead Whole Lot\"
tpath1$ = "M:\Videos\Videos 01 Simpsons\"
tpath1$ = "M:\04 Music ---\05 Whole Lot\"
tpath1$ = "M:\04 Music ---\04 My Music\02 Chart\"
tpath1$ = "M:\04 Music ---\04 My Music\01 Banging Tunes"
tpath1$ = "M:\04 Music ---\00 TBooks All"
tpath1$ = "M:\0 00 Art\00 My Pictures\03 DJ's Banging Tunes.Com"
tpath1$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files"
tpath1$ = "D:\# MY DOCS\CallerID\Text-9\Andy"
'tpath1$ = ""

tpath1$ = "M:\04 My Music"

If Mid$(tpath1$, Len(tpath1$), 1) <> "\" Then
    tpath1$ = tpath1$ + "\"
End If

txtPath.Text = tpath1$

Call cmdScan_Click

uu$ = Mid$(tpath1$, InStrRev(tpath1$, "\", Len(tpath1$) - 1))
uu$ = Mid$(uu$, 2, Len(uu$) - 2)

easy$ = "#00-" + uu$ + ".txt"


Open tpath1$ + easy$ For Output As #1

Print #1, String$(70, "-")
Print #1, "Directory to Text File Program"
Print #1, String$(70, "-")
Print #1, "Dir = " + txtPath.Text
Print #1, String$(70, "-")
Print #1, "File Count =" + Str(ListView1.ListItems.Count)
Print #1, String$(70, "-")

For we = 1 To ListView1.ListItems.Count
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)
    'b1$ = Mid$(b1$, 1, Len(b1$) - 4)
    Print #1, b1$
Next
Close #1

Set fs22 = CreateObject("Scripting.FileSystemObject")
'fs22.copyFile TPATH1$ + easy$, "D:\# MY DOCS\# 01 My Documents\00 Email Texts\Dir to Text List\" + easy$


'Shell "notepad " + "D:\# MY DOCS\# 01 My Documents\00 Email Texts\Dir to Text List\" + easy$, vbNormalFocus

Shell "notepad " + tpath1$ + easy$, vbMaximizedFocus

End

Exit Sub


tpath1$ = "E:\04 Music ---\03 My Music Zen\"
tpath2$ = "E:\04 Music ---\04 My Music\"
tpath3$ = "E:\04 Music ---\Del\"
'MkDir tpath3$
txtPath.Text = tpath1$

cboMask.Text = "*.mp3"

Call cmdScan_Click
'Dim Fs22
Set fs22 = CreateObject("Scripting.FileSystemObject")
'Fs22.copyFile a1$ + b1$, tpath2$ + c1$ + b1$
List1.AddItem "Copy On Smallest Size From Zen Folder an To Zen folder"
List1.AddItem "Then it will compare for duplicates an click them bottom list to put removed folder"

For we = 1 To ListView1.ListItems.Count
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)
c1$ = Mid$(a1$, Len(tpath1$))
rr = FileExists(tpath2$ + c1$ + b1$)
If rr = True Then
'call date_File()
tt1 = FindFileSize(a1$ + b1$)
tt2 = FindFileSize(tpath2$ + c1$ + b1$)
If tt1 <> tt2 Then
If tt1 = 0 And Mid$(b1$, 1, 8) <> "--------" Then MsgBox a1$ + b1$ + " = Zero"
If tt2 = 0 And Mid$(b1$, 1, 8) <> "--------" Then MsgBox a1$ + b1$ + " = Zero"
End If
If tt1 < tt2 And tt1 > 0 Then
snot = snot + 1
List1.AddItem "Copy From Zen Folder " + Str(we) + Str(snot) + " -- " + a1$ + b1$
DoEvents
List1.Refresh
DoEvents
fs22.copyFile a1$ + b1$, tpath2$ + c1$ + b1$
End If
If tt2 < tt1 And tt2 > 0 Then
snot = snot + 1
List1.AddItem "Copy To Zen Folder " + Str(we) + Str(snot) + " -- " + a1$ + b1$
DoEvents
List1.Refresh
DoEvents
fs22.copyFile tpath2$ + c1$ + b1$, a1$ + b1$
End If



End If
Next


tpath1$ = "E:\04 Music ---\03 My Music Zen\"
tpath2$ = "E:\04 Music ---\04 My Music\"

txtPath.Text = tpath1$

cboMask.Text = "*.mp3"

Call cmdScan_Click
'Dim Fs22
Set fs22 = CreateObject("Scripting.FileSystemObject")

For we = ListView1.ListItems.Count To 1 Step -1
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)
    tt1 = FindFileSize(a1$ + b1$)

    ListView1.ListItems.Item(we).SubItems(1) = Format$(tt1, "0000000000") + " - " + a1$ + b1$
    ListView1.SelectedItem = ListView1.ListItems(we)
    ListView1.SelectedItem.EnsureVisible
    If tt1 = 0 Then ListView1.ListItems.Remove (we)
Next

ListView1.SortOrder = lvwAscending
ListView1.Sorted = True        'Use default sorting to sort the
ListView1.SortKey = 1

For we = ListView1.ListItems.Count To 2 Step -1
    DoEvents
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)
    a2$ = ListView1.ListItems.Item(we - 1).SubItems(1)
    b1$ = ListView1.ListItems.Item(we - 1)
    'tt1 = FindFileSize(a1$ + b1$)
kk1 = Val(Mid$(ListView1.ListItems.Item(we).SubItems(1), 1, 10))
kk2 = Val(Mid$(ListView1.ListItems.Item(we - 1).SubItems(1), 1, 10))
ii1$ = "0": ii2$ = "1"

If kk1 = kk2 Then
    Reset
    fr1 = FreeFile
    Open Mid$(a1$, 14) For Binary As #fr1
    fr2 = FreeFile
    Open Mid$(a2$, 14) For Binary As #fr2
    ii1$ = Input$(4000, fr1)
    ii2$ = Input$(4000, fr2)
    Close #fr1, #fr2
End If


If ii1$ = ii2$ Then
List1.AddItem a1$
List1.AddItem a2$
ListView1.ListItems.Remove (we)
End If
Next
If Val(Mid$(ListView1.ListItems.Item(1).SubItems(1), 1, 10)) = 0 Then
    ListView1.ListItems.Remove (1)
End If

List1.AddItem "Complete.................."

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

For we = 1 To ListView1.ListItems.Count
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)
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
rr$ = Input$(LOF(f1), f1)
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


For we = 1 To ListView1.ListItems.Count
'For we = 1 To 5000
    
    
    
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)


    ii = ii + 1
'    List1.AddItem Str$(ii) + " -- " + a1$ + b1$
 '   List1.ListIndex = List1.ListCount - 1
    
    WxHex$ = Space$(8)

    LSet WxHex$ = Hex(m_CRC.CalculateFile(a1$ + b1$))
    
    ListView1.ListItems.Item(we).SubItems(1) = WxHex$ + " - - " + Format$(we, "0000000") + " - - " + a1$
    ListView1.SelectedItem = ListView1.ListItems(we)
    ListView1.SelectedItem.EnsureVisible
    Label13.Caption = Str(we)
    Label14.Caption = Str(we)

 '   Print #1, WxHex$, " - - "; a1$, b1$
DoEvents



Next
    
    ListView1.SortOrder = lvwAscending
        
    ListView1.Sorted = True        'Use default sorting to sort the
    
    ListView1.SortKey = 1

For we = 2 To ListView1.ListItems.Count
    
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)
    
    a12$ = ListView1.ListItems.Item(we - 1).SubItems(1)
    b12$ = ListView1.ListItems.Item(we - 1)




If Mid$(a1$, 1, 3) <> "G:\" Then
    
    g1$ = Mid$(a1$, 1, 8)
    g2 = Val(Mid$(a1$, 14, 7))
    g3$ = Mid$(a1$, 26)
    
    h1$ = Mid$(a12$, 1, 8)
    h2 = Val(Mid$(a12$, 14, 7))
    h3$ = Mid$(a12$, 26)

rt = 0
If InStr(rr$, g1$) Then
    If rt = 0 Then MsgBox "1st Del CRC Dont Want"
    rt = 1
    
    Kill g3$ + b1$
    aga = aga + 1
    Label14.Caption = Str(aga)
End If


'Shell "C:\Program Files\IceView\IceView.exe " + h3$ + b12$, vbNormalFocus
'Shell "C:\Program Files\IceView\IceView.exe " + g3$ + b1$, vbNormalFocus
    
    If g1$ = h1$ Then
        aga = aga + 1
        Label14.Caption = Str(aga)
        If g2 > h2 Then xxcrc = 1
        If g2 < h2 Then xxcrc = 2
        If xxcrc = 1 Then Kill g3$ + b1$
        If xxcrc = 2 Then Kill h3$ + b12$
    End If
    
    End If
    
    ListView1.SelectedItem = ListView1.ListItems(we)
    ListView1.SelectedItem.EnsureVisible
    Label13.Caption = Str(we)

Next



Next




MsgBox "End"

Exit Sub

End

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
        
        'cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
        End
    End If

frmMain.Refresh

'End
'If Command$ <> "" Then End
cmdScan.Caption = "Scan"
End Sub



Private Sub Form_Load()
    
frmMain.Caption = frmMain.Caption + App.EXEName
'scanpath.Caption = frmMain.Caption + App.EXEName
  
    
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
        .ColumnHeaders.Add , "SIZE", "Size", 0, lvwColumnRight
        .ColumnHeaders.Add , "DATE", "Date", 0, lvwColumnLeft
        
        .View = lvwReport
    End With

frmMain.Show

Call Bangers

Exit Sub
    fg1 = FreeFile
    Open App.Path + "\Scan Path.txt" For Input As #fg1
    Line Input #fg1, ae$
    Close #fg1
    txtPath.Text = ae$
'Command$ = "D:\My Webs\MatthewLan.Com Web\Photos\galleries"
'If Command$ <> "" Then
'txtPath.Text = Command$
'On Local Error Resume Next
'ChDir Command$
'If Err.Number > 0 Then MsgBox ("Error with Dir Existance"): End

'endif

Call cmdScan_Click

End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_Click()

Dim fs22
Set fs22 = CreateObject("Scripting.FileSystemObject")


tpath3$ = "E:\04 Music ---\Del\"

b1$ = Mid$(List1.List(List1.ListIndex), 14)
c1$ = b1$ + "--Todel"

'tpath1$ = "E:\04 Music ---\"

c1$ = Mid$(b1$, InStrRev(b1$, "\") + 1)
List1.RemoveItem (List1.ListIndex)
'Name b1$ As c1$

'fs22.copyFile b1$, tpath3$ + c1$
If Dir$(tpath3$ + c1$) <> "" Then Kill tpath3$ + c1$
fs22.moveFile b1$, tpath3$ + c1$



End Sub

Private Sub SP_DirMatch(Directory As String, Path As String)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################
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
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
    c1$ = Drived2$ + "Temp\Anything" + v1$ + "\" + Mid$(a1$, ets1 + 2)

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
    fs22.moveFile a1$ + b1$, c1$ + b1$
    On Local Error GoTo 0

    If errs2 <> 0 Then
        MsgBox ("Error Moving File to Temp")
        End
    End If

    List1.AddItem Format$(we, "000 ") + "Move > " + a1$ + b1$
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
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
    c1$ = Drived2$ + "My Music" + v1$ + "\" + Mid$(a1$, ets1 + 2)

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
            fs22.deletefolder d2$, True
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

    List1.AddItem Format$(we, "000 ") + "Move > " + a1$ + b1$
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
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
    c1$ = Drived2$ + "My Music" + v1$ + "\" + Mid$(a1$, ets1 + 2)
    c2$ = Drived2$ + "Temp\Anything" + v1$ + "\" + Mid$(a1$, ets1 + 2)

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
    Set fs22 = CreateObject("Scripting.FileSystemObject")
    fs22.moveFile c2$ + b1$, a1$ + b1$
    'err.number
    'err.description
    On Local Error GoTo 0

    'If errs2 <> 0 Then
    '    MsgBox ("Error Moving File to Temp")
    '    End
    'End If

    List1.AddItem Format$(we, "000 ") + "Move > " + a1$ + b1$
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

