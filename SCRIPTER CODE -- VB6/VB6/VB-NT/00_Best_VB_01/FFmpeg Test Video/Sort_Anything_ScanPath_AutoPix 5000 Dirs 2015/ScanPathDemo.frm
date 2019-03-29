VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ScanPath 
   Caption         =   "ScanPath 2.0 - Anything -- "
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12630
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   12630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   585
      Top             =   4545
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   60
      TabIndex        =   41
      Top             =   6000
      Width           =   11265
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Height          =   825
      Left            =   11595
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   885
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   2115
      Left            =   2100
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      Text            =   "ScanPathDemo.frx":1194
      Top             =   3885
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
      Top             =   2850
      Width           =   3465
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2745
      Left            =   45
      TabIndex        =   21
      Top             =   3210
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4842
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
      Default         =   -1  'True
      Height          =   825
      Left            =   10575
      Picture         =   "ScanPathDemo.frx":12CC
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   720
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
      Format          =   23461889
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
      Format          =   23461889
      CurrentDate     =   37296
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Counted"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6720
      TabIndex        =   44
      Top             =   780
      Width           =   1245
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Counting"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6720
      TabIndex        =   43
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label lblCount2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7980
      TabIndex        =   42
      Top             =   1080
      Width           =   1620
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
   Begin VB.Label lblCount1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
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
      Height          =   285
      Left            =   7980
      TabIndex        =   31
      Top             =   780
      Width           =   1620
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim G(20)
Public GTf
Public FS

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal _
   dwMilliseconds As Long)


'FFmpeg_Test_Video
'AutoPix
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




Sub FFmpeg_Test_Video()



ScanPath.Show

ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask.Text = "*.mpeg;*.mpg;*.avi;*.mp4;.wmv;*.flc;*.mov;*.qt;*.3gp;*.asf;*.m4v;*.rm;*.vob"
ScanPath.chkSubFolders = vbChecked


DirVar = "M:\0 00 VIDEO\"

ScanPath.txtPath = DirVar
Call cmdScan_Click
ScanPath.txtPath = "U:\0 00 VIDEO\"
Call cmdScan_Click
ScanPath.txtPath = "U:\00 0 VIDEO CAMERA'S\"
Call cmdScan_Click

ScanPath.txtPath = "I:\0 00 VIDEO\"
Call cmdScan_Click
ScanPath.txtPath = "I:\00 0 VIDEO CAMERA'S\"
Call cmdScan_Click

ScanPath.txtPath = "D:\0 00 Art\# 00 My Pictures\"
Call cmdScan_Click
ScanPath.txtPath = "M:\0 00 Art\# 00 My Pictures\"
Call cmdScan_Click
ScanPath.txtPath = "U:\0 00 Art\# 00 My Pictures\"
Call cmdScan_Click
ScanPath.txtPath = "I:\0 00 Art\# 00 My Pictures\"
Call cmdScan_Click

ScanPath.txtPath = "D:\Videos\"
Call cmdScan_Click
ScanPath.txtPath = "M:\Videos\"
Call cmdScan_Click
ScanPath.txtPath = "U:\Videos\"
Call cmdScan_Click
ScanPath.txtPath = "I:\Videos\"
Call cmdScan_Click


'sort of PATH then FILE
ListView1.SortOrder = lvwAscending
ListView1.SortKey = 0
ListView1.Sorted = True
ListView1.SortKey = 1
ListView1.Sorted = True
ListView1.Sorted = False

xtnow = Now

'10469 Video's 2015

lblCount1.Caption = Str(ListView1.ListItems.Count)
totalvideos = ListView1.ListItems.Count

For we = ListView1.ListItems.Count To 1 Step -1
    
    DoEvents
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)
    lblCount1.Caption = Str(ListView1.ListItems.Count - we)
    
    
    
        xf1 = a1$ + b1$ + ".txt"
        xf2 = Replace(xf1, ".txt", ".FFMpeg-Verify.txt")
        xf3 = Replace(xf1, ".txt", ".FFmpeg-Verify.txt")
    
    If FS.FileExists(xf1) = True Or FS.FileExists(xf2) Then
        
        
        If FS.FileExists(xf1) = True Then Name xf1 As xf3
        If a1$ + Dir(xf2) <> xf3 Then Name xf2 As xf3
        
        ListView1.ListItems.Remove (we)

    
    End If
Next

lblCount1.Caption = Str(ListView1.ListItems.Count)
d1$ = "": weCount = 0

For we = 1 To ListView1.ListItems.Count
    
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)
    lblCount1.Caption = Str(ListView1.ListItems.Count - we)

    xf1 = a1$ + b1$ + ".txt"
    xf2 = Replace(xf1, ".txt", ".FFmpeg-Verify.txt")
    
    'If FS.FileExists(xf1) = True Or FS.FileExists(xf2) Then
    If FS.FileExists(xf2) Then
        
        'If FS.FileExists(xf1) = True Then Name xf1 As xf2
        'Set f = FS.getfile(A1$ + B1$ + ".txt")
        'If f.Size = 0 Then Kill xf2
    Else
    'xf2 = Replace(xf1, ".txt", ".FFMpeg-Verify.txt")
        'For Batchfile - and Think Shell Cmd and My New Batchfile Method
        d1$ = d1$ + "REM ---" + Str(we) + " of" + Str(ListView1.ListItems.Count) + vbCrLf
        d1$ = d1$ + "TIME /T" + vbCrLf
        d1$ = d1$ + "REM --------------" + vbCrLf
        d1$ = d1$ + """C:\Program Files\# NO INSTALL REQUIRED\ffmpeg-20150701-git-9c010ba-win32-static\ffmpeg.exe"" -v error -i """ + a1$ + b1$ + """ -f null - >""" + xf2 + """ 2>&1" + vbCrLf
        
        weCount = weCount + 1
        lblCount2.Caption = Str(weCount)
        
    End If
    DoEvents
    'For Shell
    'd1$ = "C:\Program Files\# NO INSTALL REQUIRED\ffmpeg-20150701-git-9c010ba-win32-static\ffmpeg.exe -v error -i """ + A1$ + B1$ + """ -f null - >""" + A1$ + B1$ + ".txt""" + " 2>&1"
    
    
    'For Batchfile from shell
    'd1$ = "'C:\Program Files\# NO INSTALL REQUIRED\ffmpeg-20150701-git-9c010ba-win32-static\ffmpeg.exe' -v error -i """ + A1$ + B1$ + """ -f null - >""" + A1$ + B1$ + ".txt" + """ 2>&1"
    
    'Clipboard.Clear
    'Clipboard.SetText (d1$)
    
'    Do
'        If FindWindow("ConsoleWindowClass", vbNullString) = False Then Exit Do
'        'Sleep 1000
'        DoEvents
'    Loop Until 1 = 2
    
    'Shell "Cmd /E:ON /s /k " + d1$
'    ivar = Shell(d1$, vbNormalFocus)
    'Shell d1$, vbNormalFocus
    

Next

lblCount1.Caption = Str(ListView1.ListItems.Count)
'fr1 = FreeFile
'Open FileBatch For Append As #fr1

FILEBATCH = "D:\temp\FFmpeg Video Test -- " + Format(Now, "YYYY-MM-DD HH-MM-SS") + ".bat"
fr1 = FreeFile
Open FILEBATCH For Output As #fr1

Print #fr1, "Rem --- " + FILEBATCH
Print #fr1, "Rem --- "
Print #fr1, "Rem --- " + Format$(Now, "YYYY-MM-DD HH:MM:SS") + " -- Batch Creation Date Time --"
Print #fr1, "Rem --- " + Trim(Str(totalvideos)) + " Total Scanned on My Computer Files --"
Print #fr1, "Rem --- " + Trim(Str(ListView1.ListItems.Count)) + " Total Video Files --"
Print #fr1, "Rem --- "
Print #fr1, "Rem --- " + Trim(Str(weCount)) + " File Count Job to Do After Filter --"
Print #fr1, "Rem --- "

Print #fr1, d1$

Close #fr1


Shell "cmd /k """ + FILEBATCH + """"

Sleep 9000

Shell "C:\Program Files\Notepad++\notepad++.exe """ + FILEBATCH + """", vbNormalFocus

'Wait for Active
Do
    If FindWindow("ConsoleWindowClass", "C:\WINDOWS\system32\cmd.exe - """ + FILEBATCH + """") > 0 Then Exit Do
    Sleep 1000
    DoEvents
Loop Until 1 = 2

'Quit when Finished

Do
'    If 1 = 1 Then Exit Do
    If FindWindow("ConsoleWindowClass", "C:\WINDOWS\system32\cmd.exe - """ + FILEBATCH + """") = 0 Then Exit Do
    Sleep 1000
    DoEvents
Loop Until 1 = 2


fr1 = FreeFile
FILELOGNAME = DirVar + "FFmpeg Error Log Script -- " + Format(Now, "YYYY-MM-DD HH-MM-SS") + ".txt"
Open FILELOGNAME For Output As #fr1
    
Print #fr1, "Error's in These Files"
Print #fr1, "---------------------"

Print #fr1, Trim(Str(totalvideos)) + " -- Total Video's on My Computer"
Print #fr1, "---------------------"
Print #fr1, Trim(Str(ListView1.ListItems.Count)) + " -- Job Count Files"

Print #fr1, "---------------------"
'Print #fr1, Trim(Str(weCount)) + " Files to be Done After Naught Size Filter --"
Print #fr1, Trim(Str(weCount)) + " Files to be Done --"

Print #fr1, Format$(xtnow, "YYYY-MM-DD HH:MM:SS") + " -- Start Job"
Print #fr1, "---------------------"
Print #fr1, Format$(Now, "YYYY-MM-DD HH:MM:SS") + " -- Now Time - End Job Clean Up & Completion"
Print #fr1, "---------------------"
tagcount = 0

For we = 1 To ListView1.ListItems.Count
    
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)
    lblCount1.Caption = Str(ListView1.ListItems.Count - we)
    xf1 = a1$ + b1$ + ".txt"
    xf2 = Replace(xf1, ".txt", ".FFmpeg-Verify.txt")
    
    If FS.FileExists(xf2) = True Then
        Set f = FS.getfile(xf2)
        If f.Size = 0 Then
            'Kill A1$ + B1$ + ".txt"
        Else
            tagcount = tagcount + 1
            lblCount2.Caption = Str(tagcount)
            Print #fr1, a1$ + b1$
        End If
    End If
Next

Print #fr1, "---------------------"
Print #fr1, Trim(Str(tagcount)) + " -- Error File Count After Batch Run -- "
Print #fr1, "---------------------"

Close #fr1

Shell "C:\Program Files\Notepad++\notepad++.exe """ + FILELOGNAME + """", vbNormalFocus

End

End Sub



Sub AutoPix()


ScanPath.Show



cboMask.Text = "*.jpg"

G(1) = "S:\Art\AutoPix\AutoPix zz 000-000\"
G(1) = "S:\00 Art\AutoPix\AutoPix001\"
G(1) = ""

G(1) = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix 2004 to 6"
G(2) = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix 2007 03"
G(3) = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix 2007 04"
G(3) = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix zz 001\AutoPix001"
G(3) = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix zz 000-000\AutoPix (Main)"

G(2) = "M:\DSC\2011\1ST HALF"
G(3) = "M:\DSC\2011"
GTf = 3


'g(1) = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix 2007 05 27K Files"

For r2 = GTf To GTf


If Mid$(G(r2), Len(G(r2)), 1) <> "\" Then
    G(r2) = G(r2) + "\"
End If

Next

Call RmdirAutoPix


For r2 = GTf To GTf

txtPath.Text = G(r2)


Call cmdScan_Click

For we = ListView1.ListItems.Count To 1 Step -1
    
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)
'    If InStr(a1$, "0-Plus 800K") > 0 Then
'        ListView1.ListItems.Remove (we)
'    End If
    If InStr(a1$, "# 00 DOC") > 0 Then
        ListView1.ListItems.Remove (we)
    End If
    If InStr(a1$, "\DOC") > 0 Then
        ListView1.ListItems.Remove (we)
    End If
Next
'0-Plus 800K

'ListView1.SortKey = ListView.ColumnHeader - 1

ListView1.SortOrder = Index
ListView1.Sorted = True ' Sort the List.


'Exit Sub


'MsgBox "Ready to Go"
Dim Fs22


bignum = 5000
bignum = 2000
bignum = 1000


Set Fs22 = CreateObject("Scripting.FileSystemObject")

'd1$ = "A-02000\"
d1$ = ""
yy = bignum
tt5 = yy
ttw = yy + 1
    
    Set Fs22 = CreateObject("Scripting.FileSystemObject")

For we = 1 To ListView1.ListItems.Count
    
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)

    If we = ttw Then
        yy = yy + bignum
        ttw = ttw + bignum
    End If

    'Put back
    d1$ = "A-" + Format$(yy, "00000") + "\"

    DoEvents

    'Set Fs22 = Fs22.Getfile(a1$ + b1$)
    

    On Local Error Resume Next
    MkDir G(r2) + d1$
    On Local Error GoTo 0


    d2$ = G(r2) + d1$

    Err.Clear
    On Local Error Resume Next
    
    
    
    g1$ = b1$
    If Dir$(d2$ + b1$) <> "" Then
    g1$ = Mid$(b1$, 1, (Len(b1$) - 4)) + "_" + Trim(Str(we)) + ".jpg"
    End If
    

       
    Err.Clear
                
    If a1$ <> d2$ Then Fs22.moveFile a1$ + b1$, d2$ + b1$
        
    b4$ = ""
    countt = 0
        
    kk = 0
    rg = 0
    Do
        If Err.Number > 0 Or rg > 0 Then
            If a1$ <> d2$ Or rg > 0 Then
                countt = countt + 1
                b3$ = "--" + Format$(countt, "000")
                Err.Clear
                rg = InStrRev(b1$, ".")
                rg = InStr(b1$, "--00")
                rx = InStr(b1$, "-001-")
                If rg = 0 Then rg = InStrRev(b1$, ".")
                b4$ = Mid$(b1$, 1, rg - 1) + b3$ + ".jpg"
        
 '               Name a1$ + b1$ As a1$ + b4$
                kk = 1
                b1$ = b4$
                'If Err.Number > 0 Then MsgBox " error": Stop
                'err.description
                If Err.Number > 0 Then
                    'Stop
                End If
        
                If Err.Number = 0 Then
 '                   Fs22.moveFile a1$ + b4$, d2$ + b4$
                    Fs22.moveFile a1$ + b1$, d2$ + b1$
                End If
        
                If Err.Number = 53 Then Err.Clear
        
            End If
        End If
    Loop Until Err.Number = 0
    
    'If Err.Number <> 58 And Err.Number > 0 Then MsgBox Str$(Err.Number) + " #" + Err.Description: Stop
    If Err.Number <> 0 And Err.Number > 0 Then MsgBox Str$(Err.Number) + " #" + Err.Description: Stop
        
    On Local Error GoTo 0
       If b4$ <> "" Then b1$ = b4$
    Label13 = we
    If a1$ <> d2$ Or kk = 1 Then
        If kk = 0 Then
            List1.AddItem Format$(we, "000 ") + "Move > " + d2$ + b1$ + "--- " + a1$
        Else
            List1.AddItem Format$(we, "000 ") + "Rename > " + d2$ + b1$ + "--- " + a1$
        End If
        On Local Error Resume Next
        List1.ListIndex = List1.ListCount - 1
    End If

    If Err.Number > 0 Then
        List1.Clear
    End If
    On Local Error GoTo 0

Next



Next

Reset

Set SP = Nothing

'Clear
End

End Sub


Sub RmdirAutoPix()

Reset

For r2 = GTf To GTf
For r = 0 To 70000 Step 500

d1$ = "A-" + Format$(r, "00000") '+ "\"
On Error Resume Next
RmDir G(r2) + d1$

'err.description

Next
Next



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
    txtHelp.Visible = Not txtHelp.Visible
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
        End
    End If
   
 cmdScan.Caption = "Scan"
 
'scanpath.Refresh

'End
'If Command$ <> "" Then End

End Sub



Private Sub Form_Load()

Set FS = CreateObject("Scripting.FileSystemObject")
    
ScanPath.Caption = ScanPath.Caption + App.EXEName
    
ScanPath.Height = List1.Top + List1.Height + 450
ScanPath.Width = List1.Left + List1.Width + 150
    
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

'Call AutoPix
Call FFmpeg_Test_Video

Exit Sub
    fg1 = FreeFile
    Open App.Path + "\Scan Path.txt" For Input As #fg1
    Line Input #fg1, ae$
    Close #fg1
    txtPath.Text = ae$
'Command$ = "D:\My Webs\MatthewLan.Com Web\Photos\galleries"
If Command$ <> "" Then
txtPath.Text = Command$
On Local Error Resume Next
ChDir Command$
If Err.Number > 0 Then MsgBox ("Error with Dir Existance"): End
Call cmdScan_Click
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
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








'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------









