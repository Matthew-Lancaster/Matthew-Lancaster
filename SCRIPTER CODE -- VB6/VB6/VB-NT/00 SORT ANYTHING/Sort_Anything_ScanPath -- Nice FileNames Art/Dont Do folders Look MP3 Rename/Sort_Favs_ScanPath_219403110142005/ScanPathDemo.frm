VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ScanPath 
   AutoRedraw      =   -1  'True
   Caption         =   "ScanPath 2.0 - Demo of Class Object"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   11355
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit Operation"
      Default         =   -1  'True
      Height          =   825
      Left            =   9870
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   720
      Width           =   1485
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   585
      Top             =   4545
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   30
      TabIndex        =   41
      Top             =   5940
      Width           =   11295
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Height          =   825
      Left            =   6750
      Picture         =   "ScanPathDemo.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   2115
      Left            =   2130
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      Text            =   "ScanPathDemo.frx":15D6
      Top             =   3615
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
      ScaleWidth      =   11295
      TabIndex        =   22
      Top             =   0
      Width           =   11355
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
      Height          =   2685
      Left            =   30
      TabIndex        =   21
      Top             =   3210
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4736
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
      Left            =   7980
      Picture         =   "ScanPathDemo.frx":170E
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
      Format          =   20447233
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
      Format          =   20447233
      CurrentDate     =   37296
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
      Left            =   9450
      TabIndex        =   31
      Top             =   1620
      Width           =   1875
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim op As SHFILEOPSTRUCT


Dim ra(10)
ra(1) = "E:\Art\AutoPix\AutoPix001"
'ra(1) = "E:\Art\"

'ra(1) = "E:\04 Music ---\02 My Music Zen Removed\"
ra(2) = "E:\04 Music ---\03 My Music Zen\"
ra(3) = "E:\04 Music ---\04 My Music\"
ra(4) = "E:\04 Music ---\04 My Music Talking Books\"

'Does them all not like IdTag

ra(4) = "E:\04 Music ---\04 My Music\01 Banging Tunes\"
ra(5) = "E:\04 Music ---\04 My Music\02 Chart Songs\00 IDTag-Able\"
ra(6) = "E:\04 Music ---\04 My Music\Audio Recordings\"
ra(7) = "E:\04 Music ---\04 My Music\Comedy Mp3's\"
ra(8) = "E:\04 Music ---\04 My Music\Radio\"
ra(9) = "E:\04 Music ---\04 My Music\Sterns - Fantazia\"
'ra(10) = "E:\04 Music ---\04 My Music Talking Books\"
ra(10) = "D:\# MY DOCS\# 01 My Documents\My Music\Zen\"
ra(1) = "F:\04 Music ---\04 My Music Talking Books\_Children\"

ra(1) = "E:\05 Media\MIDI"
ra(1) = "M:\01 DVD Decrypter"



For rtu = 1 To 1

If Right$(ra(rtu), 1) <> "\" Then ra(rtu) = ra(rtu) + "\"
    
    
'Load ScanPath
ScanPath.txtPath = ra(rtu)
'ScanPath.txtPath = "E:\04 Music ---\03 My Music Zen\"
'ScanPath.txtPath = "E:\My Music\Talking Books"
ScanPath.cboMask.Text = "*.*"

Call ScanPath.cmdScan_Click

    
    
Dim outInfo As MPEGInfo

ScanPath.List1.AddItem "Syncing MP3 Bangers .."
    ScanPath.List1.Refresh
    DoEvents


Dim InTag As ID3Tag

Dim Fs22
Set Fs22 = CreateObject("Scripting.FileSystemObject")
    
tt = InStr(4, ScanPath.txtPath, "\") + 1
a3$ = Mid$(ScanPath.txtPath, tt)



For we = 1 To ScanPath.ListView1.ListItems.Count

    If we > ScanPath.ListView1.ListItems.Count Then Exit For
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ScanPath.ListView1.ListItems.Item(we)
    
    'ate7$ = "E:\ART"
 '   ate7$ = "E:\04 Music ---\04 My Music\01 Banging Tunes"
    
    'ate8$ = "E:\04 Music ---\03 My Music Zen\"
    
    tu$ = LCase(b1$)
    For r = 2 To Len(b1$)
        XC = 0
        If InStr(" _-{[]}!'@;#~><,\/1234567890", Mid$(b1$, r - 1, 1)) > 0 Then XC = 1
        If Mid$(b1$, r - 1, 1) = "_" Then XC = 1
        If XC = 1 Then Mid$(tu$, r, 1) = UCase(Mid$(tu$, r, 1))
    Next
    Mid$(tu$, 1, 1) = UCase(Mid$(tu$, 1, 1))
    
    If tu$ <> b1$ Then
        On Local Error Resume Next
        
'rr = InStrRev(tu$, ".")

'only art
'Mid$(tu$, InStrRev(tu$, ".")) = ".jpg"

        Name a1$ + b1$ As a1$ + tu$
        On Local Error GoTo 0
    End If
        tx$ = tu$
    
    
    
    
    io = 0
    
    tu$ = a1$
    ynt = InStrRev(tu$, "\")
    For r = ynt To Len(tu$)
        Mid$(tu$, ynt, 1) = LCase(Mid$(tu$, ynt, 1))
    Next
    
    For r = Len(a1$) To 2 Step -1
        XC = 0
        If InStr(" _-{[]}!'@;#~><,\/1234567890", Mid$(a1$, r - 1, 1)) > 0 Then
            XC = 1
        End If
        If XC = 1 Then Mid$(tu$, r, 1) = UCase(Mid$(tu$, r, 1))
        If Mid$(a1$, r - 1, 1) = "\" Then Exit For
    Next
    If tu$ <> a1$ Then

For r = 1 To Len(a1$)
    If Mid$(a1$, r, 1) <> Mid$(tu$, r, 1) Then
        io = r
        Exit For
    End If
Next

a12$ = Mid$(a1$, 1, Len(a1$) - 1)
tu2$ = Mid$(tu$, 1, Len(tu$) - 1)

'dir bit
If tu2$ <> a12$ Then

'Dim op As SHFILEOPSTRUCT
With op
    .wFunc = FO_RENAME ' Set function
    .pTo = tu2$ ' Set new path
    .pFrom = a12$ ' Set current path
    .fFlags = FOF_SIMPLEPROGRESS

End With
' Perform operation
On Local Error Resume Next
Err.Clear
'SHFileOperation op

Set f = Fs22.GetFolder(a12$)
ynt = InStrRev(tu2$, "\")

f.Name = Mid$(tu2$, ynt + 1)
'err.number
'err.description

Err.Clear

Name a12$ As tu2$
'err.number
'err.description
On Local Error GoTo 0



End If









'dir bit
'GoTo jump:
th = InStr(io, xx2$, "\")
If th < io And th > 0 And 1 = 2 Then



xx1$ = Left$(a1$, Len(a1$) - 1)
xx2$ = Left$(tu$, Len(tu$) - 1)

tg1 = Len(ra(rtu)) + 1
If tg1 < io Then tg1 = io
Do
th = InStr(tg1, xx2$, "\")
tg1 = th

xxt2$ = Mid$(xx2$, 1, tg1 - 1)
xxt1$ = Mid$(xx1$, 1, tg1 - 1)


'Name a1$ + b1$ As a1$ + tu$
    
'a = a
'Dim op As SHFILEOPSTRUCT
With op
    .wFunc = FO_RENAME ' Set function
    .pTo = xxt2$ ' Set new path
    .pFrom = xxt1$ ' Set current path
    .fFlags = FOF_SIMPLEPROGRESS

End With
' Perform operation
On Local Error Resume Next
SHFileOperation op
On Local Error GoTo 0

  Loop Until th = 0
    
    End If
    
End If
    
jump:
    
    On Local Error Resume Next
    ScanPath.List1.AddItem Trim(Str(we)) + "--" + a1$ + "---:---" + b1$ + "---:---" + tx$
    ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1

    If Err.Number > 0 Then
    ScanPath.List1.Clear
    End If
    
'    ScanPath.List1.Refresh
    On Local Error GoTo 0
    
    DoEvents
    
    
    
    

If Power = True Then Exit For

Next
    
If Power = True Then ScanPath.List1.AddItem "Aborted......."
ScanPath.List1.AddItem "Completed......."
ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1
ScanPath.List1.Refresh


'Call Jeepers2

Next

Endo = 1

End
Exit Sub

Stop
End

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
If Endo = 1 Then End
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

'frmMain.Refresh


'Call Jeepers



End Sub





Private Sub Form_Load()
    
ScanPath.Caption = ScanPath.Caption + App.EXEName
'frmMain.Caption = frmMain.Caption + App.EXEName
    
    
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

Power = False
cboSizeType(0).ListIndex = 0
cboSize.ListIndex = 2
Call cboSize_Click
txtSize(0).Text = "1"

'fg1 = FreeFile
'Open App.Path + "\Scan Path.txt" For Input As #fg1
'Line Input #fg1, ae$
'Close #fg1
'txtPath.Text = ae$
'txtPath.Text = App.Path

ScanPath.Show

DoEvents

Call Jeepers

End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub SP_DirMatch(Directory As String, Path As String)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################
End Sub

Private Sub SP_FileMatch(FileName As String, Path As String)
    Dim LV As ListItem
    Dim uWIN32 As WIN32_FIND_DATA
    
    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
    On Local Error GoTo GetFileError
    
    With ListView1
        Set LV = .ListItems.Add(, , FileName)
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



Sub Load_Mp3_Tag()

Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")


Dim strDrive As String
Dim strMessage As String
Dim intCnt As Integer


For intCnt = 90 To 67 Step -1
    strDrive = Chr(intCnt)


    Select Case GetDriveType(strDrive + ":\")
        Case DRIVE_REMOVABLE
        rtn = "Floppy Drive"
        Case DRIVE_FIXED
        rtn = "Hard Drive"
        Case DRIVE_REMOTE
        rtn = "Network Drive"
        Case DRIVE_CDROM
        rtn = "CD-ROM Drive"
        Case DRIVE_RAMDISK
        rtn = "RAM Disk"
        Case Else
        rtn = ""
    End Select

    If rtn = "CD-ROM Drive" Then
    On Local Error Resume Next
    e$ = Dir$(strDrive + ":\*.*")
    Checkeddrv = Checkeddrv + 1
    If Err.Number > 0 And Checkeddrv = 2 Then MsgBox ("No Disk in Last Cd Rom Drive " + strDrive + ":\"): End
    If Err.Number = 0 Then Exit For
    End If

'    If rtn <> "" Then
'       strMessage = strMessage & vbCrLf & "Drive " & strDrive & " is type: " & rtn
'    End If
Next intCnt





Load ScanPath
ScanPath.txtPath = "E:\04 Music ---\03 My Music Zen\"
'ScanPath.txtPath = "E:\My Music\Talking Books"
ScanPath.cboMask.Text = "*.mp3"

Call ScanPath.cmdScan_Click


books = Val(ScanPath.lblCount.Caption)
List1.AddItem ("BookMarks = " + Str$(books))

freef1 = FreeFile
Open "C:\# MY DOCS\CD-ROM\Kmp3.txt" For Output As #freef1

If Dir(strDrive + ":\Z_CallerId\*.*") <> "" Then alps$ = "Z_CallerId": kil = 1
If Dir(strDrive + ":\zzCallerId\*.*") <> "" Then alps$ = "zzCallerId": kil = 2
If Dir(strDrive + ":\0Rm35\*.*") <> "" Then alps$ = "Z_CallerId": kil = 3
If Dir(strDrive + ":\0Rm34\*.*") <> "" Then alps$ = "Z_CallerId": kil = 3
If Dir(strDrive + ":\0Rm33\*.*") <> "" Then alps$ = "Z_CallerId": kil = 3

'len(alpS$)

For we = 1 To ScanPath.ListView1.ListItems.Count
    a1$ = ScanPath.ListView1.ListItems.Item(we)
    b1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)


    'c1$ = Mid$(b1$, Len(ScanPath.txtPath) + 2) + a1$

'    Call Mp3_Tag.smdpick(b1$ + a1$)
    Dim outInfo As MPEGInfo

    mp32 = 0
    If InStr(UCase$(a1$), ".MP3") Then mp32 = 1
    
    outInfo.Length = 0
    outInfo.Bitrate = 0
    
    
    If mp32 = 1 Then
        mdlID3.ReadMPEGInfo b1$ + a1$, outInfo
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

    
    leng$ = """;""" + Trim(Str$(outInfo.Length \ 60)) & ":" & Format$(outInfo.Length Mod 60, "00") + """"
    kbps$ = """;""" + Trim(Str$(outInfo.Bitrate)) + " Kbps"
    c1$ = """" + b1$ + a1$
    
    
    xyz = 0
    If InStr(UCase$(a1$), ".MP3") Then xyz = 1
    
    If xyz = 0 And InStr(b1$, alps$ + "\Apps") And kil = 1 Then
        
        qwd = Len(alps$ + "\Apps") + 5
        d1$ = Mid$(b1$, 1, 3) + "Applications-\" + Mid$(b1$, qwd, InStr(qwd, b1$, "\") - qwd)
        

        'If lastpath$ = "" Then lastpath$ = d1$
        If lastpath$ <> d1$ Then
            leng$ = """;""" + """"
            Set f = fs.GetFolder(b1$)
            s22 = f.Size
            kbps$ = """;""" + Format$(s22 / (1048576), "0.0") + " MByte"
            c1$ = """" + d1$
            xyz = 1
            lastpath$ = d1$
        End If
    End If
    
    If xyz = 0 And InStr(b1$, alps$) And (kil = 2 Or kil = 3) Then
        
        qwd = Len(alps$) + 3
        If kil = 3 Then qwd = qwd + 1
        
        If InStr(qwd + 2, b1$, "\") Then
        d1$ = Mid$(b1$, 1, 3) + "Applications-" + Mid$(b1$, qwd, InStr(qwd + 2, b1$, "\") - qwd)
        End If

        'If lastpath$ = "" Then lastpath$ = d1$
        kit = 0
        If InStr(b1$, "Andy Lett") Then kit = 1
        If InStr(b1$, "My Web") Then kit = 1
        
        If lastpath$ <> d1$ And kit = 0 Then
            leng$ = """;""" + """"
            Set f = fs.GetFolder(b1$)
            s22 = f.Size
            kbps$ = """;""" + Format$(s22 / (1048576), "0.0") + " MByte"
            c1$ = """" + d1$
            xyz = 1
            lastpath$ = d1$
        End If
    End If
    
    
    
    
    
    
    
    DoEvents
    
    If InStr(b1$, Mid$(b1$, 1, 3) + alps$ + "\My Web\") Then webc = webc + 1
    Qe2$ = UCase$(Mid$(c1$, Len(c1$) - 10))
    
    
    kj = 0
    If InStrRev(Qe2$, ".EXE") Then exec = exec + 1: kj = 1
    If InStrRev(Qe2$, ".MP3") Then mp3c = mp3c + 1: kj = 1
    If InStrRev(Qe2$, ".TXT") Then txtc = txtc + 1: kj = 1
    If InStrRev(Qe2$, ".JPG") Then jpgc = jpgc + 1: kj = 1
    If InStrRev(Qe2$, ".html") Then htmc = htmc + 1: kj = 1
    If InStrRev(Qe2$, ".html") Then htmc = htmc + 1: kj = 1
    
    qe3$ = UCase$(Mid$(c1$, Len(c1$) - 3))
    wg = InStr(qe4$, qe3$)
    If wg = 0 And kj = 0 Then qe4$ = qe4$ + qe3$
    
    
    If xyz = 1 And Not (outInfo.Length = 0 And mp32 = 1) Then
        List1.AddItem (c1$ + kbps$ + leng$)
        List1.ListIndex = List1.ListCount - 1
        Print #freef1, c1$ + kbps$ + leng$
    End If
Next

others$ = Trim$(Str$(Len(qe4$) / 4))
Set f = fs.GetFolder(Mid$(b1$, 1, 3) + alps$ + "\My Web\")
s22 = f.Size / 1048576
wew$ = Format$(s22, "0.0") + " MByte"
wewc$ = Trim$(Str(webc)) + " File Count"

List1.AddItem """Total Taken By Matt's Web http://MatthewLan.com = " + wewc$ + """;""" + wew$ + """;""" + """"
Print #freef1, """Total Taken By Matt's Web http://MatthewLan.com = " + wewc$ + """;""" + wew$ + """;""" + """"

List1.AddItem """Total File's On Disk =" + Str$(books) + """;""" + """;""" + """"
List1.AddItem """Total File's *.EXE=" + Str$(exec) + " *.MP3=" + Str$(mp3c) + " *.TXT=" + Str$(txtc) + " *.JPG=" + Str$(jpgc) + " *.html=" + Str$(htmc) + """;""" + """;""" + """"
List1.AddItem """Total File's Other=" + Str$(books - exec - mp3c - txtc - jpgc - htmc) + " Of " + others$ + " Types of Other Extensions" + """;""" + """;""" + """"

Print #freef1, """Total File's On Disk =" + Str$(books) + """;""" + """;""" + """"
Print #freef1, """Total File's *.EXE=" + Str$(exec) + " *.MP3=" + Str$(mp3c) + " *.TXT=" + Str$(txtc) + " *.JPG=" + Str$(jpgc) + " *.html=" + Str$(htmc) + """;""" + """;""" + """"
Print #freef1, """Total File's Other=" + Str$(books - exec - mp3c - txtc - jpgc - htmc) + " Of " + others$ + " Types of Other Extensions" + """;""" + """;""" + """"

Set f = fs.GetFolder(Mid$(b1$, 1, 3))
s22 = f.Size

List1.AddItem """Total Size of Disk = " + Format$(s22 / (1048576), "0.0") + " Mega Bytes" + """;""" + """;""" + """"
Print #freef1, """Total Size of Disk = " + Format$(s22 / (1048576), "0.0") + " Mega Bytes" + """;""" + """;""" + """"
Close #freef1

List1.AddItem "--------Finished----------"
List1.ListIndex = List1.ListCount - 1

'Timer1.Enabled = True

End Sub

