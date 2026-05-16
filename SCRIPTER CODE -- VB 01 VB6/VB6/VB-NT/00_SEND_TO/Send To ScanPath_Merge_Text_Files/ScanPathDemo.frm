VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ScanPath 2.0 - Demo of Class Object"
   ClientHeight    =   6936
   ClientLeft      =   48
   ClientTop       =   636
   ClientWidth     =   11388
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6936
   ScaleWidth      =   11388
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Hitt To Go"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   9465
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   690
      Width           =   1875
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   585
      Top             =   4545
   End
   Begin VB.ListBox List1 
      Height          =   1584
      Left            =   30
      TabIndex        =   41
      Top             =   6930
      Visible         =   0   'False
      Width           =   11295
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Height          =   825
      Left            =   7425
      Picture         =   "ScanPathDemo.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   -300
      Visible         =   0   'False
      Width           =   1875
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
      Text            =   "ScanPathDemo.frx":1A5E
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
      Text            =   "D:\# MY DOCS\# 01 My Documents\Favorites"
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
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   6948
      TabIndex        =   22
      Top             =   0
      Width           =   7000
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single File solution to quickly add file processing to any Utility Project."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
            Size            =   7.8
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
            Size            =   7.8
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
      Height          =   3705
      Left            =   45
      TabIndex        =   21
      Top             =   3210
      Width           =   11295
      _ExtentX        =   19918
      _ExtentY        =   6541
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
      Height          =   825
      Left            =   9345
      Picture         =   "ScanPathDemo.frx":1B96
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   -345
      Visible         =   0   'False
      Width           =   1875
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   0
      Left            =   5940
      TabIndex        =   13
      Top             =   2400
      Width           =   1545
      _ExtentX        =   2731
      _ExtentY        =   550
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   107872257
      CurrentDate     =   37299
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   1
      Left            =   5940
      TabIndex        =   14
      Top             =   2760
      Width           =   1545
      _ExtentX        =   2731
      _ExtentY        =   550
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   107872257
      CurrentDate     =   37296
   End
   Begin VB.Label Label15 
      BackColor       =   &H0000FF00&
      Caption         =   "Label15"
      Height          =   165
      Left            =   7695
      TabIndex        =   45
      Top             =   1680
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " Raw Merge "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7455
      TabIndex        =   44
      Top             =   1125
      Width           =   1830
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7425
      TabIndex        =   42
      Top             =   750
      Width           =   1875
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Date/Size Filter:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Call Jeepers
Dim Tpath1$, FF$
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

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


DoEvents

'txtPath.Text = "G:\01 DVD\"
'txtPath.Text = "E:\Galidakis\"
'txtPath.Text = "D:\I\"
'txtPath.Text = "D:\VB6\"
'txtPath.Text = "D:\# MY DOCS\# 01 My Documents\My Web Sites\New Folder\"
'txtPath.Text = "D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\# DATA\"
'txtPath.Text = "C:\CallerID\Mental Health Text\Valerie\"
'txtPath.Text = "E:\ZIPS\ZIPS\Dos Programs\MS Dos 6.22 Collection\MS Dos 6 TXT DIZ NFO LOG\"
'txtPath.Text = "E:\ZIPS\ZIPS\Dos Programs\MS Dos 6.22 Collection\BlueWave"
'txtPath.Text = "E:\My Music\Talking Books"
'txtPath.Text = "M:\01 Installations\01 DVD"
'txtPath = "E:\ZIPS\ZIPS\Dos Programs\MS Dos 6.22 Collection\MS Dos 6 NFO"
'txtPath = "M:\05 Media\DSSPlayer\Message\00 My folders\2008 002"
'txtPath = "M:\05 Media\DSSPlayer\Message\00 My folders\2008 002"
'txtPath = "D:\# MY DOCS\# 01 My Documents\CALLerID\Text-9\Andy"
'txtPath = "D:\# MY DOCS\# 01 My Documents\00 Email Texts\Terrorist Files"




If Command$ <> "" Then txtPath = Command$


'If IsIDE = True Then txtPath = "D:\# MY DOCS\# 01 My Documents\00 Email Texts\Terrorist Files"

If txtPath <> "" And Mid$(txtPath, 1, 1) = """" Then
txtPath = Mid$(txtPath, 2)
txtPath = Mid$(txtPath, 1, Len(txtPath) - 1)
End If
If Len(txtPath) < 4 Then MsgBox "No Good Path": End



If Right$(txtPath.Text, 1) <> "\" Then txtPath.Text = txtPath.Text + "\"


cboMask.Text = "*.txt;*.nfo;*.jpg;*.diz;*.html,*h,*.c,*ccp,*.rc,*.dsp,*.log,*.bas;*.cls;*.frm;*.nws;*.vcf"

Call cmdScan_Click


ht$ = Mid$(txtPath.Text, InStrRev(txtPath.Text, "\", Len(txtPath.Text) - 1) + 1)
ht$ = Mid$(ht$, 1, Len(ht$) - 1)

'FF$ = "00_Scan_Merge_Alt.Philo_"
'FF$ = "00_Scan_Merge_VB6_Cls_Bas_Frm"
'FF$ = "00_Scan_Merge_matthewlan.Com.txt"
'FF$ = "00_ClipBoard_Total.txt"
'FF$ = "zz Valer-All.txt"

FF$ = "00_Scan_Merge_" + ht$ + ".txt"
If InStr(LCase(txtPath.Text), "contacts") > 0 Then
    FF$ = "..\00_Scan_Merge_" + ht$ + ".vcf"
End If

'##dir file names to text list
If Label14.BackColor <> Label15.BackColor Then
    Call Bangers
Else
    If FS.FileExists(txtPath + FF) Then Kill txtPath.Text + FF$
End If

fr1 = FreeFile
Open txtPath.Text + FF$ For Append As #fr1

For we = 1 To ListView1.ListItems.Count

DoEvents

Do
A1 = ListView1.ListItems.Item(we).SubItems(1)
B1 = ListView1.ListItems.Item(we)
xe = 0
If InStr(B1, FF$) > 0 Then
    xe = 1: we = we + 1
End If
Loop Until xe = 0

xtag = xtag + 1



Label13.Caption = Str(we)
DoEvents
fr2 = FreeFile
Open A1 + B1 For Binary As fr2
rr$ = Input$(LOF(fr2), fr2)

If Label14.BackColor <> Label15.BackColor Then
    Print #fr1, "--------------------------------------------------------------------"
    Print #fr1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS")
    Print #fr1, "--------------------------------------------------------------------"
    Print #fr1, Format$(xtag, "00000")
    Print #fr1, A1 + B1
    Print #fr1, B1
    Print #fr1, "--------------------------------------------------------------------"
    Print #fr1, rr$ + vbCrLf
    Close fr2
Else
    Print #fr1, rr$
    Close fr2
End If


Next
Close fr1



Set FS = CreateObject("Scripting.FileSystemObject")
'FS.copyFile txtPath + FF$, "D:\# MY DOCS\# 01 My Documents\00 Email Texts\Dir to Text List\" + FF$

'Shell "notepad " + txtPath.Text + FF$, vbNormalFocus
Shell "explorer /select," + txtPath.Text + FF$, vbNormalFocus


'MsgBox "Done..."

End Sub
















Sub Bangers()
'##dir file names to text list
'Bangers


uu$ = Mid$(txtPath, InStrRev(txtPath, "\", Len(txtPath) - 1))
uu$ = Mid$(uu$, 2, Len(uu$) - 2)


Open txtPath.Text + FF$ For Output As #1


Print #1, String$(70, "-")
Print #1, "Directory to Text File VB6"
Print #1, String$(70, "-")
Print #1, "Dir = " + txtPath.Text
Print #1, String$(70, "-")
Print #1, "File Count =" + Str(ListView1.ListItems.Count)
Print #1, String$(70, "-")

For we = 1 To ListView1.ListItems.Count
    A1 = ListView1.ListItems.Item(we).SubItems(1)
    B1 = ListView1.ListItems.Item(we)
    'B1 = Mid$(B1, 1, Len(B1) - 4)
    Print #1, B1
Next
Close #1

Set FS = CreateObject("Scripting.FileSystemObject")
FS.copyFile txtPath + FF$, "D:\# MY DOCS\# 01 My Documents\00 Email Texts\Dir to Text List\" + FF$


'End

Exit Sub

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
            DoEvents
        End With
        
        LockWindowUpdate 0
        
        cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
    End If

frmMain.Refresh





End Sub


Sub jeepers3()


Dim FS

Set FS = CreateObject("Scripting.FileSystemObject")

On Local Error GoTo jeep
'MkDir "R:\Temp\Favs"
Drived2$ = Mid$(txtPath.Text, 1, 3)
MkDir Drived2$ + "Temp\Favs"
On Local Error GoTo 0

List1.AddItem "Stage 1 of 3 : Make Dir's And Move Files to Temp"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For we = 1 To ListView1.ListItems.Count
A1 = ListView1.ListItems.Item(we)
B1 = ListView1.ListItems.Item(we).SubItems(1)

'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
ets1 = Len(txtPath.Text)
c1$ = Drived2$ + "Temp\Favs\" + Mid$(B1, ets1 + 2)

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
            End
        End If
        On Local Error GoTo 0
    End If
Loop Until ets2 = 0

f1$ = d2$

errs2 = 0
On Local Error GoTo jeep
'fs.Copyfile B1 + A1, c1$ + A1
FS.Movefile B1 + A1, c1$ + A1
On Local Error GoTo 0
If errs2 <> 0 Then
    MsgBox ("Error Moving File to Temp")
    End
End If


'fs.DeleteFolder comrad$ + "\" + WFD$

Next

List1.AddItem "Stage 2 of 3 : Move Files Back From Temp"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For we = 1 To ListView1.ListItems.Count
A1 = ListView1.ListItems.Item(we)
B1 = ListView1.ListItems.Item(we).SubItems(1)

'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
ets1 = Len(txtPath.Text)
'ets1 = 25

c1$ = Drived2$ + "Temp\Favs\" + Mid$(B1, ets1 + 2)

errs2 = 0
On Local Error GoTo jeep
FS.Movefile c1$ + A1, B1 + A1
On Local Error GoTo 0
If errs2 <> 0 Then
    MsgBox ("Error Moving File to Temp")
    End
End If

Next


List1.AddItem "Stage 3 of 3 : Delete Temp Folder"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

FS.deletefolder Drived2$ + "Temp\Favs\*"

List1.AddItem "Stage 3 of 3 : Complete --------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

Timer1.Enabled = True

Exit Sub
jeep:
errs2 = Err.Number
errs3$ = Err.Description
Resume Next

End Sub


Sub Jeepers2()


Dim FS

Set FS = CreateObject("Scripting.FileSystemObject")

On Local Error GoTo jeep
'MkDir "R:\Temp\Favs"
Drived2$ = Mid$(txtPath.Text, 1, 3)
MkDir Drived2$ + "Temp\Favs"
On Local Error GoTo 0

List1.AddItem "Stage 1 of 3 : Make Dir's And Move Files to Temp"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For we = 1 To ListView1.ListItems.Count
A1 = ListView1.ListItems.Item(we)
B1 = ListView1.ListItems.Item(we).SubItems(1)

'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
ets1 = Len(txtPath.Text)
c1$ = Drived2$ + "Temp\Favs\" + Mid$(B1, ets1 + 2)

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
            End
        End If
        On Local Error GoTo 0
    End If
Loop Until ets2 = 0

f1$ = d2$

errs2 = 0
On Local Error GoTo jeep
'fs.Copyfile B1 + A1, c1$ + A1
FS.Movefile B1 + A1, c1$ + A1
On Local Error GoTo 0
If errs2 <> 0 Then
    MsgBox ("Error Moving File to Temp")
    End
End If


'fs.DeleteFolder comrad$ + "\" + WFD$

Next

List1.AddItem "Stage 2 of 3 : Delete Destination Folders & Move Folder Back From Temp"
List1.ListIndex = List1.ListCount - 1
List1.Refresh


FS.deletefolder txtPath.Text + "\*"

'MkDir txtPath.Text

FS.Movefolder Drived2$ + "Temp\Favs\*", txtPath.Text


List1.AddItem "Stage 3 of 3 : Delete Temp Folder"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

FS.deletefolder Drived2$ + "Temp\Favs\*"

List1.AddItem "Stage 3 of 3 : Complete --------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

Timer1.Enabled = True

Exit Sub
jeep:
errs2 = Err.Number
errs3$ = Err.Description
Resume Next

End Sub



Private Sub Command1_Click()
Call Jeepers
End Sub

Private Sub Form_Load()
    
    Set FS = CreateObject("Scripting.FileSystemObject")
    
    
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
DoEvents

txtPath = Clipboard.GetText

End Sub


Private Sub Label14_Click()
Label14.BackColor = Label15.BackColor

End Sub


Private Sub Mnu_VB_ME_Click()

'Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'End

'    Dim MSGQ, IR, CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2
    
    On Error Resume Next
    
    Beep
    If GetSpecialfolder(38) = "" Then
        MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__<" + GetSpecialfolder(38) + ">__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
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
    
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    
    EXIT_TRUE = True
    Unload Me
    'End

End Sub
'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)

    Dim FileSpec, FILE_SPEC_GO, i
    
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    i = CODER_VBP_FILE_NAME_2
    CODER_EXE_FILE_NAME_1 = i
    
    If InStr(i, " - 64 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 64 BIT.vbp", ".vbp")
    If InStr(i, " - 32 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 32 BIT.vbp", ".vbp")
    
    CODER_EXE_FILE_NAME_1 = Replace(CODER_EXE_FILE_NAME_1, ".vbp", ".exe")
    
    FileSpec = CODER_VBP_FILE_NAME_2
    FILE_SPEC_GO = FileSpec
    If Dir(FileSpec) = "" Then
        If LCase(GetSpecialfolder(&H29)) = LCase("C:\Windows\SysWOW64") Then
            FileSpec = Replace(FileSpec, ".vbp", " - 64 BIT.vbp")
        Else
            FileSpec = Replace(FileSpec, ".vbp", " - 32 BIT.vbp")
        End If
    End If
    
    If Dir(FileSpec) <> "" Then FILE_SPEC_GO = FileSpec
    
    CODER_VBP_FILE_NAME_2 = FILE_SPEC_GO
    
End Sub


Private Sub MNU_VB_FOLDER_Click()
    Beep
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Shell "EXPLORER /SELECT, """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    Beep
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
    End If
End Sub
'-------------------------------------------
'---------------------------------------------------
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Function GetSpecialfolder(ByVal CSIDL As Long) As String
    Dim R As Long
    Dim IDL As ITEMIDLIST
    R = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If R = NOERROR Then
        Path$ = Space$(512)
        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function
'---------------------------------------------------

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
            DoEvents
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

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************





