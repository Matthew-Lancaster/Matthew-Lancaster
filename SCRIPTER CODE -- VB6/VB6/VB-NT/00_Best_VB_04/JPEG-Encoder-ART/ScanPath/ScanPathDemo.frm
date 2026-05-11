VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ScanPath 
   BackColor       =   &H00808080&
   Caption         =   "ScanPath 2.0 - Sort  Anything -"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12975
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   12975
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox Chk_SortOnDate 
      Caption         =   "Sort On Date"
      Height          =   195
      Left            =   1485
      TabIndex        =   37
      Top             =   1620
      Width           =   1275
   End
   Begin VB.CommandButton CmdScanDir 
      Caption         =   "Scan Dir"
      Height          =   675
      Left            =   6510
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   60
      Width           =   810
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Enabled         =   0   'False
      Height          =   825
      Left            =   8370
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6960
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.ComboBox cboMask 
      Height          =   315
      Left            =   375
      TabIndex        =   1
      Text            =   "*.jpg"
      Top             =   390
      Width           =   6105
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   5925
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   60
      Width           =   555
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   390
      TabIndex        =   0
      Text            =   "E:\My Music Zen"
      Top             =   45
      Width           =   5505
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "Search Sub-Folders"
      Height          =   195
      Left            =   1485
      TabIndex        =   8
      Top             =   1170
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkPatternMatching 
      Caption         =   "Pattern Matching"
      Height          =   195
      Left            =   1485
      TabIndex        =   7
      Top             =   945
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkNormal 
      Caption         =   "Normal"
      Height          =   195
      Left            =   15
      TabIndex        =   2
      Top             =   930
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      Height          =   195
      Left            =   15
      TabIndex        =   3
      Top             =   1155
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      Height          =   195
      Left            =   15
      TabIndex        =   4
      Top             =   1380
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      Height          =   195
      Left            =   15
      TabIndex        =   5
      Top             =   1605
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      Height          =   195
      Left            =   15
      TabIndex        =   6
      Top             =   1815
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   5670
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   960
      Width           =   1545
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   0
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2340
      Width           =   1365
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   5670
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1995
      Width           =   1545
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   5655
      TabIndex        =   16
      Top             =   2340
      Width           =   1530
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   1
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2700
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   5655
      TabIndex        =   18
      Top             =   2700
      Width           =   1530
   End
   Begin VB.CheckBox chkRefreshListView 
      Caption         =   "Refresh ListView (slower)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   10095
      TabIndex        =   10
      Top             =   6915
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.CheckBox chkCopyMemory 
      Caption         =   "Use CopyMemory (display Date && Size)"
      Height          =   195
      Left            =   1485
      TabIndex        =   11
      Top             =   1395
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3510
      Left            =   15
      TabIndex        =   21
      Top             =   3150
      Width           =   12915
      _ExtentX        =   22781
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
      Enabled         =   0   'False
      Height          =   195
      Left            =   10110
      TabIndex        =   9
      Top             =   7170
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Default         =   -1  'True
      Height          =   660
      Left            =   7410
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   45
      Width           =   525
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   0
      Left            =   5670
      TabIndex        =   13
      Top             =   1275
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   55181313
      CurrentDate     =   37299
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   1
      Left            =   5670
      TabIndex        =   14
      Top             =   1635
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   55181313
      CurrentDate     =   37296
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   2
      Left            =   7230
      TabIndex        =   44
      Top             =   1260
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      _Version        =   393216
      Format          =   55181314
      CurrentDate     =   37299
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
      Left            =   8610
      TabIndex        =   43
      Top             =   15
      Width           =   4335
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
      Left            =   8610
      TabIndex        =   42
      Top             =   645
      Width           =   4335
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
      Left            =   8610
      TabIndex        =   41
      Top             =   945
      Width           =   4335
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
      Left            =   8610
      TabIndex        =   40
      Top             =   1245
      Width           =   4335
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
      Left            =   8610
      TabIndex        =   39
      Top             =   330
      Width           =   4335
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
      Left            =   8610
      TabIndex        =   38
      Top             =   1545
      Width           =   4335
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Working"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8925
      TabIndex        =   34
      Top             =   7140
      Visible         =   0   'False
      Width           =   1080
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
      TabIndex        =   33
      Top             =   765
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
      Left            =   1500
      TabIndex        =   26
      Top             =   735
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mask"
      Height          =   195
      Left            =   15
      TabIndex        =   24
      Top             =   420
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      Height          =   195
      Left            =   45
      TabIndex        =   22
      Top             =   75
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
      Left            =   45
      TabIndex        =   25
      Top             =   720
      Width           =   885
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      Height          =   195
      Index           =   0
      Left            =   5055
      TabIndex        =   30
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
      Height          =   195
      Left            =   5070
      TabIndex        =   28
      Top             =   2025
      Width           =   585
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   195
      Index           =   1
      Left            =   5055
      TabIndex        =   32
      Top             =   2760
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   5055
      TabIndex        =   29
      Top             =   1305
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "File Date"
      Height          =   195
      Left            =   5055
      TabIndex        =   27
      Top             =   990
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   5055
      TabIndex        =   31
      Top             =   1605
      Width           =   195
   End
   Begin VB.Menu Mnu_Cmd 
      Caption         =   "Commands"
      Begin VB.Menu Mnu_Com1 
         Caption         =   "Com1"
      End
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------Is This Sub Get on It at start      at       Bangers
'---------------JOBS
'Not All Need Filters

Public cProcesses As New clsCnProc
Public Dog, OutPutPath, Xdate As Date, Tdate As Date, Zdate$, Ydate As Date, OldFolder, OldSize

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

'Public A1$, B1$, G1$, FF$

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


Private Sub cboSizeType_Change(index As Integer)
'cboSizeType = 0
End Sub

Private Sub chkCopyMemory_Click()
    
    chkSort.enabled = (chkCopyMemory.Value = vbUnchecked)
    
End Sub


Private Sub chkSort_Click()
    chkCopyMemory.enabled = (chkSort.Value = vbUnchecked)
End Sub

Private Sub cmdBrowse_Click()
    txtPath.Text = GetFolder(Me.hwnd, "Scan Path:", txtPath.Text)
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
       
    DoEvents
    
 
' = cProcesses.Convert(CurProcHwnd, Otss, cnFromhWnd Or cnToProcessID)
' = cProcesses.Convert(Otss, totss2, cnTohWnd Or cnFromProcessID)

Dim PID As Long
Dim HW As Long

HW = fmain.hwnd
i = cProcesses.Convert(HW, PID, cnFromhWnd Or cnToProcessID)
   
'msgbox i
   
Call Process_HIGH_PRIORITY_CLASS(PID)
    
    If FS.FolderExists(txtPath) = False Then
        If DONTWARNERROR = False Then msgbox txtPath.Text + vbCrLf + "Dont Exist"
        
        BreakScanNow = True
        NoFiles = True
    End If
    
    If BreakScanNow = True Then
        BreakScanNow = False
        ScanInProgress = False
        ScanFinished = True
        
        If cmdScan.Caption = "Stop" Then SP.StopScan
        Exit Sub
    End If
    
    F = FreeFile
    Open App.Path + "\Last Error Pix.txt" For Output As #F
    Close #F
    On Error Resume Next
    Kill App.Path + "\Error Pix.txt"
    On Error GoTo 0
   
'    On Error Resume Next


   Dim lCount As Long
    Dim lSize(1) As Long
    
    
    
    If BreakScanX = False Then
        '####################################################################################################
        'Process our Selection Criteria
        
        If FirstDirScanned = "" Then
            FirstDirScanned = txtPath.Text
        End If
  
        For lCount = 0 To 1
            lSize(lCount) = Val(txtSize(lCount).Text)
        
            Select Case cboSizeType(lCount).ListIndex
            Case 1: lSize(lCount) = lSize(lCount) * 1024
            Case 2: lSize(lCount) = (lSize(lCount) * 1024) * 1024
            End Select
        Next lCount
            
        If txtSize(1).Visible Then
            If lSize(0) = 0 Then
                msgbox "Maximium Size value must exceed 0", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                msgbox "Maximium Size value must exceed Minimum Size", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            End If
        End If
        
        '####################################################################################################
        'Reset Form
        cmdScan.Caption = "Stop"
        
        'Screen.MousePointer = vbHourglass
        'ListView1.ListItems.Clear
        
        If chkRefreshListView.Value = vbUnchecked Then
            LockWindowUpdate ListView1.hwnd
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
            .StartScan txtPath, chkSubFolders.Value, chkSort.Value And chkSort.enabled, chkPatternMatching
            
             ScanPath.lblCount1.Caption = ScanPath.ListView1.ListItems.Count

            
            'we always use direct scan as is better to sort after job then during thats all we need
            If Chk_SortOnDate.Value = vbChecked Then
                ListView1.SortOrder = lvwAscending
                ListView1.SortKey = 4
                ListView1.Sorted = True
                ListView1.Sorted = False
            Else
                ListView1.SortOrder = lvwAscending
                ListView1.SortKey = 0 ' SORT ON FILENAME THEN
                ListView1.Sorted = True
                ListView1.Sorted = False
                ListView1.SortOrder = lvwAscending
                ListView1.SortKey = 1 ' SORT ON DIR
                ListView1.Sorted = True
                ListView1.Sorted = False

            End If
        
        End With
        
       With ListView1
            If .ListItems.Count > 0 Then
                .SelectedItem = .ListItems(1)
            End If
            
            lblCount1.Caption = .ListItems.Count
        End With
        
        LockWindowUpdate 0
        
        'cmdScan.Caption = "Scan"
        'Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
        ListView1.ListItems.Clear
        ScanInProgress = False
    End If

    ScanFinished = True

    If IsIDE = False Then Call Process_LOW_PRIORITY_CLASS(PID)

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
                msgbox "Maximium Size value must exceed 0", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                msgbox "Maximium Size value must exceed Minimum Size", vbExclamation
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
            LockWindowUpdate ListView1.hwnd
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
            .StartScanDir txtPath, chkSubFolders.Value, chkSort.Value And chkSort.enabled, chkPatternMatching
        
            
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
            
            lblCount1.Caption = .ListItems.Count
        End With
        
        LockWindowUpdate 0
        
        cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
    End If


lblCount1.Caption = ListView1.ListItems.Count
    

End Sub




Private Sub Form_Load()
    
'Me.Show
    
Me.BackColor = &HC0C0C0
    
   
   
'--------
'If you use ChkMem Size Files then Know G1$ Alternate Data uses same column(2) for its data
'-
'Seem to have a problem with some program an not other when use chkmem file size and date
'Work Around done that
'-
'The Date checker box does scan for time if the var contains time
'---------
    
ScanPath.Caption = ScanPath.Caption + App.EXEName
'scanpath.Caption = frmMain.Caption + App.EXEName
  
x = 1
y = 1
On Error Resume Next
For Each Control In Controls
    If Control.enabled = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next
On Error GoTo 0


ScanPath.Width = x + 110
If mnu = 1 Then
    pluso = 720
Else
    pluso = 450
End If
ScanPath.Height = y + pluso
    
    Dim lCount As Long
    
    With cboMask
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
            .ListIndex = 0 ' set it to bytes here
        End With
    Next lCount
    
    With ListView1
        .ColumnHeaders.Add , "FILE", "File", 3000, lvwColumnLeft
        .ColumnHeaders.Add , "PATH", "Path", 5000, lvwColumnLeft
        .ColumnHeaders.Add , "DOS-8.3", "Dos", 1500, lvwColumnLeft
        .ColumnHeaders.Add , "SIZE", "Size", 1200, lvwColumnLeft
        .ColumnHeaders.Add , "DATE", "Date", 1800, lvwColumnLeft
        
        .View = lvwReport
    End With

    Call chkCopyMemory_Click

'Call DelJunkFrontPage

'ScanPath.Show

'Call INetContent

'Call Bangers

'txtPath.Text = "D:\MY WEBS\MatthewLan.Com Web\"


End Sub


Private Sub lblCount1_Change()
    
    'Main.Command1.Caption = ScanPath.lblCount1
    
    On Error Resume Next
    
    Call fmain.RunCount
    

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
            
            'CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)
            'This is not working in EXE but does in IDE Revert back to Old Way
            'CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)
                
            'Remember Public FS
            If Filename <> "" Then
                Set F = FS.getfile(Path + DFilename)
                adate1 = F.datelastmodified
                asize1 = F.Size
            Else
                Set F = FS.GetFolder(Path)
                adate1 = F.datelastmodified
                'We May want this later Slow Down is Massive
                If OldFolder <> Path Then
                asize1 = F.Size
                Else
                asize1 = OldSize
                End If
                OldFolder = Path
                OldSize = asize1
            End If

'            LV.SubItems(2) = uWIN32.nFileSizeLow
'            LV.SubItems(3) = FormatFileDate(uWIN32.ftLastWriteTime)
            LV.SubItems(3) = asize1
            LV.SubItems(4) = Format(adate1, "YYYY-MM-DD HH-MM-SS")
        End If
        
        frmain.llb1.Refresh
        Exit Sub
        
        'Refresh Count & Scroll the list every 10th file
        If (LV.index Mod 10) = 0 Then
            
            If chkRefreshListView.Value Then
                .SelectedItem = LV
                .SelectedItem.EnsureVisible
            End If
            
            lblCount1.Caption = LV.index
        End If
    End With
    Exit Sub

GetFileError:
    msgbox Err.Description, vbCritical


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
        
        
        
        
        
        
        
        
        
        
        'If chkCopyMemory.Value Then
        '    'VB does not allows UDT's to be passed directly from a Class but we can access the structure
        '    'using pointers. We use CopyMemory to place the WIN32_FIND_DATA variable from the Class into
        '    'the uWIN32 declared in this Sub.
        '
        '    'This is primarily a speed trick - we could retrieve the date & time information "ourselves"
        '    'using the standard VB functions or an API call but since we already have it...
        '    CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)
        '    LV.SubItems(3) = uWIN32.nFileSizeLow
        ''    LV.SubItems(4) = FormatFileDate(uWIN32.ftLastWriteTime)
        'End If
        
        
            If Filename <> "" Then
                Set F = FS.getfile(Path + DFilename)
                adate1 = F.datelastmodified
                asize1 = F.Size
            Else
                Set F = FS.GetFolder(Path)
                adate1 = F.datelastmodified
                'We May want this later Slow Down is Massive
                If OldFolder <> Path Then
                asize1 = F.Size
                Else
                asize1 = OldSize
                End If
                OldFolder = Path
                OldSize = asize1
            End If

'            LV.SubItems(2) = uWIN32.nFileSizeLow
'            LV.SubItems(3) = FormatFileDate(uWIN32.ftLastWriteTime)
            LV.SubItems(3) = asize1
            LV.SubItems(4) = Format(adate1, "YYYY-MM-DD HH-MM-SS")
        
        
        
        'Refresh Count & Scroll the list every 10th file
        If (LV.index Mod 10) = 0 Then
            'If chkRefreshListView.Value Then
            '    .SelectedItem = LV
            '    .SelectedItem.EnsureVisible
            'End If
            
            lblCount1.Caption = LV.index
        End If
    End With
    Exit Sub

GetFileError:
    msgbox Err.Description, vbCritical
    Resume
End Sub



Private Function FormatFileDate(CT As FILETIME) As String
    Const SHORT_DATE = "Short Date"
    
    Dim ST As SYSTEMTIME
    Dim ds1 As Single
    Dim ds2 As Single
       
    If FileTimeToSystemTime(CT, ST) Then
          ds1 = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
          ds2 = TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
          
          FormatFileDate = Format$(ds1 + ds2, "YYYY-MM-DD HH-MM-SS")
    End If

End Function


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************


