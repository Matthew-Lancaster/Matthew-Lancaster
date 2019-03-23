VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ScanPath 2.0 - Anything -- Demo of Class Object"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11355
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
      Left            =   7320
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   720
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
      Text            =   "D:\My Webs\MatthewLan.Com Web\LoveFolder\Cool\D-Words text"
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
Attribute VB_Name = "frmMain"
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



Private Sub Form_Activate()

'Call Jeepers_word
Call Jeepers_AltSZ
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




Private Sub Form_Load()
    Dim lCount As Long
    
    With cboMask
        .AddItem "*.TxT"
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
End
End Sub

Sub Jeepers_word()

txtPath.Text = "D:\My Webs\MatthewLan.Com Web\LoveFolder\Cool\D-Words text"
Call cmdScan_Click

'Dwords

'Reset

Dim Fs22

Set Fs22 = CreateObject("Scripting.FileSystemObject")

On Local Error GoTo jeep

v1$ = Mid$(txtPath.Text, InStrRev(txtPath.Text, "\"))

Statc$ = "<!-- Start of StatCounter Code -->" + vbCrLf
Statc$ = Statc$ + "<script type=""text/javascript"" language=""javascript"">" + vbCrLf
Statc$ = Statc$ + "var sc_project=585393;" + vbCrLf
Statc$ = Statc$ + "var sc_partition=4;" + vbCrLf
Statc$ = Statc$ + "var sc_security=""8832a018"";" + vbCrLf
Statc$ = Statc$ + "var sc_invisible=1;" + vbCrLf
Statc$ = Statc$ + "</script>" + vbCrLf
Statc$ = Statc$ + "<script type=""text/javascript"" language=""javascript"" src=""http://www.statcounter.com/counter/counter.js""></script><noscript><a href=""http://www.statcounter.com/"" target=""_blank""><img  src=""http://c5.statcounter.com/counter.php?sc_project=585393&amp;java=0&amp;security=8832a018&amp;invisible=1"" alt=""php hit counter"" border=""0""></a> </noscript>" + vbCrLf
Statc$ = Statc$ + "<!-- End of StatCounter Code -->" + vbCrLf

For we = 1 To ListView1.ListItems.Count
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)
    
    On Local Error GoTo jeep
'    Fs22.Movefile a1$ + b1$, c1$ + b1$
    On Local Error GoTo 0

fr4 = FreeFile
Open a1$ + b1$ For Input As #fr4
fr5 = FreeFile

c1$ = "D:\My Webs\MatthewLan.Com Web\LoveFolder\Cool\D-Words\"
Open c1$ + Mid$(b1$, 1, Len(b1$) - 4) + ".html" For Output As #fr5


Print #fr5, "<html>"
Print #fr5, ""
Print #fr5, "<head>"
Print #fr5, "<TITLE>Matthew Lan's Doubles</TITLE>"
'Print #fr5, "<TITLE>Usenet People -" + b1$ + "</TITLE>"
Print #fr5, "<META name=""Description"" content=""Matt Lan's Doubles " + Format$(Now, "DDD DD/MMM/YYYY HH:MM:SS") + """>"
Print #fr5, "<META NAME=""ROBOTS"" CONTENT=""INDEX,FOLLOW"">"
Print #fr5, ""
Print #fr5, "</head>"
Print #fr5, ""
Print #fr5, "<body>"
'Print #fr5, ""
'Print #fr5, " "

hgf1 = FreeFile
Open App.Path + "\txt.txt" For Input As #hgf1
tt$ = Input$(LOF(hgf1), hgf1)
Close hgf1
Print #fr5, tt$


Do
On Local Error Resume Next
Line Input #fr4, ks$
On Local Error GoTo 0
pig = 0
If InStr(UCase(ks$), "<FONT") > 0 Or InStr(UCase(ks$), "</FONT") > 0 Then pig = 1
If InStr(UCase(ks$), "<H1>") Then pig = 1
If InStr(UCase(ks$), "<H2>") Then pig = 1
If InStr(UCase(ks$), "<H3>") Then pig = 1
If InStr(UCase(ks$), "<H4>") Then pig = 1

If pig = 0 Then Print #fr5, ks$ + "<BR>"

Loop Until EOF(fr4)

Print #fr5, Statc$

Print #fr5, "</font>"
Print #fr5, "</body>"

Print #fr5, "</html>"


Close fr4
Close fr5
    
    
    
    
    List1.AddItem Format$(we, "000 ") + "Done > " + a1$ + b1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'Fs22.DeleteFolder comrad$ + "\" + WFD$

Next

List1.AddItem "Complete --------------"
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



Sub Jeepers_AltSZ()

txtPath.Text = "D:\My Webs\MatthewLan.Com Web\LoveFolder\Alt.UseNet\Alt_TxT\Temp Al.SZBest\01_All_Sz"
'txtPath.Text = "D:\My Webs\MatthewLan.Com Web\LoveFolder\Alt.UseNet\Alt_TxT\Temp_BorgRoidJudy\01_All_Sz"
txtPath.Text = "D:\My Webs\MatthewLan.Com Web\LoveFolder\Alt.UseNet\Alt_TxT\Temp Philo Borg\01_All_Sz"
c1$ = txtPath.Text + "\HTML\"
On Local Error Resume Next
MkDir c1$
On Local Error GoTo 0

Call cmdScan_Click

'Dwords

'Reset

Dim Fs22

Set Fs22 = CreateObject("Scripting.FileSystemObject")

On Local Error GoTo jeep

v1$ = Mid$(txtPath.Text, InStrRev(txtPath.Text, "\"))

Statc$ = "<!-- Start of StatCounter Code -->" + vbCrLf
Statc$ = Statc$ + "<script type=""text/javascript"" language=""javascript"">" + vbCrLf
Statc$ = Statc$ + "var sc_project=585393;" + vbCrLf
Statc$ = Statc$ + "var sc_partition=4;" + vbCrLf
Statc$ = Statc$ + "var sc_security=""8832a018"";" + vbCrLf
Statc$ = Statc$ + "var sc_invisible=1;" + vbCrLf
Statc$ = Statc$ + "</script>" + vbCrLf
Statc$ = Statc$ + "<script type=""text/javascript"" language=""javascript"" src=""http://www.statcounter.com/counter/counter.js""></script><noscript><a href=""http://www.statcounter.com/"" target=""_blank""><img  src=""http://c5.statcounter.com/counter.php?sc_project=585393&amp;java=0&amp;security=8832a018&amp;invisible=1"" alt=""php hit counter"" border=""0""></a> </noscript>" + vbCrLf
Statc$ = Statc$ + "<!-- End of StatCounter Code -->" + vbCrLf

For we = 1 To ListView1.ListItems.Count
    a1$ = ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ListView1.ListItems.Item(we)
    
    On Local Error GoTo jeep
'    Fs22.Movefile a1$ + b1$, c1$ + b1$
    On Local Error GoTo 0

fr4 = FreeFile
Open a1$ + b1$ For Input As #fr4
fr5 = FreeFile
c1$ = a1$ + "\HTML\"
Open c1$ + Mid$(b1$, 1, Len(b1$) - 4) + ".html" For Output As #fr5


Print #fr5, "<html>"
Print #fr5, ""
Print #fr5, "<head>"
Print #fr5, "<TITLE>Matthew Lan's Alt.Sz Best Of 1800 Users Since 2003 -" + Mid$(b1$, 1, Len(b1$) - 4) + "-</TITLE>"
'Print #fr5, "<TITLE>Usenet People -" + b1$ + "</TITLE>"
Print #fr5, "<META name=""Description"" content=""Matt Lan's Alt.SZ Best Of 1800 Users Since 2003 " + Format$(Now, "DDD DD/MMM/YYYY HH:MM:SS") + """>"
Print #fr5, "<META NAME=""ROBOTS"" CONTENT=""INDEX,FOLLOW"">"
Print #fr5, ""
Print #fr5, "</head>"
Print #fr5, ""
Print #fr5, "<body>"
'Print #fr5, ""
'Print #fr5, " "

hgf1 = FreeFile
Open App.Path + "\txt2.html" For Input As #hgf1
tt$ = Input$(LOF(hgf1), hgf1)
Close hgf1

Print #fr5, tt$


Do
On Local Error Resume Next
Line Input #fr4, ks$
On Local Error GoTo 0
pig = 0
If InStr(UCase(ks$), "<FONT") > 0 Or InStr(UCase(ks$), "</FONT") > 0 Then pig = 1
If InStr(UCase(ks$), "<H1>") Then pig = 1
If InStr(UCase(ks$), "<H2>") Then pig = 1
If InStr(UCase(ks$), "<H3>") Then pig = 1
If InStr(UCase(ks$), "<H4>") Then pig = 1

If pig = 0 Then Print #fr5, ks$ + "<BR>"

Loop Until EOF(fr4)

Print #fr5, Statc$

Print #fr5, "</font>"
Print #fr5, "</body>"

Print #fr5, "</html>"


Close fr4
Close fr5
    
    
    
    
    List1.AddItem Format$(we, "000 ") + "Done > " + a1$ + b1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'Fs22.DeleteFolder comrad$ + "\" + WFD$

Next

List1.AddItem "Complete --------------"
List1.AddItem "--------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

'Timer1.Enabled = True

End

Exit Sub
jeep:
errs2 = Err.Number
errs3$ = Err.Description
Resume Next

End Sub



