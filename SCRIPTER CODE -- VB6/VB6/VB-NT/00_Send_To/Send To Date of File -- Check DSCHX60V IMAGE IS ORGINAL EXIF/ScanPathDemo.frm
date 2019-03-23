VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ScanPath 
   Caption         =   "ScanPath 2.0 - "
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14490
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   14490
   StartUpPosition =   2  'CenterScreen
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
      ScaleWidth      =   14430
      TabIndex        =   22
      Top             =   0
      Width           =   14490
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
      Height          =   4740
      Left            =   45
      TabIndex        =   21
      Top             =   3210
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   8361
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
      Left            =   11040
      Picture         =   "ScanPathDemo.frx":12CC
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   705
      Width           =   810
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
      Format          =   22085633
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
      Format          =   22085633
      CurrentDate     =   37296
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort 02"
      Height          =   240
      Left            =   11040
      TabIndex        =   49
      Top             =   1875
      Width           =   810
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort 01"
      Height          =   240
      Left            =   11040
      TabIndex        =   48
      Top             =   1605
      Width           =   810
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remaining"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6840
      TabIndex        =   47
      Top             =   1575
      Width           =   1365
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8445
      TabIndex        =   46
      Top             =   1575
      Width           =   2220
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Count"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6840
      TabIndex        =   45
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Counting"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6840
      TabIndex        =   44
      Top             =   1005
      Width           =   1365
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Deleted"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6840
      TabIndex        =   43
      Top             =   1290
      Width           =   1365
   End
   Begin VB.Label lblCount3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8445
      TabIndex        =   42
      Top             =   1290
      Width           =   2220
   End
   Begin VB.Label lblCount2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8445
      TabIndex        =   41
      Top             =   1005
      Width           =   2220
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
      Left            =   8445
      TabIndex        =   31
      Top             =   720
      Width           =   2220
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Send To Date of File -- Check DSCHX60V IMAGE IS ORGINAL EXIF
'AutoPix

Dim EXT, A1$, B1$, GG$, H1$, XQ, S1, S2

Dim W$

'Put This in a Bas Mobule
Public m_CRC As clsCRC
Public WxHex$, ID$

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



Function GetExif(file$)
    Dim objExif As New ExifReader
    Dim txtExifInfo As String
    
    objExif.Load file$
    On Error Resume Next
    txtExifInfo = objExif.Tag(DateTimeOriginal)
    GetExif = txtExifInfo
End Function
Function GetExifModel(file$)
    Dim objExif As New ExifReader
    Dim txtExifInfo As String
    
    objExif.Load file$
    On Error Resume Next
    txtExifInfo = objExif.Tag(Make) + "-" + objExif.Tag(Model)
    GetExifModel = txtExifInfo
End Function

Function GetExif_FileSource(file$)
'    FileSource = &HA300&
'    SceneType = &HA301&
    Dim objExif As New ExifReader
    Dim txtExifInfo As String
    
    objExif.Load file$
    On Error Resume Next
    txtExifInfo = objExif.Tag(FileSource)
    GetExif_FileSource = txtExifInfo
End Function

Function GetExif_SceneType(file$)
'    FileSource = &HA300&
'    SceneType = &HA301&
    Dim objExif As New ExifReader
    Dim txtExifInfo As String
    
    objExif.Load file$
    On Error Resume Next
    txtExifInfo = objExif.Tag(SceneType)
    GetExif_SceneType = txtExifInfo
End Function



Sub AutoPix()

Dim fs22
Set fs22 = CreateObject("Scripting.FileSystemObject")

AT$ = Command$

W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2009 FaceBook London\"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Rutland Gardens\"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2007-2009 Matthew Lancaster\2008 Coretta Player"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2007-2009 Matthew Lancaster\"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2007 Dec Sqirrel Xmaz Mill View - Cyber Shop\"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2007 Mums-19th-Oct-2007-Linda Lancaster\"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2008 2010 Mill View\"
W$ = "M:\DSC\2007"
W$ = "M:\DSC\2014 NOKIA"
W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\"
W$ = "M:\DSC\"
W$ = "M:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2003-2008 Eddie-Lancaster\2005 Eddie's\"
W$ = "M:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2003-2008 Eddie-Lancaster\2004 01 27 Eddie's\"
W$ = "M:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2003-2008 Eddie-Lancaster\2006 01 Eddie's Tools & Trucks\"
W$ = "M:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2003-2008 Eddie-Lancaster\2006 Eddie's\"
W$ = "M:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2003-2008 Eddie-Lancaster\2007-12-14  Eddie\"
W$ = "M:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2003-2008 Eddie-Lancaster\2008 Brighton Beach\"
W$ = "M:\DSC\"
W$ = "D:\DSC\"
W$ = "D:\DSC\DCIM\IMAGES\"
W$ = "D:\DSC\DCIM\10050831\work\"
W$ = "D:\DSC\DCIM\"
'W$ = "D:\# MY DOCS\# 01 My Documents\My Pictures\MP Navigator EX\"
W$ = "D:\DSC\2015 Sony CyberShot DSC-HX60V\DCIM\20151200 December and After\20151220 And After\"

W$ = "D:\DSC\2015 2016\2016 Sony CyberShot DSC-HX60V\"
W$ = "D:\DSC\2015 2016\2016 Sony CyberShot DSC-HX60V\DCIM4\"
W$ = "D:\DSC\2015 2016\2016 Sony CyberShot DSC-HX60V\DCIM2\20160220\"
W$ = "D:\DSC\2015 2016\2016 Sony CyberShot DSC-HX60V\"

'PATH -- PROGRAM DATE IN FRONT OF FILENAME

'If IsIDE = True Then at$ = "J:\Music\z Root 2008-9\Phone Calls"
'If IsIDE = True Then AT$ = "U:\DSC\2011\00 25%\"
'If IsIDE = True Then AT$ = "U:\0 00 VIDEO\01 Beavis & Butthead\"
'If IsIDE = True Then AT$ = "M:\DSC\2011\"
'If IsIDE = True Then AT$ = "M:\DSC\2007\"

If IsIDE = True Then AT$ = W$


'Call Rename_Image_Folder_Proper
'Exit Sub


If AT$ <> "" And Mid$(AT$, 1, 1) = """" Then
AT$ = Mid$(AT$, 2)
AT$ = Mid$(AT$, 1, Len(AT$) - 1)
End If

If Len(AT$) < 4 Then MsgBox "Not on a Root Dir - END": End

'cboMask = "*.wav;*.wma;*.amr;*.mp3"
'cboMask.Text = "*.HTML;*.JPE;*.JPEG;*.GIF;*.BMP;*.JPG;*.ZIP;*.TXT;*.PNG"
'cboMask.Text = "*.JPE"
'cboMask.Text = "*.JPEG"
'cboMask.Text = "*.GIF"
'cboMask.Text = "*.BMP"
cboMask.Text = "*.JPG"
'cboMask.Text = "*.*"

txtPath = AT$
chkSubFolders.Value = vbChecked
Call cmdScan_Click


XST = Now
Dim Dateset As Date
Dim Dateset3 As Date

For WE = ListView1.ListItems.Count To 1 Step -1
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    ListView1.ListItems.Item(WE).EnsureVisible

    DoEvents
    
    'EXIF_GET = GetExifModel(A1 + B1$)
    
    'IF THE TWO BELOW RETURN OTHER WHEN SUPPOSED TO BE
'    -----------------------------------
'    THE TWO EXIF TYPE INFO VARIABLE
'    THAT CAN TELL A IMAGE BEEN MODIFIED
'    -----------------------------------
'    FileSource - DSC - Digital still camera
'    SceneType - A directly photographed image
'    -----------------------------------
'    -----------------------------------
'    FileSource -Other
'    SceneType -Other
'    -----------------------------------
'    -----------------------------------
    
    'FileSource = &HA300&
    'SceneType = &HA301&
    
    'GetExif_FileSource (file$)
    'GetExif_SceneType (file$)

    
    AVARA1 = GetExif_FileSource(A1 + B1$)
    AVARA2 = ""
    
'    AVARB1 = GetExif_SceneType(A1 + B1$)
'    AVARB2 = ""
    
    NOT_ORGINAL_BIT = False
    
    For R = 1 To Len(AVARA1)
    
        AVARA2 = AVARA2 + Str(Asc(Mid(AVARA1, R, 1)))
        If R = Len(AVARA1) Then Exit For
        AVARA2 = AVARA2 + " --"
    
        
        If R = 3 And Asc(Mid(AVARA1, R, 1)) = 1 Then
            NOT_ORGINAL_BIT = True
            NOT_ORGINAL_BIT_COUNT = NOT_ORGINAL_BIT_COUNT + 1
        

        End If
    Next
    
    If NOT_ORGINAL_BIT = True Then
        ListView1.ListItems.Item(WE).SubItems(2) = AVARA2
        ListView1.ListItems.Item(WE).SubItems(3) = GetExifModel(A1 + B1$)
    End If
    
    If NOT_ORGINAL_BIT = False Then
        ListView1.ListItems.Remove (WE)
        Debug.Print AVARA2 + " -- " + B1$
    End If
    
    lblCount2 = Trim(Str(ListView1.ListItems.Count)) + " NOT ORGINAL"
    lblCount3 = Trim(Str(NOT_ORGINAL_BIT_COUNT)) + " NOT ORGINAL"
    
    
'    For R = 1 To Len(AVARB1)
'
'        AVARB2 = AVARB2 + Str(Asc(Mid(AVARB1, R, 1)))
'        If R = Len(AVARB1) Then Exit For
'        AVARB2 = AVARB2 + " --"
'
'    Next
    
    'WHEN 3RD MID STRING POSITION IS = 1 THEN NOT ORGINAL
    'WHEN 3RD MID STRING POSITION IS = 0 THEN YES ORGINAL
    
    'Debug.Print GetExifModel(A1 + B1$) + " -- " + B1$
'    Debug.Print AVARA2 + " -- " + B1$
'    Debug.Print AVARB2 + " -- " + B1$
    
Next

    
    
    
FRE1 = FreeFile
Open txtPath + "\OUTPUT RESULT NOT ORGINAL IMAGE PHOTO SCRIPT -- " + Format(Now, "YYYY-MM-DD HH-MM-SS") + ".txt" For Output As #FRE1
For WE = 1 To ListView1.ListItems.Count
    
    E1$ = ListView1.ListItems.Item(WE).SubItems(3)
    D1$ = ListView1.ListItems.Item(WE).SubItems(2)
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    ListView1.ListItems.Item(WE).EnsureVisible
    'lblCount2 = ListView1.ListItems.Count
    DoEvents
    Print #FRE1, D1$ + " -- " + E1$ + " -- " + A1$ + B1$
    
Next
Close #FRE1
    
'End
Exit Sub


For WE = 1 To ListView1.ListItems.Count
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    ListView1.ListItems.Item(WE).EnsureVisible


'    If Val(Mid$(b1$, 1, 4)) < 2000 Then 'SHOULD STOP REPEAT THE ACTION
    
    'If InStr(A1$, "2007-2009 Matthew Lancaster") > 0 Then flaggo = 1
    'If flaggo = 0 Then GoTo JumpToNext
    
'    tval1 = Format$(Val(Mid(B1$, 1, InStr(B1$, ".") - 1)), "000")
'    tval2 = Format$(Val(Mid(B1$, InStr(B1$, ".") + 1, 2)), "00")
'    tval3 = Format$(Val(Mid(B1$, InStr(B1$, "-") + 1, 2)), "00")
'    tval4 = Mid(B1$, InStr(B1$, "-") + 3)
'
'    GG$ = tval1 + " S" + tval2 + "E" + tval3 + tval4
'    GG$ = B1$
'    GG$ = Replace(GG$, "  ", " ")
'
'    Name A1$ + B1$ As A1$ + GG$
                
                
'    R3 = Replace(R, "Beavis and Butthead ", " ")
'    R3 = Format(CNT, "000") + R3

'    WE = CNT - 1
    
'    R2 = Mid(B1$, 8, 11)
'    XTZ = 0
'    If Mid(R2, 3, 1) = "-" Then XTZ = DateValue(R2)
'    If XTZ <> 0 Then
'        XD1 = DateValue(R2)
'        R2 = Mid(B1$, 1, 7) + Format(XD1, "DD-MM-YYYY") + Mid(B1$, 19)
'        Name A1$ + B1$ As A1$ + R2
'    Else
'
'    End If
    
    
    'Name A1$ + B1$ As A1$ + Format(WE, "000 ") + B1$
    
    
    'fs22.MoveFILE A1$ + "3\" + R2, A1$ + R2
        
    
    
'GoTo JumpToNext
'GoTo JUMP_RENAME
    
    '-------------------------------------
    If Mid$(B1$, 1, 3) = "DSC" Or UCase(Mid$(B1$, 1, 3)) = "IMG" Then
    
    
'    If Mid$(B1$, 1, 3) = "DSC" Then 'Or Mid$(b1$, 1, 4) = "_SCF" Or Mid$(b1$, 1, 4) = "PICT" Then  'SHOULD STOP REPEAT THE ACTION
'    If Mid$(B1$, 9, 1) = "(" And Mid$(B1$, 13, 1) = ")" Then
'    If InStr(A1$, "IMAGE\SD") > 0 Then 'SHOULD STOP REPEAT THE ACTION
'    If InStr(B1$, "PICT") > 0 Then 'SHOULD STOP REPEAT THE ACTION
'    XST = Mid$(B1$, InStr(B1$, "DSC"), 8)
'        On Error Resume Next
        If fs22.FileExists(A1$ + B1$) = False Then GoTo JumpToNext
             
        Set F = fs22.GetFile(A1$ + B1$)
        
        If F.Size = 0 Then
            Clipboard.Clear
            Clipboard.SetText ("Zero Size File" + vbCrLf + A1$ + B1$)
            
            MsgBox ("Zero Size File" + vbCrLf + A1$ + B1$)
            Kill A1$ + B1$: GoTo JumpToNext
        End If
        'GG$ = Format$(F.datelastmodified, "YYYY-MM-DD HH-MM-SS") + " - " + B1$
        GG$ = Format$(F.datelastmodified, "YYYY-MM-DD HH-MM-SS") + " - " ' - "
        'Exit Sub
        Dateset2 = ""
        Dateset2 = GetExif(A1$ + B1$)
        Dateset = Now + DateSerial(9000, 1, 1)
        If IsNumeric(Mid(Dateset2, 1, 2)) = True Then
            dss1 = DateValue(Replace(Mid(Dateset2, 1, 10), ":", "-"))
            dss2 = TimeValue(Mid(Dateset2, 12, 8))
            
            Dateset = DateValue(dss1) + TimeValue(dss2)
            GG$ = Format$(Dateset, "YYYY-MM-DD HH-MM-SS") + " - " ' - "
        End If
'        If B1$ = "2007-10-17 17-18-00 - JFIF - SC01282 (1).JPG" Then Stop
        
        If F.datelastmodified < Dateset And Dateset2 = "" Then
            GG$ = Format$(F.datelastmodified, "YYYY-MM-DD HH-MM-SS") + " - " ' - "
        End If


        XB1$ = B1$
        XB1$ = Replace(XB1$, "MILL VIEW HOSPITAL-", "")
        dss1 = ""
        dss2 = ""
        Dateset3 = Now + DateSerial(9000, 1, 1)
        On Error Resume Next
'        dset = 0
'        If Mid(xb1$, 8, 1) = "-" Then dset = 1
'        If Mid(xb1$, 5, 1) = "-" Then dset = 1
'        If dset = 1 Then
        dset = 1
        If IsNumeric(Mid(XB1$, 15, 2)) = False Then dset = 0
        If IsNumeric(Mid(XB1$, 7, 2)) = False Then dset = 0
        If IsNumeric(Mid(XB1$, 1, 4)) = False Then dset = 0
        If IsNumeric(Mid(XB1$, 22, 2)) = False Then dset = 0
        If IsNumeric(Mid(XB1$, 25, 2)) = False Then dset = 0
        If IsNumeric(Mid(XB1$, 28, 2)) = False Then dset = 0
        
        dss1 = DateValue(Mid(XB1$, 15, 2) + "-" + Mid(XB1$, 7, 2) + "-" + Mid(XB1$, 1, 4))
        dss2 = TimeValue(Replace(Mid(XB1$, 22, 8), "-", ":"))
        'If dss1 <> "" And dss2 <> "" Then Stop
        On Error GoTo 0
        If dset = 1 Then
            If dss1 <> "" And dss2 <> "" Then
                Dateset3 = dss1 + dss2
                If Val(Dateset) > 0 Then
                    'If Dateset3 < F.datelastmodified And Dateset3 < Dateset Then
                    If Dateset3 < F.datelastmodified And Dateset3 < Dateset And Dateset2 <> "" Then
                        GG$ = Format$(Dateset3, "YYYY-MM-DD HH-MM-SS") + " - " ' - "
                    End If
                End If
            End If
        End If

        dss1 = ""
        Dateset4 = Now + DateSerial(9000, 1, 1)
        On Error Resume Next
        dset = 1
        If IsNumeric(Mid(XB1$, 9, 2)) = False Then dset = 0
        If IsNumeric(Mid(XB1$, 6, 2)) = False Then dset = 0
        If IsNumeric(Mid(XB1$, 1, 4)) = False Then dset = 0
        If IsNumeric(Mid(XB1$, 12, 2)) = False Then dset = 0
        If IsNumeric(Mid(XB1$, 15, 2)) = False Then dset = 0
        If IsNumeric(Mid(XB1$, 18, 2)) = False Then dset = 0
        If dset = 1 Then dss1 = DateValue(Mid(XB1$, 9, 2) + "-" + Mid(XB1$, 6, 2) + "-" + Mid(XB1$, 1, 4))
        On Error GoTo 0
        If dss1 <> "" Then
            If dset = 1 Then
                dss2 = TimeValue(Replace(Mid(XB1$, 12, 8), "-", ":"))
                Dateset4 = dss1 + dss2
                If Val(Dateset) > 0 Then
                    'If Dateset3 < F.datelastmodified And Dateset2 = "" Then
                    'If Dateset4 < F.datelastmodified And Dateset4 < Dateset3 And Dateset4 < Dateset Or Dateset = "" Then
                    If Dateset4 < F.datelastmodified And Dateset4 < Dateset3 And Dateset4 < Dateset And Dateset2 <> "" Then
                        GG$ = Format$(Dateset4, "YYYY-MM-DD HH-MM-SS") + " - " ' - "
                    End If
                End If
            End If
        End If

'        gg$ = Format$(XST, "YYYY-MM-DD HH-MM-SS") + " - " + b1$

'        gg$ = XST + b1$
'        gg$ = Mid$(b1$, 23) 'IF IT DOES REPEAT NEED A BIT OF THIS
'        Name A1$ + b1$ As A1$ + GG$
        
'        EXT = Mid$(B1$, 1, InStrRev(B1$, ".")) + "RAR"
'        If fs22.FileExists("M:\#\#D\00 Pen Drives MOBILES\DSC\" + EXT) = True Then Stop
            
'            Name A1$ + B1$ As A1$ + GG$

'---------------------------------------
'        H1$ = Mid(B1$, 1, 4) + "-" + Mid$(B1$, 7, 2) + "-" + Mid$(B1$, 15, 2) + " "
'        H1$ = H1$ + Mid(B1$, 22, 8)
        'XQ = InStrRev(B1$, " - ")
'        XQ = InStrRev(B1$, "PICT")
        'XQ = InStrRev(B1$, " ")
'        H1$ = H1$ + Mid(B1$, XQ )
        'GG$ = GG$ + Mid(B1$, XQ + 2)
'        XQ = 1
'        GG$ = GG$ + Mid(B1$, XQ)
        'GG$
        
'        If fs22.FileExists(A1$ + H1$) Then
'            Set F = fs22.GetFile(A1$ + B1$)
'            s1 = F.Size
'            Set F = fs22.GetFile(A1$ + H1$)
'            S2 = F.Size
'            If s1 > S2 Then
'                Stop
'            End If
'            If s1 < S2 Then
'                Kill A1$ + B1$
'                Name A1$ + B1$ As A1$ + H1$
'            End If
'            If s1 = S2 Then
'                Kill A1$ + B1$
'                'Name A1$ + B1$ As A1$ + H1$
'            End If
'
'        Else
'            Name A1$ + B1$ As A1$ + H1$
'
'        End If
'---------------------------------------
        
'        Stop
                    
                    
'        If InStr(B1$, "WebCapture") = 0 Then
                    
'            Name A1$ + B1$ As A1$ + GG$
'            B1$ = GG$
'            EXT = Mid$(B1$, 1, InStrRev(B1$, ".")) + "RAR"
'            DONE = False
'            If fs22.FileExists("M:\#\#D\00 Pen Drives MOBILES\DSC\" + B1$) = False And fs22.FileExists("M:\#\#D\00 Pen Drives MOBILES\DSC\" + EXT) = False Then
'                fs22.moveFile A1$ + B1$, "M:\#\#D\00 Pen Drives MOBILES\DSC\" + B1$
'                DONE = True
'                If fs22.FileExists("M:\#\#D\00 Pen Drives MOBILES\DSC\" + B1$) = True Then
'                    If Dir(A1$ + B1$) <> "" Then
'                        fs22.DeleteFile A1$ + B1$, True
'                    End If
'                End If
'            End If
'            If DONE = False Then
'                If fs22.FileExists("M:\#\#D\00 Pen Drives MOBILES\DSC\" + B1$) = True Then
'                    If Dir(A1$ + B1$) <> "" Then fs22.DeleteFile A1$ + B1$, True
'                End If
'                If fs22.FileExists("M:\#\#D\00 Pen Drives MOBILES\DSC\" + EXT) = True Then
'                    If Dir(A1$ + EXT) <> "" Then fs22.DeleteFile A1$ + EXT, True
'                End If
'            End If
        
        'M:\0 00 Art\# 00 My Pictures\01 All Phone\00 Mental\00 00 Hoss
        
'        A1$ = "M:\0 00 Art\# 00 My Pictures\01 All Phone\00 Mental\00 00 Hoss\"
        'B1$
'        If Len(A1$) > Len("M:\0 00 Art\# 00 My Pictures\01 All Phone\") Then
'            If InStr(A1$, "2008 0") > 0 Or InStr(A1$, "2007 0") > 0 Or InStr(A1$, "2009 0") > 0 Then
'        If A1$ = AT$ Then
'            H1$ = B1$
'            If InStr(B1$, "2009") = 1 Or InStr(B1$, "MILL VIEW") = 1 Then
'            If InStr(B1$, "2007") >= 1 Then
'
'            If InStr(A1$, "M:\0 00 Art\# 00 My Pictures\01 All Phone\00 Mental\2008\2008 Out an About") > 0 Then


'            If InStr(B1$, "2009 ") = 1 Then
                'GG$ = Replace(A1$, "\", "-")
''    '            GG$ = Mid$(GG$, 1 + Len("M:\0 00 ART WORK IN BOUND\Fantasy Art 1,800 Images\")) + B1$
                'GG$ = Mid$(GG$, 1 + Len("M:\0 00 Art\# 00 My Pictures\01 All Phone\")) + B1$
                
'                GoTo CLEARJUMP
'                H1$ = B1$

'                H1$ = Replace(H1$, "01 Brighton ", "")
'                H1$ = Replace(H1$, "02 Buses ", "")
'                H1$ = Replace(H1$, "03 Hove ", "")
'                H1$ = Replace(H1$, "04 Mates House ", "")
'                H1$ = Replace(H1$, "05 The Beach ", "")
'                H1$ = Replace(H1$, "06 The South Downs ", "")
                
'                H1$ = Mid(H1$, 1, 4) + "-" + Mid$(H1$, 7, 2) + "-" + Mid$(H1$, 15, 2) + " "
                 'H1$ = Mid(H1$, 1, 4) + "-" + Mid$(H1$, 6, 2) + "-" + Mid$(H1$, 9, 2) + " "
'                H1$ = H1$ + Mid(B1$, 22)
'                H1$ = Mid(B1$, 22)
                H1$ = B1$
                
'                H1$ = H1$ + Mid(B1$, InStrRev(B1$, "- ") - 17)
'                H1$ = GG$ + B1$

'
'                End If
                FG = FreeFile
                ds1$ = "    "
                ds2$ = "          "
                ds3$ = "          "
                DS4$ = Space$(4000)
                Open A1$ + B1$ For Binary As #FG
                    Get #FG, &H7, ds1$
                    Get #FG, &HB2, ds2$
                    Get #FG, &HAD, ds3$
                    Get #FG, &H1, DS4$
                    
                Close #FG
                WORK = 0
                
                '--------------------------------
                ds$ = UCase(Replace(ds2$, Chr(0), ""))
                If ds1$ = "JFIF" Then ds$ = "JFIF"
                If Len(ds$) < 4 Then ds$ = UCase(ds3$)
                
                '--------------------------------
                'CONTRADICT ABOVE
                ds$ = ""
                '--------------------------------
                
                'len(ds4$)
'                ds$ = Mid(ds$, 1, InStr(ds$, Chr(0)) - 1)

'                If InStr(DS4$, "JFIF") > 0 Then ds$ = "JFIF": WORK = 1
                
                If InStr(DS4$, "Sony Ericsson") > 0 And InStr(DS4$, "K800i") > 0 Then
                    ds$ = "Sony Ericsson K800i"
                    H1$ = Replace(H1$, " K800I ", "")
                    H1$ = Replace(H1$, " K800i ", "")
                    H1$ = Replace(H1$, " Sony Ericsson", "")
                    WORK = 1
                End If
                If InStr(DS4$, "Sony Ericsson") > 0 And InStr(DS4$, "W880i") > 0 Then
                    ds$ = "Sony Ericsson W880i"
                    H1$ = Replace(H1$, " W880I ", "")
                    H1$ = Replace(H1$, " W880i ", "")
                    H1$ = Replace(H1$, " Sony Ericsson", "")
                    WORK = 1
                End If
                If InStr(DS4$, "Sony Ericsson") > 0 And InStr(DS4$, "C905") > 0 Then
                    ds$ = "Sony Ericsson C905"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " C905 ", "")
                    H1$ = Replace(H1$, " Sony Ericsson", "")
                    WORK = 1
                End If
                If InStr(DS4$, "Sony Ericsson") > 0 And InStr(DS4$, "W595") > 0 Then
                    ds$ = "Sony Ericsson W595"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " W595 ", "")
                    H1$ = Replace(H1$, " Sony Ericsson", "")
'                    H1$ = Replace(H1$, " Sony Ericsson W595", "")
                    WORK = 1
                End If
                If InStr(DS4$, "DSC-W35") > 0 Then
                    ds$ = "SONY DSC-W35"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " SONYDSC-W35 ", "")
                    H1$ = Replace(H1$, " SONY DSC-W35 ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "BlackBerry") > 0 Then
                    ds$ = "BlackBerry 9000"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " BlackBerry 9000 ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "ViviCam") > 0 Then
                    ds$ = "ViviCam"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " ViviCam ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "NIKON D50") > 0 Then
                    ds$ = "NIKON D50"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " NIKON D50 ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "FinePix JZ") > 0 Then
                    ds$ = "FUJI FinePix JZ300"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " FinePix JZ ", "")
                    H1$ = Replace(H1$, " FUJI FinePix JZ300 ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "FinePix J100") > 0 Then
                    ds$ = "FUJI FinePix J100"
                    H1$ = Replace(H1$, " JFIF ", "")
                    'H1$ = Replace(H1$, " FinePix ", "")
                    H1$ = Replace(H1$, " FUJI FinePix J100 ", "")
                    WORK = 1
                End If
                
                
                If InStr(DS4$, "DSC-HX5V") > 0 Then
                    ds$ = "SONY DSC-HX5V"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " SONY DSC-HX5V ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "DSC-H7") > 0 Then
                    ds$ = "SONY DSC-H7"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " SONY DSC-H7 ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "USB Camera") > 0 Then
                    ds$ = "SYNTEK"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " SYNTEK ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "COOLPIX S640") > 0 Then
                    ds$ = "NIKON COOLPIX S640"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " NIKON COOLPIX S640 ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "Nokia") > 0 And InStr(DS4$, "5230") Then
                    ds$ = "Nokia"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " Nokia 5230 ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "Nokia") > 0 And InStr(DS4$, "E72-1") > 0 Then
                    ds$ = "Nokia E72-1"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " Nokia E72-1 ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "Picasa") > 0 Then
                    ds$ = "Picasa"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " Picasa ", "")
                    H1$ = Replace(H1$, " Exif 0210 ", "")
                    WORK = 1
                    Clipboard.Clear
                    Clipboard.SetText A1$ + B1$
                End If
                If InStr(DS4$, "SAMSUNG") > 0 And InStr(DS4$, "GT-P1000") > 0 Then
                    ds$ = "SAMSUNG GT-P1000"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " SAMSUNG GT-P1000 ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "SONY") > 0 And InStr(DS4$, "DSC-HX60V") > 0 Then
                    ds$ = "SONY DSC-HX60V"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " SONY DSC-HX60V ", "")
                    WORK = 1
                End If
                If InStr(DS4$, "OLYMPUS DIGITAL CAMERA") > 0 And InStr(DS4$, "FE330,X845,C550") > 0 Then
                    ds$ = "OLYMPUS DIGITAL CAMERA FE330,X845,C550"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " OLYMPUS DIGITAL CAMERA FE330,X845,C550 ", "")
                    WORK = 1
                End If
                If WORK = 0 And InStr(DS4$, "V233-00-01") > 0 Then
                    ds$ = "V233-00-01"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " V233-00-01 ", "")
                    WORK = 1
                End If
                If WORK = 0 And InStr(DS4$, String$(5, Chr(&H12))) > 0 Then
                    ds$ = "Coretta_"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " Coretta_ ", "")
                    WORK = 1
                End If
                If WORK = 0 And InStr(DS4$, "Exif") > 0 And InStr(DS4$, "0210") > 0 Then
                    ds$ = "Exif 0210"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " Exif 0210 ", "")
                    WORK = 1
                End If
                '0210
                'DSC-H7
                'COOLPIX S64
                'Nokia
                'Picasa
                'SAMSUNG GT-P1000
                'SONY DSC-HX60V
                'OLYMPUS DIGITAL CAMERA FE330,X845,C550
                'V233-00-01 - B1$ - "MILL VIEW HOSPITAL-2008-12-08 22-08-30 - V233-00-01 - MEDI0002.JPG"
                'a1$
'                If B1$ = "2009-09-22 23-09-55 - DSC03047.JPG" Then Stop
                
                'If WORK = 0 Then Stop
                'If WORK = 0 And (InStr(ds3$, Chr(0)) > 0 Or InStr(ds2$, Chr(0)) > 0) Then ds$ = DS1$
                'If WORK = 0 Then ds$ = "JFIF"
                If WORK = 0 Then ds$ = GetExifModel(A1 + B1$)
'                If WORK = 0 Then Stop
                If WORK = 0 And ds$ = "" Then ds$ = "JFIF"
                If ds$ = "" Then Stop
                
               If InStr(H1$, ds$) = 0 Or 1 = 1 Then
                    
                    H1$ = Replace(H1$, "AYMER ROAD-ROUTE ONE-BHT-", "")
                    
                    H1$ = Replace(H1$, "--", "-")
                    'H1$ = Mid(H1$, 1, 19) + " - " + ds$ + " " + Mid(H1$, InStr(H1$, "- "))
'                    If InStr(H1$, "-") > 0 Then
'                        H1$ = GG$ + ds$ + " " + Mid(H1$, InStr(H1$, "- "))
'                    Else
'                        H1$ = GG$ + ds$ + " - " + B1$
'                    End If'A1$
                    XB1$ = B1$

                    If InStrRev(XB1, "ddie  lancaster  - ") > 0 Then
'                        Stop
                        XB1$ = Replace(XB1$, "ddie  lancaster  - ", "ddie Lancaster_")
                        'xb1$ = Replace(xb1$, "SC0", "DSC0")
                    End If


'                    If B1$ = "DSC00201.JPG" Then Stop
                    If InStrRev(XB1$, " -") > 0 Then
                        XB1$ = Mid(B1$, InStrRev(B1$, " -") + 3)
                    Else
'                    Stop
                    End If
                    
                    If InStrRev(B1$, " SC0") > 0 Then
'                        Stop
                        XB1$ = Replace(XB1$, "SC0", "DSC0")
                        'xb1$ = Replace(xb1$, "SC0", "DSC0")
                    End If
                    If InStrRev(B1$, " 9DSC0") > 0 Then
'                        Stop
                        XB1$ = Replace(XB1$, "9DSC0", "DSC0")
                        'xb1$ = Replace(xb1$, "SC0", "DSC0")
                    End If
                    If InStrRev(B1$, "Icture(") > 0 Then
'                        Stop
                        XB1$ = Replace(XB1$, "Icture(", "Picture(")
                        'xb1$ = Replace(xb1$, "SC0", "DSC0")
                    End If
                    
                    If ds$ = "NIKON COOLPIX S640" And InStr(XB1$, " ") > 0 Then
                        For RCD = 0 To 9
                        XB1$ = Replace(XB1$, Trim(Str(RCD)) + " ", "")
                        Next
                    End If
                    
                    
                    
                    H1$ = GG$ + ds$ + " - " + XB1$
                    
                End If
                
                
                
'                If InStr(ds$, "Sony Ericsson W595") Then
'                    If InStrRev(B1$, "-") > 0 Then
'                        xb1$ = Mid(B1$, InStrRev(B1$, "-") + 2)
'                    End If
'                    If Mid(xb1$, 1, 2) = "SC" Then
'                        xb1$ = Replace(xb1$, "SC", "DSC")
'                    End If
'                    H1$ = GG$ + ds$ + " - " + xb1$
'                End If
'
'                If InStr(ds$, "FUJI FinePix JZ300") Then
'                    xb1$ = Mid(B1$, InStrRev(B1$, "-") + 2)
'                    H1$ = GG$ + ds$ + " - " + xb1$
'                End If
'
'                If InStr(ds$, "NIKON") Then
'                    GG$ = Format$(F.datelastmodified, "YYYY-MM-DD HH-MM-SS") + " - " ' - "
'
'                    xb1$ = Mid(B1$, InStrRev(B1$, "-") + 2)
''                    xb1$ = B1$
'                        xgg$ = Format$(F.datelastmodified, "YYYY-MM-DD") + " "
'                        If InStr(xb1$, xgg$) > 0 Then
'                            xb1$ = Replace(xb1$, xgg$, "")
'                        End If
'                    If InStr(xb1$, "- " + Trim(Str(rz)) + " ") > 0 Then Stop
'                    For rz = 0 To 9
'                        xb1$ = Replace(xb1$, "- " + Trim(Str(rz)) + " ", "- ")
'                    Next
'
'                    H1$ = GG$ + ds$ + " - " + xb1$
'                End If
'                'Stop
'
'                If InStr(H1$, "- SC") Then
'                    'Stop
'                    H1$ = Replace(H1$, "- SC", "- DSC")
'                End If
                
                
                GG$ = H1$
                GG$ = Replace(GG$, "MILL VIEW HOSPITAL_AYMER ROAD_LONDON-", "")
                GG$ = Replace(GG$, "MILL VIEW HOSPITAL-", "")
                GG$ = Replace(GG$, "AYMER ROAD-ROUTE ONE-BHT-", "")
                'GG$ = "AYMER ROAD-ROUTE ONE-BHT-" + GG$
                GG$ = Replace(GG$, "  ", " ")
                
                
CLEARJUMP:
                
                
'                GG$ = B1$
                GG$ = Replace(GG$, "MEADOWFIELD HOSPITAL", UCase("MILL VIEW HOSPITAL"))
'
                
'                GG$ = Replace(GG$, "FANTASY ART-", "")
    '            GG$ = Mid$(GG$, 13)
'                GG$ = LCase(GG$)
                GG$ = UCase(Mid(GG$, 1, 1)) + Mid(GG$, 2)
                For R = 0 To 25
                'Debug.Print Chr(R + 97)
                    GG$ = Replace(GG$, " " + Chr(R + 97), " " + UCase(Chr(R + 97)))
                    GG$ = Replace(GG$, "-" + Chr(R + 97), "-" + UCase(Chr(R + 97)))
                    GG$ = Replace(GG$, "_" + Chr(R + 97), "_" + UCase(Chr(R + 97)))
                    GG$ = Replace(GG$, "(" + Chr(R + 97), "(" + UCase(Chr(R + 97)))
                    GG$ = Replace(GG$, "[" + Chr(R + 97), "[" + UCase(Chr(R + 97)))
                    GG$ = Replace(GG$, "&" + Chr(R + 97), "&" + UCase(Chr(R + 97)))
                    GG$ = Replace(GG$, "#" + Chr(R + 97), "#" + UCase(Chr(R + 97)))
                    For R1 = 0 To 9
                        GG$ = Replace(GG$, Chr(R1 + 48) + Chr(R + 97), Chr(R1 + 48) + UCase(Chr(R + 97)))
                        GG$ = Replace(GG$, Chr(R1 + 48) + "X", Chr(R1 + 48) + "x")
                    Next
                Next
                'Ericsson Don't Cap the i
                GG$ = Replace(GG$, "0I - ", "0i - ")
                
                'GG$ = Replace(GG$, "Sony Ericsson W880I", "Sony Ericsson W880i")
                'GG$ = Replace(GG$, "", "")


'                If InStr(B1$, "  ") Then Stop
'                On Error Resume Next
'                If Dir(A1$ + GG$) <> "" Then
'                If B1$ <> GG$ And UCase(B1$) <> UCase(GG$) Then
'                    Set fs1 = fs22.GetFile(A1$ + B1$)
'                    Set fs2 = fs22.GetFile(A1$ + GG$)
'                    If fs2.Size <= fs1.Size Then
'                 '       fs2.Delete
'                    Else
'                  '   fs1.Delete
'                    End If
'                End If
'                End If
                
JUMP_RENAME:
                
                'GG$ = B1$
                '
                'GG$ = Replace(GG$, "Copy (2) of ", "")
                'GG$ = Replace(GG$, "Copy of ", "")
                'GG$ = Replace(GG$, " COPY", "")
                'GG$ = Replace(GG$, "_Old1", "")
                'GG$ = Replace(GG$, "_Old2", "")
                'GG$ = Replace(GG$, "_Old3", "")
                'GG$ = Replace(GG$, "_Old4", "")
                'GG$ = Replace(GG$, "_New1", "")
                'GG$ = Replace(GG$, "MEADOWFIELD HOSPITAL-", "")
                
'                GG$ = B1$
'                gg1$ = Mid(B1$, 7, 2)
'                gg2$ = Mid(B1$, 4, 2)
'                gg3$ = Mid(B1$, 1, 2)
'
'                Mid(GG$, 1, 2) = gg1$
'                Mid(GG$, 4, 2) = gg2$
'                Mid(GG$, 7, 2) = gg3$
                
'                If GG$ = "07" Then
                If Dir(A1$ + B1$) <> "" Or B1$ <> GG$ Then
                    
                    'Kill the Smallest or Leave
'                    If Dir(A1$ + GG$) <> "" And UCase(B1$) <> UCase(GG$) Then
'                        Set F = fs22.GetFile(A1$ + B1$)
'                        fs1 = F.Size
'
'                        Set F = fs22.GetFile(A1$ + GG$)
'                        fs2 = F.Size
'                        If fs2 < fs1 Then
'                            'Kill A1$ + GG$
'                        End If
'                    End If
                    
                    'Has To Be -Or- Not -And- or Lcase Ucase Don't Process
                    If Dir(A1$ + GG$) = "" Or B1$ <> GG$ Then
                        'GG$ = Replace(B1$, "_", "D")
'                        If Dir(A1$ + GG$) <> "" And B1$ <> GG$ Then
'
'                            Kill GG$
'                        End If
                        If Dir(A1$ + GG$) <> "" And UCase(B1$) <> UCase(GG$) Then
                            'MsgBox ("File " + GG$ + " Aready Exists" + vbCrLf + A1$ + B1$)
                        Else
                            Name A1$ + B1$ As A1$ + GG$
                        End If
                    'fs22.moveFile A1$ + B1$, A1$ + GG$
                    End If
                End If
                
                MARK = 0
                If A1$ <> AT$ Then
'                If fs22.FileExists("M:\0 00 Art\# 00 My Pictures\01 All Phone\00 Mental\2008\" + GG$) = True Then
                '    Kill A1$ + GG$
 '                   MARK = 1
  '              End If
                End If
         
                'B1$ = GG$
                If A1$ <> AT$ And MARK = 0 Then
'                    fs22.COPYFile A1$ + B1$, "M:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\" + GG$
                End If
                'End If
'            End If
        End If
        
        
JumpToNext:
        
        
'    End If
'End If
    
'        End If
'    End If
    
    Label13 = Trim(Str(WE))
    DoEvents
    

Next
'End


'------------------------------------------------
'------------------------------------------------
'RENAME FOLDER SELECT AND THEN RUN THE SUBROUTINE
'------------------------------------------------
'------------------------------------------------
result = MsgBox("Do You Want to Rename the Photo Image Folders", vbYesNo + vbMsgBoxSetForeground)
If result = vbYes Then
    Call Rename_Image_Folder_Proper
End If
'------------------------------------------------
'------------------------------------------------



MsgBox ("Done")


Exit Sub
'------------------------------------------------
'------------------------------------------------
'------------------------------------------------
'------------------------------------------------


'Put in Sub Of Code
Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32
'WxHex$ = Hex(m_CRC.CalculateString(Form3.Text1.Text))
'WxHex$ = Hex(m_CRC.CalculateFile(Filename))

txtPath.Text = "D:\VB6 Icons\Ripped"
txtPath = "M:\00 Archive Work"
txtPath = "M:\00 Installs"

If Right$(txtPath, 1) <> "\" Then txtPath.Text = txtPath.Text + "\"
chkSubFolders.Value = vbChecked
cboMask.Text = "*.*"

'GoTo jumphere

Call cmdScan_Click
    
A1$ = ListView1.ListItems.Item(1).SubItems(1)
ID$ = Mid$(A1$, Len(txtPath) + 1)
ID$ = Mid$(ID$, 1, InStr(ID$, "\"))

Dim OnlyIfSameName
OnlyIfSameName = True
    
    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True        'Use default sorting to sort the
    ListView1.SortKey = 0
    ii = ListView1.ListItems.Count

'Exit Sub

If OnlyIfSameName = True Then
tt = 1
For we2 = ListView1.ListItems.Count To 3 Step -1
    Label13.Caption = Str(WE)
    DoEvents
    A1$ = ListView1.ListItems.Item(WE - 2).SubItems(1)
    B1$ = LCase(ListView1.ListItems.Item(WE - 2))
    a2$ = ListView1.ListItems.Item(WE - 1).SubItems(1)
    b2$ = LCase(ListView1.ListItems.Item(WE - 1))
    a3$ = ListView1.ListItems.Item(WE).SubItems(1)
    b3$ = LCase(ListView1.ListItems.Item(WE))
    
    On Error Resume Next
    Err.Clear
    Set ff1 = fs22.GetFile(A1$ + ListView1.ListItems.Item(WE - 2))
    Set ff2 = fs22.GetFile(a2$ + ListView1.ListItems.Item(WE - 1))
    Set ff3 = fs22.GetFile(a3$ + ListView1.ListItems.Item(WE))
    exiterr = Err.Number
    On Error GoTo 0
    

    
    topp = 0
    If b2$ <> B1$ And b2$ <> b3$ Then
        topp = 1
        ListView1.ListItems.Remove (WE - 1)
        ScanPath.lblCount = ListView1.ListItems.Count
    End If
    If exiterr = 0 And topp = 0 And ff2.Size <> ff1.Size And ff2.Size <> ff3.Size Then
        ListView1.ListItems.Remove (WE - 1)
        ScanPath.lblCount = ListView1.ListItems.Count
    End If
    
    
    
    Next
End If
WE = ListView1.ListItems.Count
b2$ = LCase(ListView1.ListItems.Item(WE - 1))
b3$ = LCase(ListView1.ListItems.Item(WE))
If b2$ <> b3$ Then
    ListView1.ListItems.Remove (WE)
    ScanPath.lblCount = ListView1.ListItems.Count
End If
WE = 1
b2$ = LCase(ListView1.ListItems.Item(WE + 1))
b3$ = LCase(ListView1.ListItems.Item(WE))
If b2$ <> b3$ Then
    ListView1.ListItems.Remove (WE)
    ScanPath.lblCount = ListView1.ListItems.Count
End If


'Exit Sub

For WE = 1 To ListView1.ListItems.Count
    DoEvents
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    
    WxHex$ = Space$(8)
    On Error Resume Next
    LSet WxHex$ = Hex(m_CRC.CalculateFile(A1$ + B1$))
    On Error GoTo 0
    
    'Include filenumber then when sort always first file never deleted
    ListView1.ListItems.Item(WE).SubItems(1) = WxHex$ + " - - " + Format$(WE, "0000000") + " - - " + A1$
    ListView1.SelectedItem = ListView1.ListItems(WE)
    ListView1.SelectedItem.EnsureVisible
    Label13.Caption = Str(WE)

Next
    
    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True        'Use default sorting to sort the
    ListView1.SortKey = 1 ' Sort on Column 2
    ii = ListView1.ListItems.Count
    

For WE = 1 To ListView1.ListItems.Count
    DoEvents
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
'    If b1$ = "Serial.txt" Then Stop
    'a1$ = ListView1.ListItems.Item(we - 10).SubItems(1)
    tf1$ = Mid$(A1$, InStr(A1$, txtPath) + Len(txtPath))
    If InStrRev(tf1$, "\", Len(tf1$) - 1) > 0 Then
        f1$ = Mid$(tf1$, InStrRev(tf1$, "\", Len(tf1$) - 1) + 1)
    Else
        f1$ = "\"
    End If
    
    ListView1.ListItems.Item(WE) = f1$ + B1$
    ListView1.SelectedItem = ListView1.ListItems(WE)
    ListView1.SelectedItem.EnsureVisible
    Label13.Caption = Str(WE)
    
Next

jumphere:

    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True        'Use default sorting to sort the
    ListView1.SortKey = 1
    ii = ListView1.ListItems.Count

For WE = 2 To ListView1.ListItems.Count
    'Exit For
    DoEvents
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = Mid$(ListView1.ListItems.Item(WE), InStr(ListView1.ListItems.Item(WE), "\") + 1)
    
    
    a12$ = ListView1.ListItems.Item(WE - 1).SubItems(1)
    b12$ = Mid$(ListView1.ListItems.Item(WE - 1), InStr(ListView1.ListItems.Item(WE - 1), "\") + 1)
    If LCase(B1$) <> LCase(b12$) Then xc = 0
    If LCase(B1$) = LCase(b12$) And xc = 0 Then
        xc = WE
    End If
    FG = xc - 1
    If xc = 0 Then FG = WE - 1
        
'    If we - fg > 1 Then Stop
    
    For we3 = FG To WE
    For we4 = FG To WE
    DoEvents
    If we4 = we3 Then Exit For
    A1$ = ListView1.ListItems.Item(we3).SubItems(1)
    B1$ = Mid$(ListView1.ListItems.Item(we3), InStr(ListView1.ListItems.Item(we3), "\") + 1)
    
    a12$ = ListView1.ListItems.Item(we4).SubItems(1)
    b12$ = Mid$(ListView1.ListItems.Item(we4), InStr(ListView1.ListItems.Item(we4), "\") + 1)

    If LCase(B1$) <> LCase(b12$) Then Exit For


    tgs = 8 'Just Chk CRC

    g1$ = Mid$(A1$, 1, tgs)
    g3$ = Mid$(A1$, 26)
    
    H1$ = Mid$(a12$, 1, tgs)
    h3$ = Mid$(a12$, 26)

    On Error Resume Next
    'err.number
'    Set ff1 = fs22.GetFile(Mid$(a1$, 26) + b1$)
    If Err.Number > 0 Then
        qb = 0
    End If
'    Set ff2 = fs22.GetFile(Mid$(a12$, 26) + b12$)
    If Err.Number > 0 Then
        qb = 0
    End If
    
    qb = 0
'    If ff1.Size = 0 And ff2.Size = 0 Then
'        qb = 1
'    Else
'    Stop
'    End If
    If Err.Number > 0 Then
        qb = 0
    End If
    On Error GoTo 0
    
    'CRC Check
    If g1$ = H1$ Then
        qb = 1
    End If
        
    If LCase(B1$) <> LCase(b12$) Then qb = 0
    If Trim(g1$) = "" Or Trim(H1$) = "" Then qb = 0
    
    'if sub dir parent from file not same then filter for no del
    If qb = 1 Then
    f1$ = LCase(Mid$(A1$, InStrRev(A1$, "\", Len(A1$) - 1)))
    f2$ = LCase(Mid$(a12$, InStrRev(a12$, "\", Len(a12$) - 1)))
    ww = 0
    If Mid$(ListView1.ListItems.Item(we3), 1, 1) = "\" And _
        Mid$(ListView1.ListItems.Item(we4), 1, 1) = "\" Then ww = 1
    
    'Root folder
    If LCase(f1$) <> LCase(f2$) And ww = 0 Then
        qb = 0
    End If
    End If

    
    
    If InStr(g3$, ID$) > 0 And qb = 1 Then
        g3$ = h3$
    End If
    If InStr(g3$, ID$) > 0 And qb = 1 Then
        qb = 0
    End If
    
    'Exit For
    'qb = 0
    If InStr(LCase(A1$), "_vti_cnf") > 0 Then qb = 1
    If InStr(LCase(A1$), "_vti_pvt") > 0 Then qb = 1
    
    'Make sure none are deleted from same main dir
    'Root folder
    If qb = 1 Then
        tf1$ = Mid$(A1$, InStr(A1$, txtPath) + Len(txtPath))
        f1$ = Mid$(tf1$, 1, InStr(tf1$, "\"))
        tf1$ = Mid$(a12$, InStr(a12$, txtPath) + Len(txtPath))
        f2$ = Mid$(tf1$, 1, InStr(tf1$, "\"))
        If f1$ = f2$ Then
            qb = 0
        End If
    End If
    
    
    
    
    If qb = 1 Then
        
        
        On Error Resume Next
        tagger = 0
        If fs22.FileExists(g3$ + B1$) Then Kill g3$ + B1$: aga = aga + 1: tagger = 1
        If Err.Number > 0 Then
'            MsgBox "Error Deleting"
            Err.Clear
            fs22.DeleteFile g3$ + B1$, True
            aga = aga + 1: tagger = 1
            If Err.Number > 0 Then
'                MsgBox "Cant Delete" + vbCrLf + g3$ + b1$
         End If
        
        Err.Clear
        If tagger = 1 Then
            Set ff1 = fs22.GetFolder(g3$)
            If ff1.Size = 0 And Err.Number = 0 Then
                ff1.Delete True
            End If
        End If
        
        End If
        On Error GoTo 0
        
        Label14.Caption = Str(aga)
        Label13.Caption = Str(ii - aga)
    End If


    
    ListView1.SelectedItem = ListView1.ListItems(WE)
    ListView1.SelectedItem.EnsureVisible
    Label13.Caption = Str(WE)

Next
Next
Next

MsgBox "End" + vbCrLf + "Deleted =" + Str(aga)
    
'    ListView1.SortOrder = lvwAscending
'    ListView1.Sorted = True        'Use default sorting to sort the
'    ListView1.SortKey = 0
'    ii = ListView1.ListItems.Count

Exit Sub

'End

End Sub


Sub Rename_Image_Folder_Proper()

'W$ = "D:\DSC\DCIM\"
AT$ = W$

If AT$ <> "" And Mid$(AT$, 1, 1) = """" Then
    AT$ = Mid$(AT$, 2)
    AT$ = Mid$(AT$, 1, Len(AT$) - 1)
End If

If Len(AT$) < 4 Then MsgBox "Not on a Root Dir - END": End

cboMask.Text = "*.*"

txtPath = AT$
ListView1.ListItems.Clear
chkSubFolders.Value = vbChecked
Call cmdScan_Click
Dim LastPath, OldPath

For WE = 1 To ListView1.ListItems.Count
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    ListView1.ListItems.Item(WE).EnsureVisible
    
    If OldPath <> A1$ Then
        OldPath = A1$
    
        LastPath = Mid(A1$, InStrRev(A1$, "\", Len(A1$) - 1))
        LastPath = Replace(LastPath, "\", "")
        FrontofLastPath = Mid(A1$, 1, InStrRev(A1$, "\", Len(A1$) - 1))

        
        'NORMAL IMPORT FROM SDCARD CAMERA
        If IsNumeric(LastPath) = True And Mid(LastPath, 1, 3) <> "201" Then 'And Len(LastPath) <> 8 Then
'            LastPath = Mid(LastPath, 1, 8)
            LastPath = "201" + Mid(LastPath, 4)
            If Dir(FrontofLastPath + LastPath, vbDirectory) = "" And FrontofLastPath + LastPath + "\" <> A1$ Then
                
                Name A1$ As FrontofLastPath + LastPath
                wx = 0
            Else
                wx = wx + 1
                '------------------------------------------------------------------
                'IF WE PUT THE DASH IN BEFORE NUMERIC ADD END
                '------------------------------------------------------------------
                'AND THEN THAT WORK SORT THE NEW FOLDER BETTER ALONGSIDE EACH OTHER
                'RATHER THAN WITHOUT THE EXTRA CHAR NUMERIC MAKE SORT AT END
                '------------------------------------------------------------------
                Name A1$ As FrontofLastPath + LastPath + "-" + Trim(Str(wx))
            End If
        End If
        
        'BEEN IMPORTED BY WIFI
        If Mid(LastPath, 7, 3) = "201" Then
'            LastPath = Mid(LastPath, 1, 8)
            LastPath = Mid(LastPath, 7, 4) + Mid(LastPath, 4, 2) + Mid(LastPath, 1, 2)
            If Dir(FrontofLastPath + LastPath, vbDirectory) = "" And FrontofLastPath + LastPath + "\" <> A1$ Then
                
                Name A1$ As FrontofLastPath + LastPath
                wx = 0
            Else
                '------------------------------------------------------------------
                'IF WE PUT THE DASH IN BEFORE NUMERIC ADD END
                '------------------------------------------------------------------
                'AND THEN THAT WORK SORT THE NEW FOLDER BETTER ALONGSIDE EACH OTHER
                'RATHER THAN WITHOUT THE EXTRA CHAR NUMERIC MAKE SORT AT END
                '------------------------------------------------------------------
                wx = wx + 1
                Name A1$ As FrontofLastPath + LastPath + "-" + Trim(Str(wx))
            End If
        End If
    End If
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



'End
'If Command$ <> "" Then End
cmdScan.Caption = "Scan"
End Sub



Private Sub Form_Load()
    
ScanPath.Caption = ScanPath.Caption + " -- " + App.EXEName
ListView1.Left = 0

ScanPath.Height = ListView1.Top + ListView1.Height + 550
ScanPath.Width = ListView1.Width + 180
  
  
    
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
        .ColumnHeaders.Add , "PATH", "Path", 4000, lvwColumnLeft
        .ColumnHeaders.Add , "SIZE", "Size", 2000, lvwColumnRight
        .ColumnHeaders.Add , "DATE", "Date", 2000, lvwColumnLeft
        
        .View = lvwReport
    End With

ScanPath.Show

Call AutoPix

Exit Sub

End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_Click()

Exit Sub

Dim fs22
Set fs22 = CreateObject("Scripting.FileSystemObject")


tpath3$ = "E:\04 Music ---\Del\"

B1$ = Mid$(List1.List(List1.ListIndex), 14)
c1$ = B1$ + "--Todel"

'tpath1$ = "E:\04 Music ---\"

c1$ = Mid$(B1$, InStrRev(B1$, "\") + 1)
List1.RemoveItem (List1.ListIndex)
'Name b1$ As c1$

'fs22.copyFile b1$, tpath3$ + c1$
If Dir$(tpath3$ + c1$) <> "" Then Kill tpath3$ + c1$
fs22.MoveFILE B1$, tpath3$ + c1$



End Sub

Private Sub Label20_Click()
    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True        'Use default sorting to sort the
    ListView1.SortKey = 0
    ii = ListView1.ListItems.Count

'Call AutoPix

End Sub

Private Sub Label21_Click()

    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True        'Use default sorting to sort the
    ListView1.SortKey = 1
    ii = ListView1.ListItems.Count

End Sub

Private Sub SP_DirMatch(Directory As String, Path As String)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################
End Sub

Private Sub SP_FileMatch(fileName As String, Path As String)
    Dim LV As ListItem
    Dim uWIN32 As WIN32_FIND_DATA
    
    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
    On Local Error GoTo GetFileError
    
    With ListView1
        Set LV = .ListItems.Add(, , fileName)
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
    fs22.MoveFILE A1$ + B1$, c1$ + B1$
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

    List1.AddItem Format$(WE, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'Fs22.DeleteFolder comrad$ + "\" + WFD$

Next

A = A
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
    fs22.MoveFILE c2$ + B1$, A1$ + B1$
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
    fs22.copyfolder Drived2$ + "Temp\Anything", v1$, True
    
    Set fs22 = CreateObject("Scripting.FileSystemObject")
    fs22.deletefolder Drived2$ + "Temp\Anything", True


jeep:
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

