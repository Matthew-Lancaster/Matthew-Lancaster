VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ScanPath 
   BackColor       =   &H8000000A&
   Caption         =   "ScanPath 2.0 - "
   ClientHeight    =   8484
   ClientLeft      =   60
   ClientTop       =   648
   ClientWidth     =   12648
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8484
   ScaleWidth      =   12648
   Begin VB.Timer TIMER_COUNTDOWN_COMMAND_BUTTON_AUTO_GO_AH 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3396
      Top             =   1464
   End
   Begin VB.CommandButton SPECIAL_RENAME_Command_BUTTON 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SPECIAL RENAME"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   8976
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   1740
      Width           =   4416
   End
   Begin VB.Timer FORM_LOAD_Timer 
      Interval        =   1
      Left            =   2580
      Top             =   1416
   End
   Begin VB.CheckBox Check_Dont_Do_Already_Modified 
      BackColor       =   &H8000000E&
      Caption         =   "Also Do What Already Modified"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8988
      TabIndex        =   44
      Top             =   1392
      Width           =   4368
   End
   Begin VB.Timer Timer_1_MINUTE 
      Interval        =   1000
      Left            =   2988
      Top             =   1428
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "WHEN READY TO GO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   10056
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   520
      Width           =   3336
   End
   Begin VB.FileListBox File1 
      Height          =   264
      Left            =   2712
      TabIndex        =   42
      Top             =   6156
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   585
      Top             =   4752
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Height          =   825
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3876
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
      TabIndex        =   35
      Top             =   3876
      Visible         =   0   'False
      Width           =   7245
   End
   Begin VB.ComboBox cboMask 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   740
      TabIndex        =   1
      Text            =   "*.jpg"
      Top             =   520
      Width           =   5928
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   396
      Left            =   12852
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   84
      Width           =   516
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   750
      TabIndex        =   0
      Text            =   "E:\My Music Zen"
      Top             =   60
      Width           =   12012
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "Search Sub-Folders"
      Height          =   195
      Left            =   1530
      TabIndex        =   8
      Top             =   2280
      Width           =   1785
   End
   Begin VB.CheckBox chkPatternMatching 
      Caption         =   "Pattern Matching"
      Height          =   195
      Left            =   1530
      TabIndex        =   7
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkNormal 
      Caption         =   "Normal"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   2736
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   2964
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.ComboBox cboDate 
      Height          =   288
      Left            =   5940
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2220
      Width           =   1545
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   288
      Index           =   0
      Left            =   9930
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2604
      Width           =   1365
   End
   Begin VB.ComboBox cboSize 
      Height          =   288
      Left            =   8580
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2172
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   8580
      TabIndex        =   16
      Top             =   2604
      Width           =   1305
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   288
      Index           =   1
      Left            =   9930
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2964
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   8580
      TabIndex        =   18
      Top             =   2964
      Width           =   1305
   End
   Begin VB.CheckBox chkRefreshListView 
      Caption         =   "Refresh ListView (slower)"
      Height          =   195
      Left            =   1530
      TabIndex        =   10
      Top             =   2820
      Width           =   2205
   End
   Begin VB.CheckBox chkCopyMemory 
      Caption         =   "Use CopyMemory (display Date && Size)"
      Height          =   195
      Left            =   1530
      TabIndex        =   11
      Top             =   3060
      Width           =   3465
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4440
      Left            =   48
      TabIndex        =   21
      Top             =   3420
      Width           =   11916
      _ExtentX        =   21019
      _ExtentY        =   7832
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
      Top             =   2520
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   825
      Left            =   12060
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3468
      Visible         =   0   'False
      Width           =   810
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   312
      Index           =   0
      Left            =   5940
      TabIndex        =   13
      Top             =   2604
      Width           =   1548
      _ExtentX        =   2709
      _ExtentY        =   550
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   147521537
      CurrentDate     =   37299
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   312
      Index           =   1
      Left            =   5940
      TabIndex        =   14
      Top             =   2964
      Width           =   1548
      _ExtentX        =   2709
      _ExtentY        =   550
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   147521537
      CurrentDate     =   37296
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort 02"
      Height          =   240
      Left            =   12072
      TabIndex        =   41
      Top             =   4596
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort 01"
      Height          =   240
      Left            =   12072
      TabIndex        =   40
      Top             =   4332
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Count"
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
      Height          =   384
      Left            =   6744
      TabIndex        =   39
      Top             =   516
      Width           =   1500
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Count ah"
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
      Height          =   384
      Left            =   6744
      TabIndex        =   38
      Top             =   948
      Width           =   1500
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   384
      Left            =   8280
      TabIndex        =   37
      Top             =   948
      Width           =   1740
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
      Height          =   192
      Left            =   5136
      TabIndex        =   34
      Top             =   1800
      Width           =   1416
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
      Height          =   192
      Left            =   1536
      TabIndex        =   26
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mask"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   96
      TabIndex        =   24
      Top             =   540
      Width           =   636
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   96
      TabIndex        =   22
      Top             =   84
      Width           =   636
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
      Height          =   192
      Left            =   96
      TabIndex        =   25
      Top             =   1800
      Width           =   888
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      Height          =   192
      Index           =   0
      Left            =   7776
      TabIndex        =   31
      Top             =   2664
      Width           =   252
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
      Height          =   192
      Left            =   7776
      TabIndex        =   29
      Top             =   2244
      Width           =   588
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   192
      Index           =   1
      Left            =   7776
      TabIndex        =   33
      Top             =   3024
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   192
      Left            =   5136
      TabIndex        =   30
      Top             =   2664
      Width           =   348
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "File Date"
      Height          =   192
      Left            =   5136
      TabIndex        =   28
      Top             =   2244
      Width           =   636
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   192
      Left            =   5136
      TabIndex        =   32
      Top             =   2964
      Width           =   192
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   384
      Left            =   8280
      TabIndex        =   27
      Top             =   516
      Width           =   1740
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB_ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_RENAME_FROM_TEXT_FILE 
      Caption         =   "DO A RENAME FROM READ TXT FILE"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_PASTE_PATH_IN 
      Caption         =   "CLIPBOARD PASTE PATH IN"
   End
   Begin VB.Menu MNU_COMMAND_PATH_IN 
      Caption         =   "COMMAND$ PATH IN"
   End
   Begin VB.Menu MNU_NOKIA_PATH_IN 
      Caption         =   "NOKIA PATH IN"
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================
'# __ D:\VB6\VB-NT\00_Send_To\Send To Date Of File In Front FileName\#0 Send To Date Of File In Front FileName.vbp
'# __
'# __ #0 Send To Date Of File In Front FileName.vbp
'# __
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __
'# START     TIME [ ?                        ]
'# END FINAL TIME [ Fri 15-Jun-2018 03:07:45 ]
'# __
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION
'# ------------------------------------------------------------------

' 000
'# ------------------------------------------------------------------
' Dim Yes_To_Go_Rename_Folder
' FUNCTION Rename_Image_Folder_Proper_Checker = True
'# ------------------------------------------------------------------
' 2017 DECEMBER
'# ------------------------------------------------------------------

' 001
'# ------------------------------------------------------------------
' SESSSION
' WORK TO MAKE MULTI FOLDERING _ STILL HARDCODED
' BY ADD AND ARRAY IN AND MSGBOX
' LINKED WITH THE VBS AND AUTOHOTKEYS CAMERA REEL OFFLOADER
'# ------------------------------------------------------------------
' START     TIME [ Fri 15-Jun-2018 01:11:00 ]
' END       TIME [ Fri 15-Jun-2018 03:14:00 ] 2 HOURS
'# ------------------------------------------------------------------

'002
'# ------------------------------------------------------------------
' MORE WORK ADD FUJI CAMERA AND MAKE GOOD BETTER
' MORE QUESTION ON MESSAGE BOX SKIP JUMP TO NEXT ONE RATHER THAN INFO ONLY
' AND REUSE CODE
'# ------------------------------------------------------------------
' START     TIME [ Sat 16-Jun-2018 11:33:52 ]
' END       TIME [ Sat 16-Jun-2018 12:11:28 ] 38 MINUTE
'# ------------------------------------------------------------------

Dim COUNT_DOWN_TIMER_COMMAND_BUTTON

Dim IM_2
Dim ONE_DONE

Dim Z_COUNTER
Dim A_COUNTER

Dim ROUNTINE_POINTER

Dim ARRAY_I()
Dim ARRAY_DESC()

Public EXIT_TRUE

Public PASTE_PATH_SET

'PROGRAM BEGINNER ___ Sub AutoPix()
'Auto---Pix

Dim FS22
Dim fs

Dim XVB_DATE

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

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Const NULL_CHAR = 0

Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function


Sub AutoPix()

Set FS22 = CreateObject("Scripting.FileSystemObject")
Set fs = CreateObject("Scripting.FileSystemObject")

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
W$ = "D:\DSC\2015 2016\2016 Sony CyberShot DSC-HX60V\DCIM2\"

W$ = "D:\DSC\2015 2016\2016 Sony CyberShot HX60V\DCIM2\"
W$ = "D:\DSC\2015 2016\2016 Sony CyberShot HX60V\DCIM2\"
W$ = "D:\DSC\2015 2017\2016 Sony CyberShot HX60V\DCIM2\"
'W$ = "D:\DSC\2015 2017\2017 Sony CyberShot HX60V\DCIM\"

W$ = "D:\DSC\2015 2017\2017 Sony CyberShot HX60V\DCIM"
W$ = "D:\DSC\2015 2017\2017 Sony CyberShot HX60V\DCIM"
'W$ = "D:\DSC EVER\DCIM"
W$ = "D:\DSC\2015 2017\2016 CyberShot HX60V\DCIM"

W$ = "D:\DSC\2015 2017\2016 CyberShot HX60V\TEST"
W$ = "D:\VI_ DSC ME\DCIM"

W$ = "D:\DSC\2015 2017\2017 CyberShot HX60V\DCIM"

W$ = "\\7-asus-gl522vw\7_asus_gl522vw_02_d_drive\DSC\2015 2018\2017 CyberShot HX60V\DCIM"
W$ = "\\7-asus-gl522vw\7_asus_gl522vw_02_d_drive\DSC\2015-2018\2018 CyberShot HX60V"

W$ = "D:\DSC\2015-2018\2017 CyberShot HX60V\DCIM"
W$ = "D:\VI_ DSC ME\2010-2018 CyberShot\2018 CyberShot HX60V\DCIM"
W$ = "D:\VI_ DSC ME\2010-2018 CyberShot\2015 CyberShot HX60V IMAGEx83"
W$ = "D:\VI_ DSC ME\2010-2018 CyberShot\2011-06 CyberShot HX5V IMAGEx502"
W$ = "D:\VI_ DSC ME\2010-2018 CyberShot\2010 CyberShot HX5V IMAGEx146"
W$ = "D:\VI_ DSC ME\2009 K800i"
W$ = "D:\VI_ DSC ME\2008 K800i"
W$ = "D:\VI_ DSC ME\2007-010-OCT W880i"
W$ = "D:\VI_ DSC ME\2007-003-MAR K800i"
W$ = "D:\VI_ DSC ME\2010-2017 NOKIA\2017 NOKIA E72 _ Images\Camera\201706-JUNE"
W$ = "D:\VI_ DSC ME\2010-2017 NOKIA\2017 NOKIA E72 _ Images\Camera\201707-JULY"
W$ = "D:\DSC\2013 NOKIA"
W$ = "D:\DSC\2011 SONY HX5V"
W$ = "D:\DSC\2011 Nokia E72 Ebay Seller"
W$ = "D:\DSC\2010\# 2010"
W$ = "D:\DSC\2009\# 2009"
W$ = "D:\DSC\# ME"
W$ = "D:\DSC\2015-2018\2018 CyberShot HX60V\DCIM"
'W$ = "D:\DSC\DSC SD-CARD SETTER\SDHC"
'W$ = "D:\DSC\z compare\NIKON TOTAL\DCIM"
'W$ = "D:\DSC\2010 NIKON COOLPIX S640 & SONY DSC-HX5V\# 2010"
'W$ = "D:\DSC\# ME"
'W$ = "D:\DSC\2008 FinePix J100\2008_11-18"
'W$ = ""
'W$ = ""
'W$ = ""
'W$ = ""
'W$ = ""
'W$ = ""
'W$ = ""

'-----------------------------------------------------------------------------------
ReDim ARRAY_I(10)
ReDim ARRAY_I_BAK(10)
ReDim ARRAY_DESC(10)
IP1 = 0
X_YEAR = Format(Now, "YYYY")


IP1 = IP1 + 1: ARRAY_I(IP1) = "D:\DSC\2015+Sony\" + X_YEAR + " CyberShot HX60V\DCIM"
IP1 = IP1 + 1: ARRAY_I(IP1) = "D:\VI_ DSC ME 01\2010+Sony\" ' + X_YEAR + " CyberShot HX60V_#\DCIM"
IP1 = IP1 + 1: ARRAY_I(IP1) = "D:\DSC\2012+Nokia E72"
IP1 = IP1 + 1: ARRAY_I(IP1) = "D:\DSC\2017+FUJI XP90"
' 4.
IP1 = IP1 + 1: ARRAY_I(IP1) = "D:\DSC\2018 Double Screen Cam"
IP1 = IP1 + 1: ARRAY_I(IP1) = "D:\VI_ DSC ME 01\2018 Double Screen Cam"
IP1 = IP1 + 1: ARRAY_I(IP1) = "D:\DSC\2019+4K ULTRA HD\2019 4K ULTRA\DCIMA"
IP1 = IP1 + 1: ARRAY_I(IP1) = "D:\VI_ DSC ME 01\2019+4K ULTRA HD\DCIMA"

IP1 = 0

IP1 = IP1 + 1: ARRAY_DESC(IP1) = X_YEAR + " CyberShot HX60V"
IP1 = IP1 + 1: ARRAY_DESC(IP1) = "V " + X_YEAR + " CyberShot HX60V"
IP1 = IP1 + 1: ARRAY_DESC(IP1) = "NOKIA"
IP1 = IP1 + 1: ARRAY_DESC(IP1) = "FUJI XP90"
' 4.
IP1 = IP1 + 1: ARRAY_DESC(IP1) = "Double Screen Camera"
IP1 = IP1 + 1: ARRAY_DESC(IP1) = "V Double Screen Camera"
IP1 = IP1 + 1: ARRAY_DESC(IP1) = "4K-ULTRA-HD"
IP1 = IP1 + 1: ARRAY_DESC(IP1) = "V 4K-ULTRA-HD"



' NEW TO ARRIVE DASH CAM
' ONLY INTERESTED IN IMAGE CAMERA
If GetComputerName <> "4-ASUS-GL522VW" Then
End If

Set FSO = CreateObject("Scripting.FileSystemObject")

If Command$ <> "" Then
    XCOMMAND = Command$
    XCOMMAND = Replace(XCOMMAND, """", "")
    If FSO.FolderExists(XCOMMAND) = True Then
        ' CHECK ROOT DIR
        If Len(XCOMMAND) > 3 Then
            IP1 = IP1 + 1: ARRAY_I(IP1) = XCOMMAND
            ARRAY_DESC(IP1) = Command$

            Result = MsgBox("DO YOU WANT TO USER COMMAND$ " + vbCrLf + vbCrLf + ARRAY_I(IP1), vbYesNoCancel + vbMsgBoxSetForeground)
            If Result = vbNo Then ARRAY_I(IP1) = ""
            If Result = vbCancel Then ARRAY_I(IP1) = ""
        End If
    End If
End If

Dim ERROR_NOTE, XXTT, HH
ERROR_NOTE = 0
XXTT = Now + TimeSerial(0, 0, 10)
On Error Resume Next
Do
    Err.Clear
    HH = Clipboard.GetText
    If Err.Number = 0 Then Exit Do
    If XXTT < Now Then ERROR_NOTE = 1: Exit Do
    Sleep 100
Loop Until 1 = 2

If ERROR_NOTE = 0 Then
    If HH <> "" Then
        Clipboard_VAR = HH
        Clipboard_VAR = Replace(Clipboard_VAR, """", "")
        If FSO.FolderExists(Clipboard_VAR) = True Then
            ' CHECK ROOT DIR
            If Len(Clipboard_VAR) > 3 Then
                IP1 = IP1 + 1: ARRAY_I(IP1) = Clipboard_VAR
                ARRAY_DESC(IP1) = "Clipboard"
                
                Result = MsgBox("DO YOU WANT TO USER CLIBOARD PASTE" + vbCrLf + vbCrLf + ARRAY_I(IP1), vbYesNoCancel + vbMsgBoxSetForeground)
                If Result = vbNo Then ARRAY_I(IP1) = ""
                If Result = vbCancel Then ARRAY_I(IP1) = ""
                End If
        End If
    End If
End If
'-----------------------------------------------------------------------------------
        
For R_COUNTER = 1 To UBound(ARRAY_I)
    If ARRAY_I(R_COUNTER) <> "" Then
        AT$ = ARRAY_I(R_COUNTER)
        If Dir(AT$, vbDirectory) = "" Then
            Result = CreateFolderTree(AT$)
        End If
    End If
Next
        
        
For R_COUNTER = 1 To UBound(ARRAY_I)
    If ARRAY_I(R_COUNTER) <> "" Then
        AT$ = ARRAY_I(R_COUNTER)
        If Dir(AT$, vbDirectory) = "" Then
            ARRAY_I(R_COUNTER) = ""
            MsgBox "NOT VALID path" + vbCrLf + vbCrLf + AT$
            If IsIDE = True Then Stop
        End If
        If Len(AT$) < 4 Then
            MsgBox "Not on a Root Dir ABORT" + vbCrLf + vbCrLf + AT$
            ARRAY_I(R_COUNTER) = ""
        End If
    End If
Next


For R_COUNTER = 1 To UBound(ARRAY_I)
    ARRAY_I_BAK(R_COUNTER) = ARRAY_I(R_COUNTER)
Next

WORK_COUNT_1 = 0
For R_COUNTER = 1 To UBound(ARRAY_I)
    If ARRAY_I(R_COUNTER) <> "" Then
        WORK_COUNT_1 = WORK_COUNT_1 + 1
    End If
Next
        
' CHECK WHAT REQUIRE LEARN
' ------------------------
WORK_COUNT_2 = 0
For R_COUNTER = 1 To UBound(ARRAY_I)
    If ARRAY_I(R_COUNTER) <> "" Then
        AT$ = ARRAY_I(R_COUNTER)
        WORK_COUNT_2 = WORK_COUNT_2 + 1
        Label13 = Trim(Str(WORK_COUNT_2)) + " of " + Trim(Str(WORK_COUNT_1))

        chkSubFolders.Value = vbChecked
        cboMask.Text = "*.JPG"
        ListView1.ListItems.Clear
        
        ' YELLOW NOT READY TO GO
        ' -------------------------------------
        Command1.BackColor = &HC0FFFF
        'cboMask = "*.wav;*.wma;*.amr;*.mp3"
        
        ' --------------------------------------------------------------
        ' THIS SET THE PATH AND THEN THERE IS A BUTTON ON THE FORM TO GO
        ' --------------------------------------------------------------
        txtPath = AT$
        Call cmdScan_Click
        FLAG_OFF = False
        For WE = 1 To ListView1.ListItems.Count
            
            A1$ = ListView1.ListItems.Item(WE).SubItems(1)
            B1$ = ListView1.ListItems.Item(WE)
            ' ---------------------------------------------------------------------------
            ' HERE MODIFY
            ' Len(B1$) = 13
            ' SO ANY UNUSAL IMAGE GET PUT IN FOLDER ALONG WITH DAY FOLDER
            ' AND THEN MAKE SO THEY NOT INCLUDED IN RESULT OF NOT SET TO DO-ER PROCESS ON
            ' [ Friday 08:28:00 Am_17 May 2019 ]
            ' ---------------------------------------------------------------------------
            ' CHECK BEEN DONE BEFORE
            
            If Len(B1$) <= 13 Then
                FLAG_OFF = True
            
'                If Mid(B1$, 5, 1) <> "-" Or Mid(B1$, 8, 1) <> "-" Then
'                    FLAG_OFF = True
'                End If
'                If IsNumeric(Replace(Mid(B1$, 1, 10), "-", "")) = False Then
'                    FLAG_OFF = True
'                End If
            End If
            
'            If FLAG_OFF = True Then
'                Exit For
'            End If
        Next
        
        If FLAG_OFF = False Then
            ARRAY_I(R_COUNTER) = ""
        End If
    End If
Next

ListView1.ListItems.Clear

' YELLOW NOT READY TO GO
' -------------------------------------
Command1.BackColor = RGB(127, 200, 127) ' GREEN
Command1.BackColor = &HC0FFFF           ' YELLOW
        
Z_COUNTER = 0
For R_COUNTER = 1 To UBound(ARRAY_I)
    If ARRAY_I(R_COUNTER) <> "" Then
        Z_COUNTER = Z_COUNTER + 1
    End If
Next
If Z_COUNTER = 0 Then
    IM_4 = ""
    IM_4 = IM_4 + "NONE FOLDERING TO REQUIRING TO BE SCANNED" + vbCrLf + "DO YOU WANT TO DO THEM ANYWAY" + vbCrLf
    Result = MsgBox(IM_4, vbYesNoCancel + vbMsgBoxSetForeground)
    If Result = vbNo Or Result = vbCancel Then
        End
        Exit Sub
    End If
    For R_COUNTER = 1 To UBound(ARRAY_I)
        ARRAY_I(R_COUNTER) = ARRAY_I_BAK(R_COUNTER)
    Next
End If
        
Z_COUNTER = 0
For R_COUNTER = 1 To UBound(ARRAY_I)
    If ARRAY_I(R_COUNTER) <> "" Then
        Z_COUNTER = Z_COUNTER + 1
    End If
Next
        
IM = ""
IM = IM + "THESE FOLDERS WILL BE SCANNED" + vbCrLf
IM = IM + "---------------------------------------" + vbCrLf
A_COUNTER = 0
IM_2 = ""
For R_COUNTER = 1 To UBound(ARRAY_I)
    If ARRAY_I(R_COUNTER) <> "" Then
        A_COUNTER = A_COUNTER + 1
        IM_2 = IM_2 + Trim(Str(A_COUNTER)) + " OF " + Trim(Str(Z_COUNTER)) + vbCrLf + ARRAY_I(R_COUNTER) + vbCrLf + vbCrLf
    End If
Next
IM = IM + IM_2 '+ "PUSH BUTTON GO WHEN READY"

A_COUNTER = 0

Result = MsgBox(IM, vbOKCancel + vbMsgBoxSetForeground)
If Result = vbNo Or Result = vbCancel Then
    End
    Exit Sub
End If
       
ONE_DONE = False
       
Call GET_NEXT_PATH

ONE_DONE = True

End Sub


Private Sub Form_Load()
    
'reset
    
Me.BackColor = &H8000000F
    
ScanPath.Caption = ScanPath.Caption + " -- " + App.EXEName
  
ScanPath.Height = ListView1.Top + ListView1.Height + 850
'ScanPath.Width = ListView1.Width + 210
ScanPath.Width = Command1.Width + Command1.Left + 240
ListView1.Width = ScanPath.Width - 200
  
    
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

'txtPath.Top = Picture1.Height + Picture1.Top + 40
txtPath.Top = 40

cmdBrowse.Top = txtPath.Top - 20
cmdBrowse.Left = txtPath.Left + txtPath.Width + 50 + 5
cmdBrowse.Height = 280
Label1.Top = txtPath.Top + 20

'------------------------------
'Center Screen
Me.Top = 0 '(Screen.Height / 2 - Me.Height / 2)
Me.Left = Screen.Width / 2 - Me.Width / 2


ScanPath.Show

'CALL TO FORM_LOAD_Timer_Timer()
'HAPPENS

End Sub

Private Sub FORM_LOAD_Timer_Timer()

FORM_LOAD_Timer.Enabled = False
Call AutoPix

End Sub

Private Sub lblCount_Click()
' lblCount.F
End Sub

Private Sub SPECIAL_RENAME_Command_BUTTON_Click()

For WE = 1 To ListView1.ListItems.Count
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    ListView1.ListItems.Item(WE).EnsureVisible
    
    If FS22.FileExists(A1$ + B1$) = True Then
                
        GG$ = B1$
        GG$ = Replace(GG$, "E72-1 ", "E72 ")
        If GG$ <> B1$ Then
            Name A1$ + B1$ As A1$ + GG$
        End If
    End If
Next


End Sub

Sub GET_NEXT_PATH_2()

Set FSO = CreateObject("Scripting.FileSystemObject")

For R_COUNTER = 1 To UBound(ARRAY_I)
    
    If ARRAY_I(R_COUNTER) <> "" Then
        
        ROUNTINE_POINTER = R_COUNTER
        
        AT$ = ARRAY_I(R_COUNTER)
        
        'Command1.BackColor = &HC0FFFF
        ' -------------------------------------
        ' GREEN READY TO GO
        ' -------------------------------------
        Command1.BackColor = RGB(127, 200, 127)
        Command1.Caption = "WHEN READY TO GO _" + Str(40)
        COUNT_DOWN_TIMER_COMMAND_BUTTON = 40
        TIMER_COUNTDOWN_COMMAND_BUTTON_AUTO_GO_AH.Enabled = True

        'cboMask = "*.wav;*.wma;*.amr;*.mp3"
        cboMask.Text = "*.JPG"
        
        ' ------------------------------------------------------------
        ' THIS SET THE PATH AND THEN THER IS ABUTTON ON THE FORM TO GO
        ' ------------------------------------------------------------
        txtPath = AT$
        
        Exit For
        
    End If
Next

End Sub

Sub GET_NEXT_PATH()

Do
    IM = ""
    If ONE_DONE = False Then
        IM = IM + "FIRST RUN...." + vbCrLf
    Else
        IM = IM + "DONE.... AND NEXT...." + vbCrLf
    End If
    
    If ONE_DONE = False Then
        IM = IM + "THIS FOLDER WILL BE THE FIRST SCANNED" + vbCrLf
        IM = IM + "------------------------------------------------" + vbCrLf
    Else
        IM = IM + "NEXT THIS FOLDER WILL BE SCANNED" + vbCrLf
        IM = IM + "-------------------------------------------" + vbCrLf
    End If
    
    TRUE_TO_GO = False
    
    For R_COUNTER = 1 To UBound(ARRAY_I)
        If ARRAY_I(R_COUNTER) <> "" Then
            TRUE_TO_GO = True
            A_COUNTER = A_COUNTER + 1
            IM = IM + Trim(Str(A_COUNTER)) + " OF " + Trim(Str(Z_COUNTER)) + vbCrLf + ARRAY_I(R_COUNTER) + vbCrLf + vbCrLf
            Exit For
        End If
    Next
    IM = IM + "YES AND PUSH BUTTON GO WHEN READY __ OR NO = SELECT NEXT ONE"
           
    If TRUE_TO_GO = True Then
        A_VAR = txtPath
        GET_NEXT_PATH_2
        I_ANSWER = MsgBox(IM, vbYesNoCancel + vbMsgBoxSetForeground)
        If I_ANSWER = vbCancel Then
            End
            Exit Sub
        End If


        'REMOVE FIRST FOUND ENTRY
        If I_ANSWER = vbNo Then
            For R_COUNTER = 1 To UBound(ARRAY_I)
                If ARRAY_I(R_COUNTER) = txtPath Then
                    ARRAY_I(R_COUNTER) = ""
                End If
            Next
        End If
        
        'GOT SAME RESULT AFTER ASK ANOTHER MEANS NONE LEFT
        If txtPath = A_VAR Then Exit Do
    
    Else
        IM = ""
        IM = IM + "DONE COMPLETER" + vbCrLf
        IM = IM + "--------------------" + vbCrLf
        IM = IM + IM_2
        
        I_ANSWER_2 = MsgBox(IM, vbOKCancel + vbMsgBoxSetForeground)
        
        If I_ANSWER_2 = vbCancel Then
            End
            Exit Sub
        End If
        
        Exit Do
    
    End If

ONE_DONE = True

Loop Until I_ANSWER = vbYes


End Sub

Private Sub Command1_Click()

' --------------------
' GREEN IF READY TO GO
' --------------------
If Command1.BackColor <> RGB(127, 200, 127) Then Exit Sub

chkSubFolders.Value = vbChecked
cboMask.Text = "*.JPG"

ListView1.ListItems.Clear
' ----------------------
' YELLOW NOT READY TO GO
' ----------------------
Command1.BackColor = &HC0FFFF

Call cmdScan_Click

Call AUTOPIX_02_OF_02
        
Command1.BackColor = RGB(127, 200, 127)
        
ARRAY_I(ROUNTINE_POINTER) = ""

Call GET_NEXT_PATH

End Sub


Private Sub Command_CHECK_JPG_FILE_IS_REQUIRING_Click()

chkSubFolders.Value = vbChecked
cboMask.Text = "*.JPG"

ListView1.ListItems.Clear

Call cmdScan_Click
'LEADS TO __ AUTOPIX_02_OF_02()

Command1.BackColor = RGB(127, 200, 127)

' Call AUTOPIX_02_OF_02
        
' ARRAY_I(ROUNTINE_POINTER) = ""

' Call GET_NEXT_PATH

End Sub


Sub AUTOPIX_02_OF_02()

XST = Now
Dim Dateset As Date
Dim Dateset3 As Date

For WE = 1 To ListView1.ListItems.Count
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    SETGO = False
    
    ' --------------------------------------------------------------
    ' TEMP VAR TO USE WITH NEXT IF STATEMENT
    ' DEFAULT LEFT ALONE
    ' --------------------------------------------------------------
'    SETGO = True
    ' --------------------------------------------------------------
    ' THIS DOESN'T CHANGE DATE LOOK FURTHER REM STATEMENT TO DO THAT
    ' --------------------------------------------------------------

    If Check_Dont_Do_Already_Modified.Value = vbUnchecked Then
    
        If WE Mod 100 = 0 Then
            DoEvents
            ListView1.ListItems.Item(WE).EnsureVisible
        End If
        
        'If Len(B1$) <> 12 Then GoTo JumpToNext_Quicker
        If Mid$(B1$, 1, 3) = "DSC" Or UCase(Mid$(B1$, 1, 3)) = "IMG" Then SETGO = True
        'NOKIA
        FILE_WITHOUT_EXT = Mid(B1$, 1, InStr(B1$, ".") - 1)
        If IsNumeric(FILE_WITHOUT_EXT) = True Then SETGO = True
        If SETGO = False Then GoTo JumpToNext_Quicker
    
    End If
    
    ListView1.ListItems.Item(WE).EnsureVisible
    
    If FS22.FileExists(A1$ + B1$) = False Then GoTo JumpToNext

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
    
    'NOKIA
    FILE_WITHOUT_EXT = Mid(B1$, 1, InStr(B1$, ".") - 1)
    
    '-------------------------------------
    If Mid$(B1$, 1, 3) = "DSC" Or UCase(Mid$(B1$, 1, 3)) = "IMG" Or (IsNumeric(FILE_WITHOUT_EXT)) Or SETGO = True Then
    
    
'    If Mid$(B1$, 1, 3) = "DSC" Then 'Or Mid$(b1$, 1, 4) = "_SCF" Or Mid$(b1$, 1, 4) = "PICT" Then  'SHOULD STOP REPEAT THE ACTION
'    If Mid$(B1$, 9, 1) = "(" And Mid$(B1$, 13, 1) = ")" Then
'    If InStr(A1$, "IMAGE\SD") > 0 Then 'SHOULD STOP REPEAT THE ACTION
'    If InStr(B1$, "PICT") > 0 Then 'SHOULD STOP REPEAT THE ACTION
'    XST = Mid$(B1$, InStr(B1$, "DSC"), 8)
'        On Error Resume Next
             
        Set F = FS22.GetFile(A1$ + B1$)
        
        If F.Size = 0 Then
            Clipboard.Clear
            Clipboard.SetText ("Zero Size File" + vbCrLf + A1$ + B1$)
            
            MsgBox ("Zero Size File KILL AFTER MSGBOX" + vbCrLf + A1$ + B1$)
            Kill A1$ + B1$: GoTo JumpToNext
        End If
        'GG$ = Format$(F.datelastmodified, "YYYY-MM-DD HH-MM-SS") + " - " + B1$
        GG$ = Format$(F.DateLastModified, "YYYY-MM-DD HH-MM-SS") + " - " ' - "
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
        
        If F.DateLastModified < Dateset And Dateset2 = "" Then
            GG$ = Format$(F.DateLastModified, "YYYY-MM-DD HH-MM-SS") + " - " ' - "
        End If

        'TEMP OVER-RIDE IF WANT CHANGE DATE TO NEW VALUE
'        If F.DateLastModified < Dateset Then
'            GG$ = Format$(F.DateLastModified, "YYYY-MM-DD HH-MM-SS") + " - " ' - "
'        End If

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
                    If Dateset3 < F.DateLastModified And Dateset3 < Dateset And Dateset2 <> "" Then
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
                    If Dateset4 < F.DateLastModified And Dateset4 < Dateset3 And Dateset4 < Dateset And Dateset2 <> "" Then
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
                    ds$ = "SONY DSC-H70"
                    H1$ = Replace(H1$, " JFIF ", "")
                    H1$ = Replace(H1$, " SONY DSC-H70 ", "")
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
                'STRIPNULL
'                If WORK = 0 Then Stop
                ds$ = Replace(ds$, "Nokia E72-1", "Nokia E72")
                If WORK = 0 And ds$ = "" Then ds$ = "JFIF"
                If ds$ = "" Then Stop
                
                ds$ = Replace(ds$, "FUJIFILM-FinePix XP90 XP91 XP95", "FUJI-FinePix XP90")
                
                
                
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
                If InStr(GG$, "E72-1") Then Stop
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

                GG$ = Replace(GG$, "WP01-8810-CPT-0523En", "Dual Screen")
                

                
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
                            Debug.Print ("File " + GG$ + " Aready Exists" + vbCrLf + A1$ + B1$)
                        Else
                            Name A1$ + B1$ As A1$ + GG$
                            RENAMED_COUNTER = RENAMED_COUNTER + 1
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
    
    
        
    DoEvents
    
JumpToNext_Quicker:

Label13 = Trim(Str(WE)) + " + " + Trim(Str(RENAMED_COUNTER)) + " + " + Trim(Str(ListView1.ListItems.Count - WE))
    
    

Next
'End


'------------------------------------------------
'------------------------------------------------
'RENAME FOLDER SELECT AND THEN RUN THE SUBROUTINE
'------------------------------------------------
'------------------------------------------------
'----------------------------

If 1 = 2 And Rename_Image_Folder_Proper_Checker = True Then
    MSGB2 = "Do You Want to Rename Camera CyberShot Photo Image Folders" + vbCrLf
    MSGB2 = MSGB2 + "into a more Readable Date Folder Name System"
    MSGBOX_QUESTION = vbYesNo
    MSGBOX_QUESTION = MSGBOX_QUESTION_CONVERT_ISIDE(MSGBOX_QUESTION)
    Result = MsgBox(MSGB2, MSGBOX_QUESTION + vbMsgBoxSetForeground)
    If Result = vbCancel Then Stop
    Result = MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN(MSGB2, Result, MSGBOX_QUESTION)
    '----------------------------
    If Result = vbYes Then
        Call Rename_Image_Folder_Proper
    End If
    '----------------------------
End If

Call Rename_Image_Folder_Proper


'------------------------------------------------
'------------------------------------------------


'------------------------------------------------
'------------------------------------------------
If DUPE_MERGE_SORT_Checker = True Then
    
    '------------------------------------------------
    '------------------------------------------------
    README = "DUPE MERGE SORT" + vbCrLf
    README = README + "ALL FOLDER THAT SAME NAME" + vbCrLf
    README = README + "BUT MATCH ANOTHER LONGER LENGHT SAME LEVEL PATH NAME " + vbCrLf
    README = README + vbCrLf
    README = README + "EXAMPLE" + vbCrLf
    README = README + vbCrLf
    README = README + "\PATHNAME\123ABC_22" + vbCrLf
    README = README + "VS" + vbCrLf
    README = README + "\PATHNAME\123ABC_2255 " + vbCrLf
    README = README + "\123ABC_2255 WILL MATCH WITH \123ABC_22" + vbCrLf
    README = README + vbCrLf
    README = README + "WILL MOVE CONTENT COPY FILE BY FILE" + vbCrLf
    README = README + "OF THE LONGER PATH SOURCE FOUND INTO" + vbCrLf
    README = README + "THE SAME FORMER FOUND DUPE MATCH PATH" + vbCrLf
    README = README + vbCrLf
    README = README + "AND ONLY IF PARTIAL PATH NOT CALLED" + vbCrLf
    README = README + "NAME ""GOT TO HERE"" AND ALSO GOT TO HAVE " + vbCrLf
    README = README + "*.JPG IN IT" + vbCrLf
    README = README + "AFTER SUCCESS MERGE COPY DONE AND THEN" + vbCrLf
    README = README + "WILL DELETE SOURCE OF THE COPY" + vbCrLf
    README = README + "----------" + vbCrLf
    README = README + "YES AND NO"
    
    '----------------------------
    MSGBOX_QUESTION = vbYesNo
    MSGB2 = README
    MSGBOX_QUESTION = MSGBOX_QUESTION_CONVERT_ISIDE(MSGBOX_QUESTION)
    Result = MsgBox(MSGB2, MSGBOX_QUESTION + vbMsgBoxSetForeground)
    If Result = vbCancel Then Stop
    Result = MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN(MSGB2, Result, MSGBOX_QUESTION)
    '----------------------------
    If Result = vbYes Then
        Call DUPE_MERGE_Image_Folder_SORT_PROPER
    End If
    '----------------------------
End If
DoEvents
Me.SetFocus
DoEvents


Exit Sub
'------------------------------------------------
'------------------------------------------------
'------------------------------------------------
'------------------------------------------------



End Sub


Function DUPE_MERGE_SORT_Checker()

DUPE_MERGE_SORT_Checker = False

'AT$ = W$

'If AT$ <> "" And Mid$(AT$, 1, 1) = """" Then
'    AT$ = Mid$(AT$, 2)
'    AT$ = Mid$(AT$, 1, Len(AT$) - 1)
'End If

If Len(txtPath) < 4 Then MsgBox "Not on a Root Dir - END": End

cboMask.Text = "*.JPG"

'txtPath = AT$
ListView1.ListItems.Clear
chkSubFolders.Value = vbChecked
Call cmdScan_Click
Dim LastPath, OldPath

For WE = 1 To ListView1.ListItems.Count
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    If WE Mod 100 = 0 Then ListView1.ListItems.Item(WE).EnsureVisible
    
    If OldPath <> A1$ Then
            OldPath = A1$
        
            LastPath = Mid(A1$, InStrRev(A1$, "\", Len(A1$) - 1))
            LastPath = Replace(LastPath, "\", "")
            
            SET_GO = False
            If Len(LastPath) <= 8 + 3 + 1 Then SET_GO = True
            TESTER = LastPath
            TESTER = Replace(TESTER, "_", "")
            TESTER = Replace(TESTER, " ", "")
            If IsNumeric(TESTER) = True Then SET_GO = True
            If Mid(LastPath, 11, 3) <> " _ " Then SET_GO = False
            
            If SET_GO = True Then
                DUPE_MERGE_SORT_Checker = True
                Exit For
            End If
        End If
Next

End Function


Function MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN(MSGB2, MSGBOX_RESULT, MSGBOX_QUESTION)
'------------------------------
'REM OUT -- If Want to Stop in Subrountine Function
'If RESULT = vbCancel Then Stop
'------------------------------
MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN = MSGBOX_RESULT
If MSGBOX_RESULT = vbCancel Then
    'DON'T ASK AGAIN IF IS ONLY A OKAY QUESTION -------------------------
    If MSGBOX_QUESTION = vbOKCancel Then MSGBOX_QUESTION = vbOK: Exit Function
    If MSGBOX_QUESTION = vbYesNoCancel Then MSGBOX_QUESTION = vbYesNo
    Result = MsgBox("QUESTIONED ANSWERED - BY STOP -- HOW DO YOU WANT THIS QUESTION NOW " + vbCrLf + vbCrLf + MSGB2, MSGBOX_QUESTION + vbMsgBoxSetForeground)
    MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN = Result
End If


End Function

Function MSGBOX_QUESTION_CONVERT_ISIDE(MSGBOX_QUESTION)

MSGBOX_QUESTION_CONVERT_ISIDE = MSGBOX_QUESTION
If IsIDE = True Then
    If MSGBOX_QUESTION = vbOK Then MSGBOX_QUESTION = vbOKCancel
    If MSGBOX_QUESTION = vbYesNo Then MSGBOX_QUESTION = vbYesNoCancel
    MSGBOX_QUESTION_CONVERT_ISIDE = MSGBOX_QUESTION
End If

End Function


Sub DUPE_MERGE_Image_Folder_SORT_PROPER()

Dim FS22
Set FS22 = CreateObject("Scripting.FileSystemObject")

'W$ = "D:\DSC\DCIM\"
AT$ = W$

If AT$ <> "" And Mid$(AT$, 1, 1) = """" Then
    AT$ = Mid$(AT$, 2)
    AT$ = Mid$(AT$, 1, Len(AT$) - 1)
End If

If Len(AT$) < 4 Then MsgBox "Not on a Root Dir - END": End

cboMask.Text = "*.JPG"

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

        SETGO = False
        If InStr("--" + LastPath, "--" + OLastPath) > 0 And OLastPath <> "" Then SETGO = True
        TESTER = Replace(LastPath, "_", "")
        TESTER = Replace(TESTER, " ", "")
        If IsNumeric(TESTER) = False Then SETGO = False
        
        If SETGO = True Then
            
            '----------------------------
            '----------------------------
            'PREP A QUESTION IN NEXT STEP
            'AND CHECK IF NEED BE ASK FIRST
            File1.Path = FrontofLastPath + OLastPath
            File1.Filename = "*.JPG"
            File1_ListCount_BEFORE = File1.ListCount
            
            File1.Path = FrontofLastPath + LastPath
            File1.Filename = "*.JPG"
            
            FLAG_TO_GO_QUESTION = False
            For R2T = 0 To File1.ListCount - 1
                If FS22.FileExists(FrontofLastPath + OLastPath + "\" + File1.List(R2T)) = False Then
                    FLAG_TO_GO_QUESTION = True
                    Exit For
                    Else
                    PRE_COPY_CHECK_EXIST = PRE_COPY_CHECK_EXIST + 1
                End If
            Next
            '----------------------------
            '----------------------------
            
            
            If FLAG_TO_GO_QUESTION = True Then
            
                MSGB2 = "MOVE ALL" + Str(File1.ListCount) + " of *.JPG IN FOLDER ONE BY ONE FROM " + vbCrLf + vbCrLf + "FOLDER BASE PATH --- " + vbCrLf + FrontofLastPath + vbCrLf + vbCrLf
                MSGB2 = MSGB2 + "FROM --- FOLDER END PATH --- " + vbCrLf + LastPath + vbCrLf + vbCrLf
                MSGB2 = MSGB2 + "TO ----- FOLDER END PATH --- " + vbCrLf + OLastPath + vbCrLf + vbCrLf
                MSGB2 = MSGB2 + "AFTER COPY IF GOOD RESULT AND NOT ERROR AND THEN DELETE THE SOURCE FILE" + vbCrLf + vbCrLf
                
                MSGB2 = MSGB2 + "YES NO"
                
                '----------------------------
                MSGBOX_QUESTION = vbYesNo
                '----------------------------
                'MSGB2 = README
                MSGBOX_QUESTION = MSGBOX_QUESTION_CONVERT_ISIDE(MSGBOX_QUESTION)
                Result = MsgBox(MSGB2, MSGBOX_QUESTION + vbMsgBoxSetForeground)
                If Result = vbCancel Then Stop
                Result = MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN(MSGB2, Result, MSGBOX_QUESTION)
                '----------------------------

                '----------------------------
                            
                If Result = vbYes Then
                    On Error Resume Next
                    COPYCOUNT_SUCCESS = 0
                    'FS22.COPYFOLDER FrontofLastPath + LastPath, FrontofLastPath + OLASTPATH
                    
                    File1.Path = FrontofLastPath + LastPath
                    File1.Filename = "*.JPG"
                    
                    For R2T = 0 To File1.ListCount - 1
                        If FS22.FileExists(FrontofLastPath + OLastPath + "\" + File1.List(R2T)) = False Then
                            Err.Clear
                            FILE1P = FrontofLastPath + LastPath + "\" + File1.List(R2T)
                            FILE2P = FrontofLastPath + OLastPath + "\" ' + File1.List(R2T)
                            FS22.CopyFile FILE1P, FILE2P
                            If Err.Number > 0 Then
                                WHILECOPY = "WHILE COPY " + vbCrLf + vbCrLf + FILE1P + vbCrLf + vbCrLf + " -- TO -- " + vbCrLf + vbCrLf + FILE2P + vbCrLf + vbCrLf
                                '-------------------------
                                MSGB2 = "ERROR HAPPEN -- " + VBCLRF + vbCrLf + WHILECOPY + "ERR.DESCRIPTION = " + Err.Description + vbCrLf + vbCrLf + "ERR.NUMBER =" + Str(Err.Number) + vbCrLf + vbCrLf + "ABORT COPY THIS FOLDER -- YES AND NO"
                                MSGBOX_QUESTION = vbYesNo
                                '----------------------------
                                MSGB2 = MSGB2
                                MSGBOX_QUESTION = MSGBOX_QUESTION_CONVERT_ISIDE(MSGBOX_QUESTION)
                                Result = MsgBox(MSGB2, MSGBOX_QUESTION + vbMsgBoxSetForeground)
                                If Result = vbCancel Then Stop
                                Result = MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN(MSGB2, Result, MSGBOX_QUESTION)
                                '----------------------------
   
                                If Result = vbYes Then Exit For
                            Else
                                COPYCOUNT_SUCCESS = COPYCOUNT_SUCCESS + 1
                                'COPY HAPPEN GOOD -- OKAY -- TO DELETE
                                Kill FILE1P
                                If Err.Number > 0 Then
                                    WHILEKILL = "WHILE KILL " + VCRLF + vbCrLf + FILE1P + " -- TO -- " + VCRLF + vbCrLf + FILE2P + VCRLF + vbCrLf
                                    MSGB2 = "ERROR HAPPEN -- " + VBCLRF + vbCrLf + WHILEKILL + "ERR.DESCRIPTION = " + Err.Description + vbCrLf + "ERR.NUMBER =" + Str(Err.Number) + vbCrLf + vbCrLf + "MINOR PROBLEM IF FILE DON'T DEL CONTINUE ON AHEAD -- YES AND NO"
                                    MSGBOX_QUESTION = vbYesNo
                                    '----------------------------
                                    MSGB2 = MSGB2
                                    MSGBOX_QUESTION = MSGBOX_QUESTION_CONVERT_ISIDE(MSGBOX_QUESTION)
                                    Result = MsgBox(MSGB2, MSGBOX_QUESTION + vbMsgBoxSetForeground)
                                    If Result = vbCancel Then Stop
                                    Result = MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN(MSGB2, Result, MSGBOX_QUESTION)
                                    '----------------------------

                                    If Result = vbNo Then Exit For
                                End If
                            End If
                        End If
                    
                        On Error GoTo 0
                    Next R2T
                    
                    If COPYCOUNT_SUCCESS > 0 Then
                        
                        File1.Path = FrontofLastPath + LastPath
                        File1.Filename = "*.JPG"
                        COPYDONE = "SUCCESSFUL COPY AND DELETE SOURCE = " + VCRLF + vbCrLf + Str(COPYCOUNT_SUCCESS) + " of *.JPG" + vbCrLf + vbCrLf + " -- OUT OF POTENTIAL ALREADY EXIST BEFORE -- " + VCRLF + vbCrLf + Str(File1_ListCount_BEFORE) + " of *.JPG"
                        
                        File1.Path = FrontofLastPath + OLastPath
                        File1.Filename = "*.JPG"
                        COPYDONE = COPYDONE + vbCrLf + vbCrLf + " -- AND TOTAL COUNT AT DESTINATION -- " + VCRLF + vbCrLf + Str(File1.ListCount) + " of *.JPG"
                        
                        '----------------------------
                        MSGB2 = COPYDONE
                        MSGBOX_QUESTION = vbOK
                        '----------------------------
                        MSGB2 = MSGB2
                        MSGBOX_QUESTION = MSGBOX_QUESTION_CONVERT_ISIDE(MSGBOX_QUESTION)
                        Result = MsgBox(MSGB2, MSGBOX_QUESTION + vbMsgBoxSetForeground)
                        If Result = vbCancel Then Stop
                        Result = MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN(MSGB2, Result, MSGBOX_QUESTION)
                        '----------------------------
                    
                    End If
                    
                    
                    
                End If
            End If
            
            
        End If

        OLastPath = LastPath
        
    End If
Next


    If PRE_COPY_CHECK_EXIST > 0 Then
    
        
        '-------------------------------------------------------------------------------
        'IF THEN -- AND -- -- AND THEN -- THEN IF --- GREATER THAN SMALLER THAN EQUAL TO
        '-------------------------------------------------------------------------------
        
        COPYDONE = "RESULT -- PREP CHECK IF FILE EXIST " + vbCrLf
        COPYDONE = COPYDONE + "AND THEN IF COPY REQUIRE " + vbCrLf
        COPYDONE = COPYDONE + "CHECK EXIST TRUE COUNT =" + Str(PRE_COPY_CHECK_EXIST) + " of *.JPG FILE NAME Exist Already"
        '----------------------------
        MSGB2 = COPYDONE
        MSGBOX_QUESTION = vbOK
        '----------------------------
        MSGB2 = MSGB2
        MSGBOX_QUESTION = MSGBOX_QUESTION_CONVERT_ISIDE(MSGBOX_QUESTION)
        Result = MsgBox(MSGB2, MSGBOX_QUESTION + vbMsgBoxSetForeground)
        If Result = vbCancel Then Stop
        Result = MSGBOX_QUESTION_CANCEL_AND_ASK_AGAIN(MSGB2, Result, MSGBOX_QUESTION)
        '----------------------------

    
    End If



End Sub


Function Rename_Image_Folder_Proper_Checker()

Rename_Image_Folder_Proper_Checker = False

'W$ = "D:\DSC\DCIM\"
'AT$ = W$

'If AT$ <> "" And Mid$(AT$, 1, 1) = """" Then
'    AT$ = Mid$(AT$, 2)
'    AT$ = Mid$(AT$, 1, Len(AT$) - 1)
'End If

If Len(txtPath) < 4 Then MsgBox "Not on a Root Dir - END": End

cboMask.Text = "*.*"

' txtPath = AT$
ListView1.ListItems.Clear
chkSubFolders.Value = vbChecked
Call cmdScan_Click
Dim LastPath, OldPath

For WE = 1 To ListView1.ListItems.Count
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    If WE Mod 100 = 0 Then ListView1.ListItems.Item(WE).EnsureVisible
    
    If OldPath <> A1$ Then
            OldPath = A1$
        
            LastPath = Mid(A1$, InStrRev(A1$, "\", Len(A1$) - 1))
            LastPath = Replace(LastPath, "\", "")
            
            SET_GO = False
            'THE NEW FIRMWARE CAMERA UPDATE HAS FOLDER NAME LIKE 419MSDCF ONLY IF SETTING IN CAMERA NOT SET TO DATE FOLDER
            'THE FOLDER NAME LIKE                                45581020 EXAMPLE 2018 OCT 20
            If Mid(LastPath, 1, 2) <> "20" Then
                SET_GO = True
            End If
            If Len(LastPath) <> 8 Then SET_GO = False
            If IsNumeric(LastPath) = False Then SET_GO = False
            If InStr(LastPath, "MSDCF") > 0 Then
                SET_GO = True
            End If
            If A1$ = txtPath + "\" Then SET_GO = False
            
            If SET_GO = True Then
                Rename_Image_Folder_Proper_Checker = True
                Exit For
            End If
        End If
Next

End Function



Sub Rename_Image_Folder_Proper()

'W$ = "D:\DSC\DCIM\"
'AT$ = W$

'If AT$ <> "" And Mid$(AT$, 1, 1) = """" Then
'    AT$ = Mid$(AT$, 2)
'    AT$ = Mid$(AT$, 1, Len(AT$) - 1)
'End If
'
'If Len(AT$) < 4 Then MsgBox "Not on a Root Dir - END": End

cboMask.Text = "*.*"

' txtPath = AT$

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

        'NORMAL IMPORT FROM SDCARD CAMERA WHEN FIRMWARE BEEN UPDATED
        If InStr(LastPath, "MSDCF") Then
            DIR_FILE = Dir(FrontofLastPath + LastPath + "\*.JPG")
            If DIR_FILE <> "" Then
                Set F = fs.GetFile(FrontofLastPath + LastPath + "\" + DIR_FILE)
                DATE_FILE_NAME = F.DateLastModified
                LastPath = Format(DATE_FILE_NAME, "YYYY MM DD")
    '            LastPath = Mid(LastPath, 1, 8)
                'LastPath = Mid(Trim(Str(Year(Now))), 1, 3) + Mid(LastPath, 4)
                'If Mid(LastPath, 5, 1) <> " " Then LastPath = Mid(LastPath, 1, 4) + " " + Mid(LastPath, 5, 2) + " " + Mid(LastPath, 7, 2)
                If Dir(FrontofLastPath + LastPath, vbDirectory) = "" And FrontofLastPath + LastPath + "\" <> A1$ Then
                    Name A1$ As FrontofLastPath + LastPath
                    WX = 0
                Else
                    
                    WX = WX + 1
                    '------------------------------------------------------------------
                    'IF WE PUT THE DASH IN BEFORE NUMERIC ADD END
                    '------------------------------------------------------------------
                    'AND THEN THAT WORK SORT THE NEW FOLDER BETTER ALONGSIDE EACH OTHER
                    'RATHER THAN WITHOUT THE EXTRA CHAR NUMERIC MAKE SORT AT END
                    '------------------------------------------------------------------
                    Do
                        TEST_DONE = False
                        If Dir(FrontofLastPath + LastPath + "-" + Trim(Str(WX)), vbDirectory) <> "" Then
                            WX = WX + 1
                            TEST_DONE = True
                        End If
                    Loop Until TEST_DONE = False
                    
                    Name A1$ As FrontofLastPath + LastPath + " _ " + Trim(Str(WX))
                End If
            End If
        End If
        
        ' NORMAL IMPORT FROM SDCARD CAMERA
        If IsNumeric(LastPath) = True And Mid(LastPath, 1, 3) <> Mid(Trim(Str(Year(Now))), 1, 3) Then 'And Len(LastPath) <> 8 Then
'            LastPath = Mid(LastPath, 1, 8)
            LastPath = Mid(Trim(Str(Year(Now))), 1, 3) + Mid(LastPath, 4)
            If Mid(LastPath, 5, 1) <> " " Then LastPath = Mid(LastPath, 1, 4) + " " + Mid(LastPath, 5, 2) + " " + Mid(LastPath, 7, 2)
            If Dir(FrontofLastPath + LastPath, vbDirectory) = "" And FrontofLastPath + LastPath + "\" <> A1$ Then
                Name A1$ As FrontofLastPath + LastPath
                WX = 0
            Else
                
                WX = WX + 1
                '------------------------------------------------------------------
                'IF WE PUT THE DASH IN BEFORE NUMERIC ADD END
                '------------------------------------------------------------------
                'AND THEN THAT WORK SORT THE NEW FOLDER BETTER ALONGSIDE EACH OTHER
                'RATHER THAN WITHOUT THE EXTRA CHAR NUMERIC MAKE SORT AT END
                '------------------------------------------------------------------
                Do
                    TEST_DONE = False
                    If Dir(FrontofLastPath + LastPath + " _ " + Trim(Str(WX)), vbDirectory) <> "" Then
                        WX = WX + 1
                        TEST_DONE = True
                    End If
                Loop Until TEST_DONE = False
                
                Name A1$ As FrontofLastPath + LastPath + " _ " + Trim(Str(WX))
            End If
        End If
        
        'BEEN IMPORTED BY WIFI
        If Mid(LastPath, 7, 3) = Mid(Trim(Str(Year(Now))), 1, 3) Then
'            LastPath = Mid(LastPath, 1, 8)
            LastPath = Mid(LastPath, 7, 4) + Mid(LastPath, 4, 2) + Mid(LastPath, 1, 2)
            If Dir(FrontofLastPath + LastPath, vbDirectory) = "" And FrontofLastPath + LastPath + "\" <> A1$ Then
                
                Name A1$ As FrontofLastPath + LastPath
                WX = 0
            Else
                '------------------------------------------------------------------
                'IF WE PUT THE DASH IN BEFORE NUMERIC ADD END
                '------------------------------------------------------------------
                'AND THEN THAT WORK SORT THE NEW FOLDER BETTER ALONGSIDE EACH OTHER
                'RATHER THAN WITHOUT THE EXTRA CHAR NUMERIC MAKE SORT AT END
                '------------------------------------------------------------------
                WX = WX + 1
                Name A1$ As FrontofLastPath + LastPath + "-" + Trim(Str(WX))
            End If
        End If
        
        
        'IS IT THE INFO LINE
        'If Mid(LastPath, 7, 3) = Mid(Trim(Str(Year(Now))), 1, 3) Then
            
        If InStr(A1$, " -- GOT TO HERE -- ") > 0 Then
            'NOT EQUAL TO THIS DECADE
            If Mid(LastPath, 1, 2) <> "20" Then

                Mid(LastPath, 1, 3) = Mid(Format(Now, "YYYY"), 1, 3)
                If Dir(FrontofLastPath + LastPath, vbDirectory) = "" Then
                    Name A1$ As FrontofLastPath + LastPath
                End If
                WX = WX + 1
            End If
        End If
        
        
        
    End If
Next


End Sub

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
    txtExifInfo = Left$(txtExifInfo, InStr(txtExifInfo, Chr$(NULL_CHAR)) - 1)
    
    GetExifModel = txtExifInfo
End Function

Private Sub MNU_COMMAND_PATH_IN_Click()

ALT_COMMAND$ = Replace(Command$, """", "")
If FSO.FolderExists(ALT_COMMAND$) = True Then
    PASTE_PATH_SET = ALT_COMMAND$
    Call AutoPix
    Exit Sub
End If
MsgBox "COMMAND$ PATH DONT EXIST" + vbCrLf + vbCrLf + Command$, vbMsgBoxSetForeground

End Sub

Private Sub MNU_EXIT_Click()
End
End Sub

Private Sub MNU_NOKIA_PATH_IN_Click()

PASTE_PATH_SET = "\\7-asus-gl522vw\7_asus_gl522vw_02_d_drive\DSC\2013-2018 NOKIA\IMAGE NOKIA"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(PASTE_PATH_SET) = True Then
        Call AutoPix
    Else
        MsgBox "NOKIA PATH DONT EXIST" + vbCrLf + vbCrLf + PASTE_PATH_SET, vbMsgBoxSetForeground
    End If

End Sub

Private Sub MNU_PASTE_PATH_IN_Click()

Exit Sub

If Clipboard.GetText <> "" Then
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(Clipboard.GetText) = True Then
        PASTE_PATH_SET = Clipboard.GetText
        Call AutoPix
        Exit Sub
    End If
End If

ALT_COMMAND$ = Replace(Command$, """", "")
If FSO.FolderExists(ALT_COMMAND$) = True Then
    PASTE_PATH_SET = ALT_COMMAND$
    MsgBox "PRIORITY WENT TO COMMAND$ AS VALID PATH" + vbCrLf + vbCrLf + PASTE_PATH_SET, vbMsgBoxSetForeground
    Call AutoPix
    Exit Sub
End If

MsgBox "NEITHER CLIPOARD NOR COMMAND$ HAS A VALID PATH", vbMsgBoxSetForeground

End Sub


Private Sub Check_Only_Do_UnModifid_Click()

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


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_Click()

Exit Sub

Dim FS22
Set FS22 = CreateObject("Scripting.FileSystemObject")


tpath3$ = "E:\04 Music ---\Del\"

B1$ = Mid$(List1.List(List1.ListIndex), 14)
c1$ = B1$ + "--Todel"

'tpath1$ = "E:\04 Music ---\"

c1$ = Mid$(B1$, InStrRev(B1$, "\") + 1)
List1.RemoveItem (List1.ListIndex)
'Name b1$ As c1$

'fs22.copyFile b1$, tpath3$ + c1$
If Dir$(tpath3$ + c1$) <> "" Then Kill tpath3$ + c1$
FS22.MoveFile B1$, tpath3$ + c1$



End Sub

Private Sub Label20_Click()
    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True        'Use default sorting to sort the
    ListView1.SortKey = 0
    ii = ListView1.ListItems.Count

'Call Auto---Pix

End Sub

Private Sub Label21_Click()

    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True        'Use default sorting to sort the
    ListView1.SortKey = 1
    ii = ListView1.ListItems.Count

End Sub

'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    If GetSpecialfolder(&H29) = "" Then
        If MESSENGER_BOX_ABOUT_GetSpecialfolder = False Then
            MsgBox "AN ERROR HAPPEN BECUASE GET GetSpecialfolder(&H29) DOES NOT WORK UNLESS ADMIN MODE" + vbCrLf + "MEANT TO RETURN THE PATH OF 64 BIT SYSWOW64 BIT FOLDER", vbMsgBoxSetForeground
        End If
        MESSENGER_BOX_ABOUT_GetSpecialfolder = True
        Exit Sub
    End If
    
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    i = CODER_VBP_FILE_NAME_2
    If InStr(i, " - 64 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 64 BIT.vbp", ".vbp")
    If InStr(i, " - 32 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 32 BIT.vbp", ".vbp")
    
    CODER_EXE_FILE_NAME_1 = Replace(CODER_EXE_FILE_NAME_1, ".vbp", ".exe")
    
    Filespec = CODER_VBP_FILE_NAME_2
    If Dir(Filespec) <> "" Then
        FILE_SPEC_GO = Filespec
    Else
        If LCase(GetSpecialfolder(&H29)) = LCase("C:\Windows\SysWOW64") Then
            Filespec = Replace(Filespec, ".vbp", " - 64 BIT.vbp")
        Else
            Filespec = Replace(Filespec, ".vbp", " - 32 BIT.vbp")
        End If
    End If
    
    If Dir(Filespec) <> "" Then FILE_SPEC_GO = Filespec
    
    CODER_VBP_FILE_NAME_2 = FILE_SPEC_GO
End Sub


Private Sub Mnu_VB_ME_Click()

    Dim CODER_VBP_FILE_NAME_2
    Dim VB_1, VB_2, VB_3
    VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
    VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
    If Dir(VB_1) <> "" Then VB_3 = VB_1
    If Dir(VB_2) <> "" Then VB_3 = VB_2
    '------------------------------------------------------
    'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
    '------------------------------------------------------
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
    objShell.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set objShell = Nothing
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub


    Dim MSGQ, iR
    Dim CODER_EXE_FILE_NAME_1
    'Dim CODER_VBP_FILE_NAME_2
    'Dim EXIT_TRUE
    
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
        iR = MsgBox("NOT IN IDE ARE YOU SURE", MSGQ, vbMsgBoxSetForeground)
        If iR <> vbYes Then Exit Sub
    End If
    
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
    
End Sub

Private Sub MNU_VB_FOLDER_Click()
    Beep
    Dim CODER_EXE_FILE_NAME_1
    Dim CODER_VBP_FILE_NAME_2
    Dim EXIT_TRUE
    
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + CODER_EXE_FILE_NAME_1 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
    
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    Beep
    MNU_ADMINISTRATOR.Caption = "Administrator ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "Administrator ON", "Administrator OFF")
    End If
End Sub
'-------------------------------------------
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


Sub Jeepers()

Dim FS22

Set FS22 = CreateObject("Scripting.FileSystemObject")

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
    FS22.MoveFile A1$ + B1$, c1$ + B1$
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
Set FS22 = CreateObject("Scripting.FileSystemObject")
FS22.DeleteFolder txtPath.Text, True
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
Set FS22 = CreateObject("Scripting.FileSystemObject")
FS22.MoveFolder Drived2$ + "Temp\Anything\*", v2$
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
Dim FS22

Set FS22 = CreateObject("Scripting.FileSystemObject")

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
            FS22.DeleteFolder d2$, True
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
Dim FS22

Set FS22 = CreateObject("Scripting.FileSystemObject")

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
    Set FS22 = CreateObject("Scripting.FileSystemObject")
    FS22.MoveFile c2$ + B1$, A1$ + B1$
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
    Set FS22 = CreateObject("Scripting.FileSystemObject")
    FS22.CopyFolder Drived2$ + "Temp\Anything", v1$, True
    
    Set FS22 = CreateObject("Scripting.FileSystemObject")
    FS22.DeleteFolder Drived2$ + "Temp\Anything", True


jeep:
End Sub

'-------------------------------------------------
'-------------------------------------------------
Private Sub Timer_1_MINUTE_Timer()

Timer_1_MINUTE.Interval = 60000
Timer_1_MINUTE.Enabled = False
Exit Sub

Timer_1_MINUTE.Interval = 60000

Call VB_PROJECT_CHECKDATE

End Sub
'-------------------------------------------------
'-------------------------------------------------
Sub VB_PROJECT_CHECKDATE()

'TOP DELCLARE -------------
'DIM XVB_DATE
'--------------------------

If Mid(App.Path, 1, 1) <> "D" Then Exit Sub

PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE'S\")

'---------------------------------------------------
'IF A NEW PROJECT NOT BEEN SYNC YET TO MIRROR FOLDER
'---------------------------------------------------
'MINOR WORK AROUND CREATE PATH
'---------------------------------------------------
If Dir(PATH_FILE_NAME2) = "" Then
    On Error Resume Next
        MkDir Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\"))
        fs.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
    If Err.Number > 0 Then Exit Sub
End If

Set F = fs.GetFile(PATH_FILE_NAME2)

VB_DATE = F.DateLastModified

'--------
'01 OF 02
'----------------------------------
'WRITE INFO ABOUT DATE CHANGE NEWER
'----------------------------------

If XVB_DATE < VB_DATE And XVB_DATE > 0 Then
    
    '-----------------------------------
    'RUN A VBS TO COPY OVER AND RELOADER
    '-----------------------------------
    
'    FR1 = FreeFile
'    Open App.Path + "\RELOAD AND COPY.VBS" For Output As #FR1


    '----------------------------------
    'Mon 10 April 2017 16:30:50--------
    '----------------------------------
    'PROJECT REFERANCE ----------------
    'wshom.ocx
    'IF HARD FIND DO BROWSER
    'IN SYSTEM32 AND SYSWOW64
    'DIR wshom.ocx /S
    '----------------------------------
    'WINDOWS SCRIPT HOST OBJECT MODEL
    'AFTER FIND ALSO HAVE TO SELECT HER
    '----------------------------------
    '----------------------------------
    'Mon 10 April 2017 15:45:11--------
    '----
    'VISUAL BASIC wshom.ocx NOT WORKING - Google Search
    'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB721GB721&q=VISUAL+BASIC+wshom.ocx+NOT+WORKING&oq=VIUAL+BASIC+wshom.ocx+NOT+WORKING&gs_l=serp.3..30i10k1.1533.3987.0.4658.12.12.0.0.0.0.127.888.11j1.12.0....0...1c.1.64.serp..0.12.886...33i160k1j33i21k1.fYCPCpVPeVk
    '--------
    'WScript reference in VB
    'http://forums.codeguru.com/showthread.php?30458-WScript-reference-in-VB
    '----
    'Set objShell = WScript.CreateObject("WScript.Shell") - ------Not WORKER
    '----------------------------------
    '----------------------------------
    
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    WSHShell.Run """" + App.Path + "\RELOAD AND COPY.VBS" + """"
    Set WSHShell = Nothing
    Unload Me
    
End If

XVB_DATE = VB_DATE

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


Private Sub TIMER_COUNTDOWN_COMMAND_BUTTON_AUTO_GO_AH_Timer()
    If Command1.BackColor <> RGB(127, 200, 127) Then Exit Sub
    ' -------------------------------------
    ' GREEN READY TO GO
    ' -------------------------------------

    Command1.Caption = "WHEN READY TO GO _" + Str(COUNT_DOWN_TIMER_COMMAND_BUTTON)
    COUNT_DOWN_TIMER_COMMAND_BUTTON = COUNT_DOWN_TIMER_COMMAND_BUTTON - 1
    If COUNT_DOWN_TIMER_COMMAND_BUTTON = 0 Then
        Command1.Caption = "WHEN READY TO GO _" + Str(COUNT_DOWN_TIMER_COMMAND_BUTTON)
        TIMER_COUNTDOWN_COMMAND_BUTTON_AUTO_GO_AH.Enabled = False
        Call Command1_Click
        
        Exit Sub
    
    End If
    TIMER_COUNTDOWN_COMMAND_BUTTON_AUTO_GO_AH.Enabled = True
End Sub
