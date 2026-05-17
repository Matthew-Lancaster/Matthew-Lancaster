VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form aMain 
   BackColor       =   &H00000000&
   Caption         =   "Web Site Update Dates"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   13305
   Icon            =   "WebDates.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   13305
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   2250
      Left            =   4080
      TabIndex        =   14
      Top             =   1665
      Visible         =   0   'False
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   3969
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer AweSomeModeSubTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   2880
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00008000&
      Caption         =   "WK Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   11460
      TabIndex        =   13
      Top             =   15
      Width           =   1485
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00008000&
      Caption         =   "Day Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   9975
      TabIndex        =   12
      Top             =   15
      Width           =   1455
   End
   Begin VB.Timer AweSomeTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   2430
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   840
      Top             =   1590
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   750
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   6120
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   3780
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   2895
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   3435
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.ListBox List5 
      Height          =   1425
      Left            =   6090
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   2010
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.ListBox List4 
      Height          =   1035
      Left            =   2685
      TabIndex        =   8
      Top             =   2205
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Count"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   -15
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   15
      Width           =   2220
   End
   Begin VB.Timer TimerPutToWeb 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1605
      Top             =   1965
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   1170
   End
   Begin VB.CheckBox GoCrazyCheck 
      BackColor       =   &H00C00000&
      Caption         =   "Go Awesome Pleasure Crazy Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   5550
      TabIndex        =   6
      Top             =   15
      Width           =   4410
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3705
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   15
      Width           =   1845
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   1380
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
      RequestTimeout  =   20
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1350
      Top             =   990
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
      RequestTimeout  =   300
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "Hold"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2910
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   15
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   15
      Width           =   690
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   3450
      Pattern         =   "wht*.html"
      TabIndex        =   0
      Top             =   7665
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      ItemData        =   "WebDates.frx":0ABA
      Left            =   45
      List            =   "WebDates.frx":0ABC
      TabIndex        =   4
      Top             =   405
      Width           =   12990
   End
   Begin VB.Label Label2 
      BackColor       =   &H00B90000&
      Caption         =   "Label2"
      Height          =   285
      Left            =   3810
      TabIndex        =   15
      Top             =   6045
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   405
      Left            =   3615
      TabIndex        =   3
      Top             =   4890
      Width           =   1185
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "File"
      Begin VB.Menu Mnu_WinAmp 
         Caption         =   "Test WinAmp_List"
      End
   End
   Begin VB.Menu Mnu_Open_CRC 
      Caption         =   "OPEN CRC List"
   End
   Begin VB.Menu Mnu_DayRun 
      Caption         =   "Do Day Run"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_DOFAVS 
      Caption         =   "Do Favs"
   End
   Begin VB.Menu Mnu_Run_Vb 
      Caption         =   "Run VB"
   End
End
Attribute VB_Name = "aMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FAV




Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15


Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long


Dim CommadFTP, StripFTPPathTrobAY(), GOTrue
Dim IconFile
Public Easy, PuTT2, Putt3, PathBack, Paths
Public TagSH1, TagSH2, TagSH3, TagSH4, EGU, AKO, PDate, TTNow
Dim FileSpec As String
'Searches
'How this Prog starts on timer or soon as
'Timer5.Enabled = True
'or
'Call Form_Load2
Const MaxNotFLags = 3 ' Commands Plus 1 if 3 -- +1 = 4

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function TouchFileTimes Lib "imagehlp.dll" (ByVal FileHandle As Long, ByRef pSystemTime As Any) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'right order


Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Dim Tx$, Td$, StartTime
Dim Path1$(100), EQ3, RT, OldB1$, Tfn2, i, QDx$, QW3$, QX$, P, LastCommand, TimeDD
Dim Path2$(100)
Dim Path3$(100)
Dim ArrayT(30)
Dim ArrayP(30)
Dim ArrayS(30)
Dim ArrayV(30)

Dim FreeF1, Hag2, Fee, QW$, QW1$, QW5, QW2$, HR, HR3, Title$, nSize$, SSize$, Grab24$, SecondHalf, QY, BR

Dim Washad$()
Dim Date2020() As Date
Dim Ee$()

'Nice Line code
'BrowDlg.ImportExportFavorites False, "D:\MY WEBS\bookmark.html"
Public TheseToDo
Dim WriteFlag, LFO$, G1$, G2$
Dim Activate3
Dim CountD1, CountD2, CountD3
Dim Activate2
Dim Tru, InJob, InAwesome, OldCountHtmlPage2, JobRun, JobRun2, OlJobRun
Dim Mimp2Buff(), OldList3Count
Dim Mim()
Dim MimDate() As Date

'------------------------------------------
Public M, DayGo, WkGo, Togg1, ProccesGo, UploadFilet1$, QQ$
Dim DayMark2 As Date, Comm5, Ddate As Date
Public PuTT1, Force, DryRunMode, DimSize, OldHag$, OldList3
Public A1$, B1$, UpLoad, Retry, WXHex1$, ReTryTTi, ExitPro, ExitPro2
Public PathFileT1$, PathFileT2$, PathFileT3$
Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Const GW_CHILD = 5

Public Statc$, Grab22$, Header1$, Header2$, Header3$, Header4$, FileT1$

Public CommandCnt, BookDates
Public Done, Happy, OldsCommand

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Private Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long


Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Dim Adate(20) As Date
Dim Bdate(20) As Date

Public counter As Integer
Dim indexf$(4000)
Public a As Date
Public Idate As Date
'Public WebPhotComplete As Integer
Public LastUpdate As Date
Public XJag As Integer
Public dired$
Public direc$
Public PopPop As Integer
Public ButterCake As Date
Public Ledge As Integer
Public Peel As Integer
Public Puke As Date
Public Teek As Integer
Public IPaddy1$
Public We2 As Integer
Public Languard As Integer
Public Eye2
Public Books As Integer
Public Forget As Integer
Public Activated

Public BrowDlg As New SHDocVw.ShellUIHelper

Public RattleSnake01
Public RattleSnake02
Public AName01

Public YtCrc1$
Public YtCrc2$


Public Fref1
Public Fref2
Public Fref3
Public Fref4
Public Fref5
Public Freff5$
Public Freff8$, PalmPage
Public Freff9$
Public NewsGo
    
Public Asked

Public TTi

Public ButterCake1
Public ButterCake2
Public IndexMain$

Public que

Public CountHtmlPage, CountHtmlPage2



Sub SPECIALFOLDER()

sWindowsFolder = GetSpecialfolder(36)
sMyDocsFolder = GetSpecialfolder(5)

Dim R As Long
On Error Resume Next
For R = 0 To 120
'    Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
Next

'a = a

'End

End Sub


Function SpecialFoldersAre()
    LP_RESULT_CALLBACK = MsgBox("Start menu folder is located at : " + GetSpecialfolder(CSIDL_STARTMENU), vbInformation, "Start Menu Folder")
    LP_RESULT_CALLBACK = MsgBox("Favorites folder is located at : " + GetSpecialfolder(CSIDL_FAVORITES), vbInformation, "Favorites Folder")
    LP_RESULT_CALLBACK = MsgBox("Programs folder is located at : " + GetSpecialfolder(CSIDL_PROGRAMS), vbInformation, "Programs Menu Folder")
    LP_RESULT_CALLBACK = MsgBox("Desktop folder is located at : " + GetSpecialfolder(CSIDL_DESKTOP), vbInformation, "Desktop Folder")
    End
End Function



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




Function GetAttribAnClear(FileSpec)
    
    GetAttribAnClear = False
    
    Response = 0
    
    If FS.FileExists(FileSpec) = False Then
        Style = vbYesNo + vbCritical + vbDefaultButton2 + vbMsgBoxSetForeground
        Response = MsgBox("File Missing WebDates -- " + vbCrLf + FileSpec + vbCrLf + Format$(Now, "DD-MMM HH:MM:SS") + vbCrLf + "Description " + Err.Description + vbCrLf + "Err Number" + Str(Err.Number) + vbCrLf + "Errr Source " + Err.Source + vbCrLf + "STOP YES/NO", Style)
        If Response = vbYes Then Stop
        Exit Function
    End If
    
        
    
    Set F = FS.getfile(FileSpec)
    
    If (F.Attributes And FILE_ATTRIBUTE_ARCHIVE) And Err.Number = 0 Then
        GetAttribAnClear = True
        F.Attributes = F.Attributes - FILE_ATTRIBUTE_ARCHIVE
    End If

End Function


Sub Hubble()

Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32

hh$ = Mid$(PathFileT1$, InStrRev(PathFileT1$, "\") + 1)
hh$ = Mid$(hh$, 1, InStrRev(hh$, ".") - 1)
hh3$ = Mid$(PathFileT1$, 1, InStrRev(PathFileT1$, ".") - 1)
If PathFileT2$ <> "" Then hh2$ = PathFileT2$ Else hh2$ = hh$
PathFileT2$ = hh3$ + ".html"

PathFileT3$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + hh$ + ".html"


WXHex1$ = "00000000"
    
If FS.FileExists(PathFileT1$) Then
    LSet WXHex1$ = Hex(m_CRC.CalculateFile(PathFileT1$))
Else
Exit Sub
End If

'If InStr(YtCrc1$, WXHex1$) > 0 And Force = False Then UpLoad = False: Exit Sub
If Force = False Then
For R = 0 To UBound(CRCChk$)
    If InStr(CRCChk(R), WXHex1$ + " " + PathFileT1$) > 0 And CRCChk(R) <> "" Then
        UpLoad = False: Exit Sub
    End If
Next
End If

tabb = 0
For R = 0 To UBound(CRCChk$)
    ScanPath.lblCount = Trim(Str(R))
    DoEvents
    If InStr(CRCChk(R), PathFileT1$) > 0 And CRCChk(R) <> "" Then
        tabb = 1
        CRCChk(R) = WXHex1$ + " " + PathFileT1$
    End If
Next
If R > UBound(CRCChk$) Then
ReDim Preserve CRCChk$(UBound(CRCChk$) + 1)
End If
If tabb = 0 Then
    CRCChk(R) = WXHex1$ + " " + PathFileT1$
End If

FreeF1 = FreeFile
Open PathFileT1$ For Input As FreeF1
huh$ = Input$(LOF(FreeF1), FreeF1)
Close #FreeF1

r45$ = Header1$ + hh2$ + " ---" + Header2$ + hh2$ + " ---" + Header3$ + huh$ + Header4$
FreeF1 = FreeFile
Open PathFileT3$ For Output As FreeF1
Print #FreeF1, r45$
Close #FreeF1

If UpLoad = True Then
    M = M + 1: Mimp2(M) = PathFileT3$
End If
End Sub



Private Sub Check1_Click()
    If Activate2 = True Then
        Check1 = vbUnchecked
        Exit Sub
    End If
    If Check1 = vbChecked Then
        DayGo = True
        List3.AddItem "Day Mode Selected..."
    End If
End Sub

Private Sub Check2_Click()
    If Activate2 = True Then
        Check2 = vbUnchecked
        Exit Sub
    End If
    If Check2 = vbChecked Then
        WkGo = True
        DayGo = True: Check1 = vbChecked
        List3.AddItem "Week Mode Selected also Day Mode with That..."
    End If
End Sub


Public Sub Form_Load()
'End

'frmFTP.Show
'Exit Sub

Call SPECIALFOLDER


'End

Set FS = CreateObject("Scripting.FileSystemObject")

'FAVS QUICK WORK AROUND DONT RUN CODE UNLESS FAVS CHANGED
f1 = FreeFile
Open App.Path + "\" + App.EXEName + " Fav ChkSum.txt" For Input As #f1
Line Input #f1, f22$
Close #1
fg = Val(f22$)

On Error Resume Next
Err.Clear

FAV = GetSpecialfolder(6)

Set F = FS.GetFolder(FAV)
If Err.Number > 0 Then MsgBox "Error This Dir " + FAV: Stop

f33 = F.Size
If f33 = fg Then
    'List3.AddItem "Processing Favs Done Was Nothing to do - Size of Folder Same"
    'If IsIDE = False Then Exit Sub
    'test
    End
    Exit Sub
End If
'--------------------------------------------------------------------




GoTo Skip

ListView1.ListItems.Clear
ScanPath.txtPath = "D:\MY WEBS\MatthewLan.Com Web\Photos\thumbnails"
ScanPath.chkSubFolders = vbUnchecked
ScanPath.cboMask.Text = "*.jpg"
Call ScanPath.cmdScan_Click

ReDim Preserve Mimp2(ScanPath.ListView1.ListItems.Count)

For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    Mimp2(we) = A1$ + B1$
Next

ReDim Preserve Mimp2(20)

CountHtmlPage2 = 0

aMain.Show

Call ConnectToFTP
TimerPutToWeb.Enabled = True

Exit Sub



Skip:

If App.PrevInstance = True Then End
        
        
If FS.FileExists(App.Path + "\TempDone.txt") Then Kill App.Path + "\TempDone.txt"
        
        
Open App.Path + "\CRC's.txt" For Input As #1
Open App.Path + "\CRC's.txtdd" For Output As #2
Do
    Line Input #1, ll
    QQ = 0
    If InStr(LCase(ll), "\awstats") > 0 Then QQ = 1
    If InStr(LCase(ll), "\photos") > 0 Then QQ = 1
    If QQ = 0 Then Print #2, ll
Loop Until EOF(1)
Close #1, #2
Kill App.Path + "\CRC's.txt"
Name App.Path + "\CRC's.txtdd" As App.Path + "\CRC's.txt"

'LastCommand = "Start"
NotFlag = 1

If IsIDE = False And FindWinPart = True Then End

If IsIDE = False Then
    'Dim TT As Long
    'TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit ""D:\VB6\VB-NT\00_Best_VB_01\WebDates_Ace\WebDates.vbp""", vbMinimizedNoFocus)
    'Call KingMod.Process_Low(TT)
    'End
End If

CountD1 = -5: CountD2 = -5: CountD3 = -5

Open App.Path + "\StartFinishTime.txt" For Input As #1
Line Input #1, o1$
Line Input #1, o2$
Close #1

List3.AddItem o1$
List3.AddItem o2$

List3.AddItem "Select Go Awesome Mode Now If Want... Or Day Or Week Mode..."

DimSize = 15000

Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Height = List3.Height + 1130
Me.Width = List3.Width + 120


'If IsIDE = False Then
If Command$ <> "" Then
    Me.WindowState = 1
Else
    Me.WindowState = 0
End If
    
'Me.WindowState = 1

ReDim CRCChk$(DimSize)
ReDim CRCChkDate$(DimSize)
  
  
aMain.Show
DoEvents
  
'1st timer ran for Load and Run is
'Private Sub Timer1_Timer()
  
'If Command$ = "MinimizedAutoAWE" Then
'End If
'Got IT
  
End Sub

Private Sub AweSomeTimer_Timer()
Call AweSomeMode

If TimeDD = 0 Then TimeDD = Now + TimeSerial(0, 0, 20)
If TimeDD < Now Then

TimeDD = Now + TimeSerial(0, 0, 30)


List3.AddItem "CD --- /httpdocs/ - To Keep Alive"
'List3.ListIndex = List3.ListCount - 1
List3.Refresh
    
Call Execute_FTP_Command("CD /httpdocs/", ErrStat)

End If

End Sub
Sub AweSomeModeSubTimer_Timer()
TimeDD = Now + TimeSerial(0, 0, 30)

If (Inet1.StillExecuting = True) Then Exit Sub

Call PutToWeb

End Sub

Sub AweSomeMode()




Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32

'FileSpec = "D:\MY WEBS\MatthewLan.Com Web\index-main.html"
'Tru = GetAttribAnClear(FileSpec)
'If Tru = True Then fs.Copyfile FileSpec, "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Secure\index-main.html"

CountHtmlPage2 = CountHtmlPage2 + 1
ScanPath.lblCount = Str(CountHtmlPage2) + " of " + Str(UBound(Mimp2))
If CountHtmlPage2 > UBound(Mimp2) Then

If JobRun <> 0 Then
    
    If OlJobRun <> JobRun2 And JobRun2 > 0 Then
        List3.AddItem ""
        List3.AddItem "Whole Run No Jobs ... Monitoring ..."
        
        'Exit on one pass
        If Command$ = "Awesome" Then Exit Sub 'Unload Main: Exit Sub
    
    End If
    OlJobRun = JobRun2
    JobRun2 = 0
End If

CountHtmlPage2 = 1

If JobRun = 0 Then
    JobRun = 2
    List3.AddItem ""
    List3.AddItem "No Jobs On First Run... Monitoring ..."
    
    'Exit on one pass
    If Command$ = "Awesome" Then Exit Sub ':Unload Main: Exit Sub

End If
End If

'OldCountHtmlPage2 = CountHtmlPage2

PuTT1 = Mimp2(CountHtmlPage2)

Set F = FS.getfile(PuTT1)
If MimDate(CountHtmlPage2) <> F.DateLastModified Then
MimDate(CountHtmlPage2) = F.DateLastModified


WXHex1$ = "00000000"
LSet WXHex1$ = Hex(m_CRC.CalculateFile(PuTT1))
DetaiL3 = False
For R = 0 To UBound(CRCChk$)
    If CRCChk$(R) = WXHex1$ + " " + PuTT1 And CRCChk(R) <> "" Then
        DetaiL3 = True
    End If
Next
If DetaiL3 = False Then
List3.AddItem ""
List3.AddItem Format$(Now, "HH:MM:SSa/p ") + Mimp2(CountHtmlPage2)
JobRun = True
JobRun2 = 1
JobOn = True
'UploadFilet1$ = ""
Call PutToWeb
AweSomeTimer = False
AweSomeModeSubTimer = True
End If
End If

'Sub FindAndWriteCRCChkList()

End Sub


Sub AwesomeModeStart()

InAwesome = True

If counter = 0 Then counter = UBound(Mimp2)

ReDim Preserve Mimp2(counter)

M = counter
List3.AddItem "Going to Process These in Awesome Mode"
List3.AddItem "-----------------"

For r1 = 1 To UBound(Mimp2)
    List3.AddItem Trim(Str(r1)) + " -- " + Mid$(Mimp2(r1), 49)
Next
List3.AddItem "-----------------"
List3.AddItem "---------Go Awesome Monitor Files For Changes"
List3.AddItem "-----------------"

Timer1.Enabled = False
Timer2.Enabled = False

ReDim MimDate(UBound(Mimp2)) As Date

Call ConnectToFTP

'FTP
TimerPutToWeb.Enabled = True

AweSomeTimer.Enabled = True



End Sub


Sub NIceFileNames()

        If InStr(A1$ + B1$, "\autoindex") > 0 Then Exit Sub
        If InStr(A1$ + B1$, "awstats") > 0 Then
            Exit Sub
        End If
        Dim D1
        D1 = B1$
        'D1 = "Hdate1991_sdjmall.html"
        
        'If Mid$(D1, 1, 1) = "i" Then Stop
        Mid$(D1, 1, 1) = UCase(Mid$(D1, 1, 1))
        
        If Mid$(D1, 1, 1) = "I" Then
            If Mid$(LCase(D1), 1, 3) = "ind" Then Mid$(D1, 1, 3) = "ind"
        End If
        
        ef2 = Mid$(D1, 1, InStrRev(D1, ".") - 1)
        If InStr(LCase(D1), "dj") < 7 And InStr(LCase(D1), "dj") > 0 Then
            Mid$(D1, InStr(LCase(D1), "dj"), 2) = "DJ"
        End If
        
        For R = 1 To InStrRev(D1, ".") - 1
            txd$ = Mid$(D1, R, 1)
            If IsNumeric(txd$) = True Then
                If IsNumeric(Mid(ef2, R)) Then Exit For
                Err.Clear
                On Error Resume Next
                cool = 0
                txd2$ = UCase(Mid$(ef2, R))
                For R2 = 1 To Len(ef2)
                    t = Asc(Mid(txd2$, R2, 1))
                    t1 = t < 48
                    t2 = t > 57 And t < 65
                    t3 = t > 70
                    If t1 Or t2 Or t3 Then cool = 1: Exit For
                Next
                
                                
                tri = Val("&h" + Mid$(ef2, R))
                'If IsNumeric(tri) Then Stop
                If (IsNumeric(tri) And tri <> 0 Or Err.Number > 0) And cool = 0 Then
                    If Err.Number = 0 Then Exit For
                    On Error GoTo 0
                End If
                On Error GoTo 0
            End If
            
            If InStr("-_@!""£$%^&*()=}]{[~# @':;?/><,1234567890`¬¦", txd$) > 0 Then
                Mid$(D1, R + 1, 1) = UCase(Mid$(D1, R + 1, 1))
            End If
                
        
        Next
        On Error GoTo 0
    
    If D1 <> B1$ Then
        
        Name A1$ + B1$ As A1$ + D1
        
'        Name A1$ + D1 As A1$ + B1$
        LastModifiedToCurrent (A1$ + D1)
        Open App.Path + "\TempDone.txt" For Append As #1
            Print #1, B1$
        Close #1
    End If
    B1$ = D1


End Sub

Function StringPassChk(Strn2)
    
    milk = 0
    Strn2 = LCase(Strn2)
    If FS.FileExists(Strn2) = False Then StringPassChk = True: Exit Function
    
    If Mid$(Strn2, Len(Strn2) - 3) = ".txt" Then
        If FS.FileExists(Mid$(Strn2, 1, Len(Strn2) - 3) + "html") Then
            milk = 1
        End If
    End If
    If InStr(Strn2, ".") > 0 Then
        If IsNumeric(Mid$(LCase(Strn2), InStrRev(Strn2, "."))) = True Then
            milk = 1
        End If
    End If

    If GoCrazyCheck.Value = vbChecked Then
    If Mid$(Strn2, Len(Strn2) - 3) = ".txt" Then
        If FS.FileExists(Mid$(Strn2, 1, Len(Strn2) - 3) + "html") Then
            milk = 1
        End If
    End If
    If InStr(Strn2, ".") > 0 Then
        If IsNumeric(Mid$(LCase(Strn2), InStrRev(Strn2, "."))) = True Then
            milk = 1
        End If
    End If


    End If

    
    'someink else gets them
    If InStr(LCase(Strn2), ".wav") > 0 Then milk = 1
    'someink else gets them
    If InStr(LCase(Strn2), ".log") > 0 Then milk = 1
    If InStr(LCase(Strn2), ".mmf") > 0 Then milk = 1
    'Thumbs
    If InStr(LCase(Strn2), ".eml") > 0 Then milk = 1
    If InStr(LCase(Strn2), ".db") > 0 Then milk = 1
    If InStr(LCase(Strn2), ".bak") > 0 Then milk = 1
    If InStr(LCase(Strn2), ".zip") > 0 Then milk = 1
    If InStr(LCase(Strn2), ".tmp") > 0 Then milk = 1
    If InStr(LCase(Strn2), ".lnk") > 0 Then milk = 1
    If InStr(Strn2, "\autoindex") > 0 Then milk = 1
    If InStr(Strn2, "awstats") > 0 Then milk = 0
    If InStr(Strn2, "\var") > 0 Then milk = 1
    
    Set F = FS.getfile(Strn2)
    
    If F.Size < 1 Or F.Size > (1024! ^ 2) * 8 Then
        milk = 1
    End If
    If InStr(LCase(Strn2), ".rar") > 0 Then milk = 1
    If InStr(Strn2, "\-") > 0 Then milk = 1

    If milk = 1 Then StringPassChk = True
    'If milk = 1 Then Debug.Print Strn2
End Function


Sub Form_Activate2()
If InAwesome = True Then Exit Sub


On Error GoTo ErrTrp



List3.AddItem "Processing Delete Front Page Junk..."
Call Del_FrontPage_Junk
List3.AddItem "Processing Delete Front Page Junk Done..."




DryRunMode = False
'DryRunMode = True

ReDim Mimp2(DimSize)

rg = 0
eo = FreeFile
Open App.Path + "\CRC's.txt" For Input As #eo
Do
    If rg >= UBound(CRCChk$) Then
        ReDim Preserve CRCChk$(UBound(CRCChk$) + 100)
    End If
    Line Input #eo, CRCChk(rg)
    If FS.FileExists(Mid$(CRCChk(rg), 10)) = True Then
    'Set F = fs.getfile(Mid$(CRCChk(rg), 10))
    'CRCChkDate(rg) = Format$(F.DateLastModified, "YYYY-MM-DD HH:MM:SS") + " " + Mid$(CRCChk(rg), 10)
    If CRCChk(rg) <> "" Then rg = rg + 1
    End If
    
    Loop Until EOF(eo)
Close #eo
ReDim Preserve CRCChk$(rg)

rg = 0
eo = FreeFile
Open App.Path + "\CRC'sDate.txt" For Input As #eo
Do
    If rg >= UBound(CRCChkDate$) Then
        ReDim Preserve CRCChkDate$(UBound(CRCChkDate$) + 100)
    End If
    Line Input #eo, CRCChkDate(rg)
    If CRCChkDate(rg) <> "" Then
        rg = rg + 1
    End If
Loop Until EOF(eo)
Close #eo
ReDim Preserve CRCChkDate$(rg)

List3.AddItem "Processing Thumbs 1 OF 2"
Dim Thumbsx
Thumbsx = "D:\MY WEBS\MatthewLan.Com Web\Photos\thumbnails\"
ListView1.ListItems.Clear
ScanPath.txtPath = Thumbsx
ScanPath.chkSubFolders = vbUnchecked
ScanPath.cboMask.Text = "*.jpg"
Call ScanPath.cmdScan_Click

Dim HUGEJPG
For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
        
    HUGEJPG = HUGEJPG + "** " + B1$

Next

List3.AddItem "Processing Thumbs 2 OF 3"
ListView1.ListItems.Clear
ScanPath.txtPath = "D:\MY WEBS\MatthewLan.Com Web\Photos\galleries\"
ScanPath.chkSubFolders = vbChecked
ScanPath.cboMask.Text = "*.jpg"
Call ScanPath.cmdScan_Click
xxes = ScanPath.ListView1.ListItems.Count
For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    If InStr(HUGEJPG, "** " + B1$) > 0 Then
        Mid$(HUGEJPG, InStr(HUGEJPG, "** " + B1$), Len("** " + B1$)) = Space(Len("** " + B1$))
        counthh = counthh + 1
    End If
Next

Debug.Print Str(counthh)
List3.AddItem "Processing Thumbs 3 OF 3"
ListView1.ListItems.Clear
ScanPath.txtPath = Thumbsx
ScanPath.chkSubFolders = vbUnchecked
ScanPath.cboMask.Text = "*.jpg"
Call ScanPath.cmdScan_Click

For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    xxi = 0
    If InStr(HUGEJPG, "** " + B1$) > 0 Then
        FS.moveFile A1 + B1$, A1$ + "-Not_Used\" + B1$
        countjp = countjp + 1
    Else
        If InStr(HUGEJPG, "** " + B1$) > 0 Then Mid$(HUGEJPG, InStr(HUGEJPG, "** " + B1$), Len("** " + B1$)) = Space(Len("** " + B1$))
        xxi = 1
    End If
    If xxi = 0 Then dff = dff + A1$ + B1$ + vbCrLf

Next
'You know the best Leave a Empty Plate
'ww = Trim(HUGEJPG)
HUGEJPG = ""
DoEvents

'TOO HARD FOR NOW NEED LIST ONES NEEDED
'Debug.Print dff

If countjp = 0 Then
    List3.AddItem "Processing Thumbs <Zero> Were Needed to move"
Else
    List3.AddItem "Processing Thumbs --" + Str(countjp) + " -- Were Moved"
End If

If counthh = ScanPath.ListView1.ListItems.Count Then
    List3.AddItem "Processing Thumbs -- All Thumbs are Present in Photos..."
End If

List3.AddItem "Processing Photos -- You Have" + Str(xxes) + " Photos in Gallery"
List3.AddItem "Processing Photos -- You Have" + Str(ScanPath.ListView1.ListItems.Count) + " Thumbs - (¯`'•.¸ -‹(•¿•)›- ¸.•'´¯) -" + Str(countjp) + " Were Moved"



List3.AddItem "Processing File Handling..."
FileSpec = "D:\MY WEBS\MatthewLan.Com Web\index-Main.html"

'FileSpec =
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile FileSpec, "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Secure\index-Main.html"

'MsgBox YtCrc1$
Dim GoCC

If GoCrazyCheck.Value = vbChecked Then GoCC = "%"

M = 0

' % Awesome
' * Dialy
' ^ weekly


'M = M + 1: Mimp2(M) = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\sawyouonline.php"
'm = m + 1: Mimp2(m) = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\sawyouonline2.php"

10

M = M + 1: Mimp2(M) = GoCC + "D:\MY WEBS\MatthewLan.Com Web\#1-IPData.txt"
M = M + 1: Mimp2(M) = GoCC + "D:\MY WEBS\MatthewLan.Com Web\.htaccess"
M = M + 1: Mimp2(M) = GoCC + "D:\MY WEBS\MatthewLan.Com Web\index.php"
M = M + 1: Mimp2(M) = GoCC + "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\index.php"
M = M + 1: Mimp2(M) = GoCC + "D:\MY WEBS\MatthewLan.Com Web\Secure_Folder\index.php"
M = M + 1: Mimp2(M) = GoCC + "D:\MY WEBS\MatthewLan.Com Web\robots.txt"
M = M + 1: Mimp2(M) = GoCC + "D:\MY WEBS\MatthewLan.Com Web\sitemap.xml"

M = M + 1: Mimp2(M) = GoCC + "D:\MY WEBS\MatthewLan.Com Web\-0work\SiteMainMapRobots.html"
M = M + 1: Mimp2(M) = GoCC + "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\SiteMainMap.html"
    
M = M + 1: Mimp2(M) = GoCC + "D:\MY WEBS\MatthewLan.Com Web\favicon.ico"

On Error Resume Next
Open App.Path + "\TempDone.txt" For Input As #1
EGU = LCase(Input(LOF(1), 1))
Close #1
On Error GoTo 0
On Error GoTo ErrTrp



ScanPath.txtPath = "D:\MY WEBS\MatthewLan.Com Web\"
ScanPath.chkSubFolders = vbChecked
ScanPath.cboMask.Text = "*.*"
Call ScanPath.cmdScan_Click

For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    'If InStr(LCase(A1), "photos") > 0 Then Stop
    milk = 0
    If Len(ScanPath.txtPath.Text) = Len(A1$) Then milk = 1
    
    If StringPassChk(A1$ + B1$) = True Then milk = 1
    
    'ODD JOB MAN HANDSOME MAN
    
    If milk = 0 Then
        Call NIceFileNames
        M = M + 1: Mimp2(M) = A1$ + B1$
    End If
Next



'FILE COPYS


Tx$ = "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\#Wave Sounds\DOBFIG4 01.WAV"
Td$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\NHS\Abuse_At_Rutland_NHS\Loggs\DOBFIG4 01.WAV"
FileSpec = Tx$
'Tru = GetAttribAnClear(FileSpec)
'If Tru = True Then FS.copyfile Tx$, Td$
'M = M + 1: Mimp2(M) = Td$

Tx$ = "D:\VB6\VB-NT\00_Best_VB_01\WinAmp Logger\00 Music Info Logger.txt"
Td$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\00 Music Info Logger.txt"
FileSpec = Tx$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile Tx$, Td$
M = M + 1: Mimp2(M) = "*" + Td$

Tx$ = "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\2Total_Music_Day_Log.txt"
Td$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\2Total_Music_Day_Log.txt"
FileSpec = Tx$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile Tx$, Td$
M = M + 1: Mimp2(M) = "*" + Td$

Tx$ = "D:\VB6\VB-NT\#1 Documents\Programming Snippets.txt"
Td$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\Programming\Programming Snippets.txt"
FileSpec = Tx$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile Tx$, Td$
M = M + 1: Mimp2(M) = Td$


Tx$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\NHS\Abuse_At_Rutland_NHS\Loggs\"
If FS.FolderExists(Tx$) = True Then
ScanPath.txtPath = "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\MATT-555ROIDS\00_Text_Data_02"
ScanPath.chkSubFolders = vbChecked
ScanPath.cboMask.Text = "*.txt"
Call ScanPath.cmdScan_Click

For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    If InStr(B1$, "00 Knocker Boy") > 0 And Len(B1$) < 22 Then
        FileSpec = A1$ + B1$
        If Dir$(FileSpec) <> "" Then
            Tru = GetAttribAnClear(FileSpec)
            If Tru = True Then FS.copyfile A1$ + B1$, Tx$ + B1$
            M = M + 1: Mimp2(M) = "*" + Tx$ + B1$
        End If
    End If
Next
End If


20


ScanPath.txtPath = "D:\MY WEBS\MatthewLan.Com Web\awstats2\var"
ScanPath.chkSubFolders = vbUnchecked
ScanPath.cboMask.Text = "*.*"
Call ScanPath.cmdScan_Click

'new start
For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    If InStr(B1$, "MyLogg") > 0 Then
        M = M + 1: Mimp2(M) = A1$ + B1$
    End If
Next



ScanPath.txtPath = "D:\05 Media\Digital Wave Player M Drive\Bad Ass Copys Ref"
ScanPath.chkSubFolders = vbUnchecked
ScanPath.cboMask.Text = "*.wav;*.txt"
Call ScanPath.cmdScan_Click


For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    v1$ = B1$
    Set F = FS.getfile(A1$ + B1$)
    FS2 = F.Size
    If (FS2 / 1024 ^ 2) < 5 Then
        Do
        eet = Len("D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\NHS\Malice_Abuse_At_Rutland_NHS\" + B1$)
        If eet > 255 Then
            B1$ = Mid$(B1$, 1, Len(B1$) - 5) + ".wav"
        Else
            Exit Do
        End If
        Loop Until 1 = 2
    
        QQ$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\NHS\Malice_Abuse_At_Rutland_NHS\" + B1$
        Set F = FS.getfile(QQ$)
        fs3 = F.Size
        
        If FS2 <> fs3 Then
            FileSpec = A1$ + v1$
            Tru = GetAttribAnClear(FileSpec)
            If Tru = True Then FS.copyfile A1$ + v1$, QQ$
        End If
        M = M + 1: Mimp2(M) = QQ$
    End If
Next


Dim PathFile1$


CountHtmlPage2 = 0

30

FileT1$ = "01 Main Rewrite-Book.xls"
'sMyDocsFolder
PathFile1$ = sMyDocsFolder + "\CD-ROM\" + FileT1$
PathFile1$ = sMyDocsFolder + "\CD-ROM\" + FileT1$
FileSpec = PathFile1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFile1$, "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\01 Main Rewrite-Book.xls"

FileT1$ = "01 Total List.html"
PathFile1$ = sMyDocsFolder + "\CD-ROM\" + FileT1$
FileSpec = PathFile1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFile1$, "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\01 Total List.html"

FileT1$ = "2sort-lxx.txt"
PathFile1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFile2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFile1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFile1$, PathFile2$

FileT1$ = "2sort-lxx2.txt"
PathFile1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFile2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFile1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFile1$, PathFile2$

FileT1$ = "2sort-lxx3.txt"
PathFile1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFile2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFile1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFile1$, PathFile2$

FileT1$ = "BangList Total Logg.html"
PathFile1$ = Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Banging_Tunes\" + FileT1$
PathFile2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFile1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFile1$, PathFile2$

Tx$ = Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Weather\Weather Log Plain.html"
FileSpec = Tx$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile Tx$, "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\Weather Log.html"

Tx$ = Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Weather\Weather Log Mini.html"
FileSpec = Tx$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile Tx$, "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\Weather Log Mini.html"

Tx$ = Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Weather\Weather Log Text.txt"
FileSpec = Tx$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile Tx$, "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\Weather Log Text.txt"

Tx$ = Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Date1991\Date1991.html"
FileSpec = Tx$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile Tx$, "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\Date1991.html"

Tx$ = Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Date1991\Date1991_Small.html"
FileSpec = Tx$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile Tx$, "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\Date1991_Small.html"

Tx$ = Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Date1991\Date1991_Secure.html"
FileSpec = Tx$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile Tx$, "D:\MY WEBS\MatthewLan.Com Web\Secure_Folder\Cool\Date1991_Secure.html"


40

List3.AddItem "Processing File Coverts to Html..."


Statc$ = "<!-- Start of StatCounter Code -->" + vbCrLf
Statc$ = Statc$ + "<script type=""text/javascript"" language=""javascript"">" + vbCrLf
Statc$ = Statc$ + "var sc_project=585393;" + vbCrLf
Statc$ = Statc$ + "var sc_partition=4;" + vbCrLf
Statc$ = Statc$ + "var sc_security=""8832a018"";" + vbCrLf
Statc$ = Statc$ + "var sc_invisible=1;" + vbCrLf
Statc$ = Statc$ + "</script>" + vbCrLf
Statc$ = Statc$ + "<script type=""text/javascript"" language=""javascript"" src=""http://www.statcounter.com/counter/counter.js""></script><noscript><a href=""http://www.statcounter.com/"" target=""_blank""><img  src=""http://c5.statcounter.com/counter.php?sc_project=585393&amp;java=0&amp;security=8832a018&amp;invisible=1"" alt=""php hit counter"" border=""0""></a> </noscript>" + vbCrLf
Statc$ = Statc$ + "<!-- End of StatCounter Code -->" + vbCrLf


'Grab22$ = "<body style=""font-family: Arial"" link=""#00FFFF"" vlink=""#9066FF"" alink=""#FFFFFF"">" + vbCrLf
Grab22$ = "<br><br>" + vbCrLf
Grab22$ = Grab22$ + "<div align=""left"">" + vbCrLf
Grab22$ = Grab22$ + "  <table border=""5"" style=""font-family: Rockwell; font-size: 14pt"">" + vbCrLf
Grab22$ = Grab22$ + "      <tr>" + vbCrLf
Grab22$ = Grab22$ + "        <td width=""100%"">" + vbCrLf
Grab22$ = Grab22$ + "         <p align=""center""><font size=""4""><a href=""http://matthewlan.com"">Back to Home Page</a></font></td>" + vbCrLf
Grab22$ = Grab22$ + "      </tr>" + vbCrLf
Grab22$ = Grab22$ + "  </table>" + vbCrLf
Grab22$ = Grab22$ + "</div>" + vbCrLf
Grab22$ = Grab22$ + Statc$
Grab22$ = Grab22$ + "</BODY></HTML>" + vbCrLf


Header1$ = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" + vbCrLf
Header1$ = Header1$ + "<HTML><HEAD>" + vbCrLf
Header1$ = Header1$ + "<TITLE>"
Header2$ = "</TITLE>" + vbCrLf
Header2$ = Header2$ + "<META http-equiv=Content-Type content=""text/html; charset=windows-1252""></HEAD>" + vbCrLf
Header2$ = Header2$ + "<META NAME=""DESCRIPTION"" CONTENT="""
Header3$ = " Last Updated " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + """>" + vbCrLf
'Header3$ = " " + """>" + vbCrLf
Header3$ = Header3$ + "<META NAME=""robots"" content=""index,nofollow"">" + vbCrLf
Header3$ = Header3$ + "<BODY><PRE>" + vbCrLf

huh$ = "<b><font face=""Arial"" size=""4"" color=""#000000""><a href=""http://MatthewLan.com"">http://MatthewLan.com</a> --- <a href=""mailto:matt.lan@btinternet.com"">Matt.Lan@btinternet.com</a></font></b>"
huh$ = huh$ + "<font face=""Arial"" size=""3""><b>"
Header3$ = Header3$ + huh$ + "<br>" + vbCrLf

Header4$ = "</b></PRE>"
'Header4$ = Header4$ + Grab22$
Header4$ = Header4$ + "</BODY></HTML>"

PathFileT1$ = Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Weather\Weather Log Text.Txt"
PathFileT2$ = ""
UpLoad = False
Call Hubble

50

'Force = True
Tx$ = sMyDocsFolder + "\CallerID\2BusCab.txt"
FileSpec = Tx$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile Tx$, "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\2BusCab.txt"

PathFileT1$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\2BusCab.txt"
PathFileT2$ = "Matt Lan's Bus Cab Log"
UpLoad = True
Call Hubble

Tx$ = sMyDocsFolder + "\CallerID\text-9\Mental Health Text\Body\1PULSER8.TXT"
FileSpec = Tx$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile Tx$, "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\1PulseR8.txt"
PathFileT1$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\1PulseR8.txt"
PathFileT2$ = "Matt Lan's Walk and Pulse Count Log"
UpLoad = True
Call Hubble

FileT1$ = "2sort-lxx3.txt"
PathFileT1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFileT1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFileT1$, PathFileT2$
PathFileT2$ = "Double Words Sorted Only"
UpLoad = True
'UpLoad = False
Call Hubble

If UpLoad = True Then Mimp2(M) = "*" + Mimp2(M - 1)


FileT1$ = "2sort-lxx2.txt"
PathFileT1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFileT1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFileT1$, PathFileT2$
PathFileT2$ = "Double Words UnSorted Only"
UpLoad = True
'UpLoad = False
Call Hubble
If UpLoad = True Then Mimp2(M) = "*" + Mimp2(M)

FileT1$ = "2sort-lxx.txt"
PathFileT1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFileT1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFileT1$, PathFileT2$
PathFileT2$ = "Double Words Both Sort and UnSorted"
UpLoad = True
'UpLoad = False
Call Hubble
If UpLoad = True Then Mimp2(M) = "*" + Mimp2(M)
    
FileT1$ = "2Nico Gum Give Up Smoking.txt"
PathFileT1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFileT1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFileT1$, PathFileT2$
PathFileT2$ = "Matt Lan 's -Nico Gum Give Up Smoking"
UpLoad = True
Call Hubble

FileT1$ = "2CidRunMeDayCountLog.txt"
PathFileT1$ = "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFileT1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFileT1$, PathFileT2$
PathFile2$ = "Matt Lan 's - Keyboard an Mouse Log"
UpLoad = True
Call Hubble
If UpLoad = True Then Mimp2(M) = "*" + Mimp2(M)

FileT1$ = "Date1991 Next & Last.txt"
PathFileT1$ = sMyDocsFolder + "\03 NotePad Files\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFileT1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFileT1$, PathFileT2$
PathFile2$ = "Matt Lan 's - Next and Last Key Moon Sun Dates"
UpLoad = True
'Force = True
Call Hubble
'Force = False


FileT1$ = "My Email Signature Line.txt"
PathFileT1$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
'fs.Copyfile PathFileT1$, PathFileT2$
PathFile2$ = "Matt Lan 's - My Email Signature Line"
UpLoad = True
Call Hubble


'Force = True
FileT1$ = "2 Kadaitcha Man #1.txt"
PathFileT1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFileT1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFileT1$, PathFileT2$
PathFileT2$ = "Kadaitcha Man #1"
UpLoad = True
Call Hubble

FileT1$ = "2 Kadaitcha Man #2.txt"
PathFileT1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFileT1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFileT1$, PathFileT2$
PathFileT2$ = "Kadaitcha Man #2"
UpLoad = True
Call Hubble

FileT1$ = "2 Kadaitcha Man #3.txt"
PathFileT1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFileT1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFileT1$, PathFileT2$
PathFileT2$ = "Kadaitcha Man #3"
UpLoad = True
Call Hubble

FileT1$ = "2 Kadaitcha Man #4.txt"
PathFileT1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFileT1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFileT1$, PathFileT2$
PathFileT2$ = "Kadaitcha Man #4"
UpLoad = True
Call Hubble

FileT1$ = "2 Kadaitcha Man #5.txt"
PathFileT1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFileT1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFileT1$, PathFileT2$
PathFileT2$ = "Kadaitcha Man #4"
UpLoad = True
Call Hubble

FileT1$ = "2 Kadaitcha Man #6.txt"
PathFileT1$ = sMyDocsFolder + "\CallerID\" + FileT1$
PathFileT2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\" + FileT1$
FileSpec = PathFileT1$
Tru = GetAttribAnClear(FileSpec)
If Tru = True Then FS.copyfile PathFileT1$, PathFileT2$
PathFileT2$ = "Kadaitcha Man #4"
UpLoad = True
Call Hubble


60



'----------------------------------------------------------------------------------------




IndexMain$ = "index-Main.html"
'IndexMain$ = "index.html"

ReDim Detail(DimSize)
ReDim Detail1(DimSize)
ReDim DetaiL2(DimSize)
'ReDim CRCChk$(DimSize)
    

'WebPhotComplete = False

Activated = 0


Dim Aplenty20$
Dim Aplenty21$

direc1$ = "D:\MY WEBS\MatthewLan.Com Web\-Temp\Winamp Generated Playlist.html"

If Dir$(direc1$) <> "" Then
    Open direc1$ For Input As #1
    Aplenty20$ = Input(LOF(1), #1)
    Close #1
End If




direc2$ = "C:\Documents and Settings\" + GetUserName + "\Local Settings\Temp\"

File1.Path = direc2$

'this is for Winamp playlist

Call buster(Filespec2$, Idate)

If Filespec2$ <> "" Then
    Open Filespec2$ For Input As #1
    Aplenty21$ = Input(LOF(1), #1)
    Close #1
End If

If Aplenty20$ <> Aplenty21$ And Aplenty21$ <> "" Then
    List3.AddItem "Processing Winamp PlayList..."
    
    Open direc1$ For Output As #1
    Print #1, Aplenty21$;
    Close #1
    
    wedr = InStr(Aplenty21$, "<style")
    Grab223$ = "<html><head><LINK REL=stylesheet TYPE=""text/css"" HREF=""../../Style_Sheets/indexstyle.css"">"
    
    Aplenty21$ = Grab223$ + Mid$(Aplenty21$, wedr)
    
    wedr = InStr(Aplenty21$, "--BODY") - 1
    
    Aplenty21$ = Mid$(Aplenty21, 1, wedr + 2) + Mid$(Aplenty21$, wedr + 33)
    
    Tx$ = "<small><small><font face=""Arial"" color=""#FFBF00"
    
    wedr = InStr(Aplenty21$, Tx$) + Len(Tx$)
    
    zize$ = " Size=""5"""
    
    Aplenty21$ = Mid$(Aplenty21, 1, wedr) + zize$ + Mid$(Aplenty21, wedr)
    
    yg = InStr(Aplenty21$, "<blockquote><p><")
    
    
    xg = wedr
    Do
    Tx$ = "face=""Arial"""
    
    wedr = InStr(wedr + 1, Aplenty21$, Tx$) + Len(Tx$)
    
    Aplenty21$ = Mid$(Aplenty21, 1, wedr - 1) + zize$ + Mid$(Aplenty21, wedr)
    
    Loop Until wedr > yg
    '<blockquote><p><
    
    
    cool2$ = "length)</font><BR>"
    cool2$ = "</font><BR>"
    wedr = InStr(Aplenty21$, cool2$)
    wedr = InStr(Aplenty21$, "</font><BR>")
    
    If wedr = 0 Then MsgBox "Hello Winamp Instr Not Found"
    
    wedr = wedr + Len(cool2$) + 1
    
    
    
    Aplenty21$ = Mid$(Aplenty21, 1, wedr) + Mid$(Aplenty21$, InStr(Aplenty21$, "</small></td></tr></table></div>")) 'wedr + 155)
    
    wedr = InStr(Aplenty21$, "Winamp Gener")
    
    Aplenty21$ = Mid$(Aplenty21, 1, wedr + 24) + " *** http://MatthewLan.Com ***</title><META NAME=""ROBOTS"" CONTENT=""INDEX,FOLLOW"">" + Mid$(Aplenty21$, wedr + 33)
        
    wedr = InStrRev(Aplenty21$, "</body></h") - 1
    Aplenty21$ = Mid$(Aplenty21$, 1, wedr)
    
    
    Aplenty21$ = Aplenty21$ + Grab22$
    
    ecok = 1
    Do
    ecok = ecok + 1
    ecok = InStr(ecok, Aplenty21$, "[[(({{")
    If ecok = 0 Then Exit Do
    ecok2 = InStr(ecok, Aplenty21$, "}}))]]")
    
    
    If InStr(Mid$(Aplenty21$, ecok, ecok2 - ecok), "Mixes From") Or _
    InStr(Mid$(Aplenty21$, ecok, ecok2 - ecok), "DJ's From") Then
        pliut$ = "<B><font color=""#&HH963396"">"
    Else
        pliut$ = "<B><font color=""#&H00FF00"">"
    End If
    
    aplenty24$ = Mid$(Aplenty21$, 1, ecok + 5)
        
    aplenty24$ = aplenty24$ + pliut$
    aplenty24$ = aplenty24$ + Mid$(Aplenty21$, ecok + 6, ecok2 - ecok - 6)
    aplenty24$ = aplenty24$ + "</B></font>"
    ecok = Len(aplenty24$)
    aplenty24$ = aplenty24$ + Mid$(Aplenty21$, ecok2)
    Aplenty21$ = aplenty24$
    
    Loop Until ecok = 0
    
    
    
    
    direc1$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\html\Winamp_Generated_Playlist.html"
    
    Open direc1$ For Output As #1
    Print #1, Aplenty21$;
    Close #1

Call Mnu_WinAmp_Click

End If


70

DoEvents

'-------------

List3.AddItem "Processing Web Page Top Files Block with Dates..."

' ------------- Start title menu block

ReDim Washad$(100)
ReDim Date2020(100) As Date
ReDim Ee$(100)

Washad$(1) = "<a href=""LoveFolder/index.php?dir=HTML/&file=0Double.html"">Double Words</a>"
Washad$(2) = "<a href=""LoveFolder/index.php?dir=HTML/&file=0Past_Txt.html"">Rhyming + Triple an > - words</a>"
Washad$(3) = "<a href=""LoveFolder/index.php?dir=HTML/&file=Quotes_An_Stuff.html"">Quotes and Stuff</a>"
Washad$(4) = "<a href=""LoveFolder/index.php?dir=HTML/&file=My_Prescription_Charges.html"">My prescription charge's</a>"
Washad$(5) = "<a href=""LoveFolder/index.php?dir=HTML/&file=Joke's.html"">Joke's</a>"
Washad$(6) = "<a href=""LoveFolder/index.php?dir=HTML/&file=Winamp_Generated_Playlist.html"">My Winamp PlayList...</a>"
Washad$(7) = "<a href=""LoveFolder/index.php?dir=HTML/&file=Hallucinations.html"">My Hallucinations</a>"
Washad$(8) = "<a href=""LoveFolder/index.php?dir=HTML/&file=0Acron_Txt.html"">Most Common Words &amp; Most Common Acronym's&nbsp;&nbsp;</a>"
Washad$(9) = "<a href=""LoveFolder/index.php?dir=HTML/&file=Sz_An_Med_Info.html"">Schizophrenia an Med Info</a>"
Washad$(10) = "<a href=""LoveFolder/index.php?dir=HTML/&file=Delusions.html"">My Delusions and Psychoses...</a>"
Washad$(11) = "<a href=""LoveFolder/index.php?dir=HTML/&file=Death.html"">Death's & Suicide's</a>"
Washad$(12) = "<a href=""LoveFolder/index.php?dir=HTML/&file=Matthew_Lan_Cv.html"">My CV</a>"
Washad$(13) = "<a href=""LoveFolder/index.php?dir=HTML/&file=About_Me.html"">About me.</a>"
'Washad$(14) = "<a href=""LoveFolder/index.php?dir=HTML/&file=Trolls.html"">Troll Terminator...</a>"

ttd = 13
ReDim Preserve Washad$(ttd)
ReDim Preserve Date2020(ttd) As Date
ReDim Preserve Ee$(ttd)


direc$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\HTML\"
dired$ = "D:\MY WEBS\MatthewLan.Com Web\"


'For r = 1 To counter
'If InStr(indexf$(r), "<!--Updated123-->") Then Exit For
'Next

'For r1 = r To (r + que) - 2
'indexf$(r1 + 1) = indexf$(r1 + 1) + "<br>"
'Next


ScanPath.txtPath = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\HTML"
ScanPath.chkSubFolders = vbUnchecked
ScanPath.cboMask.Text = "*.txt;*.html;*.html"
Call ScanPath.cmdScan_Click

For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    M = M + 1: Mimp2(M) = A1$ + B1$
Next

For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    FileSpec = A1$ + B1$

    'getting the date here
    Call SetClearAttribbits(FileSpec)

    List5.AddItem Format$(Ddate, "YYYY-MM-DD HH:MM:SS --") + FileSpec

Next

80

Dim Date3 As String
Dim Data5 As String
Dim FileR1$
Dim FileR2$
Dim RRF As String
Dim RRF1 As String
Dim RRF2 As String
'-----------------
'Now main list freq update

'' see that = get data
FreeF1 = FreeFile
Open dired$ + IndexMain$ For Input As #FreeF1
'counter = 0
'Do
'counter = counter + 1
'Line Input #1, indexf$(counter)
'Loop Until EOF(1)
RRF = Input$(LOF(FreeF1), FreeF1)
Close #FreeF1

r1 = InStr(RRF, "<!--Updated122-->")
R2 = InStr(RRF, "<!--Updated123-->")

RRF1 = Mid$(RRF, 1, r1 - 1)
RRF2 = Mid$(RRF, R2)

Data5 = "<!--Updated122-->" + vbCrLf + vbCrLf
Date3 = vbCrLf + "<td valign=""top"" nowrap>" + vbCrLf + vbCrLf

'Puts dates dates are easy
For r1 = List5.ListCount - 1 To 0 Step -1
    If r1 <> 0 Then BR = "<br>" Else BR = ""
    Ddate = DateValue(Mid$(List5.List(r1), 1, 19)) + TimeValue(Mid$(List5.List(r1), 1, 19))
    Date3 = Date3 + "<font face=""Courier New"">" + Format$(Ddate, "DDD DDMMMYY HH:MMa/p") + "</font>" + BR + vbCrLf
    FileR1$ = Mid$(List5.List(r1), InStrRev(List5.List(r1), "\") + 1)
    
    found = False
    For ri = 1 To UBound(Washad$)
        FileR2$ = Washad$(ri)
        If InStr(LCase(FileR2$), LCase(FileR1$)) > 0 Then
            found = True
            Exit For
        End If
    Next
    If found = False Then
        FileR2$ = "<a href=""LoveFolder/index.php?dir=HTML/&file=" + FileR1$ + """>" + Mid$(FileR1$, 1, InStrRev(FileR1$, ".") - 1) + "</a>"
    End If
    
    Data5 = Data5 + FileR2$ + BR + vbCrLf

Next

Date3 = Date3 + vbCrLf


90

'On Error GoTo 0

FreeF1 = FreeFile
Open dired$ + IndexMain$ For Output As #FreeF1
Print #FreeF1, RRF1;
Print #FreeF1, Data5;
Print #FreeF1, Date3;
Print #FreeF1, RRF2;
Close #FreeF1


RRF = ""
RRF1 = ""
RRF2 = ""

'End

List3.AddItem "Processing Sort Non PHP Index..."

'--------This Sorts the non php index

FreeF1 = FreeFile
Open dired$ + IndexMain$ For Binary As #FreeF1
RRF = Input$(LOF(FreeF1), FreeF1)
Close #FreeF1

xxv = 0
Do
    xxv = InStr(xxv + 1, RRF, "LoveFolder/index.php?dir=")
    If xxv > 0 Then
        RRF2$ = Mid$(RRF, 1, xxv - 1)
        xxv2 = InStr(xxv + 23, RRF, "/") - (xxv + 23)
        xxv3 = InStr(xxv + 23, RRF, "=") - (xxv + 23)
        RRF2$ = RRF2$ + "LoveFolder/" + Mid$(RRF, xxv + 23, xxv2) + "/" + Mid$(RRF, xxv + 24 + xxv3)
        RRF = RRF2$
    End If
Loop Until xxv = 0

xxv = 0
Do
    xxv = InStr(xxv + 1, RRF, "href=""http://matthewlan.com/")
    If xxv > 0 Then
        RRF2$ = Mid$(RRF, 1, xxv + 5) + Mid$(RRF, xxv + 5 + 23)
        RRF = RRF2$
    End If
Loop Until xxv = 0


'http://matthewlan.com/
'href="http://matthewlan.com/

DoEvents

Dim IndexMainPhp$

'Reset
IndexMainPhp$ = "index-NoPhp.html"
FreeF1 = FreeFile
Open dired$ + IndexMainPhp$ For Output As #FreeF1
Print #FreeF1, RRF
Close #FreeF1


'--------- Done Sorts the non php index

List3.AddItem "Processing Do Weather..."


'------ Do Weather

'D:\VB6\VB-NT\Weather
FreeF1 = FreeFile
Open Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Weather\Weather Log Small.html" For Input As #FreeF1
ws2$ = Input(LOF(FreeF1), #FreeF1)
Close #FreeF1
ws2$ = Mid$(ws2$, InStr(ws2$, "<div align=""center"">"))
ru = InStr(ws2$, "°")
'ws2$ = StrConv(ws2$, vbUnicode)

'lj = Asc(Mid$(ws2$, (ru * 2) - 1, 1))
'ru = InStr(ws2$, Chr(0) + Chr(176))

FreeF1 = FreeFile
Open dired$ + IndexMain$ For Input As #FreeF1
freef2 = FreeFile
Open dired$ + "index2.html" For Output As #freef2
Do
    Line Input #FreeF1, indexf2$
    indextest$ = indexf2$
'    indexf2$ = StrConv(indexf2$, vbUnicode)
    If xs = 0 Then Print #freef2, indexf2$
    'If xs2 = 0 Then Print #freef2, vbCrLf
    
    If InStr(indextest$, "<!--Updated124-->") Then Print #freef2, ws2$: xs = 1: xs2 = 1
    If InStr(indextest$, "<!--Updated125-->") Then Print #freef2, indexf2$: xs = 0
Loop Until EOF(FreeF1)
Close #FreeF1
Close #freef2





'On Local Error Resume Next
Kill dired$ + IndexMain$
'On Local Error GoTo 0
Name dired$ + "index2.html" As dired$ + IndexMain$
'err.number


'------  Done Weather

List3.AddItem "Processing Weather Done..."


FreeF1 = FreeFile
Open dired$ + IndexMain$ For Input As #FreeF1
RRF = Input$(LOF(FreeF1), FreeF1)
Close #FreeF1

Do
r1 = InStr(RRF, Chr(176))
If r1 > 0 Then Mid$(RRF, r1, 1) = "*"
Loop Until r1 = 0

FreeF1 = FreeFile
Open dired$ + IndexMain$ For Output As #FreeF1
Print #FreeF1, RRF;
Close #FreeF1





'-------update time update top web page description
'Lastupdate2$ = Format$(Now, "DDD DD-MMM-YY HH:MM:SS")

'FreeF1 = FreeFile
'Open dired$ + IndexMain$ For Input As #FreeF1
'RRF = Input$(LOF(FreeF1), FreeF1)
'Close #FreeF1
'r1 = InStr(RRF, "- Updated ")

'gg$ = gg$ + "<META name=""description"" content=""English Double Word's Same First Letter - A Patient's Side of Schizophrenia"
'Mid$(RRF, RRF, Len(" - Updated " + Lastupdate2$)) = " - Updated " + Lastupdate2$
'Close #1


'SiteMap.xml

Open dired$ + "Sitemap.xml" For Input As #1
plin$ = Input(LOF(1), #1)
Close #1

plo = 1
Do
    wop = InStr(plo, plin$, ".html")
    If wop > 0 Then
    
        wop2 = InStrRev(plin$, ".com/", wop) + 5
    
        wop3 = InStr(wop, plin$, "</loc")
    
        mod2 = InStr(wop, plin$, "mod>") + 4
    
        Dim LastModDate
    
        extractfilename$ = Mid$(plin$, wop2, wop3 - wop2)
    
        RR = FS.FileExists(dired$ + extractfilename$)
        
        replacedate$ = Format$(DateSerial(1800, 1, 1), "YYYY-MM-DDTHH:MM:SSZ")
        lastmmoddate = DateSerial(1800, 1, 1)
        
        'make this better alert or remove offending file that no longer exists
        If RR = False Then
            Mid$(plin$, mod2, 20) = replacedate$
            Style = vbYesNo + vbCritical + vbDefaultButton2 + vbMsgBoxSetForeground
            MsgBox "Edit This You Need Take Out Or Rename " + vbCrLf + dired$ + extractfilename$, Style
            Shell "notePad D:\MY WEBS\MatthewLan.Com Web\Sitemap.xml"
        End If
        
        
        If RR = True Then
        Set F = FS.getfile(dired$ + extractfilename$)
    
        replacedate$ = Format$(F.DateLastModified, "YYYY-MM-DDTHH:MM:SSZ")
    
        lastmmoddate = F.DateLastModified
    
    
        Mid$(plin$, mod2, 20) = replacedate$
        
        End If
        
        plo = wop + 5
        List1.AddItem "S-Map " + Format$(LastModDate, "YYYYMMDD HH:MM:SS") + "-" + Format$(F.DateLastModified, "DD MMM YYYY HH:MM:SS Am/Pm") + "-" + extractfilename$

    End If
Loop Until wop = 0


Open dired$ + "Sitemap.xml" For Output As #1
Print #1, plin$;
Close #1

Dim IndePl
For R = 0 To List1.ListCount - 1
DoEvents
    If Mid$(List1.List(R), 1, 4) = "Inde" Then
        IndePl = IndePl + 1
    End If
Next

Dim IndeMa
For R = 0 To List1.ListCount - 1
DoEvents
    If Mid$(List1.List(R), 1, 4) = "S-Ma" Then
        IndeMa = IndeMa + 1
    End If
Next

smoke = List1.ListCount - 1
For R = 0 To smoke
DoEvents
    If Mid$(List1.List(R), 1, 4) = "Inde" Then
        List1.AddItem "ZX" + Format$(IndePl - R, "00") + "-" + List1.List(R)
    End If
Next

plotd = 0
For R = 0 To smoke
DoEvents
    If Mid$(List1.List(R), 1, 4) = "S-Ma" Then
        List1.AddItem "ZZ" + Format$(IndeMa - plotd, "00") + "-" + List1.List(R)
        plotd = plotd + 1
    End If
Next

For R = 0 To List1.ListCount - 1
DoEvents
    If Mid$(List1.List(R), 1, 2) = "ZX" Then
        List2.AddItem Mid$(List1.List(R), 6, 5) + Mid$(List1.List(R), 29)
    End If
Next

For R = 0 To List1.ListCount - 1
DoEvents
If Mid$(List1.List(R), 1, 2) = "ZZ" Then
        List2.AddItem Mid$(List1.List(R), 6, 5) + Mid$(List1.List(R), 29)
    End If
Next

ButterCake = Now + TimeSerial(0, 3, 30)

TTi = 0



If GoCrazyCheck.Value = vbChecked Then
    
    List3.AddItem "Go Crazy Mode..."
    ReDim Preserve Mimp4(M)
    For we = 1 To M
        ScanPath.lblCount = Trim(Str(we))
        'If Mimp4(we) <> "" Then
        'Set F = FS.getfile(Mimp4(we))
        'QQ$ = ""
        '^2 meg
        If Mid(Mimp2(we), 1, 1) = "%" Then Mimp2(we) = Mid$(Mimp2(we), 2)
        If Mid(Mimp2(we), 2, 1) <> ":" Then Mimp2(we) = Mid$(Mimp2(we), 2)
        Done = 0
        If StringPassChk(Mimp2(we)) = True Then Done = 1
        
        If Done = 0 Then
            milky = milky + 1
            Mimp4(milky) = Mimp2(we)
            Done = 1
        End If
        
    Next
    
    ReDim Mimp2(milky)
    For we = 1 To milky
        ScanPath.lblCount = Trim(Str(we))
        Mimp2(we) = Mimp4(we)
    Next
'-----------------------------------
    
    
100

    Open App.Path + "\FTP Logg.txt" For Append As #1
        Print #1, "***************************************************************************************************************"
        Print #1, "***************************************************************************************************************"
        Print #1, Format(Now, "DD-MM-YYYY HH:MM:SS ") + " - Process Start ---- AWESOME MODE"
        Print #1, "***************************************************************************************************************"
        Print #1, "***************************************************************************************************************"
    Close #1
    
    Open App.Path + "\WEB DATES Logg.txt" For Append As #1
        Print #1, "***************************************************************************************************************"
        Print #1, "***************************************************************************************************************"
        Print #1, Format(Now, "DD-MM-YYYY HH:MM:SS ") + " - Process Start ---- LIST BOX OUTPUT"
        Print #1, "***************************************************************************************************************"
        Print #1, "***************************************************************************************************************"
    Close #1

    
    Call AwesomeModeStart
    Exit Sub
End If

Open App.Path + "\FTP Logg.txt" For Append As #1
    Print #1, "***************************************************************************************************************"
    Print #1, "***************************************************************************************************************"
    Print #1, Format(Now, "DD-MM-YYYY HH:MM:SS ") + " - Process Start ---- NORMAL MODE"
    Print #1, "***************************************************************************************************************"
    Print #1, "***************************************************************************************************************"
Close #1

Open App.Path + "\WEB DATES Logg.txt" For Append As #1
    Print #1, "***************************************************************************************************************"
    Print #1, "***************************************************************************************************************"
    Print #1, Format(Now, "DD-MM-YYYY HH:MM:SS ") + " - Process Start ---- LIST BOX OUTPUT"
    Print #1, "***************************************************************************************************************"
    Print #1, "***************************************************************************************************************"
Close #1


110

On Local Error GoTo 0

Call Fav_an_SiteMap

On Local Error Resume Next

'How this Prog starts on timer or soon as
List3.AddItem "Enabled Start Count Down You can Go Manual Mode Now If Want..."


'This pickup here -- GoCrazyCheck_Click

'And Here --- If Activate2 = True And CountD1 = -5 Then CountD1 = 10
'This Does This
'If CountD1 <= 0 Then
'    'After Load All What To Do then Run Internet upLoad
'    Command4.Caption = "Go.."
'    Timer1.Enabled = False
'    Call ConnectToFTP
'    TimerPutToWeb.Enabled = True
'End If

Activate2 = True

Exit Sub

ErrTrp:

If Err.Number = 0 Then Resume Next
            
Style = vbYesNo + vbCritical + vbDefaultButton2 + vbMsgBoxSetForeground
Response = MsgBox("Error in webdates " + vbCrLf + Format$(Now, "DD-MMM HH:MM:SS") + vbCrLf + "Description " + Err.Description + vbCrLf + "Err Number" + Str(Err.Number) + vbCrLf + "Errr Source " + Err.Source + vbCrLf + "M = " + Str(Val(M)) + vbCrLf + "Mimp2(M)=" + Mimp2(M) + vbCrLf + "-----------" + vbCrLf + " -- OR -- " + vbCrLf + "File Missing WebDates -- " + vbCrLf + FileSpec + vbCrLf + "STOP YES/NO", Style)
If Response = vbYes Then Stop: Resume
'Resume


End

End Sub


Sub Sort_IfDoingOnceDayUpload()

Dim WkMark$
Dim DayMark$
Dim TT As Date

On Error Resume Next
    Open "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\MATT-555ROIDS\00_Text_Data_02\05WeekMark.txt" For Input As #1
        Line Input #1, WkMark$
    Close #1
    TT = DateValue(WkMark$) + TimeValue(WkMark$)
On Error GoTo 0
If TT < Now Then
    For R = 1 To 7
        If Weekday(Int(Now) + R) = 2 Then Exit For
    Next
    WkGo = True
    DayGo = True
    Check1 = vbChecked
    Check2 = vbChecked
    'List3.AddItem "Doing Week Mode ....."
    
    Open "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\MATT-555ROIDS\00_Text_Data_02\05WeekMark.txt" For Output As #1
        Print #1, Format(Int(Now) + R + TimeSerial(8, 0, 0), "DD-MM-YYYY HH:MM:SS")
    Close #1
End If

TT = 0
On Error Resume Next
    Open "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\MATT-555ROIDS\00_Text_Data_02\05DayMark.txt" For Input As #1
        Line Input #1, DayMark$
    Close #1
    TT = DateValue(DayMark$) + TimeValue(DayMark$)
On Error GoTo 0

If TT < Now Then
    DayGo = True
    Check1 = vbChecked
    Open "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\MATT-555ROIDS\00_Text_Data_02\05DayMark.txt" For Output As #1
        Print #1, Format(Int(Now) + 1 + TimeSerial(8, 0, 0), "DD-MM-YYYY HH:MM:SS")
    Close #1
End If






List3.AddItem "Processing Once a Day Or Week Code..."


For we = 1 To M
    ScanPath.lblCount = Trim(Str(we))
    If Mid$(Mimp2(we), 1, 1) = "%" Then
        Mimp2(we) = Mid$(Mimp2(we), 2)
    End If
Next

'-------- Once a Day code
If DayGo = True Then
    For we = 1 To M
        ScanPath.lblCount = Trim(Str(we))
        If Mid$(Mimp2(we), 1, 1) = "*" Then
            Mimp2(we) = Mid$(Mimp2(we), 2)
        End If
    Next
End If

For we = 1 To M
    ScanPath.lblCount = Trim(Str(we))
    If Mid$(Mimp2(we), 1, 1) = "*" Then
        Mimp2(we) = "-"
    End If
Next

'-------- Once a Week code
If WkGo = True Then
    ScanPath.lblCount = Trim(Str(we))
    For we = 1 To M
        If Mid$(Mimp2(we), 1, 1) = "^" Then
            Mimp2(we) = Mid$(Mimp2(we), 2)
        End If
    Next
End If

For we = 1 To M
    ScanPath.lblCount = Trim(Str(we))
    If Mid$(Mimp2(we), 1, 1) = "^" Then
        Mimp2(we) = "-"
    End If
Next

List3.AddItem "Processing Once a Day Or Week Code...Done..."

End Sub

Public Sub Form_Load2()

'fuck that for now web down

If Activated = 1 Then Exit Sub

'need coz this take a daymode setting
Call SiteMapList


GoTo Skip2:

Call Sort_IfDoingOnceDayUpload



Tx$ = App.Path + "\Unfinished Job Buffer for Next Time.txt"
Set F = FS.getfile(Tx$)

If F.Size > 0 Then
List3.AddItem ""
List3.AddItem "Add Files From Unfinshed Job Last Time..."
List3.AddItem "--"
Open Tx$ For Input As #1
Do
    If Not (EOF(1)) Then
        Line Input #1, linet$
        DoEvents
        If Dir$(linet$) <> "" Then
            If M >= DimSize Then
                DimSize = DimSize + 1000
                ReDim Preserve Mimp2(DimSize)
                ReDim Preserve Detail(DimSize)
                ReDim Preserve Detail1(DimSize)
                ReDim Preserve DetaiL2(DimSize)
                'ReDim Preserve CRCChk$(DimSize)
            End If
            
            If FS.FileExists(linet$) Then
                    
                M = M + 1: Mimp2(M) = linet$
            
                List3.AddItem Mid$(linet$, 49): Hag2 = Hag2 + 1
            End If
        End If
    End If
Loop Until EOF(1)
Close #1

List3.AddItem "Add Files From Last Unfinshed Job... Done" + Str(Hag2) + " Files Added"
List3.AddItem "--"
DoEvents
End If

ListView1.ListItems.Clear
ListView1.View = lvwReport
ListView1.Sorted = False        'Use default sorting to sort the

For r1 = 1 To M
        ListView1.ListItems.Add , , Mimp2(r1)
Next

Hag2 = 0
List3.AddItem "Processing Banned Files..."

For r1 = ListView1.ListItems.Count To 1 Step -1
    ScanPath.lblCount = Trim(Str(r1)) + " of " + Trim(Str(ListView1.ListItems.Count))
    'A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(r1)
    If StringPassChk(B1$) = True Then
        ListView1.ListItems.Remove (r1): Hag2 = Hag2 + 1
    End If
Next
List3.AddItem "Processing Banned Files..." + Str(Hag2) + " Banned Removed"





Hag2 = 0
List3.AddItem "Processing Extras..."

For we = ListView1.ListItems.Count To 1 Step -1
    ScanPath.lblCount = Trim(Str(we)) + " of " + Trim(Str(ListView1.ListItems.Count))
    DoEvents
    B1$ = ListView1.ListItems.Item(we)
    nameext = LCase(Mid$(B1$, 1, InStrRev(B1$, ".") - 1))
    ext = LCase(Mid$(B1$, InStrRev(B1$, ".") + 1))
'    If InStr(nameext, "ripple") Then Stop
    milk = 0
    If ext = "txt" Then
        rr1 = FS.FileExists(nameext + ".html")
        rr2 = FS.FileExists(nameext + ".htm")
        If rr1 = True Or rr2 = True Then
            milk = 1
        End If
    End If
    If InStr(B1$, "\-") Then milk = 1
    If InStr(LCase(B1$), ".lnk") > 0 Then milk = 1

    If milk = 1 Then
        ListView1.ListItems.Remove (we): Hag2 = Hag2 + 1
    End If
Next
'ListView1.Visible = True
'ScanPath.Show
'Do
'DoEvents
'Loop Until 1 = 2

List3.AddItem "Processing Extras..." + Str(Hag2) + " Extras Removed"

'-----------------------

Call CRC_Filter



Skip2:

If TheseToDo > 0 Then
    List3.AddItem "----------------------------- Go... "

    Activated = 1

    'put back web not down
    Call Command4_Click
Else
    List3.AddItem "----------------------------- None... "
    List3.AddItem "----------------------------- End... "
    
    'Call Bling2
    Command2.Enabled = True
    Command3.Enabled = True
    CountD1 = 0: CountD2 = 5: CountD3 = 0

    Timer1.Enabled = False
    Timer2.Enabled = True
    TimerPutToWeb.Enabled = False

End If
'put in web not down
'End
'or
'CountD2 = 2

'Timer2.Enabled = True

'End


End Sub


Sub CRC_Filter()


Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32

GoTo SkipChk
List3.AddItem "Processing Date Filter..."
Hag2 = 0
hx = 0
For r1 = ListView1.ListItems.Count To 1 Step -1
    ScanPath.lblCount = Trim(Str(r1)) + " of " + Trim(Str(ListView1.ListItems.Count))

    B1$ = ListView1.ListItems.Item(r1)
    DoEvents
    If Mid$(B1$, 1, 1) <> "" Then
        FileSpec = B1$
                
        If Dir$(FileSpec) <> "" Then
        
        Set F = FS.getfile(FileSpec)
        tds = Format$(F.DateLastModified, "YYYY-MM-DD HH:MM:SS")
        milk = 0
        For R = 0 To UBound(CRCChkDate$)
            If InStr(CRCChkDate(R), tds + " " + FileSpec) > 0 And CRCChkDate(R) <> "" Then
                ListView1.ListItems.Remove (r1): Hag2 = Hag2 + 1
                milk = 1
                Exit For
            End If
        Next
        If milk = 0 Then
                For r3 = 0 To UBound(CRCChkDate$)
                    If InStr(CRCChkDate(r3), FileSpec) > 1 Then
                        CRCChkDate(r3) = tds + " " + FileSpec
                        milk = 1
                        Exit For
                    End If
                Next
                If milk = 0 Then
                        ReDim Preserve CRCChkDate(UBound(CRCChkDate) + 1)
                        CRCChkDate(UBound(CRCChkDate)) = tds + " " + FileSpec
                End If
            End If
        
        End If
        If Dir$(FileSpec) = "" Then
            ListView1.ListItems.Remove (r1): Hag2 = Hag2 + 1
        End If
    End If
Next

eo = FreeFile
Open App.Path + "\CRC'sDate.txt" For Output As #eo
    For r3 = 0 To UBound(CRCChkDate$)
        Print #eo, CRCChkDate(r3)
    Next
Close eo
List3.AddItem "Processing Date Filter...Done..." + Str(Hag2) + " Removed"

SkipChk:

List3.AddItem "Processing CRC Filter..."
Hag2 = 0
hx = 0
For r1 = ListView1.ListItems.Count To 1 Step -1
    ScanPath.lblCount = Trim(Str(r1)) + " of " + Trim(Str(ListView1.ListItems.Count))

    B1$ = ListView1.ListItems.Item(r1)
    DoEvents
    If Mid$(B1$, 1, 1) <> "" Then
        FileSpec = B1$
                
                
        If Dir$(FileSpec) <> "" Then
        wxhex9$ = "00000000"
        LSet wxhex9$ = Hex(m_CRC.CalculateFile(FileSpec))
        For R = 0 To UBound(CRCChk$)
            If InStr(CRCChk(R), wxhex9$ + " " + FileSpec) > 0 And CRCChk(R) <> "" Then
                ListView1.ListItems.Remove (r1): Hag2 = Hag2 + 1: Exit For
            End If
        Next
        End If
        If Dir$(FileSpec) = "" Then
            ListView1.ListItems.Remove (r1): Hag2 = Hag2 + 1
        End If
    End If
Next
List3.AddItem "Processing CRC Filter...Done..." + Str(Hag2) + " Removed"

'-------

Dim itmX As ListItem
List3.AddItem "Processing Remove Dupes..."
Hag2 = 0
For r1 = ListView1.ListItems.Count - 1 To 1 Step -1
DoEvents
ScanPath.lblCount = Trim(Str(r1)) + " of " + Trim(Str(ListView1.ListItems.Count))
Do
    Fee = False
    'ListView1.ListItems.Item(r1+1)
    Set itmX = ListView1.FindItem(ListView1.ListItems.Item(r1), , r1 + 1, lvwWholeWord)
        If itmX Is Nothing Then Fee = False Else Fee = True
        If Fee = True Then ListView1.ListItems.Remove (itmX.Index): Hag2 = Hag2 + 1
        If r1 >= ListView1.ListItems.Count Then Exit Do
Loop Until Fee = False
Next

List3.AddItem "Processing Remove Dupes..." + Str(Hag2) + " Dupes Removed"

'-------

ReDim Preserve Mimp2(ListView1.ListItems.Count)
For we = 1 To ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    Mimp2(we) = ListView1.ListItems.Item(we)
Next

M = ListView1.ListItems.Count
   


DoEvents

'Make it store the list of job files not done yet if quit early then do again


List3.AddItem ""
List3.AddItem "These to Do....."
'List3.ListIndex = List3.ListCount - 1
DoEvents

ScanPath.lblCount = Str(M) + " To Do"

ReDim Mimp2Buff(M)
For R = 1 To M
    A1$ = Mid(Mimp2(R), 1, InStrRev(Mimp2(R), "\"))
    B1$ = Mid(Mimp2(R), InStrRev(Mimp2(R), "\") + 1)
    Call NIceFileNames
    'Test Okay
    'If B1$ <> Mid(Mimp2(r), InStrRev(Mimp2(r), "\") + 1) Then Stop
    Mid(Mimp2(R), InStrRev(Mimp2(R), "\") + 1) = B1$
    ScanPath.lblCount = Trim(Str(R))
    If Mid$(Mimp2(R), 49) <> "" Then
    List3.AddItem Mid$(Mimp2(R), 49)
    Mimp2Buff(R) = Mimp2(R)
    TheseToDo = TheseToDo + 1
    End If
Next

List3.AddItem Trim(Str(M)) + " --- These To Do -- (¯`'•.¸ -‹(•¿•)›- ¸.•'´¯)"


End Sub


Public Sub Del_FrontPage_Junk()


ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath = "D:\MY WEBS\MatthewLan.Com Web"
ScanPath.cboMask.Text = "*.*"
ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScanDir_Click
Reset
For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
'    we = ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    A1$ = A1$ + ScanPath.ListView1.ListItems.Item(we)
    outff = 0
    If InStr(A1$, "_private") > 0 Then outff = 1
    If InStr(A1$, "_vti") > 0 Then
        outff = 1
    End If
    
    If outff = 1 Then
        'ScanPath.ListView1.ListItems.Remove (we)
        On Error Resume Next
        FS.deletefolder A1$, True
        'If Err.Number > 0 Then Stop
        'err.description
        On Error GoTo 0

    End If
Next


End Sub

Public Sub Fav_an_SiteMap()

'Load ScanPath

ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath = "D:\MY WEBS\MatthewLan.Com Web"
ScanPath.cboMask.Text = "*.*"
ScanPath.chkSubFolders = vbChecked

Call ScanPath.cmdScan_Click


List3.AddItem "Processing Fav BookMarks"

f1 = FreeFile
Open App.Path + "\" + App.EXEName + " Fav ChkSum.txt" For Input As #f1
Line Input #f1, f22$
Close #1
fg = Val(f22$)


On Error Resume Next
Err.Clear

FAV = GetSpecialfolder(6)

Set F = FS.GetFolder(FAV)
If Err.Number > 0 Then MsgBox "Error This Dir " + FAV: Stop



f33 = F.Size


If f33 = fg And GOTrue = False Then
    List3.AddItem "Processing Favs Done Was Nothing to do - Size of Folder Same"
    'If IsIDE = False Then Exit Sub
    'test
    Exit Sub
End If
attx = True






ScanPath.txtPath = FAV
ScanPath.cboMask.Text = "*.URL;*.LNK;*.TXT"
ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScan_Click



f1 = FreeFile
On Error Resume Next
Open App.Path + "\" + App.EXEName + " CRC's Favs.txt" For Input As #f1
wxhex2$ = Input(LOF(1), f1)
Close #1
On Error GoTo 0
    
'GoTo Skip
For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    Command1.Caption = Trim(Str(we))
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    Set F = FS.getfile(A1$ + G1$)
    If (F.Attributes And FILE_ATTRIBUTE_ARCHIVE) And Err.Number = 0 Then
        F.Attributes = F.Attributes - FILE_ATTRIBUTE_ARCHIVE
    End If
Next

Skip:

Books = Val(ScanPath.lblCount.Caption)
List3.AddItem ("BookMarks = " + Str$(Books))

List3.AddItem "Processing Favs Done "
If attx = False Then
    List3.AddItem "Processing Favs Done Was Nothing to do"
End If

'attx = True
If attx = True Then

ScanPath.chkNormal = vbChecked
Call ScanPath.cmdScan_Click

List3.AddItem "Processing BookMarks Date Order"
Call BookMarksDate
List3.AddItem "Processing BookMarks Date Order Done..."

List3.AddItem "Processing BookMarks Main Job"
Call BookMarks
List3.AddItem "Processing BookMarks Main Job Done..."

End If

f1 = FreeFile
Open App.Path + "\" + App.EXEName + " Fav ChkSum.txt" For Output As #f1
Print #f1, Str(f33)
Close #1

End Sub

Sub SiteMapList()



List3.AddItem "Processing SiteMap And Work File Menu & Photo's PHP..."
DoEvents




fr = FreeFile
Open App.Path + "\Sitemap Count.txt" For Input As #fr
Line Input #fr, xx$
Close #fr

ScanPath.chkNormal = vbChecked
ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath = "D:\MY WEBS\MatthewLan.Com Web"
ScanPath.cboMask.Text = "*.*"
Call ScanPath.cmdScan_Click

List3.AddItem "Files on Web = " + ScanPath.lblCount.Caption

List3.AddItem "Processing SiteMap...."
DoEvents
tryim = 0

'If Now < DateSerial(2008, 11, 14) Then tryim = 1: List3.AddItem "Process Coz Date Less..."
If DayGo = True Then tryim = 1: List3.AddItem "Process In Day Mode Go..."


If tryim = 0 Then
'    Exit Sub
Else
'    List3.AddItem "Decide to Process SiteMap Coz Vars..."
End If

'End If


For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    ScanPath.lblCount = Trim(Str(we)) + " of " + Str(ScanPath.ListView1.ListItems.Count)
    DoEvents
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    Tx$ = LCase(Mid$(A1$, InStrRev(A1$, "\", Len(A1$) - 1) + 1))
    tg$ = Mid$(Tx$, 1, 1)
    
    outff = 0

    
    'If InStr(A1$, "") > 0 Then outff = 1
    'If InStr(A1$, "") > 0 Then outff = 1
    'If InStr(A1$, "") > 0 Then outff = 1
    If InStr(A1$, "Alt.UseNet") > 0 Then outff = 1
    If Tx$ = "LoveFolder\" Then outff = 1
    If Tx$ = "galleries\" Then outff = 1
    If InStr(A1$, "LoveFolder\Secure") > 0 Then outff = 1
    If InStr(A1$, "autoindex") > 0 Then outff = 1
    If InStr(A1$, "awstats") > 0 Then outff = 1
    If InStr(A1$, "awstats2") > 0 Then outff = 1
    If InStr(A1$, "ServerTest") > 0 Then outff = 1
    If InStr(A1$, "Style_Sheets") > 0 Then outff = 1
    If InStr(A1$, "Secure_Folder") > 0 Then outff = 1
    If InStr(A1$, "Web\Photos") > 0 And InStr(A1$, "\galleries") = 0 Then outff = 1
    If InStr(A1$, "\-") > 0 Then outff = 1
    If InStr(A1$, "\_") > 0 Then outff = 1
    If InStr(A1$, "Alt.Sz_Best") > 0 Then outff = 1
    If LCase(A1$) = LCase(ScanPath.txtPath) + "\" Then outff = 1
    If LCase(A1$) = LCase(ScanPath.txtPath) + "\LoveFolder\" Then outff = 1
    If InStr(B1$, "index.") > 0 Then outff = 1
    If InStr(A1$, "Alt.Sz_Borg_Roid_Judy") > 0 Then outff = 0
    
    If InStr(B1$, ".") > 0 Then ext5$ = LCase(Mid$(B1$, InStrRev(B1$, ".", Len(B1$)))) Else outff = 1
    cat$ = ".log.php.png.index.ico.css.gif.asp.css"
    If InStr(cat$, ext5$) > 0 And outff = 0 Then
        outff = 1
    End If
    
    If outff = 0 And ext5$ <> ".jpg" And InStr(A1$, "D_Words") = 0 Then
    a = a
    End If
    If outff = 1 Then
        ScanPath.ListView1.ListItems.Remove (we)
    End If



Next

Dim FileCountDiff
FileCountDiff = False

If Val(xx$) <> ScanPath.ListView1.ListItems.Count Then
    FileCountDiff = True
    fr = FreeFile
    Open App.Path + "\Sitemap Count.txt" For Output As #fr
        Print #fr, Str(ScanPath.ListView1.ListItems.Count)
    Close #fr
End If

If FileCountDiff = False Then
    List3.AddItem "Processing -- SiteMap Nothing to Do...-- File Count was Same"
End If

'List3.AddItem "Files on Web = " + ScanPath.lblCount.Caption
List3.AddItem "Files In SiteMap =" + Str(ScanPath.ListView1.ListItems.Count)

If FileCountDiff = True Then
    fn1$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\SiteMainMap.html"
    fn2$ = "D:\MY WEBS\MatthewLan.Com Web\0work\SiteMainMapRobots.html"
Else
    fn1$ = "D:\TEMP\SiteMainMap.html"
    fn2$ = "D:\TEMP\SiteMainMapRobots.html"
End If

'fn1$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\SiteMainMap.html"
'fn2$ = "D:\MY WEBS\MatthewLan.Com Web\-0work\SiteMainMapRobots.html"
'Reset
free20 = FreeFile
Open fn1$ For Output As #free20

free21 = FreeFile
Open fn2$ For Output As #free21

Print #free20, "<html>"
Print #free20, ""
Print #free20, "<head>"
Print #free20, "<TITLE>Matthew Lan's Main SiteMap</TITLE>"
Print #free20, "<META name=""Description"" content=""Matt Lan's Main SiteMap " + Format$(Now, "DDD DD/MMM/YYYY HH:MM:SS") + """>"
Print #free20, "<META NAME=""ROBOTS"" CONTENT=""INDEX,NOFOLLOW"">"
Print #free20, ""
Print #free20, "</head>"
Print #free20, ""
Print #free20, "<body>"
Print #free20, ""
Print #free20, "<H2><font face=""Arial"">Matt Lan's Main SiteMap " + Format$(Now, "DDD DD/MMM/YYYY HH:MM:SS") + "</font></H2>"
Print #free20, "<DL><p>"
Print #free20, "<font face=""Arial"" size=""2"">"

Print #free21, "<html>"
Print #free21, ""
Print #free21, "<head>"
Print #free21, "<TITLE>Matthew Lan's Main SiteMap</TITLE>"
Print #free21, "<META name=""Description"" content=""Matt Lan's Main SiteMap for Robots " + Format$(Now, "DDD DD/MMM/YYYY HH:MM:SS") + """>"
Print #free21, "<META NAME=""ROBOTS"" CONTENT=""NOINDEX,FOLLOW"">"
Print #free21, ""
Print #free21, "</head>"
Print #free21, ""
Print #free21, "<body>"
Print #free21, ""
Print #free21, "<H2><font face=""Arial"">Matt Lan's Main SiteMap For Robots " + Format$(Now, "DDD DD/MMM/YYYY HH:MM:SS") + "</font></H2>"
Print #free21, "<DL><p>"
Print #free21, "<font face=""Arial"" size=""2"">"

Dim C1$

For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)

    C1$ = Mid$(A1$, Len(ScanPath.txtPath) + 2)
    G1$ = Mid$(A1$, Len(ScanPath.txtPath) + 2) + B1$

    If InStr(A1$, "LoveFolder\") > 0 Then
    a = a
    
    machine$ = ""
    For R = 1 To Len(B1$)
    ding = Mid$(B1$, R, 1)
    If Mid$(B1$, R, 1) = " " Then ding = "%20"
    If Mid$(B1$, R, 1) = "#" Then ding = "%23"
    
    machine$ = machine$ + ding
    
    Next
    C1$ = "LoveFolder\index.php?dir=" + Mid$(C1$, 10) + "\&file=" + machine$
    End If
    
    
    
'http://www.matthewlan.com/Photos/gallery.php?
    
    If InStr(A1$, "Photos\") > 0 Then
    If A1$ <> OldA1$ Then showi = -1
    OldA1$ = A1$
    showi = showi + 1
    a = a
    machine$ = ""
    
    tth$ = Mid$(A1$, InStr(A1$, "galleries") + 10)
    tth$ = Mid$(tth$, 1, Len(tth$) - 1)
    For R = 1 To Len(tth$)
        ding = Mid$(tth$, R, 1)
        If Mid$(tth$, R, 1) = " " Then ding = "%20"
        If Mid$(tth$, R, 1) = "#" Then ding = "%23"
        machine$ = machine$ + ding
    Next
    tth$ = machine$
    machine$ = ""
    
    For R = 1 To Len(B1$)
    ding = Mid$(B1$, R, 1)
    If Mid$(B1$, R, 1) = " " Then ding = "%20"
    If Mid$(B1$, R, 1) = "#" Then ding = "%23"
    machine$ = machine$ + ding
    Next
    
    C1$ = "Photos/gallery.php?gallery=\" + tth$ + "&showimage=" + Trim(Str(showi))
    
    If M >= DimSize Then
        DimSize = DimSize + 1000
        ReDim Preserve Mimp2(DimSize)
        ReDim Preserve Detail(DimSize)
        ReDim Preserve Detail1(DimSize)
        ReDim Preserve DetaiL2(DimSize)
        'ReDim Preserve CRCChk$(DimSize)
    End If
    
    M = M + 1: Mimp2(M) = "" + A1$ + B1$
    ScanPath.lblCount = M
    End If
    
''http://www.matthewlan.com/Photos/gallery.php?gallery=/Hoss%20Photos/01%20Caburn/10%20Oct&showimage=32
    
    

        Do
            eq = InStr(C1$, "\")
            If eq > 0 Then Mid$(C1$, eq, 1) = "/"
        Loop Until eq = 0
        Do
            eq = InStr(G1$, "\")
            If eq > 0 Then Mid$(G1$, eq, 1) = "/"
        Loop Until eq = 0


        'php
        If InStr(B1$, "SiteMainMapRobots.html") = 0 Then
            Print #free20, "<DT><A HREF=""http://matthewlan.com/" + C1$ + """>" + G1$ + "</A>"
        End If
        
        'neat
        Print #free21, "<DT><A HREF=""http://matthewlan.com/" + G1$ + """>" + G1$ + "</A>"
        

Next


Print #free20, "</font>"
Print #free20, "</DL><p>"

Print #free20, Grab22$
Print #free20, "</body>"
Print #free20, "</html>"

Print #free21, "</font>"
Print #free21, "</DL><p>"

Print #free21, Grab22$
Print #free21, "</body>"
Print #free21, "</html>"

Close #free20
Close #free21



FreeF1 = FreeFile
Open dired$ + IndexMain$ For Binary As #FreeF1
RRF$ = Input$(LOF(FreeF1), FreeF1)
Close #FreeF1
xxv = 0
Do
    xxv = InStr(xxv + 1, RRF$, "LoveFolder/index.php?dir=")
    If xxv > 0 Then
        RRF2$ = Mid$(RRF$, 1, xxv - 1)
        xxv2 = InStr(xxv + 23, RRF$, "/") - (xxv + 23)
        xxv3 = InStr(xxv + 23, RRF$, "=") - (xxv + 23)
        RRF2$ = RRF2$ + "LoveFolder/" + Mid$(RRF$, xxv + 23, xxv2) + "/" + Mid$(RRF$, xxv + 24 + xxv3)
        RRF$ = RRF2$
    End If
Loop Until xxv = 0

xxv = 0
Do
    xxv = InStr(xxv + 1, RRF$, "href=""http://matthewlan.com/")
    If xxv > 0 Then
        RRF2$ = Mid$(RRF$, 1, xxv + 5) + Mid$(RRF$, xxv + 5 + 23)
        RRF$ = RRF2$
    End If
Loop Until xxv = 0

List3.AddItem "Processing SiteMap And Work File Menu & Photo's PHP Done..."




List3.AddItem "Processing Gallery Thumb Nails..."
ScanPath.chkNormal = vbChecked
ScanPath.chkSubFolders = vbUnchecked
ScanPath.txtPath = "D:\MY WEBS\MatthewLan.Com Web\Photos\thumbnails"
ScanPath.cboMask.Text = "*.jpg"
Call ScanPath.cmdScan_Click

For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    If M >= DimSize Then
        DimSize = DimSize + 1000
        ReDim Preserve Mimp2(DimSize)
        ReDim Preserve Detail(DimSize)
        ReDim Preserve Detail1(DimSize)
        ReDim Preserve DetaiL2(DimSize)
        'ReDim Preserve CRCChk$(DimSize)
    End If

    M = M + 1: Mimp2(M) = "" + A1$ + B1$
    ScanPath.lblCount = M
Next

List3.AddItem "Processing Gallery Thumb Nails -- Done..."









DimSize = M
ScanPath.lblCount = M

End Sub


Private Sub Buster2(Filespec2$, Idate As Date)
'Call buster(filespec2$, Idate)

PDate = 0
For i = 0 To File1.ListCount - 1
    DoEvents
FileSpec = Filespec2$
Set F = FS.getfile((FileSpec))
Tdate = F.DateLastModified
If Tdate > PDate Then PDate = Tdate: Filespec2$ = FileSpec
Next

Idate = PDate

End Sub

Private Sub buster(Filespec2$, Idate As Date)
'Call buster(filespec2$, Idate)


PDate = 0
For i = 0 To File1.ListCount - 1
    DoEvents
FileSpec = File1.Path + "\" + File1.List(i)
Set F = FS.getfile((FileSpec))
Tdate = F.DateLastModified
If Tdate > PDate Then PDate = Tdate: Filespec2$ = FileSpec
Next

Idate = PDate

End Sub

Private Sub Command1_Click()

Exit Sub

Timer1.Enabled = False

Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Check1.Enabled = False

CountD1 = 3

aMain.Timer1.Enabled = False

Load WebPhoto


End Sub

Private Sub Command2_Click()

ExitPro = True
If ExitPro2 = True Then Unload Me
If Comm5 > 0 Then Unload Me

'End

 'For Each Form In Forms
'If Form.Caption <> "BBC 24 Hour Weather" Then Unload Form
'Next

End Sub

Private Sub Command3_Click()

Togg1 = Not Togg1
If Togg1 = -1 Then
Timer1.Enabled = False
'Timer5.Enabled = False
TimerPutToWeb.Enabled = False
End If

If Togg1 = 0 Then
Timer1.Enabled = True
'Timer5.Enabled = True
TimerPutToWeb.Enabled = True
End If
End Sub

Private Sub Command4_Click()
If CountD3 > 0 Then CountD3 = 0
If CountD2 > 0 Then CountD2 = 0
If CountD1 > 0 Then CountD1 = 0

End Sub
Sub QQW()

QW$ = "<HTML><HEAD>" + vbCrLf
QW$ = QW$ + "<TITLE>" + Title$ + "</TITLE>" + vbCrLf
QW$ = QW$ + "<META NAME=""DESCRIPTION"" CONTENT=""" + Title$ + " - Last Updated " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + """>" + vbCrLf
QW$ = QW$ + "<META NAME=""robots"" content=""noindex,nofollow"">" + vbCrLf
QW$ = QW$ + "</HEAD>" + vbCrLf
QW$ = QW$ + "<BODY><PRE>"
QW$ = QW$ + "<a name=""top""></a>"
QW$ = QW$ + "<font face=""Arial"">"
QW$ = QW$ + "<p align=""center""><b><font size=""3"" color=""#000080"">º¤ø,¸¸,ø¤º°`°º¤ø, </font><font size=""3"" color=""#800080""> " + Title$ + " - Last Updated </font><font size=""3"" color=""#700000"">" + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + "</font><font size=""3"" color=""#000080""> ,ø¤º°`°º¤ø,¸¸,ø¤º°</font></b></p><HR>" + vbCrLf
QW$ = QW$ + "<p align=""center""><b><font size=""3"" color=""#000080"">º¤ø,¸¸,ø¤º°`°º¤ø, </font><font size=""3"" color=""#800080""> " + nSize$ + " BookMark's ---- " + SSize$ + " are Secure </font><font size=""3"" color=""#000080""> ,ø¤º°`°º¤ø,¸¸,ø¤º°</font></b></p><HR>" + vbCrLf

End Sub

Sub QQW2()

QW$ = "<HTML><HEAD>" + vbCrLf
QW$ = QW$ + "<TITLE>" + Title$ + "</TITLE>" + vbCrLf
QW$ = QW$ + "<META NAME=""DESCRIPTION"" CONTENT=""" + Title$ + " - Last Updated " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + """>" + vbCrLf
QW$ = QW$ + "<META NAME=""robots"" content=""noindex,nofollow"">" + vbCrLf
QW$ = QW$ + "</HEAD>" + vbCrLf
'QW$ = QW$ + "<a name=""top""></a>"
'QW$ = QW$ + "<font face=""Arial"">" + vbCrLf
'QW$ = QW$ + "<p align=""center""><b><font size=""3"" color=""#000080"">º¤ø,¸¸,ø¤º°`°º¤ø, </font><font size=""3"" color=""#800080""> " + Title$ + " - Last Updated </font><font size=""3"" color=""#700000"">" + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + "</font><font size=""3"" color=""#000080""> ,ø¤º°`°º¤ø,¸¸,ø¤º°</font></b></p><HR>" + vbCrLf
'QW$ = QW$ + "<p align=""center""><b><font size=""3"" color=""#000080"">º¤ø,¸¸,ø¤º°`°º¤ø, </font><font size=""3"" color=""#800080""> " + NSize$ + " BookMark's ---- " + SSize$ + " are Secure </font><font size=""3"" color=""#000080""> ,ø¤º°`°º¤ø,¸¸,ø¤º°</font></b></p><HR>" + vbCrLf + vbCrLf + "<BODY><Pre>"
QW$ = QW$ + "<BODY><PRE>"

Grab24$ = "<br><br>" + vbCrLf
Grab24$ = Grab24$ + "<div align=""left"">" + vbCrLf
Grab24$ = Grab24$ + "  <table border=""5"" style=""font-family: Rockwell; font-size: 14pt"">" + vbCrLf
Grab24$ = Grab24$ + "      <tr>" + vbCrLf
Grab24$ = Grab24$ + "        <td width=""100%"">" + vbCrLf
Grab24$ = Grab24$ + "         <p align=""center""><font size=""4""><a href=""http://matthewlan.com"">Back to Home Page</a></font></td>" + vbCrLf
Grab24$ = Grab24$ + "      </tr>" + vbCrLf
Grab24$ = Grab24$ + "  </table>" + vbCrLf
Grab24$ = Grab24$ + "</div>" + vbCrLf
Grab24$ = Grab24$ + "</BODY></HTML>" + vbCrLf

End Sub


Private Sub BookMarks()

SecondHalf = False
AName01 = 0
RattleSnake01 = 0
RattleSnake02 = 0
Freff5$ = ""
Freff8$ = ""
Freff9$ = ""

Fref1 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\Secure_Folder\Links\LinkPage_00_Text_Only_No-URL.html" For Output As Fref1
Fref2 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_00_Text_Only_No-URL.html" For Output As Fref2


Fref3 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_01_URL-Only.html" For Output As Fref3

'URL AN TEXT
Fref4 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_02_URL_and_TEXT.html" For Output As Fref4

Fref5 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_02_URL_and_TEXT.txt" For Output As Fref5

Title$ = "Matt Lan's - Bookmarks"
nSize$ = "000"
SSize$ = "000"


Title$ = "Matt Lan's - Bookmarks - Secure"
Call QQW
Print #Fref1, QW$

Title$ = "Matt Lan's - Bookmarks - Orginal"
Call QQW
Print #Fref2, QW$


Title$ = "Matt Lan's - Bookmarks - URL Only"
Call QQW
Print #Fref3, QW$

Title$ = "Matt Lan's - Bookmarks - URL & Text"
Call QQW
Print #Fref4, QW$


qwe$ = "------------------------------------------------------------------------------------------" + vbCrLf
qwe$ = qwe$ + "Matt Lan's - Bookmarks - URL & Text -- Text Version" + vbCrLf
qwe$ = qwe$ + "Last Updated --- " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + vbCrLf
qwe$ = qwe$ + "000 BookMark's ---- 000 are Secure" + vbCrLf
qwe$ = qwe$ + "------------------------------------------------------------------------------------------" + vbCrLf

Print #Fref5, qwe$



xx = 0

eq8 = 1

'paths
For we = 1 To ScanPath.ListView1.ListItems.Count
    Command1.Caption = Trim(Str(we)) + " of " + Trim(Str(ScanPath.ListView1.ListItems.Count))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    
    Call BlackList(A1$ + B1$)
    RattleSnake01 = RattleSnake01 + 1
    If Forget = 0 Then RattleSnake02 = RattleSnake02 + 1
    
    eq1 = 0
    EQ3 = 0
    
    'eq4 = InStr(A1$, "Favorites\") + 9
    
    'IS + 4 ALWAYS RIGHT
    eq4 = InStr(A1$, Mid(FAV, 3) + "\") + 4
    
    eq2 = eq4
  
    xzag = 0
    lc = 0
    Do
        DoEvents
        eq1 = InStr(eq2, A1$, "\")
        If eq1 = 0 Then Exit Do
        eq6 = InStr(eq1 + 1, A1$, "\")
        If eq6 = 0 Then Exit Do
        If eq1 > 0 Then
            EQ3 = EQ3 + 1
        End If
  
        tw1 = eq1 + 1
        tw2 = eq6 - eq1 - 1
        
        If Path1$(EQ3) <> Mid$(A1$, tw1, tw2) And xzag = 0 Then xzag = 1: lc = EQ3
        Path1$(EQ3) = Mid$(A1$, tw1, tw2)
        Path2$(EQ3) = Mid$(A1$, 1, eq6 - 1)
        Path3$(EQ3) = A1$ + B1$
        
        If InStr(eqs8$, Path1$(EQ3)) = 0 Then
            eq8t = 1
        End If
        eq2 = eq1 + 1
    Loop Until eq1 = 0


    If lc = 0 Then lc = EQ3
    
    tig = 0
    
   
    If lc < oldlc And xzag = 1 And (EQ3 <= lc) Then
        Call HTMLOut(EQ3, Path1$(EQ3))
        tig = 1
    End If
 
    If EQ3 > lc And xzag > 0 Then
        If eq8 = 0 Then eq8 = 1
        ting = 0
        If InStr(A1$, eqs8$) And eq8 > 1 Then ting = 1
        
        For RT = lc To EQ3
            Call HTMLOut(RT, Path1$(RT))
            tig = 1
        Next
    End If
    
    If EQ3 = lc And xzag > 0 And tig = 0 Then
        Call HTMLOut(EQ3, Path1$(EQ3))
    End If
        
    eq8 = EQ3
    eqs8$ = B1$
    oldlc = lc
Next

'----------------------
'----------------------
'----------------------
'----------------------
'----------------------
'----------------------
'----------------------
'----------------------

Freff8$ = Freff8$ + vbCrLf
Freff9$ = Freff9$ + vbCrLf


SecondHalf = True
RattleSnake01 = 0
RattleSnake02 = 0
AName01 = 0

qwe$ = "<HR>" + vbCrLf
qwe$ = qwe$ + "<p align=""center""><b><font size=""3"" color=""#000080"">º¤ø,¸¸,ø¤º°`°º¤ø, </font><font size=""3"" color=""#800080""> Matt Lan's - Bookmarks - Last Updated </font><font size=""3"" color=""#700000"">" + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + "</font><font size=""3"" color=""#000080""> ,ø¤º°`°º¤ø,¸¸,ø¤º°</font></b></p><HR>" + vbCrLf
qwe$ = qwe$ + "<p align=""center""><b><font size=""3"" color=""#000080"">º¤ø,¸¸,ø¤º°`°º¤ø, </font><font size=""3"" color=""#800080""> 000 BookMark's ---- 000 are Secure </font><font size=""3"" color=""#000080""> ,ø¤º°`°º¤ø,¸¸,ø¤º°</font></b></p><HR>"

Print #Fref1, qwe$
Print #Fref2, qwe$
Print #Fref3, qwe$
Print #Fref4, qwe$

'qwe$ = vbCrLf + "------------------------------------------------------------------------------------------" + vbCrLf
'qwe$ = qwe$ + "º¤ø,¸¸,ø¤º°`°º¤ø, Matt Lan's - Bookmarks - Last Updated " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + " ,ø¤º°`°º¤ø,¸¸,ø¤º°" + vbCrLf
'qwe$ = qwe$ + "º¤ø,¸¸,ø¤º°`°º¤ø, 000 BookMark's ---- 000 are Secure  ,ø¤º°`°º¤ø,¸¸,ø¤º°" + vbCrLf
'qwe$ = qwe$ + "------------------------------------------------------------------------------------------"

qwe$ = vbCrLf + "------------------------------------------------------------------------------------------" + vbCrLf
qwe$ = qwe$ + "Matt Lan's - Bookmarks - URL & TEXT Version" + vbCrLf
qwe$ = qwe$ + "Last Updated --- " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + vbCrLf
qwe$ = qwe$ + "000 BookMark's ---- 000 are Secure" + vbCrLf
qwe$ = qwe$ + "------------------------------------------------------------------------------------------" + vbCrLf
qwe$ = qwe$ + vbCrLf + "------------------------------------------------------------------------------------------"

Print #Fref5, qwe$


NewsGo = False





For we = 1 To ScanPath.ListView1.ListItems.Count
'    ScanPath.lblCount = Trim(Str(we))
    Command1.Caption = Trim(Str(we)) + " of " + Trim(Str(ScanPath.ListView1.ListItems.Count))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    
    C1$ = Mid$(B1$, Len(ScanPath.txtPath) + 2) + B1$
    
    Call BlackList(A1$ + B1$)
    RattleSnake01 = RattleSnake01 + 1
    If Forget = 0 Then RattleSnake02 = RattleSnake02 + 1
    
    eq1 = 0
    EQ3 = 0
    
    eq4 = InStr(A1$, Mid(FAV, 3) + "\") + 4

    eq2 = eq4
  
    xzag = 0
    lc = 0
    Do
        DoEvents
        eq1 = InStr(eq2, A1$, "\")
        If eq1 = 0 Then Exit Do
        eq6 = InStr(eq1 + 1, A1$, "\")
        If eq6 = 0 Then Exit Do
        If eq1 > 0 Then
            EQ3 = EQ3 + 1
        End If
  
        tw1 = eq1 + 1
        tw2 = eq6 - eq1 - 1
        
        If Path1$(EQ3) <> Mid$(A1$, tw1, tw2) And xzag = 0 Then xzag = 1: lc = EQ3
        Path1$(EQ3) = Mid$(A1$, tw1, tw2)
        Path2$(EQ3) = Mid$(A1$, 1, eq6 - 1)
        Path3$(EQ3) = A1$ + B1$
   
        If InStr(eqs8$, Path1$(EQ3)) = 0 Then
            eq8t = 1
        End If
        eq2 = eq1 + 1
    Loop Until eq1 = 0
    
    If lc = 0 Then lc = EQ3
    
    tig = 0
    
   
    If lc < oldlc And xzag = 1 And (EQ3 <= lc) Then
        Call HTMLOut(EQ3, Path1$(EQ3))
        tig = 1
    End If
 
    If EQ3 > lc And xzag > 0 Then
        If eq8 = 0 Then eq8 = 1
        ting = 0
        If InStr(A1$, eqs8$) And eq8 > 1 Then ting = 1
        
        For RT = lc To EQ3
            Call HTMLOut(RT, Path1$(RT))
            tig = 1
        Next
    End If
    
    If lc <> EQ3 And xzag = 1 And lc = oldlc And tig = 0 Then
        Call HTMLOut(EQ3, Path1$(EQ3))
    End If
    
    If EQ3 = lc And xzag > 0 And tig = 0 Then
        Call HTMLOut(EQ3, Path1$(EQ3))
    End If
        
    a2$ = Mid$(B1$, 1, InStrRev(B1$, ".") - 1)
    If EQ3 > 0 Then kily$ = Space$((EQ3 - 1) * 4) Else kily$ = " "

    
    
    
    If InStr(UCase(G1$), ".URL") > 0 Then
    
    FREFG = FreeFile
    Open A1$ + G1$ For Input As #FREFG
    On Local Error GoTo 0
    Line Input #FREFG, lin2$
    Line Input #FREFG, lin2$
    noturl = 0
    If InStr(lin2$, "BASEURL=") = 1 Then
        noturl = 1
        GOLDY$ = Mid$(lin2$, 9)
    End If
    If InStr(lin2$, "URL=") = 1 Then
        noturl = 1
        GOLDY$ = Mid$(lin2$, 5)
    End If
    If noturl = 0 Then
        Style = vbYesNo + vbCritical + vbDefaultButton2 + vbMsgBoxSetForeground
        Response = MsgBox("NOT URL" + vbCrLf + Format$(Now, "DD-MMM HH:MM:SS") + vbCrLf + "Description " + Err.Description + vbCrLf + "Err Number" + Str(Err.Number) + vbCrLf + "Errr Source " + Err.Source + vbCrLf + "STOP YES/NO", Style)
        If Response = vbYes Then Stop
    End If
    Close #FREFG
        
    End If
    
    If InStr(UCase(G1$), ".LNK") > 0 Then
    
    FREFG = FreeFile
    Open A1$ + G1$ For Binary As #FREFG
        RR$ = Space$(LOF(FREFG))
        Get #FREFG, 1, RR$
    Close #FREFG
    
    noturl = 0
    If Mid$(RR$, &H568, 3) = ":" + Chr$(0) + "\" Then
        noturl = 1
        RR$ = Mid$(RR$, &H566)
        TT$ = Mid$(RR$, InStr(105, RR$, ":\") - 1)
        D1$ = Mid$(TT$, 1, InStr(TT$, Chr$(0)) - 1)
        GOLDY$ = D1$
    End If
    
    If noturl = 0 Then
        
        OL = InStrRev(RR$, ":\")
        If OL > 0 Then
            If Mid$(RR$, OL - 2, 1) = Chr(0) Then
                GOLDY$ = Mid$(RR$, OL - 1)
                GOLDY$ = Mid$(GOLDY$, 1, InStr(GOLDY$, Chr$(0)) - 1)
                noturl = 1
            End If
            
            If Mid$(RR$, OL + 2, 1) = Chr(0) Then
                xOL = OL
                OL = InStr(OL, RR$, "\\")
                OL = InStr(OL, RR$, Chr(0))
                GOLDY$ = GOLDY$ + Mid$(RR$, OL + 1)
                GOLDY$ = Mid$(GOLDY$, 1, InStr(GOLDY$, Chr$(0)) - 1)
                noturl = 1
            End If
        
        End If
    End If
        
    'If NOTURL = 0 Then MsgBox ("NOT a Proper Link")
    
    
    
    End If
        
        
    If InStr(UCase(G1$), ".TXT") > 0 Then
    
        GOLDY$ = A1$ + B1$
        noturl = 1
       
    End If
        
        
        
    If noturl <> 0 Then
    
    QW1$ = "<font face=""Arial"" Size=""3"">" + kily$ + "<A HREF=""" + GOLDY$ + """>" + a2$ + "</font></A>"
    QW2$ = "<font face=""Arial"" Size=""3"">" + kily$ + "<A HREF=""" + GOLDY$ + """>" + GOLDY$ + "</font></A>"
    QW3$ = "<font face=""Arial"" Size=""3"">" + "<A HREF=""" + GOLDY$ + """>" + a2$ + "</font></A>"
    QW4$ = "<font face=""Arial"" Size=""3"">" + "<A HREF=""" + GOLDY$ + """>" + GOLDY$ + "</font></A>"
    
    QW3b$ = a2$
    QW4b$ = GOLDY$
    'News & Palm
    QW5 = a2$ + vbCrLf + GOLDY$
    qw5x$ = QW3$ + vbCrLf + QW4$ + vbCrLf + vbCrLf
    qw5xx$ = QW3b$ + vbCrLf + QW4b$ + vbCrLf + vbCrLf
     
    eq1 = Len(A1$)
    Tfn2 = 0
    Do
        eq2 = InStrRev(A1$, "\", eq1 - 1)
        If eq2 >= 11 Then
            Tfn2 = Tfn2 + 1
        End If
        eq1 = eq2
    Loop Until eq2 <= 11
    If EQ3 <> Tfn2 Then Stop
    
    If Tfn2 = 2 Then QX$ = "<font COLOR=""#0000B0"" face=""Arial"" Size=""3"">"
    If Tfn2 <= 1 Then QX$ = "<font COLOR=""#000000"" face=""Arial"" Size=""3"">"

    'AName01 = AName01 + 1
    
    'Only Secure and Link
    dweeb1$ = QX$ + "<b>" + Format$(RattleSnake01 + 1, "0000 ") + " --  </font></b>"
    dweeb2$ = QX$ + "<b>" + Format$(RattleSnake02 + 1, "0000 ") + " --  </font></b>"


    'RattleSnake01 = RattleSnake01 + 1
    Print #Fref1, dweeb1$ + QW1$
    
    If Forget = 0 Then
        'RattleSnake02 = RattleSnake02 + 1
        
        'This bit alright
        Print #Fref2, dweeb2$ + QW1$
        
        'url & text 1
        Print #Fref3, dweeb2$ + QW2$
        'Url & text 2
        Print #Fref4, QW3$
        Print #Fref4, QW4$
        Print #Fref4, "<BR>"
        
        Print #Fref5, QW3b$
        Print #Fref5, QW4b$
        Print #Fref5,
        
        
        If NewsGo = True Then
            Freff5$ = Freff5$ + qw5x$
        End If
        If PalmPage = True Then
            Freff8$ = Freff8$ + qw5x$
            Freff9$ = Freff9$ + qw5xx$
        End If
    End If
    
    eq8 = EQ3
    eqs8$ = B1$
    oldlc = lcn

End If '' IF VALID LINK OR URL

Next



Print #Fref1, "</pre>"
Print #Fref2, "</pre>"
Print #Fref3, "</pre>"
Print #Fref4, "</pre>"

QW1$ = Grab22$

Print #Fref1, QW1$
Print #Fref2, QW1$
Print #Fref3, QW1$
Print #Fref4, QW1$

Close #Fref1
Close #Fref2
Close #Fref3
Close #Fref4
Close #Fref5

On Local Error Resume Next
'On Local Error GoTo 0
Fref1 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\Secure_Folder\Links\LinkPage_00_Text_Only_No-URL.html" For Input As Fref1
Fref2 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_00_Text_Only_No-URL.html" For Input As Fref2
Fref3 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_01_URL-Only.html" For Input As Fref3
Fref4 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_02_URL_and_TEXT.html" For Input As Fref4
Fref5 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Date.html" For Input As Fref5

Fref6 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_02_URL_and_TEXT.txt" For Input As Fref6

Fref7 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Date.txt" For Input As Fref7


dweebs1$ = Input$(LOF(Fref1), Fref1)
dweebs2$ = Input$(LOF(Fref2), Fref2)
dweebs3$ = Input$(LOF(Fref3), Fref3)
dweebs4$ = Input$(LOF(Fref4), Fref4)
dweebs5$ = Input$(LOF(Fref5), Fref5)
dweebs6$ = Input$(LOF(Fref6), Fref6)
dweebs7$ = Input$(LOF(Fref7), Fref7)
Close #Fref1, #Fref2, #Fref3, #Fref4, #Fref5, #Fref6, #Fref7


Fref1 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\Secure_Folder\Links\LinkPage_00_Text_Only_No-URL.html" For Output As Fref1
Fref2 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_00_Text_Only_No-URL.html" For Output As Fref2
Fref3 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_01_URL-Only.html" For Output As Fref3
Fref4 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_02_URL_and_TEXT.html" For Output As Fref4

Fref5 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Date.html" For Output As Fref5

Fref6 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_02_URL_and_TEXT.txt" For Output As Fref6

Fref7 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Date.txt" For Output As Fref7


For R = 1 To 2
'are secure
wjk1 = InStr(dweebs1$, "000 are Sec")
wjk2 = InStr(dweebs2$, "000 are Sec")
wjk3 = InStr(dweebs3$, "000 are Sec")
wjk4 = InStr(dweebs4$, "000 are Sec")
wjk5 = InStr(dweebs5$, "000 are Sec")
wjk6 = InStr(dweebs6$, "000 are Sec")
wjk7 = InStr(dweebs7$, "000 are Sec")

ss$ = Trim(Str(RattleSnake01 - RattleSnake02))
If wjk1 > 0 Then dweebs1$ = Mid$(dweebs1$, 1, wjk1 - 1) + ss + Mid$(dweebs1$, wjk1 + 3)
If wjk2 > 0 Then dweebs2$ = Mid$(dweebs2$, 1, wjk2 - 1) + ss + Mid$(dweebs2$, wjk2 + 3)
If wjk3 > 0 Then dweebs3$ = Mid$(dweebs3$, 1, wjk3 - 1) + ss + Mid$(dweebs3$, wjk3 + 3)
If wjk4 > 0 Then dweebs4$ = Mid$(dweebs4$, 1, wjk4 - 1) + ss + Mid$(dweebs4$, wjk4 + 3)
If wjk5 > 0 Then dweebs5$ = Mid$(dweebs5$, 1, wjk5 - 1) + ss + Mid$(dweebs5$, wjk5 + 3)
If wjk6 > 0 Then dweebs6$ = Mid$(dweebs6$, 1, wjk6 - 1) + ss + Mid$(dweebs6$, wjk6 + 3)
If wjk7 > 0 Then dweebs7$ = Mid$(dweebs7$, 1, wjk7 - 1) + ss + Mid$(dweebs7$, wjk7 + 3)
Next

For R = 1 To 2
wjk1 = InStr(dweebs1$, "000 BookMark's")
wjk2 = InStr(dweebs2$, "000 BookMark's")
wjk3 = InStr(dweebs3$, "000 BookMark's")
wjk4 = InStr(dweebs4$, "000 BookMark's")
wjk5 = InStr(dweebs5$, "000 BookMark's")
wjk6 = InStr(dweebs6$, "000 BookMark's")
wjk7 = InStr(dweebs7$, "000 BookMark's")


ss$ = Trim(Str(RattleSnake01))

If wjk1 > 0 Then dweebs1$ = Mid$(dweebs1$, 1, wjk1 - 1) + ss + Mid$(dweebs1$, wjk1 + 3)
If wjk2 > 0 Then dweebs2$ = Mid$(dweebs2$, 1, wjk2 - 1) + ss + Mid$(dweebs2$, wjk2 + 3)
If wjk3 > 0 Then dweebs3$ = Mid$(dweebs3$, 1, wjk3 - 1) + ss + Mid$(dweebs3$, wjk3 + 3)
If wjk4 > 0 Then dweebs4$ = Mid$(dweebs4$, 1, wjk4 - 1) + ss + Mid$(dweebs4$, wjk4 + 3)
If wjk5 > 0 Then dweebs5$ = Mid$(dweebs5$, 1, wjk5 - 1) + ss + Mid$(dweebs5$, wjk5 + 3)
If wjk6 > 0 Then dweebs6$ = Mid$(dweebs6$, 1, wjk6 - 1) + ss + Mid$(dweebs6$, wjk6 + 3)
If wjk7 > 0 Then dweebs7$ = Mid$(dweebs7$, 1, wjk7 - 1) + ss + Mid$(dweebs7$, wjk7 + 3)
Next

Print #Fref1, dweebs1$;
Print #Fref2, dweebs2$;
Print #Fref3, dweebs3$;
Print #Fref4, dweebs4$;
Print #Fref5, dweebs5$;
Print #Fref6, dweebs6$;
Print #Fref7, dweebs7$;

Close #Fref1, #Fref2, #Fref3, #Fref4, #Fref5, #Fref6, #Fref7
    
    
'Fref5 = FreeFile
'Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Text3_News.html" For Output As Fref5
'Print #Fref5, Freff5$
'Close #Fref5
Title$ = "Matt Lan's - Bookmarks - Palm Page"
Call QQW2

Fref5 = FreeFile
File11 = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Secure\LinkPage_Palm.html"
Open File11 For Output As Fref5
Print #Fref5, QW$ + Freff8$ + Grab24$
Close #Fref5
   
Fref5 = FreeFile
File11 = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Palm.html"
Open File11 For Output As Fref5
Print #Fref5, QW$ + Freff8$ + Grab24$
Close #Fref5
   

Fref6 = FreeFile
File11 = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Palm.txt"

Title$ = "Matt Lan's - Bookmarks - Palm Page" + vbCrLf + vbCrLf
Open File11 For Output As Fref6
Print #Fref6, Title$ + Freff9$
Close #Fref6
   


On Error GoTo 0
   


'FS.copyfile File11, "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Palm.html"
   

End Sub



Private Sub BookMarksDate()

SecondHalf = True

Freff9$ = ""
QW$ = "<HTML><HEAD>" + vbCrLf
QW$ = QW$ + "<TITLE>Matt Lan's - Bookmarks - Date Ordered</TITLE>" + vbCrLf
QW$ = QW$ + "<META NAME=""DESCRIPTION"" CONTENT=""Matt Lan's - Bookmarks - Date Oderded - Last Updated " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + """>" + vbCrLf
QW$ = QW$ + "<META NAME=""robots"" content=""index,nofollow"">" + vbCrLf

QW$ = QW$ + "</HEAD>" + vbCrLf
QW$ = QW$ + "<BODY><PRE>"
QW$ = QW$ + "<a name=""top""></a>"
QW$ = QW$ + "<font face=""Arial"">"
QW$ = QW$ + "<p align=""center""><b><font size=""3"" color=""#000080"">º¤ø,¸¸,ø¤º°`°º¤ø, </font><font size=""3"" color=""#800080""> Matt Lan's - Bookmark's - Date Ordered - Last Updated </font><font size=""3"" color=""#700000"">" + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + "</font><font size=""3"" color=""#000080""> ,ø¤º°`°º¤ø,¸¸,ø¤º°</font></b></p><HR>" + vbCrLf
QW$ = QW$ + "<p align=""center""><b><font size=""3"" color=""#000080"">º¤ø,¸¸,ø¤º°`°º¤ø, </font><font size=""3"" color=""#800080""> 000 BookMark's - 000 are Secure</font><font size=""3"" color=""#000080""> ,ø¤º°`°º¤ø,¸¸,ø¤º°</font></b></p><HR>" + vbCrLf

'QW5 = "Matt Lan's Bookmarks Date Ordered" + vb2 seCrLf
HR3 = "------------------------------------------------------------------------------------------" + vbCrLf

QW5 = HR3 + "Matt Lan's - Bookmarks - Date Oderded" + vbCrLf
QW5 = QW5 + "Last Updated " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + vbCrLf
'QW5 = QW5 + "º¤ø,¸¸,ø¤º°`°º¤ø, 000 BookMark's - 000 are Secure ,ø¤º°`°º¤ø,¸¸,ø¤º°" + vbCrLf + HR3 + vbCrLf
QW5 = QW5 + "000 BookMark's - 000 are Secure" + vbCrLf + HR3 + vbCrLf

Fref1 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Date.txt" For Output As Fref1
Print #Fref1, QW5;
Close #Fref1



Fref1 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Date.html" For Output As Fref1
Print #Fref1, QW$;
Close #Fref1




BookDates = True

AName01 = 0
RattleSnake01 = 0
RattleSnake02 = 0
Freff5$ = ""
Freff8$ = ""
Freff9$ = ""

AName01 = 1

xx = 0
eq8 = 1



ListView1.ListItems.Clear

With ListView1
    .ColumnHeaders.Add , "-List 1-", "-List 1-", 13000, lvwColumnLeft
    .View = lvwReport
End With
'Hack = True


ListView1.SortOrder = lvwAscending
ListView1.SortKey = 0
ListView1.Sorted = False        'Use default sorting to sort the
    

For we = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount = Trim(Str(we))
    DoEvents
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    
    Set F = FS.getfile(A1$ + G1$)
    
    G2$ = Space$(13)
    Mid$(G2$, 1) = G1$
    tds = Format$(F.DateLastModified, "YYYY-MM-DD ") ' + G2$ + A1$ + B1$
    'ListView1.ListItems.Add , , tds
    With ListView1
        Set lv = .ListItems.Add(, , tds)
        lv.SubItems(1) = A1$
        lv.SubItems(2) = B1$
        lv.SubItems(3) = G2$
    End With
    'List4.AddItem Format$(F.datelastmodified, "YYYY-MM-DD ") + G2$ + A1$ + B1$
Next



ListView1.SortOrder = lvwAscending
ListView1.SortKey = 2
ListView1.Sorted = True
ListView1.Sorted = False
ListView1.SortKey = 1
ListView1.Sorted = True
ListView1.Sorted = False

ListView1.SortOrder = lvwDescending
ListView1.SortKey = 0
ListView1.Sorted = True
ListView1.Sorted = False

    
DoEvents

'SecondHalf = True
RattleSnake01 = 0
RattleSnake02 = 0
AName01 = 0


'----------------------
'----------------------
'----------------------
'----------------------
'----------------------
'----------------------
'----------------------
'----------------------

Fref1 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Date.html" For Append As Fref1

Fref2 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Date.txt" For Append As Fref2


NewsGo = False
QW4$ = ""

For we = 1 To ListView1.ListItems.Count ' To 1 Step -1
    DoEvents
    Command1.Caption = Trim(Str(we)) + " of " + Trim(Str(ListView1.ListItems.Count))
    
    ts = InStrRev(ListView1.ListItems.Item(we), "\")
    'A1$ = Mid$(ListView1.ListItems.Item(we), 12 + 13, ts - 11 - 13)
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    'B1$ = ScanPath.ListView1.ListItems.Item(we)

    'B1$ = Mid$(ListView1.ListItems.Item(we), ts + 1)
    B1$ = ListView1.ListItems.Item(we).SubItems(2)
    
    'G1$ = Mid$(ListView1.ListItems.Item(we), 12, 12)
    G1$ = ListView1.ListItems.Item(we).SubItems(3)
    Call BlackList(A1$ + B1$)
    
  
    If OldA1$ <> A1$ Then
        eq1 = Len(A1$)
        Tfn2 = 0
        Do
            eq2 = InStrRev(A1$, "\", eq1 - 1)
            If eq2 >= 11 Then
                pathl = Mid$(A1$, eq2 + 1, eq1 - eq2 - 1)
                Tfn2 = Tfn2 + 1
                Path1$(Tfn2) = pathl
            End If
            eq1 = eq2
        Loop Until eq2 <= 11
    
        For i = Tfn2 To 1 Step -1
            kily$ = Space$((Tfn2 - i) * 4)
            If Tfn2 - i > 0 Then QX$ = "<font COLOR=""#0000B0"" face=""Arial"" Size=""3"">": HR = "": HR3 = ""
            
            If Tfn2 - i = 0 Then
                QX$ = "<font COLOR=""#000000"" face=""Arial"" Size=""3"">": HR = "<HR>"
                HR3 = "------------------------------------------------------------------------------------------" + vbCrLf
            End If
            
            QW1$ = HR + kily$ + "<b>" + QX$ + Path1$(i) + "</b></font>"
            
            QW5 = HR3 + kily$ + Path1$(i)
            
            Print #Fref1, QW1$ + vbCrLf;
            Print #Fref2, QW5 + vbCrLf;
        Next
        If Forget = 0 Then
            Print #Fref1, vbCrLf;
            Print #Fref2, vbCrLf;
        End If
        kily$ = Space$((Tfn2) * 4)
    
    End If
    
    OldA1$ = A1$
    
'    If InStr(UCase(G1$), ".LNK") > 0 Then Stop
   
    
    If InStr(UCase(G1$), ".URL") > 0 Then
    
    FREFG = FreeFile
    Open A1$ + G1$ For Input As #FREFG
    On Local Error GoTo 0
    Line Input #FREFG, lin2$
    Line Input #FREFG, lin2$
    noturl = 0
    
    If InStr(lin2$, "BASEURL=") = 1 Then
        noturl = 1
        GOLDY$ = Mid$(lin2$, 9)
    End If
    If InStr(lin2$, "URL=") = 1 Then
        noturl = 1
        GOLDY$ = Mid$(lin2$, 5)
    End If
    
    IconFile = False
    If InStr(lin2$, "IconFile=") = 1 Then
        
        IconFile = True
        Style = vbYesNo + vbCritical + vbDefaultButton2 + vbMsgBoxSetForeground
        Response = MsgBox("Is a IconFile - NOT URL" + vbCrLf + A1$ + vbCrLf + B1$ + vbCrLf + Mid$(lin2$, 10) + vbCrLf + Format$(Now, "DD-MMM HH:MM:SS") + vbCrLf + "Description " + Err.Description + vbCrLf + "Err Number" + Str(Err.Number) + vbCrLf + "Errr Source " + Err.Source + vbCrLf + "STOP YES/NO", Style)
        Debug.Print A1 + vbCrLf + B1 + vbCrLf + Mid$(lin2$, 10)
        If Response = vbYes Then
        Stop
        Close #FREFG
        Shell "explorer /e, /select," + A1$ + G1$, vbNormalFocus
        End If
    
    End If
    
    If noturl = 0 And IconFile = False Then
        Style = vbYesNo + vbCritical + vbDefaultButton2 + vbMsgBoxSetForeground
        Response = MsgBox("NOT URL" + vbCrLf + Format$(Now, "DD-MMM HH:MM:SS") + vbCrLf + "Description " + Err.Description + vbCrLf + "Err Number" + Str(Err.Number) + vbCrLf + "Errr Source " + Err.Source + vbCrLf + "STOP YES/NO", Style)
        If Response = vbYes Then Stop
    End If
    
    Close #FREFG
        
    End If
    
    If InStr(UCase(G1$), ".LNK") > 0 Then
    
    FREFG = FreeFile
    Open A1$ + G1$ For Binary As #FREFG
        RR$ = Space$(LOF(FREFG))
        Get #FREFG, 1, RR$
    Close #FREFG
    
    noturl = 0
    
    If Mid$(RR$, &H568, 3) = ":" + Chr$(0) + "\" Then
        noturl = 1
        RR$ = Mid$(RR$, &H566)
        TT$ = Mid$(RR$, InStr(105, RR$, ":\") - 1)
        D1$ = Mid$(TT$, 1, InStr(TT$, Chr$(0)) - 1)
        GOLDY$ = D1$
    End If
    
    If noturl = 0 Then
        
        OL = InStrRev(RR$, ":\")
        If OL > 0 Then
'            Shell "EXPLORER /SELECT " + A1$ '+ G1$
            
            If Mid$(RR$, OL - 2, 1) = Chr(0) Then
                GOLDY$ = Mid$(RR$, OL - 1)
                GOLDY$ = Mid$(GOLDY$, 1, InStr(GOLDY$, Chr$(0)) - 1)
                noturl = 1
            End If
            
            If Mid$(RR$, OL + 2, 1) = Chr(0) Then
                xOL = OL
                xOL = InStr(OL, RR$, "\\")
                xOL = InStr(OL, RR$, Chr(0))
                GOLDY$ = GOLDY$ + Mid$(RR$, xOL + 1)
                GOLDY$ = Mid$(GOLDY$, 1, InStr(GOLDY$, Chr$(0)) - 1)
                noturl = 1
            End If
        
            If FS.FileExists(A1$ + G1$) = False Then
                Response = MsgBox("The Link Short Cut Here Dont Exist Where Point To" + vbCrLf + A1$ + vbCrLf + B1$ + vbCrLf + "Stop YES/NO", Style)
                If Response = vbYes Then Stop
            End If
        
        
        End If
    End If
    'If NOTURL = 0 Then MsgBox ("NOT a Proper Link")
    End If
        
        
        
        
    If InStr(UCase(G1$), ".TXT") > 0 Then
        GOLDY$ = A1$ + B1$
    End If
        
    If noturl <> 0 Then
        
    If InStr(GOLDY$, ":""") > 0 Then
        RR = 0
        Do
            RR = InStr(RR + 1, GOLDY$, """")
            If RR > 0 Then
                GOLDY$ = Mid$(GOLDY, 1, RR - 1) + "&quot;" + Mid$(GOLDY, RR + 1)
            End If
        Loop Until RR = 0
    End If
       
    a2$ = Mid$(B1$, 1, InStrRev(B1$, ".") - 1)
    
    QW3$ = "<font face=""Arial"" Size=""3"">" + "<A HREF=""" + GOLDY$ + """>" + a2$ + "</font></A>"
    QW4$ = "<font face=""Arial"" Size=""3"">" + "<A HREF=""" + GOLDY$ + """>" + GOLDY$ + "</font></A>"
    QW3b$ = a2$
    QW4b$ = GOLDY$
    
    Set F = FS.getfile(A1$ + G1$)
    qw51$ = Format$(F.DateLastModified, "DD-MM-YYYY  ") + QW3$ + vbCrLf
    qw52$ = Format$(F.DateLastModified, "DD-MM-YYYY  ") + QW4$ + vbCrLf
    
    qw53$ = Format$(F.DateLastModified, "DD-MM-YYYY  ") + QW3b$ + vbCrLf
    qw54$ = Format$(F.DateLastModified, "DD-MM-YYYY  ") + QW4b$ + vbCrLf
 
    If Forget = 0 Then
        Print #Fref1, qw51$; qw52$
        Print #Fref2, qw53$; qw54$
    End If
    
    End If


    'DO THE DUMP LOGG SAME TIME
    If InStr(UCase(G1$), ".URL") > 0 And Forget = 0 Then
        QW4$ = QW4$ + "<A HREF=""" + GOLDY$ + """>" + GOLDY$ + vbCrLf
    End If

Next

Print #Fref1, Grab22$;

On Local Error Resume Next
    
Close #Fref1, #Fref2

BookDates = False
    
    
'Call DumpWinHTTrack
    
Frefx = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\Link_Dump_For_WinHTTrack.html" For Output As Frefx
    Print #Frefx, QW$ + vbCrLf + vbCrLf + QW4$ + vbcrf + "</BODY></HTML>"
Close #Frefx

    
    
End Sub


Sub DumpWinHTTrack()
Exit Sub

QW4$ = ""

For we = 1 To ListView1.ListItems.Count ' To 1 Step -1
    DoEvents
    Command1.Caption = Trim(Str(we)) + " of " + Trim(Str(ListView1.ListItems.Count))
    
    ts = InStrRev(ListView1.ListItems.Item(we), "\")
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we).SubItems(2)
    G1$ = ListView1.ListItems.Item(we).SubItems(3)
    
    Call BlackList(A1$ + B1$)
    
  
    If OldA1$ <> A1$ Then
        eq1 = Len(A1$)
        Tfn2 = 0
        Do
            eq2 = InStrRev(A1$, "\", eq1 - 1)
            If eq2 >= 11 Then
                pathl = Mid$(A1$, eq2 + 1, eq1 - eq2 - 1)
                Tfn2 = Tfn2 + 1
                Path1$(Tfn2) = pathl
            End If
            eq1 = eq2
        Loop Until eq2 <= 11
    
    End If
    
    OldA1$ = A1$
    
    
    If InStr(UCase(G1$), ".URL") > 0 Then
    FREFG = FreeFile
    Open A1$ + G1$ For Input As #FREFG
    On Local Error GoTo 0
    Line Input #FREFG, lin2$
    Line Input #FREFG, lin2$
    noturl = 0
    
    If InStr(lin2$, "BASEURL=") = 1 Then
        noturl = 1
        GOLDY$ = Mid$(lin2$, 9)
    End If
    If InStr(lin2$, "URL=") = 1 Then
        noturl = 1
        GOLDY$ = Mid$(lin2$, 5)
    End If
    
    If InStr(lin2$, "IconFile=") = 1 Then noturl = 0
    
    Close #FREFG
        
    If InStr(UCase(G1$), ".TXT") > 0 Then noturl = 0
    
    'GOLDY$ = A1$ + B1$
    QW4$ = QW4$ + "<A HREF=""" + GOLDY$ + """>" + GOLDY$ + vbCrLf
    
    End If
        
    
Next


Frefx = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\Link_Dump_For_WinHTTrack.html" For Output As Frefx
    Print #Frefx, QW$ + vbCrLf + vbCrLf + QW4$ + vbcrf + "</BODY></HTML>"
Close #Frefx



End Sub


Sub BlackList(PathT)

Forget = 0
    PathT2 = UCase(PathT)
    If PathT2 = UCase("\#03 Matts\A1 Smart Links\b2 Google\Google Search\Argus Newspaper\People\") Then Forget = 1
    If PathT2 = UCase("\#03 Matts\A1 Smart Links\b2 Google\News Groups\SZ\") Then Forget = 1
    If PathT2 = UCase("\#03 Matts\A1 Smart Links\Banks\") Then Forget = 1
    If PathT2 = UCase("\#03 Matts\A1 Smart Links\My Web & Providers\") Then Forget = 1
    'If PathT2 = UCase("\#03 Matts\A1 Smart Links\South Downs Health NHS PDF's\") Then Forget = 1
    If PathT2 = UCase("\#03 Matts\A2 My Webs & Providers\") Then Forget = 1
    If PathT2 = UCase("\#03 Matts\A2 Root Favs\Dec 2005 Roots\") Then Forget = 1
    'If PathT2 = UCase("\#03 Matts\A2 Shops\b1 Electronic shops\Olympus\") Then Forget = 1
    If PathT2 = UCase("\#03 Matts\A3 Banks\") Then Forget = 1
    If PathT2 = UCase("\#03 Matts\A3 BT, BT Yahoo & The Webs\My Web\") Then Forget = 1
    If InStr(PathT2, UCase("\#03 Matts\A3 Ebay Paypal\Products\")) Then Forget = 1

    If InStr(PathT2, UCase("Secure Links")) Then Forget = 1
    
    If InStr(PathT2, "VALERIE") Then Forget = 1
    If InStr(PathT2, "VALERIE") Then Forget = 1
    If InStr(PathT2, "SPAMTASIC") Then Forget = 1
    If InStr(PathT2, "VOIDMAN") Then Forget = 1
    If InStr(PathT2, "LANCASTER") Then Forget = 1
    If InStr(PathT2, UCase$("Celsius for Brighton")) Then Forget = 1
    If InStr(PathT2, UCase$("Celsius for Hove")) Then Forget = 1
    If InStr(PathT2, UCase$("Shoreham Airport")) Then Forget = 1
    If InStr(PathT2, UCase$("#00 Local Machine")) Then Forget = 1
    If InStr(PathT2, UCase$("#00 Agents Today")) Then Forget = 1
    If InStr(PathT2, UCase$("#0 Estate Agents")) Then Forget = 1
    If InStr(PathT2, UCase$("#-MICHELLExx")) Then Forget = 1
    If InStr(PathT2, UCase$("EDGWARE")) Then Forget = 1
    
    
    Iss$ = "Banks"
    If InStr(PathT2, UCase$(Iss$)) Then Forget = 1
    Iss$ = "A3 Friends Reunited"
    If InStr(PathT2, UCase$(Iss$)) Then Forget = 1
    
    
    
End Sub





Sub HTMLOut(P, dweeb$)
            
    'secure links not display
    '--------------------
    If SecondHalf = True Then Call BlackList(Path3$(P))
    If SecondHalf = False Then Call BlackList(Path2$(P))
    If InStr(Path3$(P), "Crime Bosses Face Phone Ban - SERIOUS Criminals Could Soon Be Banned From Travelling And From Using Mobiles") Then
        'Stop
    End If
    'If InStr(dweeb$, "#00 Palm") Then
    If dweeb$ = "#00 Palm 01" Then
        PalmPage = True
    End If
            
    If dweeb$ = "#00 Palm 02" Then
        PalmPage = False
    End If
    
    If InStr(dweeb$, "#00 Secure Links") Then
        PalmPage = False
    End If
            
    If InStr(dweeb$, "#01 Root Archive") Then
        PalmPage = False
    End If
            
    If InStr(dweeb$, "ZZ News") Then
        'NewsGo = True
    End If
            
    If P > 0 Then kily$ = Space$((P - 1) * 4) Else kily$ = " "
    
    If P >= 2 Then QX$ = "<font COLOR=""#0000B0"" face=""Arial"" Size=""3"">"
    If P <= 1 Then QX$ = "<font COLOR=""#000000"" face=""Arial"" Size=""3"">"
    
    AName01 = AName01 + 1
    
    If Forget = 1 Then
        'Only Secure
        dweeb1$ = QX$ + "<b>" + Format$(RattleSnake01, "0000 ") + " --  </font></b>"
        dweeb1b$ = Format$(RattleSnake01, "0000 ") + " -- "
    Else
        'Let All
        dweeb1$ = QX$ + "<b>" + Format$(RattleSnake02, "0000 ") + " --  </font></b>"
        dweeb1b$ = Format$(RattleSnake02, "0000") + " -- "
    End If
    
    rr1 = "RR1": rr2 = "RR2"
    If SecondHalf = True Then
        rr1 = "RR2": rr2 = "RR1": HR = "<HR>": BR = vbCrLf
        HR3 = "------------------------------------------------------------------------------------------" + vbCrLf
    Else
        HR = "": BR = "": HR3 = ""
    End If
    
    If AName01 = 0 Then HR = "": HR3 = ""
    
    RS$ = "<a name=""" + rr1 + "-" + Format$(AName01, "000") + """></a>"
    
    QY = dweeb1$ + kily$ + RS$ + "<a href=""#" + rr2 + "-" + Format$(AName01, "000") + """><b>" + QX$ + dweeb$ + "</a></b></font>"
    QYb = dweeb1b$ + kily$ + dweeb$
    QW1$ = QY + BR
    qw1b$ = QYb + BR
    
    If P > 1 Then
        hr2 = ""
        HR4 = ""
    Else
        hr2 = HR
        HR4 = "------------------------------------------------------------------------------------------" + vbCrLf
    End If
    
    If P = 1 Then QW1$ = HR + QW1$
    
    QW2$ = hr2 + QY
    QYb = dweeb1b$ + kily$ + dweeb$
    QW2b$ = HR4 + QYb
    
    
    QW3$ = hr2 + RS$ + kily$ + "<a href=""#" + rr2 + "-" + Format$(AName01, "000") + """><b>" + QX$ + dweeb$ + "</a></b></font>"
    
    QW3b$ = HR4 + kily$ + dweeb$
    
    ArrayT(P) = QW2$
    ArrayP(P) = QW3$
    ArrayS(P) = QW3b$
    ArrayV(P) = QW2b$
    
    If P = 0 Then Exit Sub
    If P > 0 And Path2$(P) <> "" And SecondHalf = True Then
        
        Set F = FS.GetFolder(Path2$(P))
        Set fc = F.Files
        If fc.Count = 0 Then
            Exit Sub
        End If
        nogo = 0
        For Each f1 In fc
            If InStr(LCase(f1.Name), ".url") > 0 Then nogo = 1
        Next
        If nogo = 0 Then Exit Sub
    End If
    
    If P >= 0 Then
        If SecondHalf = True Then
                For R = 1 To P - 1
                    If Forget = 1 Then Print #Fref1, ArrayT(R)
                    
                    If Forget = 0 Then
                        Print #Fref1, ArrayT(R)
                        Print #Fref2, ArrayT(R)
                        Print #Fref3, ArrayT(R)
                        Print #Fref4, ArrayT(R)
                        Print #Fref5, ArrayV(R)
                    
                        If PalmPage = True And (R < 3 Or SecondHalf = True) Then
                            Freff8$ = Freff8$ + ArrayP(R) + vbCrLf
                            Freff9$ = Freff9$ + ArrayS(R) + vbCrLf
                        End If
                    End If
                Next
        End If
    
        If Forget = 1 Then Print #Fref1, QW1$
        
        If Forget = 0 Then
            Print #Fref1, QW1$
            Print #Fref2, QW1$
            Print #Fref3, QW1$
            Print #Fref4, QW1$
            Print #Fref5, qw1b$
                
            If PalmPage = True Then
                If SecondHalf = True Or P < 3 Then Freff8$ = Freff8$ + QW3$ + vbCrLf
                If SecondHalf = True Then Freff8$ = Freff8$ + vbCrLf
                
                If SecondHalf = True Or P < 3 Then Freff9$ = Freff9$ + QW3b$ + vbCrLf
                If SecondHalf = True Then Freff9$ = Freff9$ + vbCrLf
            End If
            
            'Bad Code
            et = DateSerial(Year(Now), Month(Now), 1)
            If NewsGo = True And Format$(et, "YYYY-MM") = Mid$(QW5, 1, 7) Then
               
                et2 = DateValue(QW5)
                hht$ = "Begin- --:-- " + Format$(Now, "DDDD MMMM YYYY -- YYYY-MM-DD HH:MM:SSa")
                hhe$ = "Finish --:-- " + Format$(et2, "DDDD MMMM YYYY -- YYYY-MM-DD")
               
                hhts$ = String$(Len(hht$), "-")
                hhes$ = String$(Len(hhe$), "-")
                If BookDates = False And NewsGo = True Then
                    Freff5$ = QW1$ + Freff5$ + hhts$ + "<br>" + vbCrLf
                    Freff5$ = Freff5$ + hht$ + "<br>" + vbCrLf
                    Freff5$ = Freff5$ + hhts$ + "<br>" + vbCrLf
               
                    Freff5$ = Freff5$ + hhes$ + "<br>" + vbCrLf
                    Freff5$ = Freff5$ + hhe$ + "<br>" + vbCrLf
                    Freff5$ = Freff5$ + hhes$ + "<br></font>" + vbCrLf
                End If
                Freff5$ = Freff5$ + QW2$ + "<br>" + vbCrLf
            End If
            
            
        End If
    End If
        
    'dweeb$ = dweeb3$

End Sub



Sub SetClearAttribbits(FileSpec)
    Set F = FS.getfile((FileSpec))
    
    
    Adate(counter) = F.DateLastModified
    Ddate = F.DateLastModified
    
    'If counter = 6 Then Adate(counter) = Idate

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
QQ = UBound(Mimp2Buff)
If Err.Number = 0 Then
    Open App.Path + "\Unfinished Job Buffer for Next Time.txt" For Output As #1
    For R = 1 To UBound(Mimp2Buff)
        If Mimp2Buff(R) <> "" Then
            Print #1, Mimp2Buff(R)
        End If
    Next
    Close #1
End If

Call Bling2

If StartTime <> 0 Then
Open App.Path + "\StartFinishTime.txt" For Output As #1
Print #1, "*** Last Start - Time *** = " + Format(StartTime, "DDD DD-MM-YYYY HH:MM:SSa/p")
Print #1, "*** Last Finish Time *** = " + Format(Now, "DDD DD-MM-YYYY HH:MM:SSa/p")
Close #1
End If

For Each Form In Forms
    Unload Form
Next

End
End Sub


Private Sub cmdSEND_Click()
   Inet1.Execute txtURL.Text, _
   "SEND C:\MyDocuments\Send.txt SentDocs\Sent.txt"
End Sub

Private Sub GoCrazyCheck_Click()
If Activate2 = True Then
    GoCrazyCheck = vbUnchecked
    Exit Sub
End If

If GoCrazyCheck.Value = vbUnchecked Then Unload Me: Exit Sub

List3.AddItem "Go Awsome Pleasure Crazy Mode Selected..."

Timer1.Enabled = False
Command4.Caption = "Awe Mode"
Call Form_Activate2

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

If State = icResponseCompleted Then XJag = 1
    
'Inet1.StillExecuting

If AKO <> Inet1.ResponseInfo And Trim(Inet1.ResponseInfo) <> "" Then
    Open App.Path + "\FTP Logg.txt" For Append As #1
        Print #1, Format(Now, "DD-MM-YYYY HH:MM:SS ") + " - " + Format(Inet1.ResponseCode, "00000") + " - " + Inet1.ResponseInfo
    Close #1
    uploadfile = (Inet1.ResponseCode = 0)
    UploadFilet1$ = Inet1.ResponseInfo

End If
AKO = Inet1.ResponseInfo

End Sub
    
Public Sub Execute_FTP_Command(ByVal sCommand As String, ErrStat)
    
    
    If sCommand <> "" Then
    'List3.ListIndex = List3.ListCount - 1
    End If
    
    'If sCommand = "" Then Stop
    
    If sCommand <> OldsCommand And sCommand <> "" Then
    Happy = 0
    If Trim(sCommand) <> "" Then
        Open App.Path + "\FTP Logg.txt" For Append As #1
            Print #1, Format(Now, "DD-MM-YYYY HH:MM:SS ") + " - " + sCommand
        Close #1
        End If
    End If
    If sCommand <> "" Then OldsCommand = sCommand
    If Asked = False Then
    ButterCake1 = Now + TimeSerial(1, 30, 0)
    ButterCake2 = Now
    
    'Happy = 0
    
    Asked = True
    
    On Local Error Resume Next
    If sCommand <> "" And DryRunMode = False Then
        Inet1.Execute , sCommand
    End If
    
    
    ErrStat = False
    If Inet1.ResponseInfo = "421 No Transfer Timeout (300 seconds): closing control connection." Or _
        UploadFilet1$ = "421 No Transfer Timeout (300 seconds): closing control connection." Then
'        List3.AddItem "421 No Transfer Timeout (300 seconds): closing control connection."
'        List3.ListIndex = List3.ListCount - 1
'        List3.BackColor = QBColor(12)
'        List3.ForeColor = QBColor(15)
        Call Bling2 'Quit slow
        ErrStat = True
    End If
    If InStr(Err.Description, "Unable to connect to remote host") Then
        'AMain.WindowState = 0
        List3.AddItem "Unable to connect to remote host xxxxx"
        'List3.ListIndex = List3.ListCount - 1
        List3.BackColor = QBColor(12)
        List3.ForeColor = QBColor(15)
        'MsgBox "Probable Not connected To Net"
        'End
        Call Bling2 'Quit slow
    
    Else
        'AMain.WindowState = 0
    
    End If
    
    If Err.Number > 0 Then ErrStat = True
    
    On Local Error GoTo 0
        
    If ErrStat = True Then Exit Sub
    
    End If
'    If Asked = False Then
    
    
    QW$ = Mid$(List3.List(List3.ListCount - 1), 1, 2)
    If QW$ <> "DE" And QW$ <> "PU" And QW$ <> "CD" And QW$ <> "MK" Then Exit Sub
    If List3.ListCount < 3 Then Exit Sub
    
    
    If InStr(List3.List(List3.ListCount - 1), "---->") > 0 Then
        Text1$ = Mid$(List3.List(List3.ListCount - 1), 1, InStr(List3.List(List3.ListCount - 1), "---->") - 5)
    Else
        Text1$ = List3.List(List3.ListCount - 1)
    End If
    
    'If List3.ListIndex <> List3.ListCount - 1 Then
    '    List3.ListIndex = List3.ListCount - 1
    '    List3.Refresh
    'End If
    
    If Now - ButterCake2 >= 0 Then
        Hag$ = Text1$ + "    ----> " + Mid$(Format$(Now - ButterCake2, "HH:MM:SS"), 4)
        Happy = Now - ButterCake2
    Else
        Hag$ = Text1$ + "    ----> " + Mid$(Format$(Happy, "HH:MM:SS"), 4)
    End If
    
    
    kl = 0
    If InStr(List3.List(List3.ListCount - 1), "-- Done") > 0 Then kl = 1
    If InStr(List3.List(List3.ListCount - 1), "-- Fail") > 0 Then kl = 1
    If InStr(List3.List(List3.ListCount - 1), "---->") > 0 Then kl = 1
    If kl = 0 Then
        List3.List(List3.ListCount - 1) = List3.List(List3.ListCount - 1) + "    ----> " + Mid$(Format$(0, "HH:MM:SS"), 4)
    End If
    
    
    If InStr(List3.List(List3.ListCount - 1), Mid$(Format$(Now - ButterCake2, "HH:MM:SS"), 4)) = 0 Then
    If List3.List(List3.ListCount - 1) <> Hag$ Then
        If OldHag$ <> Hag$ Then
            'List3.ListIndex = List3.ListCount - 1
            'DoEvents
            List3.List(List3.ListCount - 1) = Hag$
            'DoEvents
        End If
        OldHag$ = Hag$
    End If
    End If
    If ButterCake1 < Now Then
        If List3.BackColor <> QBColor(12) Then
        List3.AddItem "Time-out Error xxxxx"
        'List3.ListIndex = List3.ListCount - 1
        List3.BackColor = QBColor(12)
        End If
        'Timer2.Enabled = False
        ButterCake1 = -1
    End If
    
'Exit Sub
    
    
    
    
    
    
If (Inet1.StillExecuting = True) Then Exit Sub
    
    
    
    
    kl = 0
    If InStr(List3.List(List3.ListCount - 1), "-- Done") > 0 Then kl = 1
    If InStr(List3.List(List3.ListCount - 1), "-- Fail") > 0 Then kl = 1
    
    If kl = 1 Then Exit Sub
    
    If uploadfile = True Or DryRunMode = True Then
        plij$ = " -- Done ->>"
        If OldList3 <> List3.List(List3.ListCount - 1) Then
            List3.List(List3.ListCount - 1) = List3.List(List3.ListCount - 1) + " -- " + plij$
        End If
        OldList3 = List3.List(List3.ListCount - 1)
        
        Call FindAndWriteCRCChkList
        
        Retry = 0

    Else
        plij$ = " -- Failed -XX"
        List3.List(List3.ListCount - 1) = List3.List(List3.ListCount - 1) + " -- " + plij$
        Retry = Retry + 1
    End If
    

End Sub

Sub FindAndWriteCRCChkList()

    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32
    WXHex1$ = "00000000"
    
    If PuTT1 = "" Then Exit Sub
    
    LSet WXHex1$ = Hex(m_CRC.CalculateFile(PuTT1))

    If PuTT1 <> "" And PuTT1 <> "-" And WXHex1$ <> "" Then
        tabb = 0
        For R = 0 To UBound(CRCChk$)
'            If GoCrazyCheck.Value = vbUnchecked Then ScanPath.lblCount = Trim(Str(r))
            If InStr(CRCChk(R), PuTT1) > 0 And CRCChk(R) <> "" Then
                 tabb = 1
                CRCChk(R) = WXHex1$ + " " + PuTT1
            End If
        Next
        
        If tabb = 0 Then
            ReDim Preserve CRCChk$(UBound(CRCChk$) + 1)
            CRCChk(R) = WXHex1$ + " " + PuTT1
        End If
        
        'Exit Sub
        
        Call FileInUseDelay(App.Path + "\CRC's.txt")
        'fileinuse needed
        WriteFlag = False
        eo = FreeFile
        For R = 0 To UBound(CRCChk$)
            If CRCChk(R) <> "" Then WriteFlag = True: Exit For
        Next
        If WriteFlag = True Then
            Open App.Path + "\CRC's.txt" For Output As #eo
            For R = 0 To UBound(CRCChk$)
                Print #eo, CRCChk(R)
            Next
            Close #eo
        End If

    End If
End Sub
        

Private Sub ConnectToFTP()
    
        frmFTP.Show
    
    Exit Sub
    
    Inet1.AccessType = icUseDefault
    Inet1.Document = ""
    Inet1.Password = " "
    Inet1.Protocol = icFTP
    Inet1.RemoteHost = "matthewlan.com"
    Inet1.RequestTimeout = 300
    Inet1.UserName = "matthewlan"
    Inet1.URL = "ftp://matthewlan:interned243@matthewlan.com/httpdocs"
    
    List3.AddItem "Connecting to FTP"
 
  
    List3.AddItem "CD --- /httpdocs/ - To Connect"
    List3.ListIndex = List3.ListCount - 1
    List3.Refresh
    
    Call Execute_FTP_Command("CD /httpdocs/", ErrStat)
    
    
    
    Dim PathFile$
    Dim PathFileT1$
    Dim PathFile2$
    
    
    Dim Put1$
    Activate3 = True
    
End Sub


Private Sub List3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then TTNow = Now + TimeSerial(0, 0, 4)
End Sub

Private Sub Mnu_DayRun_Click()
DayGo = True
End Sub

Private Sub Mnu_DOFAVS_Click()
GOTrue = True
End Sub

Private Sub Mnu_Open_CRC_Click()
    Shell "notepad " + App.Path + "\CRC's.txt", vbNormalFocus
End Sub

Private Sub Mnu_Run_Vb_Click()
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" ""D:\VB6\VB-NT\00_Best_VB_01\WebDates_Ace\WebDates.vbp""", vbNormalFocus
End
End Sub

Private Sub Mnu_WinAmp_Click()
direc1$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\html\Winamp_Generated_PlayList.html"
'call Mnu_WinAmp_Click
Shell "Explorer " + direc1$
End Sub

Private Sub Timer1_Timer()

If Activate3 = True Then Exit Sub
If CountD3 = -5 Then CountD3 = 15 ': If IsIDE = True Then CountD3 = 0

If Command$ = "Awesome" Then GoCrazyCheck = vbChecked

If CountD3 >= 0 Then
Command4.Caption = "Go In" + Str$(CountD3)
CountD3 = CountD3 - 1
If CountD3 >= 0 Then Exit Sub
End If


Command4.Caption = "Go In" + Str$(CountD3 + 1)

If CountD1 = -5 And Activate1 = False Then
    StartTime = Now
    Activate1 = True
    'Main Hub Of Code Startup
    Call Form_Activate2
    '"2nd Hub Do CRC Check
    Call Form_Load2
End If

If TheseToDo = 0 Then Exit Sub

If Activate2 = True And CountD1 = -5 Then CountD1 = 10

If CountD1 > 0 Then
    CountD1 = CountD1 - 1
    Command4.Caption = "Go In" + Str$(CountD1)
End If

If CountD1 <= 0 Then
    'After Load All What To Do then Run Internet upLoad
    Command4.Caption = "Go.."
    Timer1.Enabled = False
    Call ConnectToFTP
    TimerPutToWeb.Enabled = True
End If

End Sub

Private Sub Timer2_Timer()
'---------------------------------------------------------------
'--------------------This ends The Program
'---------------------------------------------------------------
If CountD2 = -5 Then CountD2 = 10
Command4.Caption = "End In" + Str$(CountD2)
Command4.BackColor = Command2.BackColor
'Command5.BackColor = Command2.BackColor
CountD2 = CountD2 - 1
If CountD2 <= 0 Then
Unload aMain
End If
End Sub


Private Sub Timer3_Timer()
If TTNow > Now Then Exit Sub
If List3.ListCount <> OldList3Count Then
    'If List3.ListIndex > List3.ListCount - 30 Then List3.ListIndex = List3.ListCount - 1
    List3.ListIndex = List3.ListCount - 1
    DoEvents
    'List3.Refresh
    'DoEvents
    
    fr = FreeFile
    Open App.Path + "\WEB DATES Logg.txt" For Append As #fr
        For R = OldList3Count To List3.ListCount - 1
            Print #fr, Format(Now, "DD-MM-YYYY HH:MM:SS ") + " - " + List3.List(R)
        Next
    Close #fr


End If
OldList3Count = List3.ListCount
'List3.ListIndex = List3.ListCount - 1
End Sub



Sub TimerPutToWeb_timer()

    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32

    'quick exit on close prog
    If ExitPro = True Then
        Unload Me: Exit Sub
    End If
        
    CountHtmlPage2 = CountHtmlPage2 + 1
    
    If CountHtmlPage2 > UBound(Mimp2) Then
        TheseToDo = 0: Call Bling2: ExitPro2 = True: Exit Sub ' Delay exit
    End If
    
    PuTT1 = Mimp2(CountHtmlPage2)
    
    ScanPath.lblCount = Trim(Str(CountHtmlPage2)) + " of " + Trim(Str(UBound(Mimp2)))
    
    
    Call StripFTPPath
    Call StripFTPPathTrob
    
    'This is where is was
    If CommandCnt = 1 Then
            'QQ$ = "MKDIR "
            QQ$ = "CD "
            List3.AddItem QQ$ + "--- " + PuTT2
            Call Execute_FTP_Command(QQ$ + PuTT2, ErrStat)
            'LastCommand = QQ$ + PuTT2
            'ReTryTTi = TTi
    End If
    
    
    
    
Exit Sub
    
    
    If UploadFilet1$ = "" Then Exit Sub
        
    lc = Mid$(LastCommand, 1, 3)
    
    mkdir2 = False
    chdir2 = True
    kg = 0
    lg = 0
    pd = 0
    
    If InStr(UploadFilet1$, "Unable to build data connection: Connection refused") > 0 Then
       CommandCnt = 1
        kg = 1
        If GoCrazyCheck.Value = vbUnchecked Then
            CountHtmlPage2 = CountHtmlPage2 + 1
           Else
                AweSomeTimer = True
                AweSomeModeSubTimer = False
                Exit Sub
        End If
    End If
    If InStr(UploadFilet1$, "Permission denied") > 0 Then
       CommandCnt = 1
                CountHtmlPage2 = CountHtmlPage2 + 1
        kg = 1
        If GoCrazyCheck.Value = vbUnchecked Then
                CountHtmlPage2 = CountHtmlPage2 + 1
           Else
                AweSomeTimer = True
                AweSomeModeSubTimer = False
                Exit Sub
        End If
    End If
    If InStr(UploadFilet1$, "No such file or directory") > 0 And QQ$ = "DE " Then
        mkdir2 = False
               CommandCnt = CommandCnt + 1
                If CommandCnt = 6 Then
                    CountHtmlPage2 = CountHtmlPage2 + 1
                End If
        kg = 1
    End If
    If InStr(UploadFilet1$, "No such file or directory") > 0 And QQ$ = "CD " Then
        mkdir2 = True
       CommandCnt = 1
        kg = 1
    End If
    If InStr(UploadFilet1$, "File exists") > 0 And QQ$ = "CD " Then
        CountHtmlPage2 = CountHtmlPage2 + 1
        mkdir2 = False
       CommandCnt = 1
        kg = 1
    End If

    If InStr(UploadFilet1$, "No such file or directory") > 0 And QQ$ = "MKDIR " Then
        
        'got to back track through the paths
        'seems cowboy wait for error then action
        'if problem run in this routine
        If PathBack = 0 And xgo = 0 Then PathBack = Paths: xgo = 1
        '/httpdocs/awstats/icon/BROWSER/
        If xgo = 1 Then StripFTPPathTrob
        
        If PathBack = 0 And xgo = 1 Then
            xgo = 0
           CommandCnt = 1
            kg = 1
            If GoCrazyCheck.Value = vbUnchecked Then
                    CountHtmlPage2 = CountHtmlPage2 + 1
            Else
                    AweSomeTimer = True
                    AweSomeModeSubTimer = False
                    Exit Sub
            End If
        End If
    End If
    
    'If UploadFilet1$ = "" And CountHtmlPage2 > 1 Then CountHtmlPage2 = CountHtmlPage2 + 1:CommandCnt = 1
    
    If InStr(UploadFilet1$, "The operation completed successfully.") > 0 Then
        PathBack = 0
        kg = 1: lg = 1
        
        If LastCommand <> "" Then
        'flg = 0
        If InStr("PUTDELCD ", lc) > 0 Then
        'If Mid$(LastCommand, 1, 3) = "PUT" Or Mid$(LastCommand, 1, 3) = "DEL" Then
           
            'If Mid$(LastCommand, 1, 3) = "DEL" ThenCommandCnt = 1
            If lc = "PUT" And MaxNotFLags = 3 Then
                'CountHtmlPage2 = CountHtmlPage2 + 1
            Else
            End If
               CommandCnt = CommandCnt + 1
            
            
            If GoCrazyCheck.Value = vbChecked Then
                AweSomeTimer = True
                AweSomeModeSubTimer = False
                Exit Sub
            Else
            
            End If
        Else
            'If LastCommand <> "Start" ThenCommandCnt =CommandCnt + 1
            'If LastCommand = "Start" Then CountHtmlPage2 = CountHtmlPage2 + 1
        End If
        End If
    End If
    
    If CommandCnt = MaxNotFLags Then CommandCnt = 1: CountHtmlPage2 = CountHtmlPage2 + 1
    
    If InStr(UploadFilet1$, "421 No Transfer Timeout (300 seconds): closing control connection.") > 0 Then kg = 1
    If InStr(UploadFilet1$, "Connection aborted") > 0 Then kg = 1
    
    If kg = 0 And UploadFilet1$ <> "" Then
        
        Style = vbYesNo + vbCritical + vbDefaultButton2 + vbMsgBoxSetForeground
        Response = MsgBox("WebDates Unknown result " + vbCrLf + UploadFilet1$ + vbCrLf + Format$(Now, "DD-MMM HH:MM:SS") + vbCrLf + "Description " + Err.Description + vbCrLf + "Err Number" + Str(Err.Number) + vbCrLf + "Errr Source " + Err.Source + vbCrLf + vbCrLf + "STOP YES/NO", Style)
        If Response = vbYes Then Stop

    
    End If
    
    
    If InStr(UploadFilet1$, "Connection aborted") > 0 Then
        List3.AddItem "Connection aborted"
        'List3.ListIndex = List3.ListCount - 1
        TheseToDo = 0
        AweSomeTimer.Enabled = False
        AweSomeModeSubTimer = False
        Call Bling2: ExitPro2 = True: Exit Sub ' Delay exit
    End If
    
    If InStr(UploadFilet1$, "421 No Transfer Timeout (300 seconds): closing control connection.") > 0 Then
        List3.AddItem "421 No Transfer Timeout (300 seconds): closing control connection."
        'List3.ListIndex = List3.ListCount - 1
        TheseToDo = 0
        AweSomeTimer.Enabled = False
        AweSomeModeSubTimer = False
        Call Bling2: ExitPro2 = True: Exit Sub ' Delay exit
    End If

    
    
    
    If CountHtmlPage2 > UBound(Mimp2) Then
        TheseToDo = 0: Call Bling2: ExitPro2 = True: Exit Sub ' Delay exit
    End If
            
    PuTT1 = Mimp2(CountHtmlPage2)
    If (PuTT1 = "" Or PuTT1 = "-") And mkdir2 = False Then
        'CountHtmlPage2 = CountHtmlPage2 + 1: Exit Sub
    End If
    
        
    If GoCrazyCheck.Value = vbUnchecked Then
        ScanPath.lblCount = Trim(Str(CountHtmlPage2)) + " of " + Trim(Str(UBound(Mimp2)))
    End If
    
    If xog = 0 Then Call StripFTPPath
    
    WXHex1$ = "00000000"
    'PuTT1 = Mimp2(CountHtmlPage2)
    LSet WXHex1$ = Hex(m_CRC.CalculateFile(PuTT1))
    
    DetaiL3 = False
    For R = 0 To UBound(CRCChk$)
        If CRCChk$(R) = WXHex1$ + " " + PuTT1 And CRCChk(R) <> "" Then
            DetaiL3 = True
        End If
    Next
    
    If xgo = True Then
            PuTT1 = ""
            WXHex1$ = ""
            QQ$ = "MKDIR "
            List3.AddItem QQ$ + "--- " + PuTT2
            Call Execute_FTP_Command(QQ$ + PuTT2, ErrStat)
            LastCommand = QQ$ + PuTT2
            ReTryTTi = TTi
            Exit Sub
    End If
    
    If GoCrazyCheck.Value = vbChecked And DetaiL3 = True Then
       CommandCnt = 1
        AweSomeTimer = True
        AweSomeModeSubTimer = False
        Exit Sub
    End If
    
    
    If DetaiL3 = True And CommandCnt < 3 Then Exit Sub
    
    'If InStr(UploadFilet1$, "No such file or directory") > 0 And QQ$ = "CD " ThenCommandCnt = NotCommandCnt

    
    'NotFlag = 1
    If CommandCnt = 1 Then
        If DetaiL3 = False Then
            PuTT1 = ""
            WXHex1$ = ""
            If mkdir2 = False Then
            QQ$ = "CD "
            Else
            QQ$ = "MKDIR "
            End If
            List3.AddItem QQ$ + "--- " + PuTT2
            Call Execute_FTP_Command(QQ$ + PuTT2, ErrStat)
            LastCommand = QQ$ + PuTT2
            ReTryTTi = TTi
        End If
    End If
    
    'Used to be true not flag
    If CommandCnt = 2 Then
        If DetaiL3 = False Then
            QQ$ = ""
            PuTT1 = Mimp2(CountHtmlPage2)
            tagsh$ = Mid$(Mimp2(CountHtmlPage2), InStrRev(Mimp2(CountHtmlPage2), "\") + 1)
            List3.AddItem "PUT --- " + tagsh$
            CommadFTP = "PUT """ + PuTT1 + """ """ + Putt3 + """"
            Call Execute_FTP_Command(CommadFTP, ErrStat)
            LastCommand = CommadFTP
        End If
    End If

    'Have a Extra Command When Want Why Not
    If CommandCnt = 3 Then
            PuTT1 = Mimp2(CountHtmlPage2)
            
                QQ$ = "DE "
        
                TagSH1 = Mid$(Mimp2(CountHtmlPage2), InStrRev(Mimp2(CountHtmlPage2), "\") + 1)
                
                'TagSH2 = Mid$(TagSH1, 1, InStrRev(TagSH1, "."))
                
                'step1 lcase
                TagSH3 = LCase(TagSH1)
                
                'step 2 HTM
                'TagSH2 = Mid$(TagSH2, 1, InStrRev(TagSH3, ".")) + "htm"
        
                'step3 Real Old Name
                'for next stage
                If InStr(EGU, TagSH3) > 0 Then
                    TagSH4 = Mid(EGU, InStr(EGU, TagSH3), Len(TagSH3) + 1)
                End If
                                
                List3.AddItem "DELETE --- " + TagSH3
                CommadFTP = "DELETE """ + TagSH3 + """"
                Call Execute_FTP_Command(CommadFTP, ErrStat)
                LastCommand = CommadFTP
    End If

    'Have a Extra Command When Want Why Not
    If CommandCnt = 4 Then
            PuTT1 = Mimp2(CountHtmlPage2)
            
                QQ$ = "DE "
        
                TagSH1 = Mid$(Mimp2(CountHtmlPage2), InStrRev(Mimp2(CountHtmlPage2), "\") + 1)
                
                TagSH2 = Mid$(TagSH1, 1, Len(TagSH1) - 1)
                
                'step1 lcase
                TagSH3 = LCase(TagSH2)
                
                'step 2 HTM
                TagSH2 = Mid$(TagSH2, 1, InStrRev(TagSH2, ".")) + "htm"
        
                'step3 Real Old Name
                'for next stage
                If InStr(EGU, TagSH3) > 0 Then
                    TagSH4 = Mid(EGU, InStr(EGU, TagSH3), Len(TagSH3) + 1)
                End If
                If TagSH1 <> TagSH2 And InStr(TagSH1, ".htm") Then
                    List3.AddItem "DELETE --- " + TagSH2
                    CommadFTP = "DELETE """ + TagSH2 + """"
                    Call Execute_FTP_Command(CommadFTP, ErrStat)
                    LastCommand = CommadFTP
                Else
                List3.AddItem "DELETE --- " + TagSH2 + " ----X No Need"

                End If
    End If


    'Have a Extra Command When Want Why Not
    If CommandCnt = 5 Then
            PuTT1 = Mimp2(CountHtmlPage2)
            
                QQ$ = "DE "
        
                TagSH1 = Mid$(Mimp2(CountHtmlPage2), InStrRev(Mimp2(CountHtmlPage2), "\") + 1)
                
                TagSH2 = Mid$(TagSH1, 1, Len(TagSH1) - 1)
                
                'step1 lcase
                TagSH3 = LCase(TagSH2)
                
                'step 2 HTM
                TagSH2 = Mid$(TagSH2, 1, InStrRev(TagSH2, ".")) + "htm"
        
                'step3 Real Old Name
                'for next stage
                'If InStr(TagSH1, ".htm") Then
                If InStr(EGU, TagSH3) > 0 Then
                    TagSH4 = Mid(EGU, InStr(EGU, TagSH3), Len(TagSH3) + 1)
                End If
                                
                List3.AddItem "DELETE --- " + TagSH4
                CommadFTP = "DELETE """ + TagSH4 + """"
                Call Execute_FTP_Command(CommadFTP, ErrStat)
                LastCommand = CommadFTP
                'Else
                'List3.AddItem "DELETE --- " + TagSH4 + " ----> No Need"
                
                'End If
        
    End If




    'NotFlag =CommandCnt + 1
    If CommandCnt = MaxNotFLags Then CommandCnt = 1



Exit Sub
    

'Call Execute_FTP_Command("", ErrStat)
    
'If ErrStat = True Then End
    
'If (Inet1.StillExecuting = True) Then Exit Sub

'If ProccesGo = 2 Then
'    Mimp2Buff(CountHtmlPage2) = ""
'    CountHtmlPage2 = CountHtmlPage2 + 1
'    ProccesGo = 0
'    PuTT1 = ""
'End If
    
    
    

'---------------------------------------------------
'---------------------------------------------------
'---------------------------------------------------
'---------------------------------------------------
'---------------------------------------------------
'---------------------------------------------------
'---------------------------------------------------
'---------------------------------------------------
'---------------------------------------------------
'---------------------------------------------------
'---------------------------------------------------
    
    
Dim cc(100, 2)

'If Retry > 0 And Retry < 3 Then
'TTi = ReTryTTi
'End If
'If Retry > 3 Then Retry = 0


Header1$ = Header1$ + "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" + vbCrLf
Header1$ = Header1$ + "<HTML><HEAD>" + vbCrLf
Header1$ = Header1$ + "<TITLE>Matt Lan's Bus Cab Log</TITLE>" + vbCrLf
Header1$ = Header1$ + "<META http-equiv=Content-Type content=""text/html; charset=windows-1252""></HEAD>" + vbCrLf
Header1$ = Header1$ + "<META NAME=""DESCRIPTION"" CONTENT="""
'header2$ = header2$ + "Matt Lan's -  Bus Cab Log"
Header3$ = Header3$ + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + """>" + vbCrLf
Header3$ = Header3$ + "<META NAME=""robots"" content=""index,nofollow"">" + vbCrLf
Header3$ = Header3$ + "<BODY><PRE>" + vbCrLf

Header4$ = Header4$ + "</PRE>"
Header4$ = Header4$ + Grab22$
Header4$ = Header4$ + "</BODY></HTML>"



If TTi >= cc(TT, 2) + 1 Then
    Call Bling2
End If


Detail1(TTi) = False
    
'If Done = True Then
'List3.AddItem "Finished ..........."
'TimerPutToWeb.Enabled = False: Exit Sub
'End If
    
'If ExitPro = True Then End
'If ExitPro2 = True Then Exit Sub
  
    
Select Case TTi
    
    
    
     

    
    'If TTi = cc(
    'Is Multi
    'Case 5 To 11
    Case 99999
    'Case cc(4, 1) To cc(4, 2)
    
    '-----------------------------------------------------------
    '-----------------------------------------------------------
    'IP Addresses
    '-----------------------------------------------------------
    '-----------------------------------------------------------

    'flip this to number 1 get it started
    nhyes = 1
    
    freef3 = FreeFile
    If Dir$("D:\MY WEBS\MatthewLan.Com Web\AWSTATS2\IP-Addy.txt") <> "" Then
        Open "D:\MY WEBS\MatthewLan.Com Web\AWSTATS2\IP-Addy.txt" For Input As #freef3
        Do
        Line Input #freef3, IPaddy2$
        Loop Until EOF(freef3)
        Close #freef3
    End If
    
    IPaddy2$ = Mid$(IPaddy2$, 29)
    
    If (IPaddy1$ <> IPaddy2$ And We2 > 0) Or Check1.Value = vbChecked Or nhyes = 1 Then
        freef3 = FreeFile
        Open "D:\MY WEBS\MatthewLan.Com Web\AWSTATS2\IP-Addy.txt" For Append As #freef3
        Print #freef3, Format$(Now, "DDD DD/MM/YYYY HH:MM:SS"); " --- ";
        Print #freef3, IPaddy1$
        Close #freef3
        
        freef3 = FreeFile
        Open "D:\MY WEBS\MatthewLan.Com Web\AWSTATS2\awstats.matthewlan.com.conf" For Input As #freef3
        freef4 = FreeFile
        Open "D:\MY WEBS\MatthewLan.Com Web\AWSTATS2\awstats.matthewlan2.com.conf" For Output As #freef4
        Do
            Line Input #freef3, linp$
            linp$ = Trim(linp$)
            If Mid$(linp$, 1, 16) = "SkipDNSLookupFor" Or _
                Mid$(linp$, 1, 9) = "SkipHosts" Then
                If InStr(linp$, IPaddy1$) = 0 Then
                    swap22 = 1
                    ipsort$ = Mid$(linp$, 1, Len(linp$) - 1)
                    ipsort$ = ipsort$ + " " + IPaddy1$ + """"
                    linp$ = ipsort$
                End If
            End If
            Print #freef4, linp$
        Loop Until EOF(freef3)
        Close #freef3, #freef4
        
    End If
    
    If (We2 > 0 And swap22 = 1) Or Check1.Value = vbChecked Or nhyes = 1 Then
        Kill "D:\MY WEBS\MatthewLan.Com Web\AWSTATS2\awstats.matthewlan.com.conf"
        Name "D:\MY WEBS\MatthewLan.Com Web\AWSTATS2\awstats.matthewlan2.com.conf" As "D:\MY WEBS\MatthewLan.Com Web\AWSTATS2\awstats.matthewlan.com.conf"
        
        If TTi = cc(4, 1) Then
    
        List3.AddItem "CD --- /httpdocs/awstats2/"
        Call Execute_FTP_Command("CD /httpdocs/awstats2/", ErrStat)
        
        End If
        
        If TTi = cc(4, 1) + 1 Then
        List3.AddItem "Put --- IP Address " + IPaddy1$
        Call Execute_FTP_Command("PUT """ + "D:\MY WEBS\MatthewLan.Com Web\AWSTATS2\awstats.matthewlan.com.conf"" ""awstats.matthewlan.com.conf""", ErrStat)
    
        End If
        If TTi = cc(4, 1) + 2 Then
    
    
        List3.AddItem "CD --- /httpdocs/LoveFolder/"
        Call Execute_FTP_Command("CD /httpdocs/LoveFolder/", ErrStat)
        
        End If
        
    End If
    
    
    
'-----------------------------------------------------------------------
    
    
    Case cc(20, 1) To cc(20, 2)
            

End Select
    

End Sub



Sub PutToWeb()


End Sub

Sub StripFTPPath()


    Easy = InStr(PuTT1, "MatthewLan.Com Web\") + Len("MatthewLan.Com Web\")
    temp$ = Mid$(PuTT1, Easy)
    Easy = InStrRev(temp$, "\")
    PuTT2 = "/httpdocs/" + Mid$(temp$, 1, Easy)
    Putt3 = Mid$(temp$, Easy + 1)
    
    Paths = 0
    Do
        Easy = InStr(PuTT2, "\")
        If Easy > 0 Then
            Mid$(PuTT2, Easy, 1) = "/"
            Paths = Paths + 1
        End If
    Loop Until Easy = 0

End Sub
Sub StripFTPPathTrob()

    ReDim StripFTPPathTrobAY(20)
    StripFTPcnt = 0
    
    Easy = Len(PuTT2) + 1
    Do
        Easy = InStrRev(PuTT2, "/", Easy - 1)
        If Easy < 11 Then Exit Do
        If Easy > 0 Then
            StripFTPPathTrobAY(StripFTPcnt) = Mid$(PuTT2, 1, Easy)
            
        End If
    Loop Until Easy = 0


End Sub

Sub Bling2()
    
    Exit Sub
    
    If List3.List(List3.ListCount - 1) <> "Finished------------------" Then
    'If InStr(UploadFilet1$, "421 No Transfer Timeout (300 seconds): closing control connection.") = 0 Then
        If TheseToDo > 0 Then Exit Sub

        List3.AddItem "Closing FTP Site ---"
        Call Execute_FTP_Command("CLOSE", ErrStat)
    
        List3.AddItem "Finished------------------"
        List3.ListIndex = List3.ListCount - 1
        List3.Refresh
    
    End If
    
    Command2.Enabled = True
    Command3.Enabled = True

    CountD2 = 15
    Timer2.Enabled = True
    TimerPutToWeb.Enabled = False

    On Local Error Resume Next
    Kill "D:\TEMP\WeatherFlag.txt"

End Sub






Function FindWinPart() As Long

Dim ash$

Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long
Dim cText As String
'Dim Rect8 As RECT

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)

FindWinPart = False
    
Do While test_hwnd <> 0
    
    ash$ = GetWindowTitle(test_hwnd)
'    cText = Space$(255)
'    ghj$ = GetClassName(test_hwnd, cText, 255)
    
    If InStr(ash$, "WebDates - Microsoft Visual Basic") > 0 Then
        FindWinPart = True
    End If
    If InStr(ash$, "Web Site Update Dates") > 0 And test_hwnd <> Me.hWnd Then
        FindWinPart = True
    End If
    
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
If FindWinPart = True Then Exit Do
Loop

End Function



'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
'  If IsIDE Then

  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function


Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hWnd)
   S = Space(l + 1)
   GetWindowText hWnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function


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



Function LastModifiedToCurrent(dufile As String)
    Dim lngHandle As Long
    lngHandle = CreateFile(dufile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If TouchFileTimes(lngHandle, ByVal 0&) = 0 Then
        'messageBox "Error while changing file dates. Access Denied." + vbCrLf + duFile, vbCritical, "Modify Error"
    LastModifiedToCurrent = False
    Else
        'messageBox "File date changed successfully.", vbInformation, "Modify Successful"
    LastModifiedToCurrent = True
    End If
    CloseHandle lngHandle
End Function
'---------------------------

