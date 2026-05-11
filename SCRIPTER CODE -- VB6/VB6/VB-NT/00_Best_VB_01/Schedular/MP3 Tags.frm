VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Shedular For Vicea Verse"
   ClientHeight    =   8130
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11640
   Icon            =   "MP3 Tags.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer TimerMini 
      Interval        =   30
      Left            =   5925
      Top             =   2490
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Dont Mni Windows"
      Height          =   195
      Left            =   6135
      TabIndex        =   12
      Top             =   1005
      Width           =   1875
   End
   Begin VB.Timer TimerTheEnd 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5445
      Top             =   2475
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Dont Kill Re-Compare"
      Height          =   195
      Left            =   6120
      TabIndex        =   7
      Top             =   810
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.Timer TimerMainHub 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4515
      Top             =   2475
   End
   Begin VB.Timer TimerReCompareKill 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4050
      Top             =   2475
   End
   Begin VB.Timer TimerCntDownStart 
      Interval        =   1000
      Left            =   4980
      Top             =   2475
   End
   Begin VB.ListBox List1 
      Height          =   5715
      Left            =   -45
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2145
      Width           =   10875
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "#'s"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8670
      TabIndex        =   11
      Top             =   405
      Width           =   2160
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7995
      TabIndex        =   10
      Top             =   810
      Width           =   2835
   End
   Begin VB.Label Lbl8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Run Time ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4425
      TabIndex        =   9
      Top             =   1215
      Width           =   6405
   End
   Begin VB.Label Lbl7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Last Full Run Time ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      TabIndex        =   8
      Top             =   1215
      Width           =   4410
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Last Full Run Finish Time ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      TabIndex        =   6
      Top             =   810
      Width           =   6135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Last Run Finish ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5430
      TabIndex        =   5
      Top             =   405
      Width           =   3225
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Last Run Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2535
      TabIndex        =   4
      Top             =   405
      Width           =   2880
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Last Run Week"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   405
      Width           =   2520
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Waiting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   -15
      TabIndex        =   2
      Top             =   0
      Width           =   10845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Waiting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   -30
      TabIndex        =   0
      Top             =   1620
      Width           =   10860
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "File"
      Begin VB.Menu Mnu_Log 
         Caption         =   "Show Log"
      End
      Begin VB.Menu Mnu_Week 
         Caption         =   "Do Week Month"
      End
   End
   Begin VB.Menu Mnu_Show 
      Caption         =   "Show Vice"
   End
   Begin VB.Menu FullWeek 
      Caption         =   "Do Week"
   End
   Begin VB.Menu Full_Month 
      Caption         =   "Do Month"
   End
   Begin VB.Menu Start 
      Caption         =   "Start Now"
   End
   Begin VB.Menu Continue 
      Caption         =   "Continue Last One"
   End
   Begin VB.Menu ResetRun 
      Caption         =   "Reset Run"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const StartHere = 1 'Take out put zero if not want
'TimerMini.Enabled = True

'**************************************
'Windows API/Global Declarations for :Dr
'     ive Type Finder
'**************************************
Dim Test_Pro As Long, TestH, TXX, ODD, OTTi

Dim Test_Hwnd As Long
Dim CurViceHwnd As Long
Dim AbortDriveNotInSystem
Dim Arry()
Dim ResetRunFlag, Totter2(100), Totter4(100)
Public CountD, IInow, A1$, B1$
Public GetOut, Anck, HG As Long, Totter3$, OldStart2$, OldStart3$, OldStart, TrueFinish$, RunGo
Public DD1, DD2, DD3, DD4, DD5, DD6, DD7, EG, Xshow, Feed, Feeding$, Feeding2$, OldFinish2$, OldFinish As Date
Public Ff, Ffx, Rf, WeekFlag, MonthFlag, WkF$, MnF$, WkF2, MnF2, Fino$, VM, VM2, EndExit, XYZ
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Const WM_CLOSE = &H10
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)




Private Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)




Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function

Sub FileInUseDelay(Tx$)
        
    qq = Now + TimeSerial(0, 5, 0)
    Do
    '    DoEvents
        ii = FileInUse(Tx$)
        If ii = True Then Sleep (200)
    Loop Until ii = False Or qq < Now
    
    If ii = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Tx$ + vbCrLf + "Longer than 5 Min"
        Stop
    End If
End Sub






Private Sub Continue_Click()

Feed = True
On Error Resume Next
Open App.Path + "\#Loggs\Logg File For ViceVersa Schedular One.txt" For Input As #1
    Line Input #1, Feeding$
    Line Input #1, WkF$
    Line Input #1, MnF$
Close #1

If WkF$ = "True" Then WeekFlag = True
If WkF$ = "False" Then WeekFlag = False
If MnF$ = "True" Then MonthFlag = True
If MnF$ = "False" Then MonthFlag = False

Ff = 0
Call LoadList

If XYZ = False Then
'Call TimerMainHub_Timer
'old start on continue think about this
'OldStart3$ = Now
'TimerMainHub.Enabled = True
'TimerCntDownStart.Enabled = False
Call Start_Click
End If
End Sub

Private Sub Form_Load()
'Call TimerReCompareKill_Timer
'End

Set FS = CreateObject("Scripting.FileSystemObject")

IInow = Now + TimeSerial(0, 2, 0)

CountD = 35

If App.PrevInstance = True Then GetOut = True: End

If GetComputerName = "ASUS-MACHINE" Then MsgBox "NOT RUN" + vbCrLf + App.EXEName + vbCrLf + GetComputerName: End


Shell "D:\VB6\VB-NT\00_Best_VB_01\TouchDate\TouchDate TrueCrypt.exe", vbNormalNoFocus

Fino = False

WeekFlag = False
MonthFlag = False

On Error Resume Next
Open App.Path + "\#Loggs\Logg File For ViceVersa Schedular One.txt" For Input As #1
    Line Input #1, Feeding2$
    Line Input #1, WkF3$
    Line Input #1, mnF3$
    Line Input #1, Fino2$
    Line Input #1, OldStart3$
    Line Input #1, OldFinish2$
    
    'Label2 = Feeding2$
    OldFinish = DateValue(OldFinish2$) + TimeValue(OldFinish2$)
    Label6 = "Last Run Finish Time = " + Format$(OldFinish, "DDD DD-MM-YY HH:MM:SSa/p")
    'OldStart = Val(OldStart2$)
Close #1

'If OldStart = 0 Then OldStart = Now - TimeSerial(30, 10, 1)

anck2 = DateValue(OldFinish2$) + TimeValue(OldFinish2$)
anck3 = DateValue(OldStart3$) + TimeValue(OldStart3$)

Tott = DateDiff("s", anck3, anck2)
f5 = DateDiff("d", anck3, anck2)
f6 = DateDiff("h", anck3, anck2)
f7 = DateDiff("n", anck3, anck2) Mod 60
f8 = DateDiff("s", anck3, anck2) Mod 60
Lbl7 = "Last Full Run Time =" + Str(f5) + "D" + Str(f6) + "H" + Str(f7) + "M" + Str(f8) + "S"


If WkF3$ = "True" Then Label3 = "Last Run Week = Yes"
If WkF3$ = "False" Then Label3 = "Last Run Week = No"
If mnF3$ = "True" Then Label4 = "Last Run Month = Yes"
If mnF3$ = "False" Then Label4 = "Last Run Month = No"

If Fino2$ = "False" Then
    XYZ = True
    Call Continue_Click
    Label5 = "Last Run Finish = No"
Else
    Label5 = "Last Run Finish = Yes"
End If


If IsIDE = False And FindWinPart("ViceVersa Pro", 1) > 0 Then
    'MsgBox "Wont Run " + App.EXEName + vbCrLf + " Because ViceVerser Already Going"
    
    'End
End If

Shell "D:\VB6\VB-NT\00_Best_VB_01\Set NetWork Drives\Set-NetWork-Drives.exe"

On Error Resume Next
For Each Control In Controls
    If Control.Width > xx Then xx = Control.Width
    If Control.Height > YY Then YY = Control.Height
Next
On Error GoTo 0

Form1.Width = xx + 80
Form1.Height = YY + 2850


TimerReCompareKill.Enabled = True
Form1.Show
Form1.Hide
'If IsIDE = False Then Form1.WindowState = 1
DoEvents


'Kill "M:\ZIPS\TrueCrypt 01 Day\*.*"

'Shell "C:\Program Files\ViceVersa Pro 2\ViceVersa.exe """ + "D:\# MY DOCS\# 01 My Documents\ViceVersa PRO 2\Acer Backup 03 My Book 02 Not Net.fsf" + """ /dialogautoexec /autoclose /wait"

'End

'End
Form1.Show
'If IsIDE = False Then Form1.WindowState = 1

'Form1.Hide

'Exit Sub

'Load Mp3_Tag
'ScanPath.Show
'List1.AddItem "Drive List Top File Count Each Day"'
'List1.AddItem ""

'On Local Error Resume Next
'freef1 = FreeFile
'Open "C:\callerid\01_Drive_Lists\01 Drive List -- Total List All Log.txt" For Input As #freef1
'Do


'Line Input #1, a1$
'List1.AddItem a1$
'Loop Until EOF(1)
'Close #1
'On Local Error GoTo 0

'Print #freef2, "Date: " + Format$(Now, "DDD DD-MM-YYYY HH-MM-SS")
'Print #freef2, "Total Files So Far : "; TotalAll
'Close #1

'List1.AddItem ""







'yy$ = GetDriveType("J")



'Dim Drivetxt(50, 50)


Dim strDrive As String
Dim strMessage As String
Dim intCnt As Long


Load ScanPath
ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath = "D:\# MY DOCS\# 01 My Documents\ViceVersa PRO 2\#00 Vice\"
ScanPath.cboMask.Text = "*.fsf"
Call ScanPath.cmdScan_Click
For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    ttv = Mid$(A1$, InStrRev(A1$, "\", Len(A1$) - 1) + 1)
    GOAT = 0
    If InStr(UCase(A1$), "WORK") > 0 Then ScanPath.ListView1.ListItems.Remove (we): GOAT = 1
    If InStr(UCase(A1$), "NOT USED") > 0 And GOAT = 0 Then ScanPath.ListView1.ListItems.Remove (we): GOAT = 1
    If Mid$(ttv, 1, 1) = "-" And GOAT = 0 Then
        ScanPath.ListView1.ListItems.Remove (we): GOAT = 1
    End If
Next


If Feed = False Then

    On Error Resume Next
    Open App.Path + "\DateWeek.txt" For Input As #1
    Line Input #1, mweek$
    Close #1
    
    'Changes on sundays
    'Changes on sundays We Want Monday - 1
    sy = DateSerial(Year(Now), 1, 1)
    If Val(mweek$) <> DateDiff("ww", sy, (Now) - 1) Then WeekFlag = True

    Open App.Path + "\DateMonth.txt" For Input As #1
    Line Input #1, mmonth$
    Close #1

    If Val(mmonth$) <> Month(Now) Then MonthFlag = True

End If

Call LoadList

'Label8 = "#'s " + Str(Val(Mid$(Label8, 4)) + 1)


Exit Sub






End Sub

Sub LoadList()

List1.Clear

'Test code to run first one
For we = ScanPath.ListView1.ListItems.Count To 4 Step -1
'    ScanPath.ListView1.ListItems.Remove (we)
Next

DoEvents

A1$ = ScanPath.ListView1.ListItems.Item(1).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(1)

'SMALL DIALOGG STATUS SHOWN
'VM2 = "C:\Program Files\ViceVersa Pro 2\ViceVersa.exe """ + A1$ + B1$ + """ /DialogAutoExec /autoclose"
'THIS WONT WORK WITH THE DETECTION OF OF FIND WINDOWS

'FULL DIAOLOG
VM2 = "C:\Program Files\ViceVersa Pro 2\ViceVersa.exe """ + A1$ + B1$ + """ /autoexec /autoclose /wait"

For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    Rf = 0
    If WeekFlag = False Then Rf = InStr(UCase(B1$), "WEEK")
    If MonthFlag = False Then If Rf = 0 Then Rf = InStr(UCase(B1$), "MONTH")
    If Rf = 0 Then Rf = InStr(UCase(B1$), "ART WEEK")
    If Rf = 0 Then Rf = InStr(UCase(A1$), "NOT USED")
    ttv = Mid$(A1$, InStrRev(A1$, "\", Len(A1$) - 1) + 1)
    If Mid$(ttv, 1, 1) = "-" Then Rf = 1

    If Rf <> 0 Then ScanPath.ListView1.ListItems.Remove (we)
Next

For we = 1 To ScanPath.ListView1.ListItems.Count
A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(we)
List1.AddItem Format$(we, "000 -- ") + B1$

If Feeding$ <> "" Then
    If InStr(Feeding$, B1$) > 0 Then
        Ff = we + 1
        If Ff > ScanPath.ListView1.ListItems.Count Then Ff = 1
        'Display$ = ScanPath.ListView1.ListItems.Item(Ff)
    End If
End If

If Feeding2$ <> "" Then
    If InStr(Feeding2$, B1$) > 0 Then
        'hg = Mid$(Feeding2$, InStr(Feeding2$, "Versa PRO 2\") + 12)
        'hg = Mid$(hg, 1, InStr(hg, """") - 1)
        Label2 = "Last One = " + Format$(we, "000 -- ") + B1$
    End If
End If

If ResetRunFlag = True Then
    'Display$ = ScanPath.ListView1.ListItems.Item(1)
    Ff = 1
End If
    
    
    




'If InStr(b1$, "Lexar") > 0 Then
'    On Error Resume Next
'    Set y = FS.GetDrive(FS.GetDriveName(FS.GetAbsolutePathName("J:")))
'    If Err.Number > 0 Then Rf = 1
'End If



Next
Ff = Ff - 1
If StartHere > 0 Then Ff = StartHere 'start where you want
If Ff < 1 Then Ff = 1
Label1 = "Waiting -- Start With --" + Format$(Ff, "000") + " -- " + ScanPath.ListView1.ListItems.Item(Ff)

Ffx = Ff

Call FileInUseDelay(App.Path + "\#Loggs\Logger Vice Versa Run.txt")
fr = FreeFile
Open App.Path + "\#Loggs\Logger Vice Versa Run.txt" For Append As #fr
Print #fr, "Load List Schedule To Do..."
For we = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(1).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(1)
    Print #fr, Format$(Now, "DDD-DD-MM-YYYY HH:MM:SS ") + "--- " + A1$ + B1$
Next
Print #fr, "Go Start Job To Do..."
Close #fr

End Sub

Private Sub Form_Unload(Cancel As Integer)

If GetOut = True Then End: Exit Sub
If RunGo = False Then End: Exit Sub

'Shell "D:\VB6\VB-NT\00_Best_VB_01\Set NetWork Drives\Set-NetWork-Drives.exe undo"

TrueFinish$ = "no"
If EndExit = True Then
    TrueFinish$ = "yes"
    WeekFlag = False
    MonthFlag = False
    Fino = True
    BakTo$ = "M:"
    VM = VM2
End If

OldFinish = Now
Call FileInUseDelay(App.Path + "\#Loggs\Logg File For ViceVersa Schedular.txt")
Open App.Path + "\#Loggs\Logg File For ViceVersa Schedular.txt" For Append As #1
    Print #1, VM
Close #1

Call FileInUseDelay(App.Path + "\#Loggs\Logg File For ViceVersa Schedular One.txt")
Open App.Path + "\#Loggs\Logg File For ViceVersa Schedular One.txt" For Output As #1
    Print #1, VM
    Print #1, WeekFlag
    Print #1, MonthFlag
    Print #1, Fino
    Print #1, OldStart3$
    Print #1, OldFinish
Close #1


If WeekFlag = True Then
    For r = 1 To 8
    tt = Int(Now) + r
    xx = Weekday(tt)
    If xx = 2 Then Exit For
    Next
    Call FileInUseDelay(App.Path + "\DateWeek.txt")
    Open App.Path + "\DateWeek.txt" For Output As #1
    'Changes on sundays
    'Changes on sundays We Want Monday - 1
    sy = DateSerial(Year(Now), 1, 1)
    Print #1, Str(DateDiff("ww", sy, (Now) - 1))
    Close #1
End If

If MonthFlag = True Then
    For r = 1 To 32
    tt = Int(Now) + r
    xx = Day(Now)
    If xx = 1 Then Exit For
    Next
    Call FileInUseDelay(App.Path + "\DateMonth.txt")
    Open App.Path + "\DateMonth.txt" For Output As #1
    Print #1, Str(Month(Now))
    Close #1
End If

'From Part Call Mnu_Show_Click
'XShow1 = True
ShowHideVicerWindows (True)
'XShow1 = False
End
End Sub


Private Sub Full_Month_Click()
ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath = "D:\# MY DOCS\# 01 My Documents\ViceVersa PRO 2\#00 Vice\"
ScanPath.cboMask.Text = "*.fsf"
Call ScanPath.cmdScan_Click
For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    GOAT = 0
    If InStr(UCase(A1$), "WORK") > 0 Then ScanPath.ListView1.ListItems.Remove (we): GOAT = 1
    If InStr(UCase(A1$), "NOT USED") > 0 And GOAT = 0 Then ScanPath.ListView1.ListItems.Remove (we)
Next

MonthFlag = True
If WeekFlag = True Then Label3 = "Last Run Week = Yes"
If WeekFlag = False Then Label3 = "Last Run Week = No"
If MonthFlag = True Then Label4 = "Last Run Month = Yes"
If MonthFlag = False Then Label4 = "Last Run Month = No"

Call LoadList

End Sub

Private Sub FullWeek_Click()

ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath = "D:\# MY DOCS\# 01 My Documents\ViceVersa PRO 2\#00 Vice\"
ScanPath.cboMask.Text = "*.fsf"
Call ScanPath.cmdScan_Click
For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    GOAT = 0
    If InStr(UCase(A1$), "WORK") > 0 Then ScanPath.ListView1.ListItems.Remove (we): GOAT = 1
    If InStr(UCase(A1$), "NOT USED") > 0 And GOAT = 0 Then ScanPath.ListView1.ListItems.Remove (we)
Next

WeekFlag = True
If WeekFlag = True Then Label3 = "Last Run Week = Yes"
If WeekFlag = False Then Label3 = "Last Run Week = No"
If MonthFlag = True Then Label4 = "Last Run Month = Yes"
If MonthFlag = False Then Label4 = "Last Run Month = No"


Call LoadList

End Sub

Private Sub Label7_Click()
'countd
TimerCntDownStart.Enabled = False
Label7 = "P"
IInow = 0
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
IInow = 0
End Sub

Private Sub Mnu_Log_Click()

Shell "explorer D:\VB6\VB-NT\00_Best_VB_01\Schedular\#Loggs", vbNormalFocus



End Sub

Private Sub Mnu_Show_Click()

'Xshow = Not Xshow
'XShow1 = True
ShowHideVicerWindows (True)
'XShow1 = False

End Sub

Private Sub Mnu_Week_Click()
Kill App.Path + "\DateWeek.txt"
Kill App.Path + "\DateMonth.txt"

End Sub

Private Sub ResetRun_Click()

ScanPath.ListView1.ListItems.Clear
ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath = "D:\# MY DOCS\# 01 My Documents\ViceVersa PRO 2\#00 Vice\"
ScanPath.cboMask.Text = "*.fsf"
Call ScanPath.cmdScan_Click

For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
'    If InStr(UCase(a1$), "Work") > 0 Then ScanPath.ListView1.ListItems.Remove (we)
Next

WeekFlag = False
MonthFlag = False

If WeekFlag = True Then Label3 = "Last Run Week = Yes"
If WeekFlag = False Then Label3 = "Last Run Week = No"
If MonthFlag = True Then Label4 = "Last Run Month = Yes"
If MonthFlag = False Then Label4 = "Last Run Month = No"

ResetRunFlag = True
Ff = 1
LoadList

'Call Start_Click

'Call TimerMainHub_Timer
'TimerMainHub.Enabled = True
'TimerCntDownStart.Enabled = False

End Sub

Private Sub Start_Click()
    RunGo = True
    OldStart2 = Now
    OldStart3$ = Now
    TimerMainHub.Enabled = True
    TimerCntDownStart.Enabled = False
End Sub



Private Sub TimerReCompareKill_Timer()

ii = FindWindow(vbNullString, "Re-comparing ...")
If ii > 0 Then
    
    If Test_Pro > 0 And Check1.Value = vbUnchecked Then
        Process_Kill (Test_Pro)
    End If
End If

End Sub


Function GetParentHwnd(ByVal ReturnParent As Long) As String
   Dim i As Long
   Dim j As Long
   Dim k As Long
   i = ReturnParent
   If ReturnParent Then
      Do While i <> 0
         k = j
         j = i
         i = GetParent(i)
      Loop
   i = j
   End If
   GetParentHwnd = i
End Function



Private Sub TimerMini_timer()
Dim DD(17), TTi As Long, Xo

'DD(0) = "Execute: Copying to Target (Step 1 of 1)"
'DD(1) = "Execute: Copying to Target (Step 2 of 2)"
DD(0) = "Execute: Deleting from Source"
DD(1) = "Execute: Copying to Source"
DD(2) = "Execute: Deleting from Target"
DD(3) = "Execute: Copying to Target"
For r = 0 To 3
    DD(r + 4) = DD(r) + " (Step 1 of 1)"
    DD(r + 8) = DD(r) + " (Step 1 of 2)"
    DD(r + 12) = DD(r) + " (Step 2 of 2)"
Next
DD(16) = "Comparing Source(s) and Target(s)..."
DD(17) = "Re-comparing ..."

For r = 0 To UBound(DD) - 1
    TTi = FindWindow(vbNullString, DD(r))
    If TTi > 0 Then
        Xo = r
        'If r <> 16 Then Stop
        Exit For
    End If
Next

If (TTi = OTTi And ODD = Xo) Or Xo = 0 Then Exit Sub
    'If Xo = 16 Then Stop
    ODD = Xo
    OTTi = TTi
    'Debug.Print Str(Now) + " --- " + Str(TTi) + " --- " + Str(Xo)
    If InStr(DD(Xo), "Deleting") > 0 Then
        YY = WindowVisible.WindowVisible(CurViceHwnd)
        If YY = False Then ODD = 0: Exit Sub
        i = GetWindowState(CurViceHwnd)
        If i <> -1 Then ODD = 0: Exit Sub
        EGO = GetParent(TTi)
        ShowWindow EGO, SW_MINIMIZE
        ShowWindow CurViceHwnd, SW_MINIMIZE
        Exit Sub
    End If

    If InStr(DD(Xo), "Copying") > 0 Then
        YY = WindowVisible.WindowVisible(CurViceHwnd)
        If YY = False Then ODD = 0: Exit Sub
        i = GetWindowState(CurViceHwnd)
        'If i <> -1 Then ODD = 0: Exit Sub
        EGO = GetParent(TTi)
        'If i <> -1 Then ShowWindow EGO, SW_MINIMIZE
        'ShowWindow TTi, SW_MINIMIZE
        If i <> -1 Then ShowWindow CurViceHwnd, SW_NORMAL
        ShowWindow CurViceHwnd, SW_MINIMIZE
        Exit Sub
        If i = -1 Then ShowWindow CurViceHwnd, SW_MINIMIZE
        If i <> -1 Then ShowWindow EGO, SW_MINIMIZE
        Exit Sub
    End If


'If TXX = False Then
    If Check2.Value = vbUnchecked Then
        YY = WindowVisible.WindowVisible(TTi)
        If YY = False Then ODD = 0: Exit Sub
        i = GetWindowState(CurViceHwnd)
        If i <> -1 Then ODD = 0: Exit Sub
        
        EGO = GetParent(TTi)
        If i <> -1 Then ODD = 0: Exit Sub
        
        ShowWindow CurViceHwnd, SW_MINIMIZE
    End If
'End If

End Sub


Sub ShowHideVicerWindows(TF)


'TestH = GetWindowTitle(CurViceHwnd)

If TF = False Then
    If Check2.Value = vbUnchecked Then
        ShowWindow CurViceHwnd, SW_MINIMIZE
    End If
Else
    ShowWindow CurViceHwnd, SW_NORMAL
End If

End Sub

Private Sub TimerMainHub_Timer()
Dim TestH1, TestH2, Result, TT1 As Long

'If TimerCopy1.Enabled = True Then
'    TT1 = FindWindow(vbNullString, "Comparing Source(s) and Target(s)...")
'    If TT1 > 0 Then
        'gws = GetWindowState(TT1)
        'TXX = vbMinimized = True
        'TXX = gws = vbMinimized
'        TXX = WindowVisible.WindowVisible(TT1)
        'If TXX = True Then Stop
 '   End If
'End If


If Anck > 0 Then
    Tott = DateDiff("s", Anck, Now)
    HG = Int(Tott)
    f5 = DateDiff("d", Anck, Now)
    f6 = Int((HG / 60 ^ 2))
    f7 = Int(HG / 60) Mod 60
    f8 = HG Mod 60
    Lbl8 = "Run Time =" + Str(f5) + "D" + Str(f6) + "H" + Str(f7) + "M" + Str(f8) + "S"
End If

Dim CNote As String
Dim TTT3 As Long

If CurViceHwnd > 0 Then
    If GetWindowTitle(CurViceHwnd) = "" Then
        'Really Double Chk the Process has Gone as secondary messaure to the quick find window
        Test_Hwnd = -1
        Result = cProcesses.Convert(Test_Pro, Test_Hwnd, cnFromProcessID Or cnToHwnd)
        If Result = True Then
            'MsgBox "Process Chk Show Hwnd Still Runing Vice Versa":
            Exit Sub
        End If
    Else
        Exit Sub
    End If
End If

Do
    Rf = 0
    A1$ = ScanPath.ListView1.ListItems.Item(Ff).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(Ff)
    If Ff > ScanPath.ListView1.ListItems.Count Then Exit Do
Loop Until Rf = 0

If Anck = 0 Then Anck = Now

Ffx = Ffx + 1
'Shell "C:\Program Files\ViceVersa Pro 2\ViceVersa.exe """ + a1$ + b1$ + """ /dialogautoexec /autoclose /wait", vbMinimizedNoFocus
'Shell "C:\Program Files\ViceVersa Pro 2\ViceVersa.exe """ + a1$ + b1$ + """ /autoexec /autoclose", vbMinimizedNoFocus
DD1 = 0: DD2 = 0: DD3 = 0: DD4 = 0: DD5 = 0: Xshow = 0

Call EditScript

'Bet you i can exit sub here -- is good enoght that will move onto next with Ffx above
If AbortDriveNotInSystem = True Then Exit Sub

VM = "C:\Program Files\ViceVersa Pro 2\ViceVersa.exe """ + A1$ + B1$ + """ /dialogautoexec /autoclose /wait"
VM = "C:\Program Files\ViceVersa Pro 2\ViceVersa.exe """ + A1$ + B1$ + """ /dialogautoexec /autoclose /wait"

VM = "C:\Program Files\ViceVersa Pro 2\ViceVersa.exe """ + A1$ + B1$ + """ /AutoExec /autoclose /wait"

Label1 = "Now - " + Format$(Ff, "000") + " -- " + ScanPath.ListView1.ListItems.Item(Ff)
Label8 = "#'s " + Str(Val(Mid$(Label8, 4)) + 1)

'Test_Pro = Shell(VM, vbMinimizedNoFocus)
Test_Pro = Shell(VM, vbNormalNoFocus)
Dim EGO

Do
    Test_Hwnd = -1 ' Setting flag or it wont read result
    Result = cProcesses.Convert(Test_Pro, Test_Hwnd, cnFromProcessID Or cnToHwnd)
    EGO = GetParentHwnd(Test_Hwnd)
    TestH1 = GetWindowTitle(EGO)
    TestH2 = GetWindowTitle(Test_Hwnd)
    If EGO <> Test_Hwnd And EGO > 0 Then Test_Hwnd = EGO ': Stop
    If TestH1 = "" Then Sleep (500)
Loop Until TestH1 <> ""


CurViceHwnd = Test_Hwnd

Do
    et = FindWindow("#32770", "Output")
    count5 = count5 + 1
    DoEvents
    If et = 0 Then Sleep (500)
Loop Until et > 0
Call CloseProgramHwnd(et)







Call FileInUseDelay(App.Path + "\#Loggs\Logger Vice Versa Run.txt")
fr = FreeFile
Open App.Path + "\#Loggs\Logger Vice Versa Run.txt" For Append As #fr
Print #fr, Format$(Now, "DDD-DD-MM-YYYY HH:MM:SS ") + "--- " + VM
Close #fr


TimerReCompareKill.Enabled = True


Ff = Ffx
If Ff > ScanPath.ListView1.ListItems.Count Then
    TimerMainHub.Enabled = False
    TimerTheEnd.Enabled = True
Exit Sub
End If



End Sub

Sub EditScript()

Dim LACIE, DC, HH$, EQ
        
ReDim ART1(50)
EQ = 0: ART1(EQ) = "LACIE"

'NOT USED AT MOMENT THE 2.5 INCH DRIVE
EQ = EQ + 1: ART1(EQ) = "ACER_C_BAK"
EQ = EQ + 1: ART1(EQ) = "ACER_D_BAK"

EQ = EQ + 1: ART1(EQ) = "WD PASSPORT"
EQ = EQ + 1: ART1(EQ) = "UDISK 8GB"
EQ = EQ + 1: ART1(EQ) = "ICY BOX ART"
EQ = EQ + 1: ART1(EQ) = "MEDIA"

ReDim Preserve ART1(EQ)
ReDim ART2(EQ)

'ReDim ART1(EQ, 2)

'For r = 0 To EQ
'    ART1(r, 1) = ART2(r)
'Next

'It Dont Like this With Preserve dont use it  ----- what a nightmare over this array
'ReDim Preserve ART1(EQ + 1, 2)
'Might as well used 2 arrays wont do that again
'Wank
    
Set DC = FS.Drives
For Each d In DC
    If d.DriveType <> 4 Then
        If d.DriveType = 3 Then
            n = d.ShareName
        Else
            n = d.VolumeName
        End If
    End If
    
    For r = 0 To UBound(ART1)
        If InStr(UCase(n), ART1(r)) > 0 Then
            ART2(r) = d.DriveLetter
            'only chk here if drive in question about to back up wants to know
            'if drive not available one way or another see drive space if
            'source empty better than chk in program vicer
            If InStr(UCase(B1$), "[" + ART1(r) + "]") > 0 Then
                If ART2(r) = "" Then AbortDriveNotInSystem = True: Exit Sub
                If d.isready = False Then AbortDriveNotInSystem = True: Exit Sub
                'd.TotalSize/1024, -- 'd.FreeSpace/1024
            End If
        End If
    Next
Next



Dim KK$, XSEX, fp
Open A1$ + B1$ For Input As #1
Do
    Line Input #1, KK$
    If InStr(KK$, "; ******* Log File") > 0 Then key1 = 1
    If InStr(KK$, "; ******* values for Empty Log: (0=No) (1=Yes)") > 0 Then
        If key1 = 0 Then
            HH$ = HH$ + "; ******* Log File" + vbCrLf
            HH$ = HH$ + "Log File=D:\# MY DOCS\# 01 My Documents\ViceVersa PRO 2\Loggs\Vice Versa Loggs.log" + vbCrLf
        End If
    End If
    If Mid$(KK$, 1, 1) <> ";" Then
        If InStr(KK$, "Speed=") > 0 Then KK$ = "Speed=90"
        If InStr(KK$, "Log File=") > 0 Then KK$ = "Log File=D:\# MY DOCS\# 01 My Documents\ViceVersa PRO 2\#01 Loggs\Vice_Versa_Loggs_" + Trim(Str(Year(Now))) + ".log"
        If InStr(KK$, "Empty Log=") > 0 Then KK$ = "Empty Log=0"
        If InStr(KK$, "Log File Size=") > 0 Then KK$ = "Log File Size=0"
        If InStr(KK$, "Log Only Summary and Exceptions=") > 0 Then KK$ = "Log Only Summary and Exceptions=1"
        '******* values for Priority Class: (0=Default) (1=Idle (recommended)) (2=Low) (3=Below Normal) (4=Normal)
        '          (5=Above Normal) (6=High) (7=Realtime)
        If InStr(KK$, "Priority Class=") > 0 Then KK$ = "Priority Class=1"
        
        'You need this coz only in multi folder mode do you need this it dont like set if not
        If InStr(KK$, "Target Folder=") > 0 And InStr(KK$, " | ") > 0 Then
            XSEX = 1
        End If
        If InStr(KK$, "Continue even if sources/targets do not exist=") > 0 And XSEX = 1 Then
            KK$ = "Continue even if sources/targets do not exist=1"
        End If
        If InStr(KK$, "Continue even if sources/targets do not exist=") > 0 And XSEX = 0 Then
            KK$ = "Continue even if sources/targets do not exist=0"
        End If
        
        If InStr(KK$, "Create sources/targets that do not exist=") > 0 Then KK$ = "Create sources/targets that do not exist=1"
        '=No Confirmations
        If InStr(KK$, "UnatSync=") > 0 Then KK$ = "UnatSync=1"
        'About Middle of Five
        If InStr(KK$, "Copy Buffer=") > 0 Then KK$ = "Copy Buffer=262144"
    
        uu = 0
        If InStr(KK$, "Target Folder=") > 0 Or InStr(KK$, "Source Folder=") > 0 Then
            Do
                uu = InStr(uu + 1, KK$, "\")
                If uu > 0 Then ll = ll + 1
            Loop Until uu = 0
        End If
        If fp = 0 And ll = 1 Then fp = True
        
    
    
        For r = 0 To UBound(ART1)
            If InStr(UCase(B1$), "[" + ART1(r) + "]") > 0 And ART2(r) <> "" Then
                If InStr(UCase(B1$), "[TO]") > 0 Then kiss = "Target Folder="
                If InStr(UCase(B1$), "[FROM]") > 0 Then kiss = "Source Folder="
                
                If InStr(KK$, kiss) > 0 Then
                    uu = 0
                    Do
                        uu = InStr(uu + 1, KK$, ":\")
                        If uu > 0 Then
                            Mid$(KK$, uu - 1, 3) = ART2(r) + ":\"
                        End If
                    Loop Until uu = 0
                End If
            End If
        Next
    
    End If
    HH$ = HH$ + KK$ + vbCrLf
Loop Until EOF(1)
Close #1
Call FileInUseDelay(A1$ + B1$)
Open A1$ + B1$ For Output As #1
Print #1, HH$;
Close #1

Dim HH1, HH2, NumF
HH1 = ""

Call FileInUseDelay(A1$ + B1$)
Open A1$ + B1$ For Input As #1
Do
    Line Input #1, KK$
    HH1 = HH1 + KK$ + vbCrLf
Loop Until KK$ = "[FiltersDirs]"

Line Input #1, NumF

Dim NumFF
NumFF = Val(Mid$(NumF, InStrRev(NumF, "=") + 1)) - 1

NumofFilters = NumF

If NumFF = -1 Then NumFF = 0
ReDim Arry(NumFF)

'If FP = False Then NumFF = 0

'Better than if then Saves Indent
Select Case NumFF

Case Is > 0

Dim KS
KS = 0
kick = 0
For r = 0 To NumFF
    Line Input #1, KK$
    If KK$ = "[ArchiveFilters]" Then kick = 1: Exit For
    Arry(r) = Arry(r) + KK$
Next


'dupe effect code
If NOOOO <> NOOOO Then
For r = 0 To NumFF
    If InStr(LCase(Arry(r)), LCase("=Exclude")) > 0 Then
        If InStr(LCase(Arry(r + 1)), LCase(":\Recycled")) > 0 Then
            Arry(r) = "": Arry(r + 1) = "": Arry(r + 2) = ""
        End If
        If InStr(LCase(Arry(r + 1)), LCase(":\RECYCLER")) > 0 Then
            Arry(r) = "": Arry(r + 1) = "": Arry(r + 2) = ""
        End If
        If InStr(LCase(Arry(r + 1)), LCase(":\System Volume Information")) > 0 Then
            Arry(r) = "": Arry(r + 1) = "": Arry(r + 2) = ""
        End If
    End If
Next
End If


End Select
If kick = 0 Then
Do
    Line Input #1, KK$
Loop Until KK$ = "[ArchiveFilters]"
End If
HH2 = "[ArchiveFilters]" + vbCrLf
Do
    Line Input #1, KK$
    HH2 = HH2 + KK$ + vbCrLf
Loop Until EOF(1)
Close #1

Dim R1, R2, R3
For r = 0 To NumFF
    If InStr(LCase(Arry(r)), LCase("=Exclude")) > 0 Then
        If InStr(LCase(Arry(r + 1)), LCase("=\Recycled")) > 0 Then R1 = True
        If InStr(LCase(Arry(r + 1)), LCase("=\RECYCLER")) > 0 Then R2 = True
        If InStr(LCase(Arry(r + 1)), LCase("=\System Volume Information")) > 0 Then R3 = True
    End If
Next


If fp = True Then
If R1 + R2 + R3 <> -3 Or NumFF = 0 Then

    If R1 = False Then
        ReDim Preserve Arry(UBound(Arry) + 3)
        NumFF = NumFF + 1
        Arry(NumFF) = "Filter " + Trim(Str(NumFF)) + "=Exclude"
        NumFF = NumFF + 1
        Arry(NumFF) = "Filter " + Trim(Str(NumFF)) + "=\Recycled"
        NumFF = NumFF + 1
        Arry(NumFF) = "Filter " + Trim(Str(NumFF)) + "="
    End If
    If R2 = False Then
        ReDim Preserve Arry(UBound(Arry) + 3)
        NumFF = NumFF + 1
        Arry(NumFF) = "Filter " + Trim(Str(NumFF)) + "=Exclude"
        NumFF = NumFF + 1
        Arry(NumFF) = "Filter " + Trim(Str(NumFF)) + "=\RECYCLER"
        NumFF = NumFF + 1
        Arry(NumFF) = "Filter " + Trim(Str(NumFF)) + "="
    End If
    If R3 = False Then
        ReDim Preserve Arry(UBound(Arry) + 3)
        NumFF = NumFF + 1
        Arry(NumFF) = "Filter " + Trim(Str(NumFF)) + "=Exclude"
        NumFF = NumFF + 1
        Arry(NumFF) = "Filter " + Trim(Str(NumFF)) + "=\System Volume Information"
        NumFF = NumFF + 1
        Arry(NumFF) = "Filter " + Trim(Str(NumFF)) + "="
    End If
End If


Dim NumFF2
For r = 0 To NumFF
    If Arry(r) <> "" Then
        Arry(NumFF2) = Arry(r)
        NumFF2 = NumFF2 + 1
    End If
Next
        
ReDim Preserve Arry(NumFF2 - 1)

'test if last arry is data in correct end data if check next one la la
'test = Arry(UBound(Arry))

Dim A7, A8
For r = 0 To UBound(Arry)
    i = InStr(Arry(r), " ")
    ii = InStr(Arry(r), "=")
    'NumFF = Val(Mid$(Arry(r), i, ii - i))
    A7 = Mid$(Arry(r), 1, i - 1)
    A8 = Mid$(Arry(r), ii)
    Arry(r) = A7 + Str(r) + A8
Next

NumofFilters = "Num. of Filters=" + Trim(Str(UBound(Arry)) + 1)

For r = 0 To UBound(Arry)
    Tx = InStr(Arry(r), "=")
    If Tx > 0 Then
        If Mid$(Arry(r), Tx + 2, 1) = ":" Then
            Arry(r) = Mid$(Arry(r), 1, Tx) + Mid$(Arry(r), Tx + 3)
        End If
    End If
Next


End If 'FP=true

ttg = ""
For r = 0 To UBound(Arry)
ttg = ttg + Arry(r) + vbCrLf
Next

'MsgBox ttg
'MsgBox Mid$(ttg, Len(ttg) - 200)

Dim C1, A5, CountF, A9, OA9

CountF = 9999

C1 = Mid$(B1$, 1, InStrRev(B1$, ".") - 1)
A9 = A1$ + C1 + "--(" + Format(CountF, "0000") + ").fsf"
   
FS.copyfile A1$ + B1$, A9

'If FP = False Then

Call FileInUseDelay(A1$ + B1$)
Open A1$ + B1$ For Output Lock Write As #1
    'MsgBox Mid$(HH1, Len(HH1) - 200)
    Print #1, HH1;
    
    Print #1, NumofFilters
    If UBound(Arry) > 0 Then Print #1, ttg;
    Print #1, HH2;
Close #1

'Make Backup
Open A1$ + B1$ For Input Lock Write As #1
Open A9 For Input Lock Write As #2
    FC1 = Input(LOF(1), 1)
    FC2 = Input(LOF(2), 2)
    Close #1, #2
    If FC1 <> FC2 Then
        
        On Error Resume Next
    MkDir A1$ + "-Backup_Auto_Script_Edit"
    On Error GoTo 0
    OA9 = A9

    CountF = 1
    C1 = Mid$(B1$, 1, InStrRev(B1$, ".") - 1)
    Do
        A9 = A1$ + "-Backup_Auto_Script_Edit\" + C1 + "--(" + Format(CountF, "0000") + ").fsf"
        A5 = Dir(A9)
        CountF = CountF + 1
    Loop Until A5 = ""

    FS.copyfile OA9, A9
    Kill OA9
    Else
    Kill A9
End If

'End

A = A

'len(rr1)
'len(rr2)


'Format typical empty file after my edit
'[FiltersDirs]
'Num. of Filters=9
'Filter 0=Exclude
'Filter 1=\Recycled
'Filter 2=
'Filter 3=Exclude
'Filter 4=\RECYCLER
'Filter 5=
'Filter 6=Exclude
'Filter 7=\System Volume Information
'Filter 8=

End Sub

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function

Private Sub TimerCntDownStart_Timer()
CountD = CountD - 1
Label7 = Trim(Str(CountD))

If CountD > 0 Then Exit Sub

TimerCntDownStart.Enabled = False

Call Start_Click
End Sub


Private Sub TimerTheEnd_Timer()

TimerMainHub.Enabled = False

If IInow < Now And IInow > 0 Then Form1.WindowState = 1: IInow = 0


If FindWinPart("ViceVersa Pro -", 1) > 0 Then Exit Sub

'Shell "D:\VB6\VB-NT\00_Best_VB_01\Set NetWork Drives\Set-NetWork-Drives.exe"
If IsIDE = False Then
    Shell "D:\VB6\VB-NT\00_Best_VB_01\Drives_File_Lister\Drives_File_Lister.exe", vbMinimizedNoFocus
    Shell "C:\WINDOZE\system32\ntbackup.exe backup ""@M:\# BACKUPS\BackupNT\BackupNT COMMAND FILE.bks"" /n ""Backup System State Plus Some.bkf created 03/10/2007 at 00:21"" /d ""Set created 03/10/2007 at 00:21"" /n ""Backup System State Plus Some.bkf created 03/10/2007 at 00:21"" /v:no /r:no /rs:no /hc:off /m differential /j ""NtBackup"" /l:s /f ""M:\# BACKUPS\BackupNT\BACKUP FILE.bkf", vbMinimizedNoFocus
End If

TimerTheEnd.Enabled = False
EndExit = True
Unload Form1

End Sub


'---------

Public Sub CloseProgram(ByVal Caption As String)

 Handle = FindWindow(vbNullString, Caption)
 If Handle = 0 Then Exit Sub
 SendMessage Handle, WM_CLOSE, 0&, 0&

End Sub

Public Sub CloseProgramHwnd(ByVal Handle As Long)

 If Handle = 0 Then Exit Sub
 SendMessage Handle, WM_CLOSE, 0&, 0&

End Sub

'---------------------



