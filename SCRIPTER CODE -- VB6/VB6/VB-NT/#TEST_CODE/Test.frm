VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13596
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   13596
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   10440
      Top             =   360
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   780
      Left            =   285
      TabIndex        =   10
      Top             =   5580
      Width           =   2415
      _ExtentX        =   4255
      _ExtentY        =   1376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Timer SCROLL_TIMER_Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6480
      Top             =   1155
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   2475
      TabIndex        =   9
      Top             =   420
      Width           =   870
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4740
      Top             =   1155
   End
   Begin RichTextLib.RichTextBox RTB1 
      DragMode        =   1  'Automatic
      Height          =   870
      Left            =   2445
      TabIndex        =   8
      Top             =   795
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   1524
      _Version        =   393217
      ScrollBars      =   1
      TextRTF         =   $"Test.frx":0000
   End
   Begin VB.TextBox Text1 
      CausesValidation=   0   'False
      Height          =   390
      Left            =   3495
      TabIndex        =   6
      Text            =   "50"
      Top             =   345
      Width           =   2085
   End
   Begin VB.DirListBox Dir1 
      Height          =   1350
      Left            =   7230
      TabIndex        =   3
      Top             =   180
      Visible         =   0   'False
      Width           =   1485
   End
   Begin SHDocVwCtl.WebBrowser WebB1 
      Height          =   4440
      Left            =   2880
      TabIndex        =   2
      Top             =   1950
      Visible         =   0   'False
      Width           =   10080
      ExtentX         =   17780
      ExtentY         =   7832
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2100
      Left            =   780
      ScaleHeight     =   2052
      ScaleWidth      =   1884
      TabIndex        =   1
      Top             =   3330
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   744
      Left            =   300
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2820
      Visible         =   0   'False
      Width           =   6555
   End
   Begin VB.Label Label3 
      Caption         =   "GO"
      Height          =   480
      Left            =   5700
      TabIndex        =   7
      Top             =   375
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   22.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   420
      TabIndex        =   5
      Top             =   885
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   22.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   450
      TabIndex        =   4
      Top             =   195
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   9405
      Top             =   1560
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TO MAKE --- PUT DATE OF LINK
'TO MAKE --- RESUME



Public FS

Dim GG
Dim DayCount1, DayCountMonth, WholeYear1, WeeksSinceYear, WeeksSinceInput, ResultYearDate
Dim InputDate, TestDate
Dim QQ1 As Long
Private Declare Function TouchFileTimes Lib "imagehlp.dll" (ByVal FileHandle As Long, ByRef pSystemTime As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As BOOL
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowDLL Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'DLL
Enum BOOL
   FALSE_ = 0
   TRUE_ = 1
End Enum



Sub SCROLL_OF_DELETE_QUE()

'
Dim FS

Set FS = CreateObject("Scripting.FileSystemObject")
    FS.DeleteFolder "M:\0 00 Art\00 My Pictures", True
    FS.DeleteFolder "M:\01 Loggers", True
    'K\0 00 Art Loggers
Stop


End Sub

'
'B.E.N. was Here 04/28/08

Private Sub GETIP()

'Make sure the "Microsoft winsock 6.0" component is enabled, then put one on the form
'Open "C:\Documents and Settings\All Users\Desktop\" & Winsock1.LocalHostName & ".txt" For Output As #1
MsgBox Winsock1.LocalIP & "  " & Winsock1.LocalHostName
'Close #1



'Unload Me

Stop
End Sub



Sub Percent_1_100()


'## TEST CODE
Dim PER_i, PER_i2, PER_T, PER_OT, PER_x, PER_x1, PER_MULTI


PER_i2 = 0
PER_x = -1
PER_OT = -1

'REM HERE FINAL IMPLIMENT CODE
'------------------------------------
PER_T = 94
PER_i = -0.01
If PER_i > 0 Then PER_i2 = 1 Else PER_i2 = -1
'------------------------------------



'REM HERE FINAL IMPLIMENT CODE
Me.Show

'REM HERE FINAL IMPLIMENT CODE
Do

Sleep 50
DoEvents


PER_MULTI = 10
If PER_T >= 95 Then PER_MULTI = 10
'If PER_T >= 99 Then PER_MULTI = 100
If (PER_T - 1) < 5 Then PER_MULTI = 10
'If (PER_T - 1) < 1 Then PER_MULTI = 100


If PER_OT > 0 And PER_i2 <> 0 Then
    If PER_i2 > 0 Then
        If Int(PER_T * PER_MULTI) <> PER_x Then
            PER_x = Int(PER_T * PER_MULTI)
            
            'REM HERE FINAL IMPLIMENT CODE
            Label1 = Format(Int(PER_T * PER_MULTI) / PER_MULTI, "0.00")
        
        End If
    End If

    If PER_i2 < 0 Then
        If Int(PER_T * PER_MULTI) <> PER_x Then
            PER_x = Int(PER_T * PER_MULTI)
            
            'REM HERE FINAL IMPLIMENT CODE
            Label1 = Format((PER_x / PER_MULTI) + (1 / PER_MULTI), "0.00")
        
        End If
    End If
End If

'REM HERE FINAL IMPLIMENT CODE
Label2 = Format(PER_T, "0.00")



'REM OUT WHEN IMPLIMENT CODE
'-------------------------
If PER_T > 100 Then
    PER_i = PER_i * -1: PER_T = 100: PER_x1 = PER_x1 + 1
End If

If PER_T < 0 Then
    PER_i = Abs(PER_i): PER_T = 0
End If
PER_T = PER_T + PER_i
'-------------------------



'USE THIS WITH IMPLIMENT AND TEST CODE
'---------------------------------------
If PER_OT < PER_T Then PER_i2 = 1 Else PER_i2 = -1
If PER_OT = -1 Then
    If PER_i > 0 Then PER_i2 = 1 Else PER_i2 = -1
End If

PER_OT = PER_T
'---------------------------------------




'REM HERE FINAL IMPLIMENT CODE
'---------------------------------------
Loop Until PER_x1 = 3
End
'---------------------------------------



End Sub


Sub CreateDirAnotherDrive()

'1 LEVEL PATH
'Dir1.H
'Dir1.Path = "M:\"
'Dir1.Path = "M:\Videos"
'Dir1.Path = "M:\04 Music ---"

Dir1.Path = "D:\"
DriveToGoTo = "E"

For R = 0 To Dir1.ListCount - 1

Debug.Print Dir1.List(R)
If FS.FolderExists(DriveToGoTo + Mid(Dir1.List(R), 2)) = False Then
    MkDir DriveToGoTo + Mid(Dir1.List(R), 2)
End If


Next

End


End Sub


Sub StartIndentStartTXTFile()


'Open "D:\# MY DOCS\# 01 My Documents\www.lightningnews.com - Usenet_Groups_List grouplist.txt" For Input As #1
'Open "D:\# MY DOCS\# 01 My Documents\www.lightningnews.com - Usenet_Groups_List grouplist1.txt" For Output As #2




Open "D:\# MY DOCS\# 01 My Documents\New Text Document.txt" For Output As #1
Print #1, Clipboard.GetText;
Close #1



Open "D:\# MY DOCS\# 01 My Documents\New Text Document.txt" For Input As #1
Open "D:\# MY DOCS\# 01 My Documents\New Text Document2.txt" For Output As #2




Do

Line Input #1, LL
If Mid(LL, 1, 2) = "| " Then
LL = Mid(LL, 3)
End If

Print #2, LL


Loop Until EOF(1)


Close #1, #2

Open "D:\# MY DOCS\# 01 My Documents\New Text Document2.txt" For Input As #1
TT = Input(LOF(1), 1)
Close #1

Clipboard.Clear
Clipboard.SetText TT

'Kill "D:\# MY DOCS\# 01 My Documents\www.lightningnews.com - Usenet_Groups_List grouplist.txt"
'Name "D:\# MY DOCS\# 01 My Documents\www.lightningnews.com - Usenet_Groups_List grouplist1.txt" As "D:\# MY DOCS\# 01 My Documents\www.lightningnews.com - Usenet_Groups_List grouplist.txt"




End Sub

Sub Start2RAR()
'###.rar


ScanPath.Show
DoEvents

'ScanPath.txtPath = "D:\03 MICROSOFT\# VB 6\02 VB6 An MSDN 05"
ScanPath.txtPath = "D:\#0 1 INSTALLATIONS\Games Installations\"

ScanPath.chkSubFolders = vbChecked

ScanPath.cmdScan_Click
On Error Resume Next

For we = 1 To ScanPath.ListView1.ListItems.Count
    
    ScanPath.lblCount = str(we)
    DoEvents
    
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
    d1 = Replace(LCase(B1), ".###.rar", "")
    Err.Clear
    R = FS.FileExists(A1 + d1)
    'TT = Dir(A1 + B1)
    'Set F = FS.GetFile(A1 + B1)
    'Idate = F.DateLastModified
    
    
    If Err.Number > 0 Then
        Debug.Print Err.Number
        Debug.Print Err.Description
        Debug.Print A1
        Debug.Print B1
        Stop
    End If
    If InStr(B1, ".###.rar") > 0 Then
        ChDrive A1
        ChDir A1
        
            'Shell "C:\Program Files\WinRAR\WinRAR.exe E """ + A1 + "\*###.rar""", vbMinimizedNoFocus
        Shell "C:\Program Files\WinRAR\WinRAR.exe E -CFG- -Y -pmattkingofkingskingratmatt """ + A1 + B1 + """", vbMinimizedNoFocus
        Kill A1 + B1
    End If
    
Next
End

End Sub


Sub Start1()




ScanPath.txtPath = "M:\0 00 Art\Google Downloader"

ScanPath.chkSubFolders = vbChecked

ScanPath.cmdScan_Click
On Error Resume Next

For we = 1 To ScanPath.ListView1.ListItems.Count
    
    DoEvents
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
    
    Err.Clear
    R = FS.FileExists(A1 + B1)
    TT = Dir(A1 + B1)
    Set F = FS.GetFile(A1 + B1)
    Idate = F.DateLastModified
    
    
    If Err.Number > 0 Then
    Debug.Print Err.Number
    Debug.Print A1
    Debug.Print B1
    
    Stop
    
    
    
    End If
    
Next
End

End Sub


Sub ReSizeForm()
'Code to Auto Size Form based on controls used Including Menu Bar Sketchy Style
'Make sure form set to scale Twips
'put in your form load

X = 1
Y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > X Then X = Control.Width + Control.Left
        If Control.Height + Control.Top > Y Then Y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next
On Error GoTo 0

Me.Width = (X + 80)
If mnu = 1 Then
    pluso = 720: pluso = 1100 'Sometimes Different
Else
    pluso = 450
End If

Me.Height = (Y + pluso)
Me.Refresh
DoEvents

'Me.Left = (Screen.Width - Me.Width)
'Me.Top = (Screen.Height - Me.Height)

End Sub


Sub WSCRIPT()


'WshNetwork Object

'WshNetwork.ComputerName

'WshNetwork.UserDomain

'WshNetwork.UserName

Dim VbsObj As Object
Set VbsObj = CreateObject("WScript.Network")
  
'Set objNet = WSCRIPT.CreateObject("WScript.Network")
Debug.Print
Debug.Print VbsObj.ComputerName
Debug.Print VbsObj.USERDOMAIN
Debug.Print VbsObj.UserName

'WSCRIPT.Echo "Your Computer Name is " & objNet.ComputerName
'WSCRIPT.Echo "Your Username is " & objNet.UserName


End Sub


Sub EBAY45DAYS()

    'howe8458 ( 25827)    Contact seller
    'Paid on 31-Dec-09 via PayPal
    'Willy Wonka And The Chocolate Factory - DVD - NEW
    x1 = DateSerial(2009, 12, 31)
    x2 = Now
    X3 = DateDiff("d", x1, x2)
    X4 = Format(x1 + 45, "DDD DD-MMM-YYYY")
    Debug.Print "Days Since" + str(X3)
    Debug.Print "Date due for 45 Days - " + X4

    
    Debug.Print "Corrospondance's -- Before"
    x1 = DateSerial(2009, 12, 31)
    x2 = DateSerial(2010, 2, 18)
    X3 = DateDiff("d", x1, x2)
    X4 = Format(x1 + 45, "DDD DD-MMM-YYYY")
    Debug.Print "Days Since" + str(X3)
    Debug.Print "Date Corro - " + X4
    
    '---DEBUG OUTPUT
    'Days Since 64
    'Date due for 45 Days - Sun 14-Feb-2010
    'Corrospondance 's -- Before
    'Days Since 49
    'Date Corro - Sun 14-Feb-2010

    'PAYPAL WAS ENTERED 24-02-10
    'SEEMS HARD BELIVE 45 DAY WAS PASSED
    
    
    'KILL MOCKING BIRD 8 JAN
    x1 = DateSerial(2010, 1, 8)
    x2 = Now
    X3 = DateDiff("d", x1, x2)
    X4 = Format(x1 + 45, "DDD DD-MMM-YYYY")
    Debug.Print "KILL MOCK"
    Debug.Print "Days Since" + str(X3)
    Debug.Print "Date due for 45 Days - " + X4
    
    'ZAPPER
    x1 = DateSerial(2010, 1, 29)
    x2 = Now
    X3 = DateDiff("d", x1, x2)
    X4 = Format(x1 + 45, "DDD DD-MMM-YYYY")
    Debug.Print "ZAPPER"
    Debug.Print "Days Since" + str(X3)
    Debug.Print "Date due for 45 Days - " + X4
    'DUE Sun 14-Feb-2010

    'CALENDAR CARDS
    x1 = DateSerial(2010, 2, 2)
    x2 = Now
    X3 = DateDiff("d", x1, x2)
    X4 = Format(x1 + 45, "DDD DD-MMM-YYYY")
    Debug.Print "CALENDAR CARD"
    Debug.Print "Days Since" + str(X3)
    Debug.Print "Date due for 45 Days - " + X4
    'Date due for 45 Days - Fri 19-Mar-2010

    'POWER CONNECTORS
    x1 = DateSerial(2010, 2, 5)
    x2 = Now
    X3 = DateDiff("d", x1, x2)
    X4 = Format(x1 + 45, "DDD DD-MMM-YYYY")
    Debug.Print "POWER CONNECTORS"
    Debug.Print "Days Since" + str(X3)
    Debug.Print "Date due for 45 Days - " + X4
    'Date due for 45 Days - Mon 22-Mar-2010


End Sub

Sub SCROLLWIN()

On Error Resume Next
For Each Control In Controls
Control.Visible = False

Next

RTB1.Visible = True



End Sub



Sub SUSPENDCODE()

'Call Process_Idle(i)
'i = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE", vbNormalFocus)
i = ExecCmd("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE")


Do

t = Now + TimeSerial(0, 0, 2)
Do
'A = GetWindowTitle(i)
'a = GetActiveWindowTitle(i)
'A = GetActiveWindowClassParent(i)
'A = GetActiveWindowClass(i)
If a <> oa Then B = B + a + vbCrLf
oa = a
DoEvents

Loop Until t < Now
MsgBox B

Call Thread_Suspend(test_thread_id)

o = MsgBox("Hi", vbOKCancel)

Call Thread_Resume(test_thread_id)

Loop Until o = vbCancel

End


End Sub

Sub GOODSYNC_EDITOR()

FR1 = FreeFile
Open "C:\SCRIPTER\jobs-groups-options - Copy.tic" For Input As FR1
TT_EDITOR = StrConv(Input(LOF(FR1), FR1), vbUnicode)
Debug.Print "----"
Debug.Print LOF(FR1)
Debug.Print Len(TT_EDITOR)
Close #FR1

X = 1
Do
    T1 = InStr(X, TT_EDITOR, StrConv("wait", vbUnicode))
    
    If T1 > 0 Then
        T3 = InStrRev(TT_EDITOR, StrConv("|M", vbUnicode), T1)
        
        T2 = InStr(T3, TT_EDITOR, StrConv("|i=", vbUnicode))
        N3 = Mid(TT_EDITOR, T3, (T2 - T3) + Len(StrConv("|i=", vbUnicode)))
        
        REASSEMBLE_VAR = StrConv("|M:0:|U=0|B:0:|J:0:|i=", vbUnicode)
        TT_EDITOR = Replace(TT_EDITOR, N3, REASSEMBLE_VAR, , 1)
        
        Counter_VAR = Counter_VAR + 1
    End If
    
    X = T1 + 8
Loop Until T1 = 0

FR1 = FreeFile
Open "C:\SCRIPTER\jobs-groups-options.tic" For Output As FR1
Print #FR1, StrConv(TT_EDITOR, vbFromUnicode);
Close #FR1

MsgBox "Done Total " & Counter_VAR & " Proecesses"

End Sub


Private Sub Form_Load()


' --------------------------------------------------
' VIDEO RATIO CONVERSION BIT RATE
' RESOLUTION CHANGER FROM 1440x1080 TO 1280x720
' FIND THE BIT RATE FOR VALUE COST
' --------------------------------------------------
x1 = 1440               ' IN RATIO
y1 = 1080
' --------------------------------------------------
x2 = 1280               ' OUT RATIO
y2 = 720
' --------------------------------------------------
v1 = x1 * y1            ' PIXEL MULTIPLER
' --------------------------------------------------
v2 = x2 * y2            ' PIXEL MULTIPLER
VT = (v2 / v1) * 100    ' 59.2592592592593 SMALLER %
' --------------------------------------------------
BR_1 = 11997            ' BIT RATE
BR_T = BR_1 / 100 * VT  ' RATIO BIT RATE AFTER
' --------------------------------------------------
Debug.Print "" & VT
Debug.Print BR_T & vbCr ' 11997 BIT RATE -- NOW AFTER -- 7109.33333333333
' METHOD USER BIT RATE 9000 AND RESULT WHEN RECODE -- 9044
' ALL VALUE BIT RATE WHEN DONE A RESULT HIGHER
' CONVERT WHOLE FILE SEE WHAT MEG AT
' G-PHOTO NOT WORK FILE BIGGER THAN 4GB
' AND HERE WAS 6.5GB
' 2020_04_29 Wed_Apr 14_08_26__MAH00164__14_08_26_PM__RAIN.MP4
' LOAD FILE IN AGAIN AFTER REDITOR
' PUT 8200 IN AND GOT 8200 OUT
' CLAIM FILE SIZE BE 4579 GB -- TEST DRIVE
' OR LOWER
' ADJUST IN 7800 RESULT GET 7825 -- 4362 GB -- BIGGER
' ADJUST IN 7200 RESULT GET 7153 -- 4037GB -- BITRATE VS GB -- SMALLER
' DEPEND SIZE FILE AND VARIABLE BITRATE TYPE THING -- EXAMPLE CUT SHORT BEGIN TO RESULT FIND
' CHECK VISUAL QUALITY
' 7200 MIGHT BE NEW STANDARD OR UP DEPEND ON SIZE
' FINAL DONE 3.9 TB 3975 GB  7200 REQUEST NUMBER BIT RATE AND DO SAME
' FINAL DONE FROM 6.5 GB
' FINAL CLAIM BEFORE ADJUST IN 7200 RESULTER 4037GB ACTUALLY RESULT WAS 3.9 TB 3975 GB
' --------------------------------------------------
End
Exit Sub



Set FS = CreateObject("Scripting.FileSystemObject")


Call GOODSYNC_EDITOR


End

Exit Sub

    
    
    a = 240 / 3
    B = 10 / 3
    Debug.Print B
    d = 240
    
    
    
    Exit Sub


    Call Start2RAR

Exit Sub

'    266.93
Debug.Print 265.93 - 250 'BANK


ShowTaskBar
    
    
'Call SCROLL_OF_DELETE_QUE

End
    
Set FS = CreateObject("Scripting.FileSystemObject")

Call GETIP
Exit Sub

Call SUSPENDCODE

Exit Sub

Call SCROLLWIN

Me.Show


On Error Resume Next
For Each Control In Controls
    Control.Visible = False
Next
On Error GoTo 0


Text2.Top = 0
Text2.Left = 0
RTB1.Left = 0
RTB1.Top = Text1.Top + Text1.Height + 15
RTB1.Width = Me.Width
RTB1.Height = Me.Height

RTB1.LoadFile "M:\# W880i\Other\Letter Andy 2009-07-11.txt"
RTB1.Visible = True
Text2.Visible = True


Exit Sub


Call CreateDirAnotherDrive
End

Exit Sub

    Call EBAY45DAYS

End



Call WSCRIPT


End

Call Percent_1_100





Dim i As Long


ir = DateValue("2009-12-21 00:05:38") + TimeValue("2009-12-21 00:05:38")
ir = ir - TimeSerial(0, 58, 26)
I2 = Format(ir, "DDD DD-MMM-YYYY HH:MM:SS")
Clipboard.Clear
Clipboard.SetText I2

'=Sun 20-Dec-2009 23:07:12

End


Call StartIndentStartTXTFile

End
Exit Sub

Call Start1

Exit Sub
End

Dim glong As Double, glat As Double, tz As Double
glat = 50.82845: glong = -0.180663
glong = 50.82845: glat = -0.180663
tz = 0
Dim RR As String

RR = sunevent(year(Now), month(Now), day(Now), glong, glat, tz, 1)

Debug.Print RR
Debug.Print "----------"

RR = sunevent(2000, 1, 3, -1.91667, 52.5, 0, 1)
Debug.Print RR
Debug.Print "----------"

'Result 02 -- Data Where I Live Compare with True Data Not Correct
'Result 01
'0841 ---- Should Be 06:43
'----------
'Result 02 -- Same Data Provide as your example used with Qbasic but not same result
'0035 ---- This Result Should Be
'----------




End


'GoTo Bush

ee = Clipboard.GetText

ee = LCase(ee)

Do
    tt2 = InStr(tt2 + 1, ee, " ")
    If tt2 > 0 Then Mid$(ee, tt2 + 1, 1) = UCase(Mid$(ee, tt2 + 1, 1))
    tt3 = InStr(tt3 + 1, ee, vbCrLf)
    If tt3 > 0 Then Mid$(ee, tt2 + 1, 1) = UCase(Mid$(ee, tt2 + 1, 1))
Loop Until tt2 = 0

tt2 = 0
Do
    tt2 = InStr(tt2 + 1, ee, vbLf)
    If tt2 > 0 And tt2 + 1 <= Len(ee) Then Mid$(ee, tt2 + 1, 1) = UCase(Mid$(ee, tt2 + 1, 1))
Loop Until tt2 = 0

tt2 = 0
Do
    tt2 = InStr(tt2 + 1, ee, "-")
    If tt2 > 0 And tt2 + 1 <= Len(ee) Then Mid$(ee, tt2 + 1, 1) = UCase(Mid$(ee, tt2 + 1, 1))
Loop Until tt2 = 0
tt2 = 0
Do
    tt2 = InStr(tt2 + 1, ee, "(")
    If tt2 > 0 And tt2 + 1 <= Len(ee) Then Mid$(ee, tt2 + 1, 1) = UCase(Mid$(ee, tt2 + 1, 1))
Loop Until tt2 = 0

Clipboard.Clear
Clipboard.SetText (ee)



End


Bush:

Open "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\BUSH TORNADO OUT.txt" For Input As #1
Open "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\BUSH TORNADO OUT2.txt" For Output As #2

Do
'RR = ""
RR2 = ""
For R = 1 To 4
    'If EOF(1) Then Exit Do
    Line Input #1, OO
    'RR = RR + OO
    If R = 1 Then RR3 = OO + vbCrLf
    If R > 1 Then RR2 = RR2 + OO
    'If R < 4 Then RR = RR + vbCrLf
    If R < 4 And R > 1 Then RR2 = RR2 + vbCrLf
Next
'LEN(RR)
If InStr(LCase(RR), LCase(RR2)) = 0 Then
    Print #2, RR3 + RR2
    cc = cc + 1
End If
RR = RR + RR2

Loop Until EOF(1)
Close #1, #2









End





End
Dim A1, B1

Form1.Show
DoEvents

'WebB1.LocationURL = "http://news.bbc.co.uk/weather/"






Exit Sub

'http://news.bbc.co.uk/weather/



'For Each Control In Controls
'Control.Visible = False
'Next

'Picture1.Visible = True
Form1.Show
Form1.Refresh
DoEvents
With Picture1
    .Top = 0
    .Left = 0
    .Height = Screen.Height - 1550
    .Width = Screen.Width - 100
End With
Me.Left = 0
Me.Top = 600
'Picture1.Refresh
'DoEvents
Call ReSizeForm

ScanPath.txtPath = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Vintage"
ScanPath.txtPath = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Ultra Sexi"
'ScanPath.txtPath = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Ultra Sexi\#Bad Pix"
ScanPath.txtPath = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Cartoon"
'ScanPath.txtPath = "M\0 00 Art Loggers\Screen-Shots"
'ScanPath.txtPath = "M\0 00 Art Loggers\Web_Cam\WebCapture_2009-08-06-Thu"
ScanPath.txtPath = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Cartoon"

ScanPath.cmdScan_Click
ScanPath.chkSubFolders = vbUnchecked
Dim lPic As Picture
For we = 1 To ScanPath.ListView1.ListItems.Count
    If we > 50 Then Exit For
    DoEvents
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
    On Error Resume Next
    
    
    'Image1.Picture = LoadPicture(A1 + B1)
    'Picture1.Width = Image1.Width / 2
    'Picture1.Height = Image1.Height / 2
    
    
    
    'Picture1.Picture = LoadPicture(A1 + B1)
    'Picture1.ScaleMode = 3
    'Picture1.AutoRedraw = True
    'Picture1.PaintPicture Picture1.Picture, _
    '0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
    '0, 0, Picture1.Picture.Width / 26.46, _
    'Picture1.Picture.Height / 26.46
    'Picture1.Picture = Picture1.Image

    Me.Picture1.AutoRedraw = True
    Set lPic = LoadPicture(A1 + B1) 'Use the correct path and filename here
    ResizePicture Me.Picture1, lPic
    
    
    
    'Call ReSizeForm
    
    
    
    'Picture1.Picture = Image1.Picture
Next
Beep

Exit Sub



'A = FindAvailablePorts
'Exit Sub
'A = EnumeratingConnectionsStates
'Exit Sub
'A = GetIPAddress
'A = FindNetworkAddress
'A = GetMACAddress
'A = IsInternetConnected
'A = MemoryStatus
'A = ChangeGrayTextColor(0) ' dont work on argus grey net
'A = ProcessTimings
'i = FindWindow(vbNullString, "Windows Task Manager")
'A = DisableClose(i)
'a = EnableClose(i)
'CaptureScreen() As Picture
'A = Show_TrayIcons ' didnt see much


'lot nice screen captures
'no logg off or go to welcome screen

'End
'Exit Sub
'frmMain.Show

'Exit Sub



'Sleep 400

'call
'Call Thread_Suspend(i)



Exit Sub

Open "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\WinAmp Music Lenght Logg Day Total.txt" For Input As #1
Open "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\WinAmp Music Lenght Logg Day Total --.txt" For Output As #2
Do
If ST Mod 100 = 0 Then DoEvents
Line Input #1, kk
easy = DateValue(Mid$(kk, 1, 19)) + TimeValue(Mid$(kk, 1, 19))
If easy > oeasy Then
Print #2, kk
Else
ST = ST + 1
End If
If easy > oeasy Then oeasy = easy
'04-01-2008 01:32:51
Loop Until EOF(1)
Close #1, #2
MsgBox str(ST)

End
'----------------------

List1.Width = Form1.Width
List1.Height = Form1.Height


'1st RockPort Boots MARANG2B £124.99
InputDate = DateValue("02-12-2000")
TestDate = DateValue("27-12-2002")
Call FindTimeInfo
TT$ = Format(InputDate, "DD-MM-YYYY") + " -- 1st Rockport Boots Black £124.99  -- " + ResultYearDate + " Years" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Weeks" + vbCrLf
List1.AddItem TT$
tt1$ = tt1$ + TT$


'2nd
'MyRockPorts
'XCS cross condition system rockport boots
'£170
'27 Dec 2002
'5 years ago he was doing that

InputDate = DateValue("27-12-2002")
TestDate = DateValue("06-09-2008")
Call FindTimeInfo
TT$ = Format(InputDate, "DD-MM-YYYY") + " -- 2st Rockport Boots Gold £160 -- " + ResultYearDate + " Years" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Weeks" + vbCrLf
List1.AddItem TT$
tt1$ = tt1$ + TT$

InputDate = DateValue("27-12-2002")
TestDate = Now
Call FindTimeInfo
TT$ = Format(InputDate, "DD-MM-YYYY") + " -- 2st Rockport Boots Gold £160 Still Got -- " + ResultYearDate + " Years" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Weeks" + vbCrLf
List1.AddItem TT$
tt1$ = tt1$ + TT$

'3rd
'Description Colour Size Price Qty Item Status
'320 137  Black and Silver Male 420 ( UK 8 ) £124.99 1 Ready to deliver
'Order Summary
'Sub Total(): £124.99
'Delivery Charge: £2.99
'Order Total: £127.98
'06 September 2008 01:27
'helpdesk@shoe-shop.com <helpdesk@shoe-shop.com>

InputDate = DateValue("06-09-2008")
TestDate = Now
Call FindTimeInfo
TT$ = Format(InputDate, "DD-MM-YYYY") + " -- 3st Rockport Boots Black/Silver £127.99 -- " + ResultYearDate + " Years" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Weeks" + vbCrLf
List1.AddItem TT$
tt1$ = tt1$ + TT$


Clipboard.Clear
Clipboard.SetText (tt1$)


    
    
'Calll DateCalc
    
Exit Sub
    


    'Dim myDateTime As Date = New Date(2000, 1, 1, 12, 34, 56)
    'Dim myFileInfo As FileInfo

    ' Get a handle on the file we're manipulating
    'myFileInfo = New FileInfo(Server.MapPath("test.html"))

    ' Change timestamps to new values
    myFileInfo.CreationTime = myDateTime
    myFileInfo.LastWriteTime = myDateTime

Exit Sub

internetaddys = 255! * 255! * 255! * 255!
List1.AddItem Format$(internetaddys, "###,###,###,###,###,###") + " Internet Adresses"

Exit Sub

Call HalfYear
Exit Sub
End
Call Redate
Exit Sub
End

Call Test3
Exit Sub
End
'Call Test2
'Exit Sub
Call Test1
End
Exit Sub

'1 pound = 0.0714285714 stone
'    More about calculator.

'1 pound = 0.0714285714 stone
stone = 112 * 0.0714285714
stone = 212 * 0.0714285714
Debug.Print stone
'= 7.9999999968
'= 15.1428571368
 
 
 'More than a year ago, in July 2007, International Space Station astronauts threw an obsolete, refrigerator-sized ammonia reservoir overboard. Ever since, the 1400-lb piece of space junk has been circling Earth in a decaying orbit--and now it is about to reenter.
 
stone = 1400 * 0.0714285714
Debug.Print stone

 
 
 
 
' That 's cheating! You still don't know HOW it's done.

'256 ^ 0 = 1
'256 ^ 1 = 256
'256 ^ 2 = 65,536
'256 ^ 3 = 16,777,216

'So, 194.247.44.146 is ...

'146 * 256^0 = 146
'44 * 256^1 = 11,264
'247*256^2 = 16,187,392
'194*256^3 = 3,254,779,904
'
'Add those up, and ...

'146 + 11,264 + 16,187,392 + 3,254,779,904 = 3270978706
'
'3410960384  3410964479  UK  UNITED KINGDOM

Dim a0

a0 = 3270978706#
A1 = a0 / (256 ^ 3)
A2 = A1 - Int(A1) / 256 ^ 2
A3 = 3270978706# / 256 ^ 1
A4 = 3270978706# / 256 ^ 0

 
'203.79.32.0
'203.79.47.255
 
 
 
End Sub


Sub HalfYear()
Dim qw As Double
'OMM = 1
'MON = 1
'DD = 1
'qw = DateDiff("s", DateValue("01-01-2009"), DateValue("1-1-2010"))
'For r = 1 To (qw / 2)
'ss = ss + 1
'If ss = 60 Then mm = mm + 1: ss = 0
'If mm = 60 Then hh = hh + 1: mm = 0
'If hh = 24 Then DD = DD + 1: hh = 0
'tt = DateSerial(2008, MON, DD)
'If Month(tt) <> OMM Then
'MON = MON + 1: DD = 1
'End If
'OMM = Month(tt)
'DoEvents
'Next

qw1 = DateDiff("s", DateValue("01-01-2009"), DateValue("1-1-2010"))
qw = DateDiff("d", DateValue("01-01-2009"), DateValue("1-1-2010"))
'List1.AddItem Str(qw)
tt1 = Int(DateSerial(2008, 1, 1) + qw / 2)
List1.AddItem Format$(tt1, "DD-MM-YYYY") + " Half a Normal Year Days"
qw2 = DateDiff("s", DateValue("01-01-2008"), tt1)
qw3 = DateDiff("s", tt1 + 1, DateValue("01-01-2009"))
tt2 = (qw1 - qw2 - qw3) / 2
'tt2 = qw1 / 2

tt3 = (tt2 / 60) Mod 60
tt4 = tt2 / (60 * 60)
List1.AddItem Format$(tt1 + TimeSerial(tt4, tt3, tt2 Mod 60), "DD-MM-YYYY HH:MM:SS") + " Half a Normal Year Days and Time"

qw1 = DateDiff("s", DateValue("01-01-2008"), DateValue("1-1-2009"))
qw = DateDiff("d", DateValue("01-01-2008"), DateValue("1-1-2009"))
'List1.AddItem Str(qw)
tt1 = Int(DateSerial(2008, 1, 1) + qw / 2)
List1.AddItem Format$(tt1, "DD-MM-YYYY") + " Half a Leap Year Days"
qw2 = DateDiff("s", DateValue("01-01-2008"), tt1)
qw3 = DateDiff("s", tt1 + 1, DateValue("01-01-2009"))
tt2 = (qw1 - qw2 - qw3) / 2
'tt2 = qw1 / 2

tt3 = (tt2 / 60) Mod 60
tt4 = tt2 / (60 * 60)
List1.AddItem Format$(tt1 + TimeSerial(tt4, tt3, tt2 Mod 60), "DD-MM-YYYY HH:MM:SS") + " Half a Leap Year Days and Time"



End Sub

Sub Redate()
Open "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info_Ace\00 TextData\Idle and Active Logger Idle Over Time.txt" For Input As #1
Open "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info_Ace\00 TextData\Idle and Active Logger Idle Over Time.txt.tmp" For Output As #2
Do
Line Input #1, TT$
RR$ = Format$(DateValue(Mid$(TT$, 1, 9)), "DDD ")
Print #2, RR$ + TT$
Loop Until EOF(1)
Close #1, #2

End Sub

Sub Test2()

Dim Idate As Date, Hdate As String, File1 As String, File2 As String

Call GetFileDates

File1 = Mid$(List1.List(List1.ListCount - 1), 21)
For R = List1.ListCount - 2 To 0 Step -1
File2 = Mid$(List1.List(R), 21)
If Mid$(List1.List(List1.ListCount - 1), 1, 19) <> Mid$(List1.List(R), 1, 19) Then
FS.CopyFile File1, File2
TT = LastModifiedToCurrent(File2)
End If
Next

Call GetFileDates

End Sub

Public Sub GetFileDates()
Dim FS, Idate As Date, Hdate As String, File1 As String, File2 As String
Set FS = CreateObject("Scripting.FileSystemObject")
List1.Clear
File1 = "D:\VB6\VB-NT\00_Best_VB_01\Banging_Tunes\BangList Total Logg.html"
Call Date_File(File1, Idate, Hdate)
'dd = Hdate + " " + file1
List1.AddItem Hdate + " " + File1

For R = 3 To 25
On Error Resume Next
Err.Clear
Set g = FS.GetDrive(FS.GetDriveName(FS.GetAbsolutePathName(Chr$(R + 64) + ":")))
On Error GoTo 0
If InStr(RR, g.VolumeName) = 0 And Err.Number = 0 Then
RR = RR + g.VolumeName + "-"
If Mid(g.VolumeName, Len(g.VolumeName) - 1, 2) = "GB" Then

File2 = Chr$(R + 64) + ":\Banging\BangList Total Logg.html"
Call Date_File(File2, Idate, Hdate)
List1.AddItem Hdate + " " + File2

'fs.copyFile file1, file2


End If
'Nuke4 = Nuke4 + g.totalsize / 1024 ^ 3
'Nuke5 = Nuke5 + g.freespace / 1024 ^ 3
End If
Next

End Sub

Public Sub Date_File(Filespec2, Idate As Date, Hdate As String)
'Call Date_File(filespec2$, Idate)

Dim FS, F
Set FS = CreateObject("Scripting.FileSystemObject")

pdate = 0
FileSpec = Filespec2
Idate = DateSerial(1980, 1, 1)
If FS.FileExists(FileSpec) Then
Set F = FS.GetFile(FileSpec)
Idate = F.DateLastModified
End If
Hdate = Format$(Idate, "YYYY-MM-DD HH:MM:SS")


End Sub


Public Function FindFileSize(FileName As String) As Long
        'On Error Resume Next
        'Dim filedata As WIN32_FIND_DATA

        'filedata = Findfile(FileName)
        
        'If filedata.nFileSizeHigh = 0 Then
        '    FindFileSize = Format$(filedata.nFileSizeLow, "###,###,###")
        'Else
        '    FindFileSize = Format$(filedata.nFileSizeHigh, "###,###,###")
        'End If
End Function

Function LastModifiedToCurrent(duFile As String)
    Dim lngHandle As Long
    lngHandle = CreateFile(duFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If TouchFileTimes(lngHandle, ByVal 0&) = 0 Then
        'MsgBox "Error while changing file dates. Access Denied." + vbCrLf + duFile, vbCritical, "Modify Error"
    LastModifiedToCurrent = False
    Else
        'MsgBox "File date changed successfully.", vbInformation, "Modify Successful"
    LastModifiedToCurrent = True
    End If
    CloseHandle lngHandle
End Function

Sub Test1()

'restore header on wave fle
Open "m:\DWR 2008-11-27 00-36-03 -- 2008-11-28 13-02-14 -#- Janice Paul Short.wav" For Binary As #1
LL$ = "TTTT"
Get #1, 5, LL$
QQ1 = LOF(1)
Close #1
A1 = Asc(Mid$(LL, 1, 1))
A2 = Asc(Mid$(LL, 2, 1))
A3 = Asc(Mid$(LL, 3, 1))
A4 = Asc(Mid$(LL, 4, 1))
Mid(LL, 1, 1) = Chr(&H78)
Mid(LL, 2, 1) = Chr(&H10)
Mid(LL, 3, 1) = Chr(&H83)
Mid(LL, 4, 1) = Chr(&H23)

Mid(LL, 1, 1) = Chr(&H78)
Mid(LL, 2, 1) = Chr(&H94)
Mid(LL, 3, 1) = Chr(&H45)
Mid(LL, 4, 1) = Chr(&H0)

Open "M:\05 Media\Digital Wave Player M Drive\DR2008\DWR 2008-11-26 07-34-18 -- 2008-11-26 07-41-10 -#- New Anita Dingo Missed Steve On Linda On.wav" For Binary As #1
qq2 = LOF(1)
Close #1
Open "M:\05 Media\Digital Wave Player M Drive\DR2008\DWR 2008-11-26 09-39-09 -- 2008-11-27 00-35-50 -#- Steve Ibrag Linda Itching For Attack But Lets Clearly See Who Is Attacking First.wav" For Binary As #1
'qq2 = LOF(1)
JJ$ = Space$(&H86)
Get #1, 1, JJ$
Close #1

HH = &H78108323
HH = &H457894
gg1 = Hex$(QQ1 - 8)
gg2 = Hex$(qq2 - 8)
'4,560,000

Mid(LL, 1, 1) = Chr(&H78)
Mid(LL, 2, 1) = Chr(&H94)
Mid(LL, 3, 1) = Chr(&H45)
Mid(LL, 4, 1) = Chr(&H0)

Mid(LL, 1, 1) = Chr(&H78)
Mid(LL, 2, 1) = Chr(&H60)
Mid(LL, 3, 1) = Chr(&H8C)
Mid(LL, 4, 1) = Chr(&H56)

'QQ1 = QQ1 * 1.03025 '37.26,11 - correct from time files

'QQ1 = QQ1 * 1.032

QQ1 = QQ1 - 29000 '* 1.0000001 '36:20:11 - correct run lenght

gg3 = Hex$(Int(QQ1))
LL = ""
lDataSize = Len(gg3)
Dim i As Integer
I2 = 1
For i = lDataSize - 1 To 1 Step -2
'For i = 1 To lDataSize Step 2
    'bytearray(i2) = Val("&h" + Mid$(gg3, i, 2))
    LL = LL + Chr(Val("&h" + Mid$(gg3, i, 2)))
    I2 = I2 + 1
Next



'T1 = DateSerial(bytearray(1), bytearray(2), bytearray(3),bytearray(4))

T1 = DateDiff("h", DateSerial(2008, 11, 27) + TimeSerial(0, 36, 3), DateSerial(2008, 11, 28) + TimeSerial(13, 2, 14))
T2 = DateDiff("n", DateSerial(2008, 11, 27) + TimeSerial(0, 36, 3), DateSerial(2008, 11, 28) + TimeSerial(13, 2, 14)) Mod 60
T3 = DateDiff("s", DateSerial(2008, 11, 27) + TimeSerial(0, 36, 3), DateSerial(2008, 11, 28) + TimeSerial(13, 2, 14)) Mod 60

A1 = Asc(Mid$(LL, 1, 1))
'a2 = Asc(Mid$(ll, 2, 1))
'a3 = Asc(Mid$(ll, 3, 1))
'a4 = Asc(Mid$(ll, 4, 1))
getb$ = " "
Open "m:\DWR 2008-11-27 00-36-03 -- 2008-11-28 13-02-14 -#- Janice Paul Short.wav" For Binary As #1
'll$ = "TTTT"
Put #1, 1, JJ$
Put #1, 5, LL
Put #1, &H7D, LL
'Get #1, LOF(1), getb$
'Put #1, LOF(1), Chr$(0)
Close #1

a = Hex$(Asc(getb$))
a = Len(LL)

Shell "C:\Program Files\Windows Media Player\mplayer2.exe ""m:\DWR 2008-11-27 00-36-03 -- 2008-11-28 13-02-14 -#- Janice Paul Short.wav""", vbNormalFocus

End
End Sub

Sub Test3()
Exit Sub
Form1.Hide
ScanPath.Show
ScanPath.txtPath.Text = "M:\04 Music ---\04 My Music\01 Banging Tunes"
ScanPath.txtPath.Text = "M:\04 Music ---\03 My Music Zen\01 Banging Tunes"
ScanPath.cmdScanDir_Click

For we = 1 To Form1.ListView2.ListItems.Count
    A1 = Form1.ListView2.ListItems.Item(we).SubItems(1)
    B1 = Form1.ListView2.ListItems.Item(we)
    A1 = A1 + B1
    
    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    
    et = InStr(A1, "\200")
    If et > 0 Then
    'If et = Len(A1) - 4 Then
    'ef = InStrRev(A1, "\", et - 1)
    ef1 = InStr(et + 1, A1, "\")
    ef2 = InStr(ef1 + 1, A1, "\")
    If ef1 > 0 And ef2 = 0 Then
    'If Len(A1) > Len(ScanPath.txtPath.Text) + 1 And InStr(Len(ScanPath.txtPath.Text) + 2, A1, "\") > 0 Then
    
    tg1 = Mid$(A1, ef1 + 1)
    tg2 = Mid$(A1, et + 1, 4)
    tg3 = Mid$(A1, 1, et + 5) + Mid$(tg1, 1, 7) + tg2 + " - " + Mid$(tg1, 8)
    'ef = InStr(et + 6, A1, "\")
    'If ef = 0 Then
    ty = Mid$(A1, 1, et - 1)
    'tx = Mid$(A1, 1, Len(ScanPath.txtPath.Text) + 12) + " " + tg1 + " - " + Mid$(A1, Len(ScanPath.txtPath.Text) + 14)
    'ty = A1 + "\" + Mid$(A1,et+5 , 5)
    On Error GoTo 0
'    MkDir ty
    Name A1 As tg3
'    MkDir tg3
    
    End If
    
    Set FS = CreateObject("Scripting.FileSystemObject")
    'ScanPath.txtPath.Text = A1
    'ScanPath.cmdScan_Click
    A1 = tg3

    Set F = FS.GetFolder(A1)
        Set fc = F.Files
        For Each f1 In fc
            's = s & f1.Name
            's = s & vbCrLf
    '    Next
    
    
    'For we2 = 1 To ScanPath.ListView1.ListItems.Count
    '    A1 = ScanPath.ListView1.ListItems.Item(we2).SubItems(1)
    '    B1 = ScanPath.ListView1.ListItems.Item(we2)
    
    
    Set f1 = FS.GetFile(A1 + "\" + f1.Name)
'    f1.Move tg3 + "\" + f1.Name
    Next
    'F.Move tg3
    Set fc = F.SubFolders
        For Each f1 In fc
            Set f1 = FS.GetFolder(A1 + "\" + f1.Name)
'            f1.Move tg3 + "\" + f1.Name
        
        Next
    
    'RmDir A1
    
    'End If
    End If
Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub




Sub FindTimeInfo()

'InPutDate = DateValue("01-12-2008")
'TestDate = DateValue("31-05-2009")

Year5 = 0
For r5 = year(InputDate) + 1 To year(TestDate) + 1
    If DateSerial(r5, month(InputDate), day(InputDate)) < Now Then Year5 = Year5 + 1
Next
For r5 = 1 To -2 Step -1
    xx = DateSerial(year(TestDate), month(TestDate) + r5, day(InputDate))
    MonthStart = xx
    If xx <= TestDate Then Exit For
Next
Month5 = 0
Month7 = 0
For r5 = 1 To 10000
    xx = DateSerial(year(InputDate), month(InputDate) + r5, day(InputDate))
    If year(xx) <> oxx Then Month7 = 0
    oxx = year(xx)
    If xx <= TestDate + 1 Then Month5 = Month5 + 1: Month7 = Month7 + 1
    If xx >= TestDate + 1 Then Exit For
Next

For r5 = year(TestDate) To 1 Step -1
    If DateSerial(r5, month(InputDate), day(InputDate)) < Now Then
        DayCount1 = DateDiff("d", DateSerial(r5, month(InputDate), day(InputDate)), TestDate)
        DayCountMonth = DateDiff("d", MonthStart, TestDate)
        WholeYear1 = DateDiff("d", DateSerial(r5, month(InputDate), day(InputDate)), DateSerial(r5 + 1, month(InputDate), day(InputDate)))
        'Month5 = DateDiff("m", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate)
        WeeksSinceYear = DateDiff("w", DateSerial(r5, month(InputDate), day(InputDate)), TestDate) - 1
        WeeksSinceInput = DateDiff("w", InputDate, TestDate) - 1
        ResultYearDate = Format$(Year5 + DayCount1 / WholeYear1, ".000")
        Exit For
    End If
Next

End Sub



Private Sub Label3_Click()

R = FindWindow(vbNullString, "Find Message")
Call SetForegroundWindow(R)

For R = 1 To Val(Text1)

    SendKeys "^{DOWN}", True
DoEvents
    SendKeys "^{DOWN}", True
DoEvents
    SendKeys "^ ", True
DoEvents

Next

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print Now; Str(Timer)

End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Debug.Print Now; str(Timer)

End Sub

Public Sub Process_Idle(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    SetPriorityClass hProcess, NORMAL_PRIORITY_CLASS
    
End Sub

Private Sub Picture1_Click()
End
End Sub

Private Sub RichTextBox1_Change()

End Sub

Private Sub RTB1_Change()
RTB1.LoadFile "C:\Program Files\WordNet\2.0\DICT\DICTIO2.TXT"



End Sub

Private Sub Text2_Click()
Text2.Visible = False
RTB1.Left = 0
RTB1.Top = 0 'Text1.Top + Text1.Height + 15
RTB1.Width = Me.Width
RTB1.Height = Me.Height
End Sub

Private Sub Timer1_Timer()
Text2 = str(GG)
If GG = 0 Then GG = 1
RTB1.SelStart = GG
GG = GG + 10
End Sub

Private Sub Timer2_Timer()

End Sub
