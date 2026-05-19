VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WinRar_Notes_VB6"
   ClientHeight    =   5664
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   9636
   Icon            =   "DSS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5664
   ScaleWidth      =   9636
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3495
      Top             =   4110
   End
   Begin VB.FileListBox File1 
      Height          =   1416
      Hidden          =   -1  'True
      Left            =   240
      Pattern         =   "*.dss"
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.ListBox List1 
      Height          =   2928
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9615
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Dim LK1, LK2
Dim T1, F1, E$, Orig$, DSS1$, DwPRe, TimerCount, NotFolder, ExcudeSpec, Clue, Icute As Long

Private Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long  'MODULE 1141
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142
Private Const GW_HWNDNEXT = 2

Dim LoveFolder
Dim RarFileSpec
Dim Pass, Exe4, Path, Path1
Dim HackCut1, HackCut2, Hack, FileSpec, FlagEnd, FlagSize
Dim DayFlag, ITime$, IvTime As Date, ISW



Private Sub Form_Load()

    If App.PrevInstance = True Then End

    Set Fs = CreateObject("Scripting.FileSystemObject")

    If Fs.DRIVEEXISTS("m") = False Then End
    Dim d
    Set d = Fs.GetDrive("M:\")
    If d.ISREADY = False Then End


    TimerCount = 1
    
    Form1.Height = 80
    Form1.Width = 4000
    
    
    Form1.WindowState = 1
    
End Sub


Sub Timer1_Timer()
    
If FlagEnd = True Then
    End
End If

If FindProcessEXERar = True Then Exit Sub
    
'---------------------
'For example, use switch -to15d to process files older than 15 days and -to2h30m to process files older than 2 hours 15 minutes.
'-to1d
    
'Switch -TB<date>  -  process files modified before the specified date       'Oldest Then Stuff
'Switch -TA<date>  -  process files modified after the specified date         'Newsest stuff
'Switch -TN<time>  -  process files newer than the specified time
'For example, use switch -tn15d to process files newer than 15 days and -tn2h30m to process files newer than 2 hours 30 minutes.
    
'Switch -TO<time>  -  process files older than the specified time

'Switch -RI<p>[:<s>]  -  set priority and sleep time
'If <p> is 0, WinRAR uses the default task priority. <p> equal to 1 sets the lowest possible priority, 15 - the highest possible.
'Sleep time <s> is a value from 0 to 1000 (milliseconds). This is a period of time that WinRAR gives back to system after every
'read or write operation while compressing or extracting. Non-zero <s> may be useful if you need to reduce system load even
'more than can be achieved with <p> parameter.
'execute WinRAR with default priority and 10 ms sleep time:
'WinRAR a -ri0:10 backup *.*
    
'Switch -X@<listfile>  -  exclude files using a specified list file
'Exclude files that are too big
'WinRAR a -x@list.txt bin *.exe
    
'-s Switch -S  -  create solid archive
'-ep Switch -EP  -  exclude paths from names
'-ep1 Exculde base folder from names
'-inul no error msg
'-cfg- no profile usage
'-ibck in background
'-Hp encrypt filenames also
'-y answer yes to all dont stop
'-df  'Del afterwards'
'-x*\temp\*
'-x*\temp\*
On Error Resume Next
Open App.Path + "\TimeLogg.txt" For Input As #1
Line Input #1, ITime$
Close #1
On Error GoTo 0
If ITime$ <> "" Then
    IvTime = DateValue(ITime$)
End If

If IvTime < Int(Now) Then DayFlag = True
DayFlag = True
Open App.Path + "\TimeLogg.txt" For Output As #1
Print #1, Now
Close #1

Path1 = "M:\WinRar Archives\"

FlagSize = False
NotFolder = ""

If TimerCount = 1 Then
    ExcudeSpec = "-x*\Recycled\* -x""*\System Volume Information\*"""
    Path = "MY DOCS"
    RarFileSpec = Path1 + "WinRar Archive - MY DOCS"
    'LoveFolder = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\*.*"
    LoveFolder = "D:\# MY DOCS\*.*" 'Better
End If
    
If TimerCount = 2 Then
    ExcudeSpec = "-x*\VBDataNoArchive\* -x*\0TextData\* -x*\00_Text_Data*\*"
    Path = "VB6 Files"
    RarFileSpec = Path1 + "WinRar Archive - VB6"
    LoveFolder = "D:\VB6\*.*"
   FlagEnd = True
End If


If TimerCount = 3 And 1 = 2 Then
    ExcudeSpec = ""
    Path = "MatthewLan.Com"
    RarFileSpec = Path1 + "WinRar Archive - MatthewLan.Com"
    LoveFolder = "D:\MY WEBS\MatthewLan.Com Web\*.*"
End If

'TimerCount=TimerCount+1
If TimerCount = 4 And 1 = 2 Then
    ExcudeSpec = "-x*\Recycled\* -x""*\System Volume Information\*"""
    Path = "True 5G "
    RarFileSpec = Path1 + "WinRar Archive - True Crypt 5G X"
    'NotFolder = "-x""*\True Back*\*"" -x""*\# 01 My Documents\*"""
    'True Back W-Drive Docs not not be able to do wild card -x""*\True Back*\*""
    LoveFolder = "X:\*.*"
End If

'If TimerCount = 5 And DayFlag = False Then TimerCount = TimerCount + 1
If TimerCount = 5 And DayFlag = True Then
    
    Open App.Path + "\FEED-DEMION-DAYS-SINCE.TXT" For Input As #1
    Line Input #1, LK1
    Close #1
    LK2 = DateValue(LK1)
    
    ExcudeSpec = ""
    Path = "Feed-Demon3"
    RarFileSpec = Path1 + "WinRar Archive - Feed-Demon3"
    'NotFolder = "-x""*\True Back*\*"" -x""*\# 01 My Documents\*"""
    'True Back W-Drive Docs not not be able to do wild card -x""*\True Back*\*""
    LoveFolder = "D:\Feed-Demon3\*.*"

    Open App.Path + "\FEED-DEMION-DAYS-SINCE.TXT" For Output As #1
    Print #1, Int(Now)
    Close #1

End If

If TimerCount = 6 And 1 = 2 Then
    ExcudeSpec = ""
    Path = "UDisk 8GB"
    RarFileSpec = Path1 + "WinRar Archive - UDisk 8GB"
    'NotFolder = "-x""*\True Back*\*"" -x""*\# 01 My Documents\*"""
    'True Back W-Drive Docs not not be able to do wild card -x""*\True Back*\*""
    LoveFolder = "S:\*.*"
    FlagEnd = True
End If


Dim ZipRar
ZipRar = ".rar"

If LoveFolder <> "" Then
    
    RarFileSpec = RarFileSpec + "\" + Format$(Now, "YYYY") + "\" + Format$(Now, "MM MMM") + "\"
    Hack = RarFileSpec
    cute = 4
    Do
        HackCut1 = InStr(cute, Hack, "\")
        If HackCut1 = 0 Then Exit Do
        HackCut2 = Mid(Hack, 1, HackCut1 - 1)
        On Error Resume Next
            MkDir HackCut2
        On Error GoTo 0
        cute = HackCut1 + 1
    Loop
    FileSpec = Path
    
    Pass = LCase(" ") + " "
    
    TestRun = False
    'If TimerCount = 4 Or TimerCount = 1 Then TestRun = True
'    ISW = " -tn48h "
    If DayFlag = True And TestRun = False And TimerCount = 5 Then
        RarFileSpec = RarFileSpec + FileSpec + " # " + Format$(Now, "YYYY-MM-DD @ HH-MM") + " - Day Mode" + ZipRar
        ISW = " -tn48h "

        If LK2 > 0 Then
            ISW = " -tn" + Trim(str(Int(Now) - LK2)) + "d"
        End If
        If LK2 = 0 Then
            ISW = " -tn10d"
        End If
    End If
    
    If DayFlag = False And TestRun = False Then
        ListFilesTooBig
        RarFileSpec = RarFileSpec + FileSpec + " # " + Format$(Now, "YYYY-MM-DD @ HH-MM") + ZipRar
        ISW = "-tn15h "
        ISW = "-x@""" + App.Path + "\RarExcludeList.txt"" " + ISW
    End If
   If ISW = "" = True Then ISW = "-tn15h ":    ISW = "-x@""" + App.Path + "\RarExcludeList.txt"" " + ISW
    If TestRun = True Then
        RarFileSpec = RarFileSpec + FileSpec + " # " + Format$(Now, "YYYY-MM-DD @ HH-MM") + " - Test Mode Whole Lot" + ZipRar
        ISW = " "
    End If
    
    speed = "-ri1:400 "
'    speed = " "
    If IsIDE = True Then speed = " "
    Dim SW
    If IsIDE = False Then SW = " -ibck -inul -y " + speed
    SW = SW + NotFolder + " "
    
    'ISW how many days hours
    Exe4 = "C:\Program Files\WinRAR\winrar.exe A -cfg- -av " + ISW + ExcudeSpec + SW + " -m5 -ep -s -ep1 -r -p" + Pass + " """ + RarFileSpec + """ """ + LoveFolder + """"
    Icute = Shell(Exe4, vbMinimizedNoFocus) ', vbNormalNoFocus)
    Clue = True

    
    If FlagEnd = True Then
        End
    End If
End If

TimerCount = TimerCount + 1

End Sub


Sub ListFilesTooBig()

'ScanPath.Show

ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath.Text = Mid$(LoveFolder, 1, Len(LoveFolder) - 4)
ScanPath.cboMask = "*.*"
ScanPath.cboSize.ListIndex = 2 ' -- Select Files Greater than
ScanPath.cboSizeType(0).ListIndex = 2 'select Mega Bytes
ScanPath.txtSize(0).Text = "2" ' ---- For our List Not Allowed over 5 Meg

Call ScanPath.cmdScan_Click

Open App.Path + "\RarExcludeList.txt " For Output As #1
    For we = 1 To ScanPath.ListView1.ListItems.Count
        a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(we)
        Print #1, a1$ + B1$
    Next
Close #1



DoEvents
'Timer1.Enabled = False
End Sub


Function Nice_Names01(StringToDo As String) As String

'---------------------------------------------------------------
'Nice Name one string


E$ = StringToDo

Tu$ = LCase$(E$)
Orig$ = E$
For r = 2 To Len(E$)
    XC = 0
    If InStr(" _-{[]}!'@;#~><,\/1234567890", Mid$(E$, r - 1, 1)) > 0 Then XC = 1
    If InStr(":\", Mid$(E$, r - 1, 1)) > 0 Then XC = 0
    If Mid$(E$, r - 1, 1) = "_" Then XC = 1
    
    If XC = 1 Then Mid$(Tu$, r, 1) = UCase(Mid$(Tu$, r, 1))
Next
Mid$(Tu$, 1, 1) = UCase(Mid$(Tu$, 1, 1))

For r = 3 To Len(E$)
    XC = 0
    If InStr(":?", Mid$(E$, r, 1)) > 0 Then XC = 1
    If XC = 1 Then
        Mid$(Tu$, r, 1) = "-"
    End If
Next
dr = InStrRev(Tu$, "\") + 1

If Mid$(Tu$, dr, 3) = "dwr" Then Mid$(Tu$, dr, 3) = "DWR"
If Mid$(Tu$, dr, 2) = "ds" Then
    Mid$(Tu$, dr, 2) = "DS"
End If

E$ = Tu$
If InStrRev(E$, ".") > 0 Then
ff$ = Trim(Mid$(E$, 1, InStrRev(E$, ".") - 1)) + Mid$(E$, Len(E$) - 3)
E$ = ff$
End If
dr = InStrRev(Orig$, "\") + 1

If Mid$(Orig$, dr) <> Mid$(E$, dr) Then
On Error Resume Next
'Name Orig$ As E$

End If

Nice_Names01 = E$

End Function

Sub Nice_Names()

'---------------------------------------------------------------
'Nice Names

For we = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    ScanPath.lblCount = ScanPath.ListView1.ListItems.Count - we

    E$ = a1$ + B1$

Tu$ = LCase$(E$)
Orig$ = E$
For r = 2 To Len(E$)
    XC = 0
    If InStr(" _-{[]}!'@;#~><,\/1234567890", Mid$(E$, r - 1, 1)) > 0 Then XC = 1
    If InStr(":\", Mid$(E$, r - 1, 1)) > 0 Then XC = 0
    If Mid$(E$, r - 1, 1) = "_" Then XC = 1
    
    If XC = 1 Then Mid$(Tu$, r, 1) = UCase(Mid$(Tu$, r, 1))
Next
Mid$(Tu$, 1, 1) = UCase(Mid$(Tu$, 1, 1))

For r = 3 To Len(E$)
    XC = 0
    If InStr(":?", Mid$(E$, r, 1)) > 0 Then XC = 1
    If XC = 1 Then
        Mid$(Tu$, r, 1) = "-"
    End If
Next
dr = InStrRev(Tu$, "\") + 1

If Mid$(Tu$, dr, 3) = "dwr" Then Mid$(Tu$, dr, 3) = "DWR"
If Mid$(Tu$, dr, 2) = "ds" Then
    Mid$(Tu$, dr, 2) = "DS"
End If

E$ = Tu$

ff$ = Trim(Mid$(E$, 1, InStrRev(E$, ".") - 1)) + Mid$(E$, Len(E$) - 3)
E$ = ff$
dr = InStrRev(Orig$, "\") + 1

If Mid$(Orig$, dr) <> Mid$(E$, dr) Then
On Error Resume Next
Name Orig$ As E$

Pleasure$ = Mid$(E$, dr)
ScanPath.ListView1.ListItems.Item(we) = Pleasure$

End If

'ScanPath.ListView1.ListItems.Remove (1)

Next

End Sub


Sub DWP_ReName_Move()

Dim Idate As Date
DwPRe = True
ScanPath.ListView1.ListItems.Clear
'Nice_Names_DWP
ScanPath.cboMask = "*.*"
ScanPath.txtPath.Text = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files"
Call ScanPath.cmdScan_Click
'ScanPath.txtPath.Text = "D:\05 Media\"
'Call ScanPath.cmdScan_Click
'Call Nice_Names




Dim Fs, Fs2, F
Set Fs = CreateObject("Scripting.FileSystemObject")
Set Fs2 = CreateObject("Scripting.FileSystemObject")

If 1 = 2 Then
For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    If InStr(B1$, "DWR") = 0 Then
        ScanPath.ListView1.ListItems.Remove (we)
    End If
Next
End If

For we = 1 To ScanPath.ListView1.ListItems.Count
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)

    'Do
    'a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    'B1$ = ScanPath.ListView1.ListItems.Item(we)

    'Set F = Fs2.getfile(a1$ + B1$)
    'moddate = F.datelastmodified
    'If moddate <= Now - 10 Then we = we + 1
    'If we > ScanPath.ListView1.ListItems.Count Then Exit Sub
    'Loop Until moddate > Now - 10



FileSpec = a1$ + B1$
Set F = Fs.getfile((FileSpec))

moddate = F.datelastmodified
credate = F.datecreated

xx$ = Trim(Mid$(B1$, 9))
d1$ = a1$ + "DWR " + Format$(credate, "YYYY-MM-DD HH-MM-SS") + " -- " + Format$(moddate, "YYYY-MM-DD HH-MM-SS") + " -#- " + xx$
G1$ = "D:\05 Media\Digital Wave Player M Drive\DR2008\"
d2$ = G1$ + "DWR " + Format$(credate, "YYYY-MM-DD HH-MM-SS") + " -- " + Format$(moddate, "YYYY-MM-DD HH-MM-SS") + " -#- " + xx$


'this only run move if one need rename
tt = FileInUse(a1$ + B1$)
If tt = False Then
    Fs.MoveFile a1$ + B1$, d2$
    'Name a1$ + B1$ As d1$


'    t5$ = d2$
'    t3$ = Trim(d2$) + ".###.rar"

    'For example, use switch -to15d to process files older than 15 days and -to2h30m to process files older than 2 hours 15 minutes.
    '-to2d
'    DoEvents
'    pass$ = LCase(" ")
'    Shell "C:\Program Files\WinRAR\winrar.exe A -CFG- -av -m5 -ibck -ep -s -y -df -ep1 -r -hp" + pass$ + " """ + t3$ + """ """ + t5$ + """", vbNormalFocus
'    DoEvents
End If
Next

'ScanPath.ListView1.ListItems.Clear
'ScanPath.txtPath.Text = "O:\05 Media\Digital Wave Player M Drive"
'ScanPath.cboMask = "*.*"
''ScanPath.Show
'Call Nice_Names

ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask = "*.wav;*.mp3;*.amr;*.rar"
'ScanPath.txtPath.Text = "D:\05 Media\Digital Wave Player M Drive\"
ScanPath.txtPath.Text = "O:\05 Media\"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "D:\05 Media\"
Call ScanPath.cmdScan_Click
Call Nice_Names

Timer1.Enabled = True


End Sub

Sub DSS_Rename_Move()



Dim Fs, Fs2, F
Set Fs = CreateObject("Scripting.FileSystemObject")
Set Fs2 = CreateObject("Scripting.FileSystemObject")

Dim LDate As Date, MoveToFolder

ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask = "*.dss;*.amr;*.wav;*.mp3"
ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\00 My folders\2008\"
MoveToFolder = ScanPath.txtPath.Text
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\FolderA\"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\FolderB\"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\FolderC\"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\FolderD\"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\FolderE\"
Call ScanPath.cmdScan_Click

Call Nice_Names

Set Fs2 = CreateObject("Scripting.FileSystemObject")
'Call Date_File(a1$ + b1$, Idate)

'ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1


'if you want date eless
For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)

    Set F = Fs2.getfile(a1$ + B1$)
    moddate = F.datelastmodified
    If moddate <= Now - 30 Then
'        ScanPath.ListView1.ListItems.Remove (we)
    End If
Next


For we = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    ScanPath.lblCount = we
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    

    'Call Date_File(a1$ + b1$, Idate)



'e$ = "DS-330-" + Format$(we, "00000") + "-" + Format$(Idate, "YYYY-MM-DD---HH-MM-SS") + ".dss"
'E$ = "DS-330-" + Format$(Idate, "YYYY-MM-DD---HH-MM-SS") + ".dss"

tt = FileInUse(a1$ + B1$)
If tt = True Then
    a = a
End If

If tt = False Then
freef = FreeFile
Open a1$ + B1$ For Binary As #freef

freef = FreeFile
Open a1$ + B1$ For Binary As #freef
Seek #freef, &H26 + 1
Value = Input(&H3D - &H25, #freef)

lDataSize = Len(Value)
  ReDim ByteArray(lDataSize)
  Dim i As Integer
i2 = 1
For i = 1 To lDataSize Step 2
    ByteArray(i2) = Val(Mid$(Value, i, 2))
    i2 = i2 + 1
Next

T1 = DateSerial(ByteArray(1), ByteArray(2), ByteArray(3))
T1 = T1 + TimeSerial(ByteArray(4), ByteArray(5), ByteArray(6))
F1 = DateSerial(ByteArray(7), ByteArray(8), ByteArray(9))
F1 = F1 + TimeSerial(ByteArray(10), ByteArray(11), ByteArray(12))
E$ = B1$
'Nice_Names01(StringToDo As String)

ext = LCase(Mid$(B1$, InStrRev(B1$, ".")))
If InStr("*.dss", ext) > 0 Then
    Seek #freef, &H31E + 1
    DSS1$ = Input(&H381 - &H31D, #freef)
    xs = InStr(DSS1$, Chr$(0))
    If xs > 0 Then DSS1$ = Mid$(DSS1$, 1, xs - 1)
    Close #freef
    rename = False
    If LTrim$(DSS1$) <> "" Then
        DSS1$ = Nice_Names01(DSS1$)
        'e$ = "DS-330-" + Format$(we, "00000") + "-" + Format$(Idate, "YYYY-MM-DD---HH-MM-SS") + " #- " + dss$ + ".dss"
        E$ = "DS-330-" + Format$(T1, "YYYY-MM-DD HH-MM-SS") + " -- " + Format$(F1, "YYYY-MM-DD HH-MM-SS") + " -#- " + DSS1$ + ".dss"
        rename = True
    End If
End If
Close #freef
End If


If E$ = B1$ Then rename = False

tt = FileInUse(a1$ + B1$)

If rename = True And tt = False Then
    If ext = ".dss" Then
        If rename = True Then Name a1$ + B1$ As a1$ + E$
        If a1$ <> MoveToFolder Then Fs.MoveFile a1$ + E$, MoveToFolder + E$
    
        B1$ = ScanPath.ListView1.ListItems.Item(we)
        d1$ = Mid$(B1$, 1, InStrRev(B1$, ".") - 1) + ".wav"
        d2$ = Mid$(B1$, 1, InStrRev(B1$, ".") - 1) + ".mp3"
        d3$ = Mid$(B1$, 1, InStrRev(B1$, ".") - 1) + ".amr"
        If Dir$(a1$ + d1$) <> "" Then
            g$ = Mid$(E$, 1, InStrRev(E$, ".") - 1) + ".wav"
            Name a1$ + d1$ As a1$ + g$: E = g$
            If a1$ <> MoveToFolder Then Fs.MoveFile a1$ + E$, MoveToFolder + E$
        End If
        If Dir$(a1$ + d2$) <> "" Then
            g$ = Mid$(E$, 1, InStrRev(E$, ".") - 1) + ".wav"
            Name a1$ + d2$ As a1$ + g$: E = g$
            If a1$ <> MoveToFolder Then Fs.MoveFile a1$ + E$, MoveToFolder + E$
        End If
        If Dir$(a1$ + d3$) <> "" Then
            g$ = Mid$(E$, 1, InStrRev(E$, ".") - 1) + ".wav"
            Name a1$ + d3$ As a1$ + g$: E = g$
            If a1$ <> MoveToFolder Then Fs.MoveFile a1$ + E$, MoveToFolder + E$
        End If
    End If
End If

tt = FileInUse(a1$ + B1$)
If rename = True And tt = False Then
End If


DoEvents
Next
ScanPath.ListView1.ListItems.Clear


ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask = "*.dss;*.amr;*.wav;*.mp3"
ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\00 My folders\2008\"
MoveToFolder = ScanPath.txtPath.Text
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\FolderA\"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\FolderB\"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\FolderC\"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\FolderD\"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\FolderE\"
Call ScanPath.cmdScan_Click

Timer1.Enabled = True

End Sub



Sub OldNiceNameCode()
'---------------------------------------------------------------
'Nice Names
Tu$ = LCase$(E$)
For r = 2 To Len(E$)
    XC = 0
    If InStr(" _-{[]}!'@;#~><,\/1234567890", Mid$(E$, r - 1, 1)) > 0 Then XC = 1
    If Mid$(E$, r - 1, 1) = "_" Then XC = 1
    If XC = 1 Then Mid$(Tu$, r, 1) = UCase(Mid$(Tu$, r, 1))
Next
Mid$(Tu$, 1, 1) = UCase(Mid$(Tu$, 1, 1))

For r = 1 To Len(E$)
    XC = 0
    If InStr(":?", Mid$(E$, r, 1)) > 0 Then XC = 1
    If XC = 1 Then
        Mid$(Tu$, r, 1) = "-"
    End If
Next

E$ = Tu$

If a1$ + B1$ <> a1$ + E$ And LTrim$(DSS1$) <> "" Then

Name a1$ + B1$ As a1$ + E$

mp3$ = "Ds-" + Mid$(B1$, 14)
mp3$ = Mid$(mp3$, 1, InStrRev(mp3$, ".")) + "mp3"
If Dir$(a1$ + mp3$) <> "" Then
    mp3t$ = Mid$(E$, 1, InStrRev(E$, ".")) + "mp3"
    Name a1$ + mp3$ As a1$ + mp3t$
End If

wav$ = B1$
wav$ = Mid$(wav$, 1, InStrRev(wav$, ".")) + "wav"
If Dir$(a1$ + wav) <> "" Then
    wavt$ = Mid$(E$, 1, InStrRev(E$, ".")) + "wav"
    Name a1$ + wav$ As a1$ + wavt$
End If

End If

End Sub

Private Sub Date_File(Filespec2$, Idate As Date)
'Call Date_File(filespec2$, Idate)

Dim Fs, F
Set Fs = CreateObject("Scripting.FileSystemObject")

pdate = 0
FileSpec = Filespec2$
Set F = Fs.getfile((FileSpec))
Idate = F.datelastmodified

End Sub


Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
'End
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





Function FindProcessEXERar() As Boolean
Dim cProcesses As New clsCnProc
Icute = -1
FindProcessEXERar = cProcesses.GetEXEID(Icute, "WinRAR.exe")
End Function


Function FindWinPartRar() As Long

Dim Ash$
Dim Test_hwnd As Long, _
    Test_pid As Long, _
    Test_thread_id As Long

Dim NeroArray(30)
            
Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME
'an another one
Test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)

FindWinPartRar = False
    
NeroDupe$ = ""
    
neros = 0
    
Do While Test_hwnd <> 0
    
    
    Ash$ = GetWindowTitle(Test_hwnd)
    cText = Space$(255)
    ghj$ = GetClassName(Test_hwnd, cText, 255)
    'StripNulls ghj$
    
    ff = 0
    If InStr(Ash$, "Preparing Files") Then ff = 1: Stop
    If InStr(Ash$, "Updating archive") Then ff = 1
    If InStr(Ash$, "Updating archive") Then ff = 1
    If InStr(Ash$, "Creating archive") Then ff = 1
    If InStr(Ash$, "Extracting from") Then ff = 1
    If InStr(Ash$, "Paused, press ""Continue"" to resume") Then ff = 1
  
'    If ff = 1 Then Stop
    
    If ff = 1 And InStr(cText, "#32770") > 0 Then
        Clue = False
        FindWindowPart = Test_hwnd
        
        
        neros = neros + 1
'        NeroArray(neros) = ash$
'        If neros > 1 Then
'            For R = 1 To neros
'                If NeroArray(R) = ash$ And R <> neros And ash$ <> "Nero Wave Editor" Then
'                    NeroDupe$ = ash$
'                    EXITDUPE = 1
'                Exit For
'                End If
'            Next
'        End If
        
        'If EXITDUPE = 1 Then Exit Do
    End If
    
'retrieve the next window
Test_hwnd = GetWindow(Test_hwnd, GW_HWNDNEXT)

Loop

FindWinPartRar = neros

End Function


Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hWnd)
   S = Space(l + 1)
   GetWindowText hWnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function

Public Function GetFileFromProc(lngProcess) As String

'MsgBox getfilefromhwnd(Me.hwnd)
'Dim lngProcess&, hProcess&, bla&, C&
Dim hProcess&, bla&, C&
Dim strFile As String
Dim x

strFile = String$(256, 0)
'x = GetWindowThreadProcessId(lngHwnd, lngProcess)
'hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
'x = EnumProcessModules(hProcess, bla, 4&, C)
'C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromProc = Left(strFile, C)

End Function

Public Function GetProcExist(lngProcess) As Long
If lngProcess = 0 Then Exit Function
'MsgBox getfilefromhwnd(Me.hwnd)
'Dim lngProcess&, hProcess&, bla&, C&
Dim hProcess&, bla&, C&
Dim strFile As String
Dim x

strFile = String$(256, 0)
'x = GetWindowThreadProcessId(lngHwnd, lngProcess)
'hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
'x = EnumProcessModules(hProcess, bla, 4&, C)
'C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))

GetProcExist = Left(strFile, C) <> ""

End Function




'***********************************************
'# Check, whether we are in the IDE
Public Function IsIDE() As Boolean
'  If IsIDE Then

  Debug.Assert Not TestIDE(IsIDE)
'End Function

'# In the IDE, the forwarding of
'# keystrokes will be suppressed
If IsIDE Then

End If

End Function



Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function

