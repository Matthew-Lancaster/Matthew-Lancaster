VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DSS"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3495
      Top             =   4110
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Hidden          =   -1  'True
      Left            =   240
      Pattern         =   "*.dss"
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.ListBox List1 
      Height          =   2985
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

Dim T1, F1, E$, Orig$, DSS1$, DwPRe, NewerThan, NotIfDSS
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

Sub FixRenameERrors()

ScanPath.ListView1.ListItems.Clear

ScanPath.cboMask = "*.wav"
ScanPath.txtPath.Text = "D:\05 Media\02 Olympus VN-2100PC\DR2009"
Call ScanPath.cmdScan_Click
ScanPath.Show

For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    aa1 = InStr(B1$, "-#-")
    aa2 = InStrRev(B1$, "-#-")
    If aa1 <> aa2 Then
        tt$ = Mid$(B1$, 1, aa1 - 1) + Mid$(B1$, aa2)
        Name a1$ + B1$ As a1$ + tt$
    End If
Next
Exit Sub



End Sub

Private Sub Form_Load()




If App.PrevInstance = True Then End

ScanPath.Show
DoEvents
Me.WindowState = 1
ScanPath.WindowState = 1

NewerThan = 30 '30 days back
NewerThan = DateDiff("d", DateSerial(2008, 10, 1), Now)


End

Call DSS_Rename_Move


'End

End Sub

Sub DWP_ReName_Move()







Dim Idate As Date
DwPRe = True
ScanPath.ListView1.ListItems.Clear
'Nice_Names_DWP
ScanPath.cboMask = "*.wav" ';*.rar;*.mp3;*.amr"
'ScanPath.cboMask = "*.wav;*.rar;*.mp3;*.amr"
ScanPath.cboDate.ListIndex = 0
ScanPath.DTPicker1(0).Value = Now - NewerThan ' any newer than this date

ScanPath.txtPath.Text = "D:\# MY DOCS\# 01 My Documents\Digital Wave Player\"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "D:\05 Media\"
Call ScanPath.cmdScan_Click
Call Nice_Names

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
    DoEvents
    ScanPath.lblCount = ScanPath.ListView1.ListItems.Count - we
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



Filespec = a1$ + B1$
Set F = Fs.getfile((Filespec))

moddate = F.datelastmodified
credate = F.datecreated


xx$ = Trim(Mid$(B1$, 9))

gocrazy = 0
If InStr(B1$, "DWR") = 0 Then
    d1$ = a1$ + "DWR " + Format$(credate, "YYYY-MM-DD HH-MM-SS") + " -- " + Format$(moddate, "YYYY-MM-DD HH-MM-SS") + " -#- " + xx$
    G1$ = "D:\05 Media\Digital Wave Player M Drive\DR2009\"
    d2$ = G1$ + "DWR " + Format$(credate, "YYYY-MM-DD HH-MM-SS") + " -- " + Format$(moddate, "YYYY-MM-DD HH-MM-SS") + " -#- " + xx$
    gocrazy = 1
End If

gg$ = ""
If gocrazy = 1 Then
    gg$ = Dir$(a1$ + Mid$(B1$, 1, 8) + "*.txt")
    If gg$ <> "" Then
        xx2$ = Trim(Mid$(gg$, 9))
        d3$ = G1$ + "DWR " + Format$(credate, "YYYY-MM-DD HH-MM-SS") + " -- " + Format$(moddate, "YYYY-MM-DD HH-MM-SS") + " -#- " + xx2$
        gg$ = a1$ + Dir$(a1$ + Mid$(B1$, 1, 8) + "*.txt")
    End If
End If


'this only run move if one need rename should be
tt = FileInUse(a1$ + B1$)
If tt = False And xx$ <> "" And gocrazy = 1 Then
    ScanPath.Refresh
    DoEvents
    On Error Resume Next
    
    
    Fs.MoveFile a1$ + B1$, d2$
    If gg$ <> "" Then Fs.MoveFile gg$, d3$
    On Error GoTo 0
    DoEvents
    
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

ScanPath.cboDate.ListIndex = 0
ScanPath.DTPicker1(0).Value = Now - NewerThan ' any newer than this date

ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask = "*.wav;*.mp3;*.amr" '';*.rar"
'ScanPath.txtPath.Text = "D:\05 Media\Digital Wave Player M Drive\"

'ScanPath.txtPath.Text = "O:\05 Media\"
'Call ScanPath.cmdScan_Click

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

'ScanPath.txtPath.Text = "O:\05 Media\DSSPlayer\Message\00 My folders\2008\"
ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask = "*.dss;*.amr;*.wav;*.mp3"

ScanPath.cboDate.ListIndex = 0
ScanPath.DTPicker1(0).Value = Now - NewerThan ' any newer than this date

ScanPath.txtPath.Text = "D:\05 Media\01 DSSPlayer Olympus DS-330\Message\"
MoveToFolder = "D:\05 Media\01 DSSPlayer Olympus DS-330\Message\00 My folders\2009\"
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
    If moddate <= Now - NewerThan Then
'        ScanPath.ListView1.ListItems.Remove (we)
    End If
Next


For we = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    ScanPath.lblCount = we
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    E$ = B1$
    

    'Call Date_File(a1$ + b1$, Idate)

'If InStr(B1$, "Test") > 0 Then Stop

'e$ = "DS-330-" + Format$(we, "00000") + "-" + Format$(Idate, "YYYY-MM-DD---HH-MM-SS") + ".dss"
'E$ = "DS-330-" + Format$(Idate, "YYYY-MM-DD---HH-MM-SS") + ".dss"

tt = FileInUse(a1$ + B1$)
If tt = True Then
    a = a
End If

ext = LCase(Mid$(B1$, InStrRev(B1$, ".")))

If tt = False And NotIfDSS = False And ext = ".dss" Then
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

'tt = FileInUse(a1$ + B1$)
'Reset

dss = "": amr = "": wav = "": mp3 = ""
If rename = True Then
    If ext = ".dss" Then
        If rename = True Then
            Name a1$ + B1$ As a1$ + E$
            dss = E$
        End If
    
        B1$ = ScanPath.ListView1.ListItems.Item(we)
        d1$ = Mid$(B1$, 1, InStrRev(B1$, ".") - 1) + ".wav"
        d2$ = Mid$(B1$, 1, InStrRev(B1$, ".") - 1) + ".mp3"
        d3$ = Mid$(B1$, 1, InStrRev(B1$, ".") - 1) + ".amr"
        If Dir$(a1$ + d1$) <> "" Then
            g$ = Mid$(E$, 1, InStrRev(E$, ".") - 1) + ".wav"
            Name a1$ + d1$ As a1$ + g$: e1 = g$
            wav = e1
            'If a1$ <> MoveToFolder Then Fs.MoveFile a1$ + d1$, MoveToFolder + e1
        End If
        If Dir$(a1$ + d2$) <> "" Then
            g$ = Mid$(E$, 1, InStrRev(E$, ".") - 1) + ".mp3"
            Name a1$ + d2$ As a1$ + g$: e1 = g$
            mp3 = e1
            'If a1$ <> MoveToFolder Then Fs.MoveFile a1$ + d2$, MoveToFolder + e1
        End If
        If Dir$(a1$ + d3$) <> "" Then
            g$ = Mid$(E$, 1, InStrRev(E$, ".") - 1) + ".amr"
            Name a1$ + d3$ As a1$ + g$: e1 = g$
            amr = e1
            'If a1$ <> MoveToFolder Then Fs.MoveFile a1$ + d3$, MoveToFolder + e1
        End If
    End If
End If

If a1$ <> MoveToFolder Then
    If dss <> "" And FileInUse(a1$ + dss) = False Then Fs.MoveFile a1$ + dss, MoveToFolder + dss
    If mp3 <> "" And FileInUse(a1$ + mp3) = False Then Fs.MoveFile a1$ + mp3, MoveToFolder + mp3
    If wav <> "" And FileInUse(a1$ + wav) = False Then Fs.MoveFile a1$ + wav, MoveToFolder + wav
    If amr <> "" And FileInUse(a1$ + amr) = False Then Fs.MoveFile a1$ + amr, MoveToFolder + amr
End If

If a1$ <> MoveToFolder And Mid$(B1$, 1, 4) <> "DS33" Then
    If FileInUse(a1$ + B1$) = False And Dir$(a1$ + B1$) <> "" Then
    ww = 0
    If Dir$(MoveToFolder + B1$) <> "" Then
    'MsgBox "File Already Exists" + vbCrLf + MoveToFolder + B1$+vbcrlf +"Del
        ww = 1
    End If
    On Error Resume Next
    Fs.MoveFile a1$ + B1$, MoveToFolder + B1$
    On Error GoTo 0
    End If
    If ww = 1 Then Kill a1$ + B1$
End If


DoEvents
Next
ScanPath.ListView1.ListItems.Clear

ScanPath.cboDate.ListIndex = 0
ScanPath.DTPicker1(0).Value = Now - NewerThan ' any newer than this date

ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask = "*.dss;*.amr;*.wav;*.mp3"
ScanPath.txtPath.Text = "D:\05 Media\01 DSSPlayer Olympus DS-330\Message\"
MoveToFolder = ScanPath.txtPath.Text

Timer1.Enabled = True

End Sub


Sub Timer1_Timer()
    
    Timer1.Interval = 1
    'If FindWinPartRar > 0 Then Exit Sub
    'Exit Sub
    If ScanPath.ListView1.ListItems.Count = 0 Then
            DoEvents
            ScanPath.lblCount = ScanPath.ListView1.ListItems.Count
            Timer1.Enabled = False
            If DwPRe = False Then
            Call DWP_ReName_Move
        Else
            End
            Exit Sub
        End If
    End If
    
    If ScanPath.ListView1.ListItems.Count = 0 Then
        End
        Exit Sub
    End If
    
    a1$ = ScanPath.ListView1.ListItems.Item(1).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(1)
    
    'For example, use switch -to15d to process files older than 15 days and -to2h30m to process files older than 2 hours 15 minutes.
    '-to2d
    
    If Mid$(B1$, 1, 3) = "DS3" Then
        ScanPath.lblCount = ScanPath.ListView1.ListItems.Count
        ScanPath.ListView1.ListItems.Remove (1)
    Exit Sub
    End If
    
    t3$ = ""
    If InStr(B1$, ".###.rar") = 0 Then
        t3$ = a1$ + Trim(B1$) + ".###.rar"
        t5$ = a1$ + B1$
        tt = FileInUse(t5$)
    End If
    On Error Resume Next
    wd = Dir$(t3$)
    'If Err.Number > 0 Then MsgBox "Error FileName Continue"
    
    If Dir$(t3$) <> "" Or tt = True Or t3$ = "" Or Err.Number > 0 Then
        ScanPath.lblCount = ScanPath.ListView1.ListItems.Count
        ScanPath.ListView1.ListItems.Remove (1)
    On Error GoTo 0
    Exit Sub
    End If

    On Error GoTo 0

    'older than 3 h "-to3h"
    '-ri0:10 slow it down
    '-ep1 Exculde base folder from names
    '-inul no error msg
    '-cfg- no profile usage
    '-ibck in background
    '-Hp encrypt filenames also
    
    
    pass$ = LCase(" ")
    
    exe4$ = "C:\Program Files\WinRAR\winrar.exe A -CFG- -av -m1 -ibck -ri1:100 -to12h -ep -s -y -df -ep1 -r -hp" + pass$ + " """ + t3$ + """ """ + t5$ + """"

    
    'No More Winrars
    'Shell exe4$, vbNormalFocus
    
    
    ScanPath.ListView1.ListItems.Remove (1)
    ScanPath.lblCount = ScanPath.ListView1.ListItems.Count

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
Filespec = Filespec2$
Set F = Fs.getfile((Filespec))
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
    
Do While Test_hwnd <> 0
    
    
    Ash$ = GetWindowTitle(Test_hwnd)
    cText = Space$(255)
    ghj$ = GetClassName(Test_hwnd, cText, 255)
    
    ff = 0
    If InStr(LCase(Ash$), "preparing file") Then ff = 1 ': Stop
    If InStr(Ash$, "Updating archive") Then ff = 1
    If InStr(Ash$, "Creating archive") Then ff = 1
    If InStr(Ash$, "Extracting from") Then ff = 1
    If InStr(Ash$, "Paused, press ""Continue"" to resume") Then ff = 1
  
    'If ff = 1 And InStr(cText, "#32770") > 0 Then Stop
    
    If ff = 1 And InStr(cText, "#32770") > 0 Then
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




