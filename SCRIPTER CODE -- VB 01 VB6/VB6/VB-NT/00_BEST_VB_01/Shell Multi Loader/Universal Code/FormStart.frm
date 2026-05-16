VERSION 5.00
Begin VB.Form FormStart 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FormStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public QQ$, LL$
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Sub FormStartLoader()


FontSizez = 10

'ScanPath.ListView1.ListItems.Clear



ScanPath.cboMask.Text = "*.*"
ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath.Text = "D:\# MY DOCS\# 01 My Documents\02_GOOGLE_EARTH"
Call ScanPath.cmdScan_Click




ScanPath.cboMask.Text = "*.*"
ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath.Text = "E:\01 VB Shell Folders\00 " + App.EXEName
Call ScanPath.cmdScan_Click


'ScanPath.cboMask.Text = "*.txt"
'ScanPath.chkSubFolders = vbChecked
'ScanPath.txtPath.Text = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files"
'Call ScanPath.cmdScan_Click

Exit Sub


dd = ScanPath.ListView1.ListItems.Count + 1


ScanPath.cboMask.Text = "*.txt"
ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath.Text = "D:\# MY DOCS\# 01 My Documents\00 Email Texts\Encrypt Sends"
Call ScanPath.cmdScan_Click

ww = 0
For r = ScanPath.ListView1.ListItems.Count To dd Step -1
    A1$ = ScanPath.ListView1.ListItems.Item(r).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(r)
    yy = 1
    If InStr(LCase(B1$), LCase("#Top Message")) > 0 Then
    yy = 0
    If ww = 1 Then
    q = q
    ScanPath.ListView1.ListItems.Remove (r)
    End If
    ww = 1
    End If
    If yy = 1 Then ScanPath.ListView1.ListItems.Remove (r)
Next

End Sub

Sub LabelClick(index)

If SetTrueToLoadLast = False Then
    A1$ = ScanPath.ListView1.ListItems.Item(index).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(index)
    C1$ = Form1.Label1(index).Caption
End If


'Call Form1.LoadLoggs

'Load File From Link
d1$ = A1$ + B1$
If InStr(LCase(B1$), ".lnk") > 0 Then 'And InStr(LCase(B1$), ".txt") Then
Open A1$ + B1$ For Binary As #1
rr$ = Space$(LOF(1))
Get #1, 1, rr$
Close #1
'If Mid$(rr$, &H568, 3) = ":" + Chr$(0) + "\" Then
'rr$ = Mid$(rr$, &H566)
'End If
tt$ = Mid$(rr$, InStr(105, rr$, ":\") - 1)
d1$ = Mid$(tt$, 1, InStr(tt$, Chr$(0)) - 1)
If InStr(d1$, "Blue Tooth Sync Folder") > 0 Then MsgBox ("Wrong Link Points to Blue Tooth folder"): End
End If

If LoadFolderFile = True And InStr(LCase(B1$), ".txt") Then
    d1$ = Mid$(tt$, 1, InStr(tt$, Chr$(0)) - 1)
    'd1$ = Mid$(d1$, 1, InStrRev(d1$, "\"))
    Shell "explorer /e, /select, " + d1$, vbNormalFocus
    End
End If

If InfoRapid = True Then
    d1$ = Mid$(tt$, 1, InStr(tt$, Chr$(0)) - 1)
    'd1$ = Mid$(d1$, 1, InStrRev(d1$, "\"))
    'Shell "explorer /e, /select, " + d1$, vbNormalFocus
    
    A1 = Mid$(d1$, 1, InStrRev(d1$, "\"))
    B1 = Mid$(d1$, 1 + InStrRev(d1$, "\"))
    RT = 0
    Do
        RT = InStr(A1, " ")
        If RT > 0 Then Mid$(A1, RT, 1) = "*"
    Loop Until RT = 0
    
    RT = 0
    Do
        RT = InStr(B1, " ")
        If RT > 0 Then Mid$(B1, RT, 1) = "*"
    Loop Until RT = 0
    
    Shell "C:\Program Files\seRapid\seRapid.exe /nologo /Dir" + A1 + " /File" + B1, vbNormalFocus
    End
End If




xxy = 1
'If A1$ + B1$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Hosso\Hosso Weather.txt" Then xxy = 1
'If A1$ + B1$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Hosso\Talk About In Front of Me.txt" Then xxy = 1
'If InStr(A1$ + B1$, "D:\# MY DOCS\# 01 My Documents\03 NotePad Files") Then xxy = 1

'If InStr(B1$, "Date1991") > 0 Then xxy = 0

If B1$ = "WinAmp Logger.exe.lnk" Then
d1$ = Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\WinAmp Logger\00 Music Info Logger.txt": xxy = 0
End If
xa = 0
If xa = 0 Then xa = FindWindow(vbNullString, "* " + d1$ + " - Notepad2")
If xa = 0 Then xa = FindWindow(vbNullString, d1$ + " - Notepad2")


If xa > 0 Then
    'MsgBox "Already Open" + vbCrLf + B1$
    SetForegroundWindow (xa)
    
    End
    Exit Sub
End If


If InStr(B1$, "Clued Up") Then
    Open d1$ For Binary As #1
    rr$ = Space$(LOF(1))
    Get #1, 1, rr$
    Close #1

    'all on the run hur rats

    ff$ = ">>-- " + Format$(Now, "DDD DD-MM-YY") + " --<<"
    If InStr(rr$, ff$) = 0 Then
        rr$ = rr$ + vbCrLf + ff$ + vbCrLf
        Open d1$ For Output As #1
        Print #1, rr$
        Close #1
    End If
End If


If xxy = 1 Then
'C1$ = Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\NotePad Abbrev\NotePad instr top LF 2abbrev txt.exe"
'd1$ = GetShortName(a1$ + b1$)
'd1$ = A1$ + B1$
QQ$ = d1$
'Call AbbrevCode


'Shell A1$ + B1$, vbNormalFocus
'vLaunch A1$ + B1$, vbNullString ', d1$

'vLaunch C1$, d1$
'End
End If

vLaunch A1$ + B1$, vbNullString ', d1$

End

'ends here --------------------

xa = 0
If xa = 0 Then xa = FindWindow(vbNullString, "* " + A1$ + B1$ + " - Notepad2")
If xa = 0 Then xa = FindWindow(vbNullString, A1$ + B1$ + " - Notepad2")






If xa > 0 Then
    'MsgBox "Already Open" + vbCrLf + B1$
    SetForegroundWindow (xa)
    
    End
    Exit Sub
End If


vLaunch A1$ + B1$, vbNullString

End



End Sub

Public Sub AbbrevCode()


If QQ$ = "" Then
QQ$ = "D:\# MY DOCS\# 01 My Documents\CALLerid\2ABBREV.TXT"
QQ$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Hosso\Hosso Weather.txt"
QQ$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\00 Opinion Clip.txt"
QQ$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Alt.SZ Members.txt"
QQ$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Hosso\Complaints.txt"
'QQ$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Hosso\Imposters.txt"
QQ$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\00 Opinion Clip.txt"
QQ$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Hosso\RutLand Hosso.txt"
QQ$ = "D:\# MY DOCS\# 01 My Documents\CALLerid\2ABBREV.TXT"
QQ$ = "D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\# DATA\#ClipBoard.Txt"
End If

'If Command$ <> "" Then QQ$ = Command$
'If Mid$(QQ$, 1, 1) = """" Then QQ$ = Mid$(QQ$, 2)
'MsgBox QQ$

Set fs22 = CreateObject("Scripting.FileSystemObject")
Set F = fs22.GetFile(QQ$)

scansize = (1024! * 30)

If F.Size > (1024! * 1000) Then
    'AddDateToTopLogg
    Exit Sub
End If

tt$ = ""

AddDateToTopLogg

'If F.Size > (1024! * 150) Then
'    AddDateToTopLogg
'    Exit Sub
'End If

If tt$ = "" Then
    Open QQ$ For Binary As #1
        tt$ = Input$(LOF(1), #1)
    Close #1
End If

tt2$ = LCase(tt$)

Do
    DoEvents
    vf = InStr(tt$, " i ")
    If vf = 0 Then vf = InStr(tt$, vbLf + "i ")
    If vf = 0 Then vf = InStr(tt$, vbCr + "i ")
    If vf > 0 Then
        Mid$(tt$, vf + 1, 1) = "I"
        cute3 = 1
    End If
Loop Until vf = 0 Or vf > scansize

If Capp1 = True Then
    vf = 0
    Do
        DoEvents
        vf1 = InStr(vf + 1, tt$, vbCrLf)
        vf2 = InStr(vf + 1, tt$, vbLf)
        vf3 = InStr(vf + 1, tt$, " ")
        If vf1 = 0 And vf2 = 0 And vf3 = 0 Then Exit Do
        If vf1 < vf2 Then vf = vf1 + 1
        If vf2 < vf1 Then vf = vf2
        If vf3 < vf Then vf = vf2
        
        If vf > 0 Then
            If vf + 1 >= Len(tt$) Then
                Exit Do
            End If
            Mid$(tt$, vf + 1, 1) = UCase(Mid$(tt$, vf + 1, 1))
        End If
    Loop Until vf = 0 Or vf + 1 >= Len(tt$)
End If

'Capp2 = True
If Capp2 = True Then
    vf = 0
    Do
        DoEvents
        If vf > 0 Then
            vf1 = InStr(vf + 1, tt$, vbCrLf)
            vf2 = InStr(vf + 1, tt$, vbLf)
        End If
        
        If vf1 = 0 And vf2 = 0 And vf > 0 Then Exit Do
        
        If vf = 0 Then vf = 1
        
        If vf1 < vf2 And vf1 > 0 Then vf = vf1 + 1
        If vf2 < vf And vf2 > 0 Then vf = vf2
        
        tg = 1
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Mid$(tt, vf + 1, 1))) = 0 Then
            vf3 = InStr(vf + 1, tt$, " ")
            vf1 = InStr(vf + 1, tt$, vbCrLf)
            vf2 = InStr(vf + 1, tt$, vbLf)
            If vf3 < vf1 Or vf3 < vf2 Then
            tg = 1: vf = vf3
            Else
                tg = 0
            End If
        End If
        
        If vf > 0 And tg = 1 Then
            If vf + 1 >= Len(tt$) Then
                Exit Do
            End If
            Mid$(tt$, vf + 1, 1) = UCase(Mid$(tt$, vf + 1, 1))
        
        End If
    Loop Until vf = 0 Or vf + 1 >= Len(tt$)
End If

On Local Error GoTo 0
On Error GoTo 0

Capp3 = False
If Capp3 = True Then
    vf = 0
    Do
        DoEvents
        If vf > 0 Then
            vf1 = InStr(vf + 1, tt$, vbCrLf)
            vf2 = InStr(vf + 1, tt$, vbLf)
        End If
        
        If vf1 = 0 And vf2 = 0 And vf3 = 0 Then Exit Do
        If vf1 < vf2 Then vf = vf1 + 1
        If vf2 < vf1 Then vf = vf2
        If vf3 < vf Then vf = vf2
        
        If vf > 0 Then
            If vf + 1 >= Len(tt$) Then
                Exit Do
            End If
            If tc < 3 Then Mid$(tt$, vf + 1, 1) = UCase(Mid$(tt$, vf + 1, 1))
            tc = tc + 1
        End If
    
        If vf > 0 Then
        vf3 = InStr(vf + 1, tt$, " ")
        Do
        eq2 = InStr("ABCDEFGHoIJKLMNOPQRSTUVWXYZ", UCase(Mid$(tt, vf3 + 1, 1)))
        If eq2 = 0 Then vf3 = vf3 + 1
        Loop Until eq2 > 0 Or vf3 > Len(tt$) Or vf3 > vf1
        
        If tc < 3 Then Mid$(tt$, vf + 1, 1) = UCase(Mid$(tt$, vf + 1, 1))
        End If
    
    Loop Until vf = 0 Or vf + 1 >= Len(tt$)
End If

Capp5 = False
If Capp5 = True Then

    vf = 0
    Do
        DoEvents
        If vf > 0 Then
            vf1 = InStr(vf + 1, tt$, vbCrLf)
            vf2 = InStr(vf + 1, tt$, vbLf)
        End If
        
        If vf1 = 0 And vf2 = 0 And vf > 0 Then Exit Do
        
        If vf = 0 Then vf = 1
        
        If vf1 < vf2 And vf1 > 0 Then vf = vf1 + 1
        If vf2 < vf And vf2 > 0 Then vf = vf2
        
        tg = 1
        vf3 = vf 'InStr(vf + 1, tt$, " ")
        Do
            'eq1 = Mid$(tt, vf3, 1) = " "
            eq2 = InStr("ABCDEFGHoIJKLMNOPQRSTUVWXYZ", UCase(Mid$(tt, vf3 + 1, 1))) > 0
            If eq2 = False Then vf3 = vf3 + 1
            vf = vf3
        Loop Until eq2 = True Or vf3 > vf1
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Mid$(tt, vf + 1, 1))) = 0 Then
            vf3 = InStr(vf + 1, tt$, " ")
            vf1 = InStr(vf + 1, tt$, vbCrLf)
            vf2 = InStr(vf + 1, tt$, vbLf)
            If vf3 < vf1 Or vf3 < vf2 Then
            tg = 1: vf = vf3
            Else
                tg = 0
            End If
        End If
        
        If vf > 0 And tg = 1 Then
            If vf + 1 >= Len(tt$) Then
                Exit Do
            End If
            Mid$(tt$, vf + 1, 1) = UCase(Mid$(tt$, vf + 1, 1))
        
        End If
            
        If tg = 1 Then
        
        vf4 = vf3
        vf3 = InStr(vf + 1, tt$, " ")
        Do
            'eq1 = Mid$(tt, vf3, 1) = " "
            eq2 = InStr("ABCDEFGHoIJKLMNOPQRSTUVWXYZ", UCase(Mid$(tt, vf3 + 1, 1))) > 0
            If eq2 = False Then vf3 = vf3 + 1
            vf = vf3
        Loop Until eq2 = True Or vf3 > vf1
        
        vf = vf3
        If tg = 1 And vf3 < vf2 Then
            If vf + 1 >= Len(tt$) Then
                Exit Do
            End If
            Mid$(tt$, vf + 1, 1) = UCase(Mid$(tt$, vf + 1, 1))
        
        End If
        End If
        
        'If tg = 0 Then vf = vf + 1
    
    If vf < oldvf Then Exit Do
    oldvf = vf
    
    Loop Until vf = 0 Or vf + 1 >= Len(tt$)

End If
'Mid$(tt$, vf + 1, 30)

FixIS = True
If FixIS = True Then

Dim Ary1(30)
Dim Ary2(30)
    
cc = 1: Ary1(cc) = "what": Ary2(cc) = " wat "
cc = cc + 1: Ary1(cc) = " Joanna ": Ary2(cc) = " jo "
cc = cc + 1: Ary1(cc) = "Sophia": Ary2(cc) = "sophie"
cc = cc + 1: Ary1(cc) = "callerID": Ary2(cc) = "callerid"
cc = cc + 1: Ary1(cc) = "frequently": Ary2(cc) = "freq "
cc = cc + 1: Ary1(cc) = "favourite ": Ary2(cc) = "fav "
cc = cc + 1: Ary1(cc) = "Sharon": Ary2(cc) = "shaz"
cc = cc + 1: Ary1(cc) = "accommodation ": Ary2(cc) = "accom "
cc = cc + 1: Ary1(cc) = "accommodation ": Ary2(cc) = "accomo "
cc = cc + 1: Ary1(cc) = "accommodation ": Ary2(cc) = "accommo "
cc = cc + 1: Ary1(cc) = "alison": Ary2(cc) = "alice"
cc = cc + 1: Ary1(cc) = "didn't": Ary2(cc) = "didnt"
cc = cc + 1: Ary1(cc) = "wasn't": Ary2(cc) = "wasnt"
cc = cc + 1: Ary1(cc) = "can't": Ary2(cc) = "cant"
cc = cc + 1: Ary1(cc) = "coundn't": Ary2(cc) = "couldnt"
cc = cc + 1: Ary1(cc) = "office": Ary2(cc) = "offy"
cc = cc + 1: Ary1(cc) = "Dawn Gluck": Ary2(cc) = "dawn gluck"
cc = cc + 1: Ary1(cc) = "message": Ary2(cc) = "msg"
cc = cc + 1: Ary1(cc) = "that's": Ary2(cc) = "thats"
cc = cc + 1: Ary1(cc) = "maximum": Ary2(cc) = "max "
cc = cc + 1: Ary1(cc) = "program": Ary2(cc) = "prog "
cc = cc + 1: Ary1(cc) = "who's": Ary2(cc) = "whos"
cc = cc + 1: Ary1(cc) = "tries": Ary2(cc) = "trys"
cc = cc + 1: Ary1(cc) = "here's": Ary2(cc) = "heres"
cc = cc + 1: Ary1(cc) = " another": Ary2(cc) = " nother"
cc = cc + 1: Ary1(cc) = "Rutland": Ary2(cc) = "rut "
cc = cc + 1: Ary1(cc) = "Rutland": Ary2(cc) = "rutt "
'cc = cc + 1: Ary1(cc) = "Hospital": Ary2(cc) = "hoss"
'cc = cc + 1: Ary1(cc) = "": Ary2(cc) = ""


vf = 0
Do
    DoEvents
    
    For ii = 1 To cc
    If Ary1(cc) = "" Then Exit For
    vf = InStr(tt2$, Ary2(ii))
    If vf = 0 Then vf = InStr(tt2$, vbLf + Ary2(ii))
    If vf = 0 Then vf = InStr(tt2$, vbCr + Ary2(ii))

    If vf > 0 Then
        'xx = Mid$(tt2$, vf - 1, Len(Ary1(ii)) + 1)
        tt$ = Mid$(tt$, 1, vf - 1) + Ary1(ii) + Mid$(tt$, vf + Len(Ary2(ii)))
        tt2$ = Mid$(tt2$, 1, vf - 1) + Space$(Len(Ary1(ii))) + Mid$(tt2$, vf + Len(Ary2(ii)))
        cute3 = 1
    End If

    Next

Loop Until vf = 0 Or vf > scansize




vf = 0
Do
    DoEvents
    vf = InStr(vf + 1, tt$, vbCrLf + vbCrLf)
    If vf > 0 Then
        If vf + 4 >= Len(tt$) Then
            Exit Do
        End If
        Mid$(tt$, vf + 4, 1) = UCase(Mid$(tt$, vf + 4, 1))
        cute3 = 1
    End If
Loop Until vf = 0 Or vf > scansize

tt2$ = LCase(tt$)
Do
    DoEvents
    tq$ = "sophia": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "matthew": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "oct": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "nov": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "dec": vf = InStr(tt2$, tq$)
'    If vf = 0 Then tq$ = "rut": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "hur": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "jo ": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "john": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "janice": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "paul": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "shaz": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "sarah": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "sophia": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "alice": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "alison": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "shaz": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "sharon": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "clive": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "andy": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "luv": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "steve": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "stephen": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "joanna": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "i'll": vf = InStr(tt2$, tq$)
    For r = 0 To 7
        If vf = 0 Then
            tq$ = LCase(Format$(DateSerial(2000, 1, r), "DDDD")): vf = InStr(tt2$, tq$)
            If vf > 0 Then
                Exit For
            End If
        End If
    Next
    For r = 0 To 7
        If vf = 0 Then tq$ = LCase(Format$(DateSerial(2000, 1, r), " DDD ")): vf = InStr(tt2$, tq$)
    Next
    For r = 0 To 12
        If vf = 0 Then tq$ = LCase(Format$(DateSerial(2000, r, 1), "MMMM")): vf = InStr(tt2$, tq$)
    Next
    For r = 0 To 12
        If vf = 0 Then tq$ = LCase(Format$(DateSerial(2000, r, 1), " MMM ")): vf = InStr(tt2$, tq$)
    Next
    
    
    If vf > 0 Then Mid$(tt$, vf, 1) = UCase(Mid$(tt$, vf, 1)): Mid$(tt2$, vf, 2) = UCase(Mid$(tt2$, vf, 2))

Loop Until vf = 0 Or vf > scansize

'All Caps
tt2$ = LCase(tt$)
Do
    DoEvents
    tq$ = " id ": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = " bc ": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = " bc" + vbCr: vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "bht ": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "bht" + vbCr: vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "pot ": vf = InStr(tt2$, tq$)
    If vf = 0 Then tq$ = "cod ": vf = InStr(tt2$, tq$)
    'If Mid$(tq$, 1, 1) = " " Then vf = vf + 1
    If vf > 0 Then Mid$(tt$, vf, Len(tq$)) = UCase(Mid$(tt$, vf, Len(tq))): Mid$(tt2$, vf, 1) = "x"
Loop Until vf = 0 Or vf > scansize



'cute3 = True


End If

'If cute3 > 0 Then
    Open QQ$ For Output As #1
        Print #1, tt$
    Close #1
'End If

'End If 'file size too big to do


End Sub


Sub AddDateToTopLogg()

Open "D:\VB6\VB-NT\00_Best_VB_01\NotePad Abbrev\Abbrev NotePad Logg.txt " For Input As #1
    LL$ = Input$(LOF(1), 1)
Close #1
If InStr(LL$, QQ$) = 0 Then
    Open "D:\VB6\VB-NT\00_Best_VB_01\NotePad Abbrev\Abbrev NotePad Logg.txt " For Append As #1
    Print #1, Format$(Now, "dd-mm-yyyy hh:mm:ss ") + QQ$
    Close #1
End If


Open QQ$ For Binary As #1
    tt$ = Input$(LOF(1), #1)
Close #1

gx = 1
Dim RT$(1005)
For r2 = 1 To 1000
gh = InStr(gx, tt$, vbCrLf)
If gh = 0 Then
    RT$(r2) = Mid$(tt$, gx)
    Exit For
End If
RT$(r2) = Mid$(tt$, gx, gh - gx)
gx = gh + 2
Next

For r2 = 1 To 1000
If Mid$(RT$(r2), 1, 2) = "--" And Len(RT$(r2)) = 28 Then
If RT$(r2 + 1) = "" And RT$(r2 + 2) = "" And RT$(r2 + 3) = "-----" Then
ty = InStr(tt$, RT$(r2))
If ty > 0 Then
ti = Len(RT$(r2)) + Len(RT$(r2 + 1)) + Len(RT$(r2 + 2)) + Len(RT$(r2 + 3)) + 8 '8=vbcrlf's*4
If ty = 1 Then
    yy$ = Mid$(tt$, ty + ti)
Else
    yy$ = Mid$(tt$, 1, ty) + Mid$(tt$, ty + ti)
End If
tt$ = yy$
Exit For
End If
End If
End If
Next

'End If


tt$ = Format$(Now, "-- DDD DD-MMM-YY HH:MM:SS --") + vbCrLf + vbCrLf + vbCrLf + String$(5, "-") + vbCrLf + tt$

ffg = 0
If InStr(LL$, QQ$) > 0 Then ffg = 1
If InStr(QQ$, "NotePad Files") = 0 Then ffg = 1
If ffg = 0 Then
tx = 0
tuff = Len(tt$)
Do
tx = InStr(tx + 1, tt$, vbCrLf)
If tx = 0 Then Exit Do
If tx + 1 >= tuff Then Exit Do
hy = InStrRev(tt$, "---------------------" + vbCrLf + "Count =", tx)
hy2 = InStrRev(tt$, vbCrLf + "-----" + vbCrLf, tx + 2)
pup = 0
If hy2 > 0 Then
hy3 = InStrRev(tt$, " --" + vbCrLf, tx)
hy4 = InStrRev(tt$, "-- ", hy3)
hy5 = InStr(hy4, tt$, ":")
If hy4 + 25 = hy3 Then pup = 1
If hy4 + 19 = hy5 And pup = 1 Then pup = 2
End If
If pup = 2 And hy4 > hy Then pup = 3

'If hy > 0 Then Stop
If tx > 0 And pup = 3 Then
    Mid$(tt$, tx + 2, 1) = UCase(Mid$(tt$, tx + 2, 1))
End If

'Mid$(tt$, tx + 2, 10)
If (InStr("0123456789 -:.,)(*&^%$£""!@;#}{[]+_-=/\?><`¬" + vbTab, Mid$(tt$, tx + 2, 1)) > 0 And pup = 3) Then
td = InStr(tx + 1, tt$, vbCrLf)
If td = 0 Then td = Len(tt$)
tj = tx + 1
pip = 0
Do
tk = InStr("0123456789 -:.,)(*&^%$£""!@;#}{[]+_-=/\?><`¬" + vbTab, Mid$(tt$, tj + 1, 1))
If tj + 1 >= td Then pip = 1: Exit Do
If tk = 0 Then Exit Do
tj = tj + 1
Loop Until tx = 0
If pip = 0 Then
    Mid$(tt$, tj + 1, 1) = UCase(Mid$(tt$, tj + 1, 1))
End If
End If

Loop Until tx = 0

End If

Open QQ$ For Output As #1
Print #1, tt$;
Close #1

'Shell "C:\windoze\system32\notepad.exe D:\# MY DOCS\# 01 My Documents\CALLerid\2ABBREV.TXT"
'Shell "C:\windoze\system32\notepad.exe " + QQ$, vbNormalFocus

End Sub


Private Sub Label_Click()

End Sub

