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

Public Sub FormStartLoader()


FontSizez = 7.9

ScanPath.cboMask.Text = "*.*"
ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath.Text = "E:\01 VB Shell Folders\00 " + App.EXEName
Call ScanPath.cmdScan_Click

ScanPath.cboMask.Text = "*.txt;*.lnk"
ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath.Text = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files"
Call ScanPath.cmdScan_Click

'WHY DUPELICATE AND BIG
'ScanPath.cboMask.Text = "*.txt"
'ScanPath.chkSubFolders = vbUnchecked
'ScanPath.txtPath.Text = "D:\# MY DOCS\# 01 My Documents"
'Call ScanPath.cmdScan_Click

'not much use this one
ScanPath.cboMask.Text = "*.txt"
ScanPath.chkSubFolders = vbUnchecked
ScanPath.txtPath.Text = "D:\# MY DOCS\# 01 My Documents\00 Email Texts"
Call ScanPath.cmdScan_Click



End Sub

Sub LabelClick(index)


If SetTrueToLoadLast = False Then
    A1$ = ScanPath.ListView1.ListItems.Item(index).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(index)
    C1$ = Form1.Label1(index).Caption
End If

UU$ = B1$

If LoadFolder2 = True Then
    
    If InStr(LCase(B1$), ".lnk") = 0 Then
        Shell "Explorer.exe /select, " + A1$ + B1$, vbNormalFocus
        Unload Form1
        Exit Sub
    Else
        
        Call GETSHORTLINK(A1$ + B1$)
        
        If InStr(txtTargetPath, ":\") = 0 Then
            i = MsgBox("INVAILD LINK" + vbCrLf + A1$ + B1$ + vbCrLf + "DO YOU WANT ME TO TAKE YOU TO THE LINK", vbYesNo + vbMsgBoxSetForeground)
            If i = vbYes Then
                Shell "Explorer.exe /select, " + A1$ + B1$, vbNormalFocus
            End If
            Unload Form1
            Exit Sub
        End If
        If FS.FileExists(txtTargetPath) = False Then
            i = MsgBox("BROKEN LINK" + vbCrLf + txtTargetPath + vbCrLf + "IN" + vbCrLf + A1$ + B1$ + vbCrLf + "DO YOU WANT ME TO TAKE YOU TO THE LINK", vbYesNo + vbMsgBoxSetForeground)
            If i = vbYes Then
                Shell "Explorer.exe /select, " + A1$ + B1$, vbNormalFocus
            End If
            Unload Form1
            Exit Sub
        End If
        
        
        Shell "Explorer.exe /select, " + txtTargetPath, vbNormalFocus
        Unload Form1
        Exit Sub
    End If
End If

'If InStr(LCase(B1$), ".lnk") > 0 Then vLaunch A1$ + B1$, vbNullString: End
UU$ = Mid$(B1$, 1, Len(B1$) - 4)
xa = 0
If xa = 0 Then xa = FindWindow(vbNullString, "* " + A1$ + UU$ + " - Notepad2")
If xa = 0 Then xa = FindWindow(vbNullString, A1$ + UU$ + " - Notepad2")

If xa > 0 Then
    'MsgBox "Already Open" + vbCrLf + B1$
    SetForegroundWindow (xa)
    'END
    Unload Form1
    Exit Sub
End If

If InStr(LCase(B1$), ".lnk") > 0 Then vLaunch A1$ + B1$, vbNullString: End


'Call Form1.LoadLoggs

xxy = 0
If A1$ + B1$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Hosso\Hosso Weather.txt" Then xxy = 1
If A1$ + B1$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Hosso\Talk About In Front of Me.txt" Then xxy = 1
If InStr(A1$ + B1$, "D:\# MY DOCS\# 01 My Documents\03 NotePad Files") Then xxy = 1
If InStr(B1$, "Date1991") > 0 Then xxy = 0


'TOOK OUT DON'T RUN THE BLOGGER CODE ANYMORE
'If xxy = 1 Then
'C1$ = "D:\VB6\VB-NT\00_Best_VB_01\NotePad Abbrev\NotePad instr top LF 2abbrev txt.exe"
''d1$ = GetShortName(a1$ + b1$)
'D1$ = A1$ + B1$
'vLaunch C1$, D1$
'End
'End If


vLaunch A1$ + B1$, vbNullString
Unload Form1
Exit Sub
End


End Sub

Private Sub Form_Unload(Cancel As Integer)

'SET ALL TIMERS IN ALL FORMS TO ENABLED = FALSE
On Error Resume Next
    For i = 0 To Forms.Count - 1
        For Each Control In Forms(i).Controls
            'If InStr(UCase(Control.Name), "TIMER") > 0 Then
                Control.Enabled = False
            'End If
        Next
    Next i
On Error GoTo 0
   
Dim Form As Form
For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form
Unload Form1
End Sub
