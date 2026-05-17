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


FontSizez = 10

ScanPath.cboMask.Text = "*.*"
'ScanPath.cboMask.Text = "*.lnk;*.pif"
ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath.Text = "E:\01 VB Shell Folders\00 Shell " + OIP2$ + " Loader"
Call ScanPath.cmdScan_Click
A = ScanPath.ListView1.ListItems.Count
ScanPath.cboMask.Text = "*.vbp;*.exe"
ScanPath.chkSubFolders = vbChecked
ScanPath.txtPath.Text = "D:\VB6\VB-NT\00_Best_VB"
Call ScanPath.cmdScan_Click

'Set fs22 = CreateObject("Scripting.FileSystemObject")
Dim WFD As WIN32_FIND_DATA
Dim lResult As Long
    

For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    If InStr(A1$, "00_Best_VB") > 0 And InStr(B1$, ".vbp") > 0 Then
        
        lResult = FindFirstFile(A1$ + "*.exe", WFD)
        If lResult = INVALID_HANDLE_VALUE Then
            ScanPath.ListView1.ListItems.Remove (we)
        End If
        
    End If
Next


For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    If InStr(A1$, "00_Best_VB") > 0 And InStr(B1$, ".exe") > 0 Then
        
        lResult = FindFirstFile(A1$ + "*.vbp", WFD)
        If lResult = INVALID_HANDLE_VALUE Then
            ScanPath.ListView1.ListItems.Remove (we)
        End If
        
    End If
Next




End Sub

Sub LabelClick(index)


'If Check1.Value = vbUnchecked Then
'Open App.Path + "\Winloads.txt" For Append As #1
'Print #1, Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " -- " + Label1(Index)
'Close #1
'Shell Dir1.List(Index) + "\Winamp.exe", vbNormalFocus
A1$ = ScanPath.ListView1.ListItems.Item(index).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(index)

If Sqidge = True Then Shell "Explorer.exe /e," + A1$, vbNormalFocus: End


vLaunch A1$ + B1$, vbNullString
End

que = 0
If A1$ = "--Drive" Then
que = 1
Shell "Explorer.exe /e," + Mid$(B1$, 1, 3), vbNormalFocus
Else
Open A1$ + B1$ For Binary As #1
rr$ = Space$(LOF(1))
Get #1, 1, rr$
If Mid$(rr$, &H568, 3) = ":" + Chr$(0) + "\" Then
rr$ = Mid$(rr$, &H566)
End If

'que = 0
If B1$ = "My Computer.lnk" Then
vLaunch A1$ + B1$, vbNullString
End
Exit Sub
End If

If que = 0 Then
tq = InStrRev(rr$, ":\")
If tq = 0 Then
    tq = InStrRev(rr$, ":" + Chr(0) + "\")
    If tq > 0 Then tq = tq - 1
End If
If tq = 0 Then
'MsgBox "Not Right Path Not Found in shortcut Prog Launch Explorer"
vLaunch A1$ + B1$, vbNullString
End
End If
rr$ = Mid$(rr$, tq - 1)
End If

tq = InStr(rr$, Chr$(0) + Chr$(0))
rr$ = Mid$(rr$, 1, tq - 1)

Do
tq = InStr(rr$, Chr$(0))
If tq > 0 Then rr$ = Mid$(rr$, 1, tq - 1) + Mid$(rr$, tq + 1)
Loop Until tq = 0

Shell "Explorer.exe /e," + rr$, vbNormalFocus
'vLaunch a1$ + b1$

End If

End


End Sub

