VERSION 5.00
Begin VB.Form FormStart 
   Caption         =   "Form2"
   ClientHeight    =   3192
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3192
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FormStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub FormStartLoader()

FontSizez = 7

Set fs = CreateObject("Scripting.FileSystemObject")

ScanPath.cboMask.Text = "*.lnk;*.bat"

ScanPath.txtPath.Text = "E:\01 VB Shell Folders\00 " + App.EXEName
Call ScanPath.cmdScan_Click

End Sub

Sub LabelClick(index)

If SetTrueToLoadLast = False Then
    A1$ = ScanPath.ListView1.ListItems.Item(index).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(index)
    C1$ = Form1.Label1.Item(index)
End If

'Call Form1.LoadLoggs

'----------------- Keep Above




If LoadFolderFile = True Then

    d1$ = A1$ + B1$
    If InStr(LCase(B1$), ".lnk") > 0 Then
        Open A1$ + B1$ For Binary As #1
        rr$ = Space$(LOF(1))
        Get #1, 1, rr$
        Close #1
        'If Mid$(rr$, &H568, 3) = ":" + Chr$(0) + "\" Then
        'rr$ = Mid$(rr$, &H566)
        'End If
        tt$ = Mid$(rr$, InStr(105, rr$, ":\") - 1)
        d1$ = Mid$(tt$, 1, InStr(tt$, Chr$(0)) - 1)
        'If InStr(d1$, "Blue Tooth Sync Folder") > 0 Then MsgBox ("Wrong Link Points to Blue Tooth folder"): End


        d1$ = Mid$(tt$, 1, InStr(tt$, Chr$(0)) - 1)
        'd1$ = Mid$(d1$, 1, InStrRev(d1$, "\"))
        Shell "explorer /e, /select, " + d1$, vbNormalFocus
        End
    End If
End If


'If Check1.Value = vbUnchecked Then
'Open App.Path + "\Winloads.txt" For Append As #1
'Print #1, Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " -- " + Label1(Index)
'Close #1
'Shell Dir1.List(Index) + "\Winamp.exe", vbNormalFocus
A1$ = ScanPath.ListView1.ListItems.Item(index).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(index)

Dim fs, F
Set fs = CreateObject("Scripting.FileSystemObject")

If fs.FileExists(A1$ + B1$) = False Then
    MsgBox "NOT FOUND" + vbCrLf + A1$ + B1$
End If

Form1.vLaunch A1$ + B1$, vbNullString
End
Exit Sub

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
Form1.vLaunch A1$ + B1$, vbNullString
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
Form1.vLaunch A1$ + B1$, vbNullString
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
'End

'Else
'Shell "notepad " + Dir1.List(Index) + "\Current Play To Text File Append.txt", vbNormalFocus

'End

'End If

End



End Sub

