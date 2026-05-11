Attribute VB_Name = "Module1"



Public Sub MoveFilesOlderandMakeDIR()

'If IsIDE = True Or Command$ = "" Then
If Command$ = "" Then
    ScanPath.Show
    DoEvents
    ScanPath.WindowState = 0
    DoEvents
Else
    ScanPath.Show
    DoEvents
    ScanPath.WindowState = 1
    DoEvents
End If

Dim RSS

For RSS = 1 To 3

ScanPath.lblCount1 = ""
ScanPath.lblCount2 = ""
ScanPath.lblCount3 = ""
ScanPath.lblCount4 = ""
ScanPath.lblCount5 = ""
ScanPath.lblCount6 = ""


If RSS = 1 Then ScanPath.txtPath.Text = "M:\WinRar Archives\"
If RSS = 2 Then ScanPath.txtPath.Text = "M\0 00 Art Loggers\Screen-App-Shots\"
If RSS = 3 Then ScanPath.txtPath.Text = "M\0 00 Art Loggers\Screen-Shots\"


'NEWER
'ScanPath.cboDate.ListIndex = 0 'MODIFIED
'ScanPath.DTPicker1(0) = Int(Now - 50)

'OLDER
ScanPath.cboDate.ListIndex = 0 'MODIFIED
'ScanPath.cboDate.ListIndex = 1 'CREATED
ScanPath.DTPicker1(1) = Int(Now - 50)

ScanPath.cboMask.Text = "*.*"
'cboSize.ListIndex = 1 'Less than
'cboSize.ListIndex = 2 'Bigger Than
'cboSizeType(lCount).ListIndex = 0 'Byte
'cboSizeType(lCount).ListIndex = 2 'Meg
'cboSizeType(lCount).ListIndex = 1 'K
'txtSize(0) = 3

ScanPath.lblCount6 = "*DO* 01 OF 04 - Scan List"

ScanPath.ListView1.ListItems.Clear

Call ScanPath.cmdScan_Click


On Error Resume Next

'Exit Sub
OA1 = ""

ScanPath.lblCount6 = "*DO* 02 OF 04 - Make Dir's"

'MAKE DIRS
For we = 1 To ScanPath.ListView1.ListItems.Count
    
    ScanPath.lblCount2 = Str(we) + " Make Dir's"
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    'ICYBOX
    d1 = "L:\" + Mid$(ScanPath.txtPath, 4, Len(ScanPath.txtPath) - 4)
    d2 = d1 + Mid$(A1$, Len(d1) + 1)
        
    Err.Clear
    DoEvents
    
    'OA1 = A1
    If OA1 <> A1 Then
    If FS.FolderExists(d2) = False Then
        ar = 3
        d2 = d1 + Mid$(A1$, Len(d1) + 1)
        
        Do
            DoEvents
            ar = InStr(ar + 1, d2, "\")
            
            D3 = Mid$(d2, 1, ar - 1)
            If ar = 0 Then D3 = d2
            
            Err.Clear
            MkDir D3
        Loop Until FS.FolderExists(d2) = True Or ar = 0
        
        'ANY ERROR REPORTED SHOULD BE NONE
        If Err.Number > 0 Then Stop
        If FS.FolderExists(d2 + "\") = False Then Stop
    End If
    End If

Next



ScanPath.lblCount6 = "*DO* 03 OF 04 - Move Files"


ScanPath.DTPicker1(1) = vbNullString


'MOVE FILES
For we = 1 To ScanPath.ListView1.ListItems.Count
    
    ScanPath.lblCount3 = Str(we) + " Move Files"
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    'd1 = "J:\WinRar Arcives"
    d1 = "L:\" + Mid$(ScanPath.txtPath, 4, Len(ScanPath.txtPath) - 4)
    d2 = d1 + Mid$(A1$, Len(d1) + 1) + B1$
    'd2 = Mid(d2, 1, Len(d2) - 1)
        
    Err.Clear
    DoEvents
    If FS.FileExists(d2) = False Then
    Set F = FS.getfile(A1 + B1)
    D3 = d1 + Mid$(A1$, Len(d1) + 1) + B1$
    
    On Error Resume Next
    F.Move (D3)
    On Error GoTo 0
    
    'ERR.DESCRIPTION
    'THIS ERROR IF ENCRYPTED = "Method 'Move' of object 'IFile' failed"
    
    End If
    
    
    
    


Next


ScanPath.lblCount6 = "*DO* 04 OF 04 - Del Emptys"


If ScanPath.ListView1.ListItems.Count > 0 Then

'DEL EMPTYS
'ScanPath.txtPath.Text = "M:\WinRar Arcives\"

ScanPath.ListView1.ListItems.Clear

Call ScanPath.cmdScanDir_Click

On Error Resume Next
'On Error GoTo 0


For we = 1 To ScanPath.ListView1.ListItems.Count
    
    DoEvents
    
    ScanPath.lblCount5 = "Count *" + Trim(Str(we)) + " Del Dirs *" + Trim(Str(tt)) + "Err *" + Trim(Str(tt1))
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)

    If FS.FolderExists(A1) = True Then
    Set F = FS.GetFolder(A1)
    If F.Size = 0 Then
    
    D3 = Mid$(A1$, 1, Len(A1) - 1)
    Err.Clear
    FS.deletefolder D3, True
    tt = tt + 1
    If Err.Number > 0 Then tt1 = tt1 + 1
        
    End If
    End If


Next

End If


ScanPath.lblCount6 = "-- *DONE* --"

Next RSS

If ScanPath.WindowState = 1 Then End


End Sub


