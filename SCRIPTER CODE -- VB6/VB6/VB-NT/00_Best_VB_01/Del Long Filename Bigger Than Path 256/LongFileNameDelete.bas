Attribute VB_Name = "LongNameDelete"
Public m_CRC As clsCRC

'Put in Sub Of Code

Sub LongNameDeleteSub()

ScanPath.Show

Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32



ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath = Clipboard.GetText
ScanPath.cboMask.Text = "*.*"
ScanPath.chkSubFolders = vbChecked

If IsIDE = True And 1 = 2 Then


'    ScanPath.txtPath = "M:\0 00 ART LOGGERS - WEBCAM\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
    
'    ScanPath.txtPath = "U:\0 00 ART LOGGERS\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'    ScanPath.txtPath = "Y:\0 00 Art Loggers\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click

'    ScanPath.txtPath = "U:\0 00 ART LOGGERS\# TEMP INTERNET\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click


w$ = "U:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\"
w$ = "M:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\"
w$ = "f:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\"
w$ = "F:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\02 VB6 An MSDN 02\"
w$ = "F:\03 MICROSOFT\# VB 6\"
w$ = "M:\03 MICROSOFT\# OP SYS\Test\"
'w$ = "M:\DSC\"
'w$ = "I:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\02 VB6 An MSDN 02\"
'    'ScanPath.txtPath = "U:\00 0 VIDEO CAMERA'S\"
'    'ScanPath.txtPath = "M:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
''w$ = "F:\#0 1 INSTALLATIONS\Installation programs\03 MICROSOFT\# VB 6\"
    
    
    
'2015 jul 11
'    w$ = "I:\03 MICROSOFT\# VB 6\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
'2015 jul 11
'    w$ = "M:\DSC\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
'2015 JUL 13 done
'    w$ = "U:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
 
''2015 jul 13 DONE
'    w$ = "I:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
''2015 jul 13 DONE
'    w$ = "I:\00 0 VIDEO CAMERA'S\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'
''2015 jul 13 DONE
'    w$ = "I:\0 00 MEDIA\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
    
''2015 JUL 14
'    w$ = "I:\0 00 ART LOGGERS\# FEED STATION\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "M:\0 00 ART LOGGERS\# FEED STATION\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "U:\0 00 ART LOGGERS\# FEED STATION\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    

'2015 JUL 14 start working on one drive - DONE JUN 16
'    w$ = "I:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "M:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "U:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
'2015 JUL 16 start working NEXT WAIT FOR SUPERCOPY
'subst i: w:
    
'    w$ = "T:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "w:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "X:\0 00 ART LOGGERS\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "T:\0 00 Art\# 00 My Pictures - LARGE COUNT\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "W:\0 00 Art\# 00 My Pictures - LARGE COUNT\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "X:\0 00 Art\# 00 My Pictures - LARGE COUNT\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'
'    w$ = "T:\0 00 Art\# 00 My Pictures - LARGE COUNT\AUTOPIX 01\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "w:\0 00 Art\# 00 My Pictures - LARGE COUNT\AUTOPIX 01\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "T:\0 00 HTTrack\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "W:\0 00 HTTrack\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    w$ = "X:\0 00 HTTrack\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'

'    w$ = "D:\0 00 Art\# 00 My Pictures - LARGE COUNT\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'    w$ = "T:\0 00 Art\# 00 My Pictures - LARGE COUNT\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'    w$ = "i:\0 00 Art\# 00 My Pictures - LARGE COUNT\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'    w$ = "u:\0 00 Art\# 00 My Pictures - LARGE COUNT\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'    w$ = "m:\0 00 Art\# 00 My Pictures - LARGE COUNT\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'

    w$ = "X:\0 00 Art Loggers\"
    ScanPath.txtPath = w$
    If Fs.FolderExists(ScanPath.txtPath) = False Then End
    Call ScanPath.cmdScan_Click

'Next
'    w$ = "T:\0 00 ART WORK IN BOUND\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click

    
    
    
'----------------- done
'T:\0 00 Art\# 00 My Pictures\02 HTTRACK\NASA
'M: SUBST As U:
'U:\0 00 Art\# 00 My Pictures\02 HTTRACK\NASA
'X:\0 00 Art\00 HTTrack\00 NASA
    
'    w$ = "T:\0 00 Art\# 00 My Pictures\02 HTTRACK\NASA\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'    w$ = "U:\0 00 Art\# 00 My Pictures\02 HTTRACK\NASA\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'    w$ = "X:\0 00 Art\00 HTTrack\00 NASA\"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
    
    
'
    
    
    'NEXT
'    w$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental"
'    ScanPath.txtPath = w$
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click



'
'
'    ScanPath.txtPath = "U:\0 00 MEDIA\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'

'
'    ScanPath.txtPath = "U:\KAT MP3 RECORDER\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'
'    ScanPath.txtPath = "Y:\KAT MP3 RECORDER\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
'

'    ScanPath.txtPath = "Y:\#\"
'    If Fs.FolderExists(ScanPath.txtPath) = False Then End
'    Call ScanPath.cmdScan_Click
Else
    If Mid(ScanPath.txtPath, Len(ScanPath.txtPath)) <> "\" Then ScanPath.txtPath = ScanPath.txtPath + "\"
    If Fs.FolderExists(ScanPath.txtPath) = False Then End
    Call ScanPath.cmdScan_Click
End If

ScanPath.Show
Reset
yy = "--"
    
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " -- Total"


'Delete LONG'S is Naught Size's Only
'Exit Sub

For WE = ScanPath.ListView1.ListItems.Count To 1 Step -1
    
    
    'If WE Mod 10 = 0 Then ScanPath.lblCount2 = Trim(Str(WE)): DoEvents
    ScanPath.lblCount2 = Trim(Str(WE)): DoEvents
    
    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(WE)
    G1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
    If Fs.FileExists(G1$) Then
        Set F = Fs.getfile(G1$)
        FSize = F.Size
        'Set F = Nothing
    
        'xgo = 1
        
        'DONT DELETE FROM MASTER AND NEVER CRC DUPE WITH OTHER FOLDERS
    '    If InStr(A1$, "# COPY FAMILY 1991-2006") > 0 Then xgo = 0
    '    If InStr(A1$, "\DSC\# ME") > 0 Then xgo = 0
    '    If InStr(A1$, "\DSC\# Docus Proofs Texts") > 0 Then xgo = 0
        
        If FSize = 0 And Len(A1$ + B1$) > 255 Then
            'ScanPath.ListView1.ListItems.Remove (WE)
            'Set F = Fs.getfile(A1$ + G1$)
            F.Delete
            wex4 = wex4 + 1
            ScanPath.lblCount5 = Trim(Str(wex4)) + " -- Del - Long = " + Format((we2 / wex4) * 100, "0.00") + "%"
            
            BB$ = BB$ + A1$ + B1$ + vbCrLf
        End If
    End If

Next

NB = "DELETED TOO LONG PATH" + vbCrLf
NB = NB + "---------------------" + vbCrLf
NB = NB + BB
Clipboard.Clear
Clipboard.SetText NB
lblCount6 = "Finished"
Exit Sub

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " -- Total"

For WE = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    ScanPath.lblCount2 = Trim(Str(WE))
    DoEvents
    
    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(WE)
    G1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
'    Shell "explorer.exe /select," + A1$ + B1$, vbNormalFocus
    
    If Fs.FileExists(A1$ + G1$) = True Then
        Set F = Fs.getfile(A1$ + G1$)
        RS = Format(F.Size, "000000000000000000")
        Set F = Nothing
        ScanPath.ListView1.ListItems.Item(WE).SubItems(3) = RS
    Else
        RS = Format(0, "000000000000000000")
        ScanPath.ListView1.ListItems.Item(WE).SubItems(3) = RS
    
    End If

'    Kill A1$ + G1$
'    If d1$ <> A1$ Then
'        d1$ = A1$
'        ChDrive A1$
'        ChDir A1$
'        Shell "C:\Program Files\WinRAR\WinRAR.exe e -inul *.rar"
'    End If

Next
'Exit Sub


'----------------------------------------------------
'REMOVE NAUGHT SIZE AND RAR'S

For WE = ScanPath.ListView1.ListItems.Count To 1 Step -1
    DoEvents
    ScanPath.lblCount2 = Trim(Str(WE))
    DoEvents
    X1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)

    XDONE = 0

    If Val(X1$) = 0 Then
        A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(WE)
        'Kill A1$ + B1$
        ScanPath.ListView1.ListItems.Remove (WE)
        XDONE = 1
    End If
    

'    B1$ = ScanPath.ListView1.ListItems.Item(WE)
'
'    If InStr(LCase(B1$), ".rar") > 0 And XDONE = 0 Then
''        ScanPath.ListView1.ListItems.Remove (WE)
'    End If
'    If InStr(UCase(A1$), "\ME") > 0 And XDONE = 0 Then
'        ScanPath.ListView1.ListItems.Remove (WE)
'    End If

Next


'----------------------------------------------------
'SORT ON SIZE

'ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " - After 0 and *.Rar"
ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " - After 0 Size"
DoEvents

ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 3
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False

DoEvents

'--------------------------------------------------
'REMOVE NON DUPE SIZE'S ONE ABOVE AND ONE BELOW pass 01 of 02

TOPWE = ScanPath.ListView1.ListItems.Count

For WE = ScanPath.ListView1.ListItems.Count To 1 Step -1
    DoEvents
    ScanPath.lblCount3 = Trim(Str(WE))
    DoEvents
    
    If WE - 1 >= 1 Then
        X1$ = ScanPath.ListView1.ListItems.Item(WE - 1).SubItems(3)
    Else
        X1$ = "X"
    End If
    X2$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(3)
    If WE + 1 > TOPWE Then
        X3$ = "X"
    Else
        X3$ = ScanPath.ListView1.ListItems.Item(WE + 1).SubItems(3)
    End If


    If X2$ = X1$ Or X2$ = X3$ Then
        RESULT = True
    Else
        ScanPath.ListView1.ListItems.Remove (WE)
        TOPWE = ScanPath.ListView1.ListItems.Count
    
    End If
Next

ScanPath.lblCount3 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " - After Removed Non Dupes"


'------------------------------------------------
'Main Routine

we3 = ScanPath.ListView1.ListItems.Count

ScanPath.chkRefreshListView.Value = vbChecked


For WE = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    ScanPath.lblCount4 = Trim(Str(WE))
    ScanPath.ListView1.ListItems(WE).Selected = True
    ScanPath.ListView1.SelectedItem.EnsureVisible
    
    G1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(WE)
    
    FSize = Val(ScanPath.ListView1.ListItems.Item(WE).SubItems(3))
    ScanPath.lblCount6 = Trim(Str(FSize))
    'WxHex$ = Hex(m_CRC.CalculateString(Form3.Text1.Text))
    On Error Resume Next

'    Err.Clear
'    If we = 1 Then Stop
'    If Err.Number > 0 Then GoTo endb
    
    If OldFSize <> FSize Then yy = ""
    OldFSize = FSize
    
    If FSize = 0 Then Stop
'    If FSize > 0 Then
        If Fs.FileExists(A1$ + G1$) = True Then
            WXHEX$ = Hex(m_CRC.CalculateFile(A1$ + G1$))
            If oldfilename <> "" Then
                Set F = Fs.getfile(oldfilename)
                xd1 = F.datelastmodified
            End If
            
            Set F = Fs.getfile(A1$ + G1$)
            xd2 = F.datelastmodified
            
            
        
            If InStr(yy, WXHEX$) > 0 Then 'And Last$ <> A1$ + B1$ Then
                we2 = we2 + 1

                'BLACKLIST
                'DONT DELETE FROM MASTER FOLDER BUT DO FOR DUPES
                xgo = 1
                If InStr(A1$, "\03 MICROSOFT\# VB 6\01 VB6 An MSDN 01") > 0 Then xgo = 0
                If InStr(A1$, "\03 MICROSOFT\# VB 6\01 VB6 An MSDN 02") > 0 Then xgo = 0
                If InStr(A1$, "\03 MICROSOFT\# OP SYS\Test\00 Windows Xp Pro") > 0 Then xgo = 0

                          
                If xgo = 1 Then
                    wex4 = wex4 + 1
                    F.Delete
                End If
                
'                wex4 = wex4 + 1
                ScanPath.lblCount5 = Trim(Str(we2)) + " / " + Str(wex4) + " -- Del - Dupes = " + Format((we2 / we3) * 100, "0.00") + "%"

            End If
            
            Set F = Nothing
            
            yy = yy + " - " + WXHEX$
        End If

    oldfilename = A1$ + B1$

endb:
    
Next


ScanPath.WindowState = vbNormal

End Sub

'***********************************************
'# Check, whether we are in the IDE
Public Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Public Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************

