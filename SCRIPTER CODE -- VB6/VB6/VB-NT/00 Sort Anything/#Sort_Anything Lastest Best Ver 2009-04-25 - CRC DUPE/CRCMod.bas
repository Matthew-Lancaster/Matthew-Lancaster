Attribute VB_Name = "CRCMod"
Public m_CRC As clsCRC, OFSI1

'Put in Sub Of Code

Sub CRCDUPE()

ScanPath.Show

ScanPath.lblCount1 = ""
ScanPath.lblCount2 = ""
ScanPath.lblCount3 = ""
ScanPath.lblCount4 = ""
ScanPath.lblCount5 = ""
ScanPath.lblCount6 = ""


Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32



ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask.Text = "*.jpg;*.png;*.gif;*.ico;*.3gp;*.avi;*.mp4"


If Fs.FolderExists(ScanPath.txtPath) = False Then Exit Sub


'ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count))
'ScanPath.txtPath = "M\0 00 Art Loggers\Temp Inet Files Gif Ico Png\"
Call ScanPath.cmdScan_Click


'sort of FILE then PATH ' reverse - most important at front
ScanPath.ListView1.SortOrder = lvwDescending
ScanPath.ListView1.SortKey = 1
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.SortKey = 0
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False

'Sort On Size
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 3
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False
    
wx8 = Val(ScanPath.ListView1.ListItems.Count)
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " Total"

'SORT SAME SIZE FILES
For we = ScanPath.ListView1.ListItems.Count To 2 Step -1
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count))
    ScanPath.lblCount3 = Trim(Str(we))
    
    FSI1 = ScanPath.ListView1.ListItems.Item(we).SubItems(3)
    fsi2 = ScanPath.ListView1.ListItems.Item(we - 1).SubItems(3)
    'Debug.Print Str(WE) + " --- " + FSI1 + " ---- " + fsi2
    If Val(FSI1) <> Val(fsi2) Then
        ScanPath.ListView1.ListItems.Remove (we)
    End If
    If Val(FSI1) = Val(fsi2) Then
        we = we - 1
    End If
    If Len(FSI1) > TINK Then TINK = Len(FSI1)
Next

'PAD ZEROS IN FRONT BASD ON BIZZEST SIZE
For we = 1 To ScanPath.ListView1.ListItems.Count
    FSI1 = ScanPath.ListView1.ListItems.Item(we).SubItems(3)
    TY = String(TINK - Len(FSI1), "0") + FSI1
    
    ScanPath.ListView1.ListItems.Item(we).SubItems(3) = TY
Next

'Exit Sub


'Sort On Size - RESORT AFTER PADDED
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 3
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False

'DEL ZERO SIZE FILES AND SMALL
For we = 1 To ScanPath.ListView1.ListItems.Count
    DoEvents
    ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count))
    ScanPath.lblCount3 = Trim(Str(we))
    FSI1 = ScanPath.ListView1.ListItems.Item(we).SubItems(3)
    'If Val(FSI1) > 70 Then Stop
    If Val(FSI1) >= 70 Then Exit For
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
'    FileExt = Mid$(G1, Len(G1) - 3)
        
    we8 = we8 + 1
    ScanPath.lblCount5 = Trim(Str(we8)) + " Del as Too Small <70 Bytes"
    
    If ScanPath.Label5.BackColor <> ScanPath.Label6.BackColor Then Kill A1$ + G1$
    
    ScanPath.ListView1.ListItems.Remove (we)
    we = we - 1

Next



'Exit Sub



ScanPath.Show
Reset
YY = "--"
    
'ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count))

For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    ScanPath.lblCount3 = Trim(Str(we))
    
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    
    If InStr(B1$, ".rar") > 0 Then
        ScanPath.ListView1.ListItems.Remove (we)
    End If

Next

we4 = ScanPath.ListView1.ListItems.Count
ScanPath.lblCount2 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " After Same Size"

we2 = 0
ScanPath.lblCount4 = Trim(Str(we2)) + " Del - Dupes = " + Format((we2 / wx8) * 100, "0.00") + "%"

we = 1
we2 = ScanPath.ListView1.ListItems.Count
On Error Resume Next
Do

'For we = 1 To ScanPath.ListView1.ListItems.Count
    
    we2 = we2 - 1
    
    ScanPath.lblCount3 = Trim(Str(we2))
    
    FSI1 = ScanPath.ListView1.ListItems.Item(1).SubItems(3)
    If FSI1 > 5000 Then ScanPath.lblCount3.Refresh
    If FSI1 > 50000 Then DoEvents
    
    'DoEvents
    
    If we2 Mod 10 = 0 Then ScanPath.lblCount3.Refresh
    If we2 Mod 50 = 0 Then DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(1).SubItems(2)
    A1$ = ScanPath.ListView1.ListItems.Item(1).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(1)
    
    
    
    If FSI1 <> OFSI1 Then
        YY = ""
    End If
    OFSI1 = FSI1

    ScanPath.ListView1.ListItems.Remove (we)



    Err.Clear
    WXHEX$ = Hex(m_CRC.CalculateFile(A1$ + G1$))
    
    
    If InStr(YY, WXHEX$) > 0 And Err.Number = 0 Then
        we3 = we3 + 1

        ScanPath.lblCount4 = Trim(Str(we3)) + " Del - Dupes = " + Format((we3 / we4) * 100, "0.00") + "%"
        If ScanPath.Label5.BackColor <> ScanPath.Label6.BackColor Then
            'If A1$ = "M:\0 00 Art Loggers\# APP AND SCREEN -- SHOT\Hot-Key-App-Shots\" Then
                Fs.DeleteFile A1$ + G1$
            'End If
        End If
    End If
    
    YY = YY + " - " + WXHEX$
   
   
    
'Next
Loop Until ScanPath.ListView1.ListItems.Count = 0


ScanPath.lblCount6 = "DONE *"
ScanPath.WindowState = 0




End Sub


