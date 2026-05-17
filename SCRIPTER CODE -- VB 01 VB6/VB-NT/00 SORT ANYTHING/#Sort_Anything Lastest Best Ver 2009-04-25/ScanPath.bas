Attribute VB_Name = "ScanPathMod"
Sub WINRAR_EXTRACT()

'THIS EXTRACT RAR'S IN MULTI PATH FOLDERS



ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath.Text = "M:\01 Installations\Games Installations\"

ScanPath.cboMask.Text = "*.RAR"
ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScanDir_Click
ScanPath.Show

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count))
For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    ScanPath.lblCount2 = Trim(Str(WE))
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    A1 = A1 + B1
    
    ChDrive A1
    ChDir A1
    
    outff = 0

    If Dir(A1 + "\*###.rar") <> "" Then
        'Shell "C:\Program Files\WinRAR\WinRAR.exe E """ + A1 + "\*###.rar""", vbMinimizedNoFocus
        'Kill A1 + "\*###.rar"

        'YY2 = Now + TimeSerial(0, 0, 10)
        'Do
        '    DoEvents
        'Loop Until YY2 < Now

   End If

DoEvents




Next


End

End Sub
