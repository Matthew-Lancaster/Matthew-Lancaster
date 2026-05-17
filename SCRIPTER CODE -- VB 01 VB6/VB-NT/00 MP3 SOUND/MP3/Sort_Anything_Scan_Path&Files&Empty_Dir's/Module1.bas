Attribute VB_Name = "Module1"


Sub Runner()


ra$ = "E:\04 Music ---\"

'Load ScanPath
FrmMain.txtPath = ra$
'ScanPath.txtPath = "E:\04 Music ---\03 My Music Zen\"
'ScanPath.txtPath = "E:\My Music\Talking Books"
FrmMain.cboMask.Text = "*.mp3"

Call FrmMain.cmdScan_Click
'Call FrmMain.cmdScanDir_Click



'cmdScanDir_Click


'cmdScan_Click

    
    
    





End Sub
