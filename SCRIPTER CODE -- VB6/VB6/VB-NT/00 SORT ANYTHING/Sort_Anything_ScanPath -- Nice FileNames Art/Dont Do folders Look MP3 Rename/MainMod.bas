Attribute VB_Name = "MainMod"
Public Endo

Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
    
Public Idate As Date

Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2

Public Power






Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long ' only used if FOF_SIMPLEPROGRESS, sets dialog title
End Type

Public Const FO_COPY = &H2 ' Copy File/Folder
Public Const FO_DELETE = &H3 ' Delete File/Folder
Public Const FO_MOVE = &H1 ' Move File/Folder
Public Const FO_RENAME = &H4 ' Rename File/Folder
Public Const FOF_ALLOWUNDO = &H40 ' Allow to undo rename, delete ie sends to recycle bin
Public Const FOF_FILESONLY = &H80  ' Only allow files
Public Const FOF_NOCONFIRMATION = &H10  ' No File Delete or Overwrite Confirmation Dialog
Public Const FOF_SILENT = &H4 ' No copy/move dialog
Public Const FOF_SIMPLEPROGRESS = &H100 ' Does not display file names

Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
                        (lpFileOp As SHFILEOPSTRUCT) As Long

 

Sub Jeepers2()
    
''Load ScanPath
'ScanPath.txtPath = "E:\04 Music ---\03 My Music Zen\"
''ScanPath.txtPath = "E:\My Music\Talking Books"
'ScanPath.cboMask.Text = "*.mp3"

'Call ScanPath.cmdScan_Click

    
    
Dim outInfo As MPEGInfo

ScanPath.List1.AddItem "Syncing MP3 Bangers .."
    ScanPath.List1.Refresh
    DoEvents

'On Local Error Resume Next
'f1 = FreeFile
'Open App.Path + "CRC Such.txt" For Input As #f1
'wdm$ = Input(LOF(f1), f1)
'Close #f1
'On Local Error GoTo 0


Dim InTag As ID3Tag

Dim Fs22
Set Fs22 = CreateObject("Scripting.FileSystemObject")
    
tt = InStr(4, ScanPath.txtPath, "\") + 1
a3$ = Mid$(ScanPath.txtPath, tt)


For we = 1 To ScanPath.ListView1.ListItems.Count

    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ScanPath.ListView1.ListItems.Item(we)
    
    ate7$ = "E:\04 Music ---\04 My Music\"
    ate8$ = "E:\04 Music ---\03 My Music Zen\"

    ate9$ = ate7$ + Mid$(a1$, Len(ate8$) + 1)

    file1$ = a1$ + b1$
    Call FileDateBuster(file1$, Idate)
    tdate1 = Idate
    
    file2$ = ate9$ + b1$
    Call FileDateBuster(file2$, Idate)
    
    'idate=0 file not found
    
    yy$ = "Copy No. "
    If tdate1 > Idate And Idate > 0 Then
'    Stop
    yy$ = "Copy Yes "
    Fs22.CopyFile file1$, file2$


'    Stop
    
    End If
    
    po = po + 1
    
    ScanPath.List1.AddItem yy$ + Trim(Str(po)) + " -- " + file1$
    ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1
    ScanPath.List1.Refresh
    DoEvents


If Power = True Then Exit For

Next
    
If Power = True Then ScanPath.List1.AddItem "Aborted......."
ScanPath.List1.AddItem "Completed......."
ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1
ScanPath.List1.Refresh


Exit Sub
Stop
End

End Sub


Private Sub FileDateBuster(filespec2$, Idate As Date)

Dim fs, f
Set fs = CreateObject("Scripting.FileSystemObject")
Idate = 0

pdate = 0
On Local Error Resume Next
Set f = fs.GetFile((filespec2$))
Idate = f.datelastmodified


End Sub


Sub CopyDir()

'This will copy c:\backup to c:\backup2 and will not show filenames:

Dim op As SHFILEOPSTRUCT
With op
    .wFunc = FO_COPY ' Set function
    .pTo = "C:\backup2" ' Set new path
    .pFrom = "C:\backup" ' Set current path
    .fFlags = FOF_SIMPLEPROGRESS
End With
' Perform operation
SHFileOperation op

End Sub
'Not all functions require all the


