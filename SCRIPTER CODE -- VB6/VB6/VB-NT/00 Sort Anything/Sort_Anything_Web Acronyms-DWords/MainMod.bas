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
    
'MainMody


Dim Tdate1 As Date
Dim Tdate2 As Date
Dim Fs22
Set Fs22 = CreateObject("Scripting.FileSystemObject")

    file1$ = App.Path + "\Acronym Header.html"
    Call FileDateBuster(file1$, Tdate1)
    
    file2$ = App.Path + "\Acronym Header.txt"
    Call FileDateBuster(file2$, Tdate2)

    If Tdate1 > Tdate2 Then Fs22.CopyFile file1$, file2$
    If Tdate1 < Tdate2 Then Fs22.CopyFile file2$, file1$

'ScanPath.txtPath = "D:\MY WEBS\00-Lists-Common-Words\Acronyms.txt\"

Load ScanPath
ScanPath.Show

GoTo Jump

ScanPath.txtPath = "D:\MY WEBS\00-Lists-Common-Words\Acronyms.txt\"

ScanPath.cboMask.Text = "*.txt"

Call ScanPath.cmdScan_Click

    
    
DoEvents



'Dim Fs22
'Set Fs22 = CreateObject("Scripting.FileSystemObject")
    
'tt = InStr(4, ScanPath.txtPath, "\") + 1
'a3$ = Mid$(ScanPath.txtPath, tt)

Open file2$ For Input As #1
tt$ = Input(LOF(1), 1)
Close #1

file3$ = App.Path + "\statcounter.txt"
Open file3$ For Input As #1
statcounter$ = Input(LOF(1), 1)
Close #1





r44 = InStr(tt$, "Last Update") + 12

Mid$(tt$, r44, Len("mmm DD-MM-YYYY HH-MM-SS")) = Format$(Now, "DDD DD-MM-YYYY HH:MM:SS")

'Open file2$ For Output As #1
'Print #1, tt$
'Close #1


For we = 1 To ScanPath.ListView1.ListItems.Count
again:
'we = we + 1
    po = 0
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ScanPath.ListView1.ListItems.Item(we)
    
        If LCase$(b1$) = "key-.txt" Then Stop
'    Reset
    f2 = FreeFile
    Open a1$ + b1$ For Input As #f2
    If LOF(1) > 100000 Then we = we + 1:     Close #f2: GoTo again
    Close #f2
    ScanPath.List1.AddItem Format$(we, "0000") + " --- " + a1$ + b1$
    ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1
    
    DoEvents

    hap$ = ""
    f1 = FreeFile
    Open a1$ + b1$ For Input As #f1
    
    Do
    
    '#@**************---------------------------------************
    Line Input #f1, ll$
        
    DoEvents
    If EOF(f1) Then Exit Do
    Loop Until ll$ = "#@**************---------------------------------************"
        
    If EOF(f1) Then
        we = we + 1
        Close #f1
        GoTo again
    End If
    Do
    DoEvents
    
    
    If Not (EOF(1)) Then Line Input #1, ll$
    po = po + 1
        
    DoEvents
    If InStr(ll$, "<") Then
hh$ = ll$
   
   Do
    DoEvents
    Y1 = InStr(ll$, "<")
    Y2 = InStr(Y1 + 1, ll$, ">")
     X1 = InStr(ll$, " ")
     
    
    If X1 > Y1 And Y1 > 0 Then
       
        'tR$ = Mid$(ll$, 1, Y1 - 2) + " " + Mid$(ll$, X1 + 1)
        'll$ = tR$
    End If
    
     
     If Y1 = 0 Then Y2 = 0
       
    If Y2 = 0 And Y1 > 0 Then
        Mid$(ll$, Y1, 1) = "-"
    End If
    If Y1 > 0 And Y2 > 0 Then
        If X1 > Y1 Then
            tR$ = Mid$(ll$, 1, Y1 - 1) + " " + Mid$(ll$, X1)
        
        Else
            tR$ = Mid$(ll$, 1, Y1 - 1) + Mid$(ll$, Y2 + 1)
        
        End If
        
        ll$ = tR$
    End If
    Loop Until Y1 = 0
   
   Do
    DoEvents
    Y1 = InStr(ll$, ">")
       
    If Y1 > 0 Then
        tR$ = Mid$(ll$, 1, Y1 - 1) + " " + Mid$(ll$, Y1 + 1)
        ll$ = tR$
    End If
    Loop Until Y1 = 0
         
    '     ScanPath.List1.AddItem ll$
    ' ScanPath.List1.AddItem hh$
     
     End If
       
     hap$ = hap$ + Trim(ll$) + "<br>" + vbCrLf
               
        
    Loop Until EOF(1)
    Close #1
    
    
    If hap$ <> "" Then
    yl = InStrRev(b1$, ".")
    c1$ = Mid$(b1$, 1, yl - 1) + ".html"
    c1$ = LCase(c1$)
    Mid$(c1$, 1, 1) = UCase$(Mid$(c1$, 1, 1))
    
    r44 = InStr(tt$, ">Last Update") + 13

    Mid$(tt$, r44, Len("ddd DD-MM-YYYY HH-MM-SS")) = Format$(Now, "DDD DD-MM-YYYY HH:MM:SS")
   
    
    we1$ = "D:\MY WEBS\00-Lists-Common-Words\HTM\"
    Open we1$ + c1$ For Output As #1
    
hap$ = hap$ + "<BR>" + Str$(po) + " Lines Displayed<BR>" + vbCrLf
    
    
    Print #1, tt$ + vbCrLf + hap$ + vbCrLf + statcounter$
    Close #1
    End If
    
If Power = True Then Exit For

Next
    
If Power = True Then ScanPath.List1.AddItem "Aborted......."
ScanPath.List1.AddItem "Completed......."
ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1
ScanPath.List1.Refresh

'Jump:


ScanPath.txtPath = "D:\MY WEBS\00-Lists-Common-Words\htm\"

ScanPath.cboMask.Text = "*.html"
    
Call ScanPath.cmdScan_Click

For we = 1 To ScanPath.ListView1.ListItems.Count
    po = 0
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ScanPath.ListView1.ListItems.Item(we)
    
    
    ScanPath.List1.AddItem Format$(we, "0000") + " --- " + a1$ + b1$
    ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1

    '    If LCase$(b1$) = "key-.html" Then Stop
   
    file1$ = a1$ + b1$
    
    file2$ = "D:\# MY DOCS\httpdocs\LoveFolder\Cool\D-Words\All_of_Them_2_Letter\"
    file3$ = "D:\# MY DOCS\httpdocs\LoveFolder\Cool\D-Words\All_of_Them_3_Letter\"
   
    file7$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\D-Words\All_of_Them_2_Letter\"
    file8$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\D-Words\All_of_Them_3_Letter\"

    file9$ = "D:\MY WEBS\00-Lists-Common-Words\HTM\All_of_Them_2_Letter\"
    file10$ = "D:\MY WEBS\00-Lists-Common-Words\HTM\All_of_Them_3_Letter\"

    file4$ = file3$
    file5$ = file8$
    file11$ = file10$
    If Mid$(b1$, 3, 2) = "--" Then
        file4 = file2$: file5$ = file7$: file11$ = file9$
    End If
    If Mid$(b1$, 4, 1) = "-" Then
     '   file4 = file2$: file5$ = file7$: file11$ = file9$
    End If
    
    file4$ = file4$ + b1$
    file5$ = file5$ + b1$
    file11$ = file11$ + b1$
    
    'If LCase$(b1$) = "key-.html" Then Stop
    
file14$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\D-Words\" + b1$
If Dir$(file14$) <> "" Then
    
    'Kill file14$
    Fs22.CopyFile file1$, file14$
    
End If

    Fs22.CopyFile file1$, file5$
'    Fs22.moveFile file1$, file11$
DoEvents
    'ScanPath.txtPath = "D:\MY WEBS\00-Lists-Common-Words\Acronyms.txt\"

Next
ScanPath.List1.AddItem "...........End.........."

Jump:


file7$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\D-Words\"


ScanPath.txtPath = file7$

ScanPath.cboMask.Text = "*.html"
    
Call ScanPath.cmdScan_Click

For we = 1 To ScanPath.ListView1.ListItems.Count
    
    po = 0
again2:
    
Do
    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    b1$ = ScanPath.ListView1.ListItems.Item(we)
x25 = 1

If we > ScanPath.ListView1.ListItems.Count Then Exit Do
If InStr(a1$, "_vti") > 0 Then we = we + 1: GoTo again2
x25 = 0
Loop Until x25 = 0

If x25 = 1 Then Exit For

    ScanPath.List1.AddItem Format$(we, "0000") + " --- " + a1$ + b1$
    ScanPath.List1.ListIndex = ScanPath.List1.ListCount - 1

file2$ = "D:\# MY DOCS\httpdocs\LoveFolder\Cool\D-Words\"

tte = InStr(a1$, "D-Words")
If tte = 0 Then MsgBox "Not Called D-words anymore Acronyms"

file3$ = Mid$(a1$, tte + 8)
DoEvents

    Fs22.CopyFile a1$ + b1$, file2$ + file3$ + b1$

DoEvents


Next


End
Exit Sub


End

Exit Sub
Stop
End

End Sub


Private Sub FileDateBuster(filespec2$, Idate As Date)
    
    'idate=0 file not found

Dim fs, f
Set fs = CreateObject("Scripting.FileSystemObject")
Idate = 0

pdate = 0
On Local Error Resume Next
Set f = fs.GetFile((filespec2$))
Idate = f.datelastmodified

'On Local Error Resume Next
'f1 = FreeFile
'Open App.Path + "CRC Such.txt" For Input As #f1
'wdm$ = Input(LOF(f1), f1)
'Close #f1
'On Local Error GoTo 0

'    file1$ = a1$ + b1$
'    Call FileDateBuster(file1$, Idate)
'    tdate1 = Idate

'    file2$ = ate9$ + b1$
'    Call FileDateBuster(file2$, Idate)

'    yy$ = "Copy No. "
'    If tdate1 > Idate And Idate > 0 Then
'    Stop
'    yy$ = "Copy Yes "
'    Fs22.CopyFile file1$, file2$

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


