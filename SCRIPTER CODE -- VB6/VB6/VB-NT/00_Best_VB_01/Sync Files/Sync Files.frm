VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sync Files"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   2070
      TabIndex        =   1
      Top             =   2415
      Width           =   4215
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   210
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   405
      Width           =   6555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RS()
Dim A1 As String, B1 As String
Dim AX1 As String, BX1 As String

Dim QQ1 As Long, F1$, Tx1 As String, Tx2 As String
Private Declare Function TouchFileTimes Lib "imagehlp.dll" (ByVal FileHandle As Long, ByRef pSystemTime As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As BOOL
Enum BOOL
   FALSE_ = 0
   TRUE_ = 1
End Enum
Const MAX_PATH = 260

Private Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type


Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Type SYSTEMTIME
    wYear             As Long
    wMonth            As Long
    wDayOfWeek        As Long
    wDay              As Long
    wHour             As Long
    wMinute           As Long
    wSecond           As Long
    wMilliseconds     As Long
End Type



Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15


Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long




Private Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Private Sub Form_Load()

On Error Resume Next

If App.PrevInstance = True Then End

If IsIDE = False Then
    'Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" /runexit """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
    'End
End If


Set FS = CreateObject("Scripting.FileSystemObject")
'Set F = FS.GETDRIVE("M")
'If F.ISREADY = True Then
'    'FS.MOVEFILES DR + ":\DCIM\100MSDCF", "D:\# W880i\VIDEO - ASUS"
''    Call ShellFileMove("D:\00 Art\01 Loggers\*.*", "M:\0 00 Art Loggers\")
'End If



'ERROR IN THE SORTED SCAN TO DEBUG
'Call KATMP3


'Call M2CARD

'End


'Shell """D:\VB6\VB-NT\00 Sort Anything\#Sort_Anything Lastest Best Ver 2009-04-25 - MULTI USE MENU\Sort_AnyThing - MULTI USE MENU.EXE"" MOVE_IN_DATE_FOLDERS_M2CARD", vbMinimizedNoFocus

'
'
'Call EMAIL_RAPID_LAST_3_MONTHS
'Call EMAIL_RAPID_LAST_2_MONTHS
'Call EMAIL_RAPID_LAST_3_MONTHS_JUST_MINE
'
'AX1 = "D:\00 Art\01 Loggers\Web_Cam"
'BX1 = "M\0 00 Art Loggers\Web_Cam"
'
''FS.COPYFOLDER A1, B1
'Call TestCOPY_IF_EXISTS(AX1, BX1)


'BOOTLOGG ARCHIVE

'C:\WINDOZE\ntbtlog.txt
On Error Resume Next

MkDir "C:\WINDOWS BOOT UP Archive Loggs"
'On Error GoTo 0


Set F = FS.GETFILE("C:\WINDOWS BOOT UP Archive Loggs\## ntbtlog02.txt")

IDATE01 = F.datelastmodified
On Error GoTo 0

'Set F = FS.GETFILE("C:\WINDOZE\ntbtlog.txt")
Set F = FS.GETFILE("C:\WINDOWS\ntbtlog.txt")
IDATE02 = F.datelastmodified
If IDATE01 <> IDATE02 Then
    F.Copy "C:\WINDOWS BOOT UP Archive Loggs\ntbtlog " + Format(IDATE02, "YYYY-MM-DD - hh-mm-ss") + ".txt"
    F.Copy "C:\WINDOWS BOOT UP Archive Loggs\## ntbtlog02.txt"
End If




'GetSpecialFolder_Show_Script_Debug (1)


i = 0
ReDim RS(200)
i = i + 1: RS(i) = GetSpecialfolder(6) + "\#00 Palm 01\#000 New"
i = i + 1: RS(i) = GetSpecialfolder(6) + "\#00 Palm 02\xWiki"
i = i + 1: RS(i) = "D:\#0 1 INSTALLATIONS\Downloads"
i = i + 1:  RS(i) = "C:\Program Files\# NO INSTALL REQUIRED\DOWNLOADS"
i = i + 1: RS(i) = GetSpecialfolder(5) + "\Downloads"
i = i + 1:  RS(i) = "X:\DOCUMENTS\"
'i = i + 1: RS(i) = "T:\DOCUMENTS\"

ReDim Preserve RS(i)
For i = 1 To UBound(RS)

If FS.FolderExists(RS(i)) Then

    Dir1.Path = RS(i)
    
    For R = 0 To Dir1.ListCount - 1
    
    'Debug.Print Dir1.List(R)
    
        yg = Mid(Dir1.List(R), InStrRev(Dir1.List(R), "\"))
        
        'If InStr(Dir1.List(R), "# 09 USENET") Then Stop
        
        If Mid$(yg, 1, 3) = "\# " Then  'SPACE
        
        
        
            'TO CREATE BEFORE THAN USE HERE
            For R2 = DateSerial(Year(Now), Month(Now) - 2, 1) To Now
                'R2 = Now'      REM LINE
                '----------
                
                ii = Dir1.List(R)
                II2 = "\" + Format(R2, "YYYY-0MM MMM")
                
                
                If FS.FolderExists(ii) = False Then MkDir ii
                If FS.FolderExists(ii + II2) = False Then
                    MkDir ii + II2
                End If
            Next R2
        End If
        
        
        Next
    
    'Me.Show
    'Exit Sub
    End If
Next i

End
'Exit Sub

'End
'Call Test2
'Exit Sub



'For R = 3 To 26
'    On Error Resume Next
''    Set z = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(Chr$(r + 64) + ":")))
'    Set z = FS.GETDRIVE(FS.GetDriveName(Chr$(R + 64) + ":"))
'    a = ""
'    a = z.VolumeName
'    If Err.Number > 0 Then a = ""
'    On Error GoTo 0
'
'    'REM OUT
'    'If a = "UDISK 8GB" Then h1 = Chr$(r + 64)
'    'If a = "W880i" Then h2 = Chr$(r + 64)
'Next

'Call Test("D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\# DATA\00_ClipBoard_Tot_Strip.txt", "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Word_Lists\D_Words_Text\")
A1 = "D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\# DATA\55-88-HAPPY\00_ClipBoard_Tot_Strip.txt"
B1 = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\00_ClipBoard_Tot_Strip-55-88-HAPPY.txt"
ShellFileCopy A1, B1

'fs.COPYFILE A1, B1


A1 = "D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\# DATA\MATT-555ROIDS\00_ClipBoard_Tot_Strip.txt"
B1 = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\00_ClipBoard_Tot_Strip-MATT-555ROIDS.txt"
ShellFileCopy A1, B1

'fs.COPYFILE A1, B1

A1 = "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\00_Text_Data_KeyLogg\Key Logger Text-Stripe.txt"
B1 = "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\00_Text_Data_KeyLogg\00_ClipBoard_Tot_Strip.txt"

ShellFileCopy A1, B1
'fs.COPYFILE A1, B1

A1 = "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\URL Logging.txt"
B1 = "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\00_Text_Data_KeyLogg\URL Logging.txt"

ShellFileCopy A1, B1

'fs.COPYFILE A1, B1




'Call Test5("X:\00 Lists-Common-Words\Acronyms.txt\", "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Word_Lists\D_Words_Text\")
'Call Test4("X:\00 Lists-Common-Words\Double Word Examples Text", "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Word_Lists\Word_Usage_Examples")

'NEVER
'If h1 <> "" Then
'    Call Test4(h1 + ":\Other\Word-Net\", "D:\# MY DOCS\# 01 My Documents\CALLerID\Text-9\Word-Net\")
'    Call Test4(h1 + ":\Other\Word-Net\", "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Word_Lists\Word_Net\")
'
'    'Call Test4(h1 + ":\Other\", "D:\# W880i\Other\")
'
''    Tx1 = h1 + ":\Music\02 Smalls\Smalls Mobile Rings"
''    Tx2 = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\Techno_Voice_Sample_Clips"
''    Call Test4(Tx1, Tx2)
'End If

End

Exit Sub

If h2 <> "" Then
    Tx1 = h2 + ":\Music\02 Smalls\Smalls Mobile Rings"
    Tx2 = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\Techno_Voice_Sample_Clips"
    Call Test4(Tx1, Tx2)

    Tx1 = h2 + ":\Other\"
    Tx2 = "D:\# W880i\Other\"
    Call Test4(Tx1, Tx2)
End If

End


Call Test1

Exit Sub

 
End Sub




Function SpecialFoldersAre()
    LP_RESULT_CALLBACK = MsgBox("Start menu folder is located at : " + GetSpecialfolder(CSIDL_STARTMENU), vbInformation, "Start Menu Folder")
    LP_RESULT_CALLBACK = MsgBox("Favorites folder is located at : " + GetSpecialfolder(CSIDL_FAVORITES), vbInformation, "Favorites Folder")
    LP_RESULT_CALLBACK = MsgBox("Programs folder is located at : " + GetSpecialfolder(CSIDL_PROGRAMS), vbInformation, "Programs Menu Folder")
    LP_RESULT_CALLBACK = MsgBox("Desktop folder is located at : " + GetSpecialfolder(CSIDL_DESKTOP), vbInformation, "Desktop Folder")
    End
End Function




Function GetSpecialfolder(ByVal CSIDL As Long) As String
    Dim R As Long
    Dim IDL As ITEMIDLIST
    R = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If R = NOERROR Then
        Path$ = Space$(512)
        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function
'GetSpecialfolder(6) +



Sub CHK_M2_DRIVES_AND_MOVE_FILES()

On Local Error Resume Next
Set dc = FS.Drives
For Each D In dc
    Err.Clear
    DR = D.DriveLetter
    
    If D.DriveType = 3 Then
        If D.ISREADY = True Then
            If FS.FolderExists(DR + ":\DCIM\100MSDCF") Then
                    
                If Dir(DR + ":\DCIM\100MSDCF\*.*") <> "" Then
                    'FS.MOVEFILES DR + ":\DCIM\100MSDCF", "D:\# W880i\VIDEO - ASUS"
                    Call ShellFileMove(DR + ":\DCIM\100MSDCF\*.*", "D:\# W880i\VIDEO - ASUS\")
                End If
    
            End If
        End If
    End If


    'Select Case d.DriveType
    'Case 0: t = "Unknown"
    'Case 1: t = "Removable"
    'Case 2: t = "Fixed"
    'Case 3: t = "Network"
    'Case 4: t = "CD-ROM"
    'Case 5: t = "RAM Disk"

    
Next

End


End Sub


Sub M2CARD()

'Call CHK_M2_DRIVES_AND_MOVE_FILES

ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask.Text = "*.3GP"

Dim DsTx

Tx = "D:\# W880i\"
ScanPath.txtPath.Text = Tx
Call ScanPath.cmdScan_Click

Set F = FS.GETDRIVE("M")
If F.ISREADY = True Then
    DsTx = "M:\# W880i\VIDEO\"
    Tx = "M:\# W880i\"
    ScanPath.txtPath.Text = Tx
    Call ScanPath.cmdScan_Click
Else
    DsTx = "D:\# W880i\VIDEO\"
End If

'If d.ISREADY = True Then





For we = 1 To ScanPath.ListView1.ListItems.Count
    
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
      
    Set F = FS.GETFILE(A1 + B1)
    
    ADATE1 = F.datelastmodified
    
    'II2 = "\" + Format(F.datelastmodified, "YYYY-0MM MMM DD")
    II2 = Format(ADATE1, "YYYY") + " 0" + Format(ADATE1, "MM") + "(" + Format(ADATE1, "MMM") + ") " + Format(ADATE1, "DD DDD")

    
    ii = DsTx
    If FS.FolderExists(ii) = False Then MkDir ii
    If FS.FolderExists(ii + II2) = False Then
        MkDir ii + II2
    End If
    

    
    'F.Move ii + II2 + "\" + B1
    If FS.FileExists(ii + II2 + "\" + B1) = False And FileInUse(A1 + B1) = False Then
        
        
        Call ShellFileMove(A1 + B1, ii + II2 + "\" + B1)
    End If
      
Next


'Shell """D:\VB6\VB-NT\00 Sort Anything\#Sort_Anything Lastest Best Ver 2009-04-25 - MULTI USE MENU\Sort_AnyThing - MULTI USE MENU.EXE"" MOVE_IN_DATE_FOLDERS_M2CARD", vbMinimizedNoFocus


'-------------------------------



End Sub


Sub KATMP3()

ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask.Text = "*.MP3"
ScanPath.DTPicker1(0).Value = Now - 10
ScanPath.txtPath.Text = "M:\05 Media\KAT RECORD"
If FS.FolderExists(ScanPath.txtPath.Text) = "" Then Stop

Call ScanPath.cmdScan_Click




'COPY FILES
For we = 1 To ScanPath.ListView1.ListItems.Count
    
    'Label13 = ScanPath.ListView1.ListItems.Count
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
        
    Err.Clear
    Set F = FS.GETFILE(A1 + B1)
    
    If F.Size <= 1253 And Err.Number = 0 Then
        Kill A1 + B1
    End If
Next

ScanPath.DTPicker1(0).Value = Nothing

End Sub


Sub EMAIL_RAPID_LAST_3_MONTHS_JUST_MINE()


'2009-12-Dec


ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask.Text = "*.TXT;*.EML"
Tx = "M:\00 Loggers Text\#00 EMAIL DATA POP3\"
T = DateSerial(Year(Now), Month(Now), 1)
TTY = Tx + Format(T, "YYYY-MM-MMM")
ScanPath.txtPath.Text = TTY
Call ScanPath.cmdScan_Click

T = DateSerial(Year(Now), Month(Now) - 1, 1)
TTY = Tx + Format(T, "YYYY-MM-MMM")
ScanPath.txtPath.Text = TTY
Call ScanPath.cmdScan_Click

'T = DateSerial(Year(Now), Month(Now) - 2, 1)
'TTY = Tx + Format(T, "YYYY-MM-MMM")
'ScanPath.txtPath.Text = TTY
'Call ScanPath.cmdScan_Click

P2 = "M:\00 Loggers Text\#00 EMAIL DATA POP3\z--EMAIL_RAPID_SEARCH_LAST_3_MONTHS_JUST_MINE\"

If FS.FolderExists(P2) = False Then
    MkDir P2
End If

'COPY FILES
For we = 1 To ScanPath.ListView1.ListItems.Count
    
    'Label13 = ScanPath.ListView1.ListItems.Count
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
        
    
    XXI = 0
    ii = InStr(B1, "EMAIL-matt.lan@btinternet.com-")
    If ii > 0 Then XXI = 1
    ii = InStr(B1, "SUBJECT-MattL.com")
    If ii > 0 Then XXI = 0


    If FS.FileExists(P2 + B1) = False And XXI = 1 Then
        FS.COPYFILE A1 + B1, P2 + B1
    End If

Next







ScanPath.ListView1.ListItems.Clear

ScanPath.txtPath.Text = P2


'ScanPath.Show


Call ScanPath.cmdScan_Click





'DEL OLDS
For we = 1 To ScanPath.ListView1.ListItems.Count
    
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
        
    'Email--2009-11-01
    TTY = DateValue(Mid(B1, 8, 10))
    
    If TTY < DateSerial(Year(Now), Month(Now) - 2, 1) Then
        Kill A1 + G1$
    End If

Next




End Sub




Sub EMAIL_RAPID_LAST_3_MONTHS()


'2009-12-Dec


ScanPath.ListView1.ListItems.Clear

ScanPath.cboMask.Text = "*.TXT;*.EML"

Tx = "M:\00 Loggers Text\#00 EMAIL DATA POP3\"
T = DateSerial(Year(Now), Month(Now), 1)
TTY = Tx + Format(T, "YYYY-MM-MMM")
ScanPath.txtPath.Text = TTY
Call ScanPath.cmdScan_Click

T = DateSerial(Year(Now), Month(Now) - 1, 1)
TTY = Tx + Format(T, "YYYY-MM-MMM")
ScanPath.txtPath.Text = TTY
Call ScanPath.cmdScan_Click

'T = DateSerial(Year(Now), Month(Now) - 2, 1)
'TTY = TX + Format(T, "YYYY-MM-MMM")
'ScanPath.txtPath.Text = TTY
'Call ScanPath.cmdScan_Click

P2 = "M:\00 Loggers Text\#00 EMAIL DATA POP3\z--EMAIL_RAPID_SEARCH_LAST_3_MONTHS\"

'COPY FILES
For we = 1 To ScanPath.ListView1.ListItems.Count
    
    'Label13 = ScanPath.ListView1.ListItems.Count
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
        

    If FS.FileExists(P2 + B1) = False Then
        FS.COPYFILE A1 + B1, P2 + B1
    End If

Next







ScanPath.ListView1.ListItems.Clear

ScanPath.txtPath.Text = P2


'ScanPath.Show


Call ScanPath.cmdScan_Click





'DEL OLDS
For we = 1 To ScanPath.ListView1.ListItems.Count
    
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
        
    'Email--2009-11-01
    TTY = DateValue(Mid(B1, 8, 10))
    
    If TTY < DateSerial(Year(Now), Month(Now) - 2, 1) Then
        Kill A1 + G1$
    End If

Next




End Sub





Sub EMAIL_RAPID_LAST_2_MONTHS()


'2009-12-Dec


ScanPath.ListView1.ListItems.Clear

ScanPath.cboMask.Text = "*.TXT;*.EML"

Tx = "M:\00 Loggers Text\#00 EMAIL DATA POP3\"
T = DateSerial(Year(Now), Month(Now), 1)
TTY = Tx + Format(T, "YYYY-MM-MMM")
ScanPath.txtPath.Text = TTY
Call ScanPath.cmdScan_Click

T = DateSerial(Year(Now), Month(Now) - 1, 1)
TTY = Tx + Format(T, "YYYY-MM-MMM")
ScanPath.txtPath.Text = TTY
Call ScanPath.cmdScan_Click

'T = DateSerial(Year(Now), Month(Now) - 2, 1)
'TTY = TX + Format(T, "YYYY-MM-MMM")
'ScanPath.txtPath.Text = TTY
'Call ScanPath.cmdScan_Click

P2 = "M:\00 Loggers Text\#00 EMAIL DATA POP3\z--EMAIL_RAPID_SEARCH_LAST_2_MONTHS\"
If FS.FolderExists(P2) = False Then MkDir P2

'COPY FILES
For we = 1 To ScanPath.ListView1.ListItems.Count
    
    'Label13 = ScanPath.ListView1.ListItems.Count
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
        

    If FS.FileExists(P2 + B1) = False Then
        FS.COPYFILE A1 + B1, P2 + B1
    End If

Next







ScanPath.ListView1.ListItems.Clear

ScanPath.txtPath.Text = P2


'ScanPath.Show


Call ScanPath.cmdScan_Click





'DEL OLDS
For we = 1 To ScanPath.ListView1.ListItems.Count
    
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
        
    'Email--2009-11-01
    TTY = DateValue(Mid(B1, 8, 10))
    
    If TTY < DateSerial(Year(Now), Month(Now) - 1, 1) Then
        Kill A1 + G1$
    End If

Next




End Sub




Sub Test4(Dir1$, Dir2$)


On Error Resume Next

If Mid$(Dir1$, Len(Dir1$), 1) <> "\" Then Dir1$ = Dir1$ + "\"
If Mid$(Dir2$, Len(Dir2$), 1) <> "\" Then Dir2$ = Dir2$ + "\"


Dim FS, Idate As Date, Hdate As String, File1 As String, File2 As String
Set FS = CreateObject("Scripting.FileSystemObject")

For IY = 1 To 2

If IY = 2 Then
    ig$ = Dir1$
    Dir1$ = Dir2$
    Dir2$ = ig$
End If

ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath.Text = Dir1$
ScanPath.cboMask.Text = "*.txt;*.mp3"

Call ScanPath.cmdScan_Click

For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    
    Label13 = ScanPath.ListView1.ListItems.Count
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
    
    'Spark of Life
    
    Call Date_File(A1 + B1, Idate, Hdate)
    xdate = Idate
    F1$ = Dir2$
    If Len(A1) <> Len(Dir1$) Then
        F1$ = Dir2$ + Mid$(A1, Len(Dir1$) + 1)
    End If
    ii = MakeDir(F1$)
    'err.number
    If ii = 76 Then Exit Sub
    Call Date_File(F1$ + B1, Idate, Hdate)
    If xdate > Idate Then
    FS.COPYFILE A1 + B1, F1$ + B1
    End If
Next
Next

End Sub





'#SOURCE DESTINATION ONLY ONE WAY
Sub TestCOPY_IF_EXISTS(Dir1$, Dir2$)


On Error Resume Next

If Mid$(Dir1$, Len(Dir1$), 1) <> "\" Then Dir1$ = Dir1$ + "\"
If Mid$(Dir2$, Len(Dir2$), 1) <> "\" Then Dir2$ = Dir2$ + "\"


Dim FS, Idate As Date, Hdate As String, File1 As String, File2 As String
Set FS = CreateObject("Scripting.FileSystemObject")

ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath.Text = Dir1$
ScanPath.cboMask.Text = "*.*"

Call ScanPath.cmdScan_Click

For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    
    Label13 = ScanPath.ListView1.ListItems.Count
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
    
    'Spark of Life
    
    F1$ = Dir2$
    If Len(A1) <> Len(Dir1$) Then
        F1$ = Dir2$ + Mid$(A1, Len(Dir1$) + 1)
    End If
    
    If FS.FolderExists(F1$) = False Then ii = MakeDir(F1$):    If ii = 76 Then Exit Sub
    
    If FS.FileExists(F1$ + B1) = False Then
        FS.COPYFILE A1 + B1, F1$ + B1
    End If
Next

End Sub


Sub Test5(Dir1$, Dir2$)


On Error Resume Next


If Mid$(Dir1$, Len(Dir1$), 1) <> "\" Then Dir1$ = Dir1$ + "\"
If Mid$(Dir2$, Len(Dir2$), 1) <> "\" Then Dir2$ = Dir2$ + "\"

Dim FS, Idate As Date, Hdate As String, File1 As String, File2 As String
Set FS = CreateObject("Scripting.FileSystemObject")

For IY = 1 To 1

If IY = 2 Then
    ig$ = Dir1$
    Dir1$ = Dir2$
    Dir2$ = ig$
End If

ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath.Text = Dir1$

ScanPath.cboMask.Text = "*.txt;*.mp3"

Call ScanPath.cmdScan_Click

For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
    
    Label13 = ScanPath.ListView1.ListItems.Count
    DoEvents
    
    G1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(2)
    A1 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(we)
    
    'Spark of Life
    
    Call Date_File(A1 + B1, Idate, Hdate)
    xdate = Idate
    F1$ = Dir2$
'    If Len(A1) <> Len(Dir1$) Then
'        F1$ = Dir2$ + Mid$(A1, Len(Dir1$) + 1)
'    End If
'    MakeDir (F1$)
    ha = 0
    If InStr(B1, "--.txt") Then ha = 1: tu = "2-Letter\"
    If InStr(B1, "-.txt") And ha = 0 Then: tu = "3-Letter\"
    Call Date_File(F1$ + tu + B1, Idate, Hdate)
    If xdate > DateSerial(2008, 1, 1) And xdate > Idate Then
        FS.COPYFILE A1 + B1, F1$ + tu + B1
    End If
Next
Next

End Sub


Function MakeDir(Path$)
    
    text1 = Path$
    On Error Resume Next
    R = 4
    Do
                
        T = InStr(R + 1, text1, "\")
        dd = Mid$(text1, 1, T - 1)
        R = T
        Err.Clear
        If T > 0 Then MkDir dd
        yy = Err.Description
        MakeDir = Err.Number
        'Drive Not Exist
        If MakeDir = 76 Then Exit Function
    Loop Until T = 0

End Function

Sub Test2()
On Error GoTo Ende

Dim FS, Idate As Date, Hdate As String, File1 As String, File2 As String
Set FS = CreateObject("Scripting.FileSystemObject")

Call GetFileDates

File1 = Mid$(List1.List(List1.ListCount - 1), 21)
For R = List1.ListCount - 2 To 0 Step -1
File2 = Mid$(List1.List(R), 21)
If Mid$(List1.List(List1.ListCount - 1), 1, 19) <> Mid$(List1.List(R), 1, 19) Then
FS.COPYFILE File1, File2
tt = LastModifiedToCurrent(File2)
tg = 1
End If
Next

Call GetFileDates

If tg = 1 Then MsgBox "Done Some Sync Files at " + Now
'If tg = 0 Then MsgBox "Done Some Sync Files at " + Str(Now)

Exit Sub
Ende:
End

End Sub

Public Sub GetFileDates()
On Error GoTo Ende
Dim FS, Idate As Date, Hdate As String, File1 As String, File2 As String
Set FS = CreateObject("Scripting.FileSystemObject")
List1.Clear
File1 = "M:\VB6\VB-NT\00_Best_VB_01\Banging_Tunes\BangList Total Logg.html"
Call Date_File(File1, Idate, Hdate)
'dd = Hdate + " " + file1
List1.AddItem Hdate + " " + File1

For R = 3 To 25
On Error Resume Next
Err.Clear
Set g = FS.GETDRIVE(FS.GetDriveName(FS.GetAbsolutePathName(Chr$(R + 64) + ":")))
On Error GoTo 0
If InStr(rr, g.VolumeName) = 0 And Err.Number = 0 Then
rr = rr + g.VolumeName + "-"
If Mid(g.VolumeName, Len(g.VolumeName) - 1, 2) = "GB" Then

File2 = Chr$(R + 64) + ":\Banging\BangList Total Logg.html"
Call Date_File(File2, Idate, Hdate)
List1.AddItem Hdate + " " + File2

'fs.copyFile file1, file2


End If
'Nuke4 = Nuke4 + g.totalsize / 1024 ^ 3
'Nuke5 = Nuke5 + g.freespace / 1024 ^ 3
End If
Next

Exit Sub
Ende:
End

End Sub

Public Sub Date_File(Filespec2, Idate As Date, Hdate As String)
'Call Date_File(filespec2$, Idate)

Dim FS, F
Set FS = CreateObject("Scripting.FileSystemObject")

pdate = 0
FileSpec = Filespec2
Idate = DateSerial(1980, 1, 1)
If FS.FileExists(FileSpec) Then
Set F = FS.GETFILE(FileSpec)
Idate = F.datelastmodified
End If
Hdate = Format$(Idate, "YYYY-MM-DD HH:MM:SS")


End Sub


Public Function FindFileSize(Filename As String) As Long
        On Error Resume Next
        Dim filedata As WIN32_FIND_DATA

        filedata = Findfile(Filename)
        
        If filedata.nFileSizeHigh = 0 Then
            FindFileSize = Format$(filedata.nFileSizeLow, "###,###,###")
        Else
            FindFileSize = Format$(filedata.nFileSizeHigh, "###,###,###")
        End If
End Function

Private Function Findfile(xstrfilename) As WIN32_FIND_DATA
    Dim Win32Data As WIN32_FIND_DATA
    Dim plngFirstFileHwnd As Long
    Dim plngRtn As Long
    plngFirstFileHwnd = FindFirstFile(xstrfilename, Win32Data)
    ' Get information of file using API call
    If plngFirstFileHwnd = 0 Then
        Findfile.cFileName = "Error"   ' If file was not found, return error as name
    Else
        Findfile = Win32Data    ' Else return results
    End If
    plngRtn = FindClose(plngFirstFileHwnd) ' It is important that you close the handle
                'for FindFirstFile
End Function



Function LastModifiedToCurrent(duFile As String)
    Dim lngHandle As Long
    lngHandle = CreateFile(duFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If TouchFileTimes(lngHandle, ByVal 0&) = 0 Then
        'MsgBox "Error while changing file dates. Access Denied." + vbCrLf + duFile, vbCritical, "Modify Error"
    LastModifiedToCurrent = False
    Else
        'MsgBox "File date changed successfully.", vbInformation, "Modify Successful"
    LastModifiedToCurrent = True
    End If
    CloseHandle lngHandle
End Function

Sub Test1()

'restore header on wave fle
Open "D:\DWR 2008-11-27 00-36-03 -- 2008-11-28 13-02-14 -#- Janice Paul Short.wav" For Binary As #1
ll$ = "TTTT"
Get #1, 5, ll$
QQ1 = LOF(1)
Close #1
A1 = Asc(Mid$(ll, 1, 1))
a2 = Asc(Mid$(ll, 2, 1))
a3 = Asc(Mid$(ll, 3, 1))
a4 = Asc(Mid$(ll, 4, 1))
Mid(ll, 1, 1) = Chr(&H78)
Mid(ll, 2, 1) = Chr(&H10)
Mid(ll, 3, 1) = Chr(&H83)
Mid(ll, 4, 1) = Chr(&H23)

Mid(ll, 1, 1) = Chr(&H78)
Mid(ll, 2, 1) = Chr(&H94)
Mid(ll, 3, 1) = Chr(&H45)
Mid(ll, 4, 1) = Chr(&H0)

Open "D:\05 Media\Digital Wave Player M Drive\DR2008\DWR 2008-11-26 07-34-18 -- 2008-11-26 07-41-10 -#- New Anita Dingo Missed Steve On Linda On.wav" For Binary As #1
qq2 = LOF(1)
Close #1
Open "D:\05 Media\Digital Wave Player M Drive\DR2008\DWR 2008-11-26 09-39-09 -- 2008-11-27 00-35-50 -#- Steve Ibrag Linda Itching For Attack But Lets Clearly See Who Is Attacking First.wav" For Binary As #1
'qq2 = LOF(1)
jj$ = Space$(&H86)
Get #1, 1, jj$
Close #1

hh = &H78108323
hh = &H457894
gg1 = Hex$(QQ1 - 8)
gg2 = Hex$(qq2 - 8)
'4,560,000

Mid(ll, 1, 1) = Chr(&H78)
Mid(ll, 2, 1) = Chr(&H94)
Mid(ll, 3, 1) = Chr(&H45)
Mid(ll, 4, 1) = Chr(&H0)

Mid(ll, 1, 1) = Chr(&H78)
Mid(ll, 2, 1) = Chr(&H60)
Mid(ll, 3, 1) = Chr(&H8C)
Mid(ll, 4, 1) = Chr(&H56)
QQ1 = QQ1 * 1.03025
gg3 = Hex$(Int(QQ1))
ll = ""
lDataSize = Len(gg3)
Dim i As Integer
i2 = 1
For i = lDataSize - 1 To 1 Step -2
'For i = 1 To lDataSize Step 2
    'bytearray(i2) = Val("&h" + Mid$(gg3, i, 2))
    ll = ll + Chr(Val("&h" + Mid$(gg3, i, 2)))
    i2 = i2 + 1
Next



'T1 = DateSerial(bytearray(1), bytearray(2), bytearray(3),bytearray(4))

t1 = DateDiff("h", DateSerial(2008, 11, 27) + TimeSerial(0, 36, 3), DateSerial(2008, 11, 28) + TimeSerial(13, 2, 14))
t2 = DateDiff("n", DateSerial(2008, 11, 27) + TimeSerial(0, 36, 3), DateSerial(2008, 11, 28) + TimeSerial(13, 2, 14)) Mod 60
t3 = DateDiff("s", DateSerial(2008, 11, 27) + TimeSerial(0, 36, 3), DateSerial(2008, 11, 28) + TimeSerial(13, 2, 14)) Mod 60

A1 = Asc(Mid$(ll, 1, 1))
'a2 = Asc(Mid$(ll, 2, 1))
'a3 = Asc(Mid$(ll, 3, 1))
'a4 = Asc(Mid$(ll, 4, 1))
Open "D:\DWR 2008-11-27 00-36-03 -- 2008-11-28 13-02-14 -#- Janice Paul Short.wav" For Binary As #1
'll$ = "TTTT"
Put #1, 1, jj$
Put #1, 5, ll
Put #1, &H7D, ll
Close #1

a = Len(ll)

Shell "C:\Program Files\Windows Media Player\mplayer2.exe ""D:\DWR 2008-11-27 00-36-03 -- 2008-11-28 13-02-14 -#- Janice Paul Short.wav""", vbNormalFocus

End
End Sub

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************






Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function

Sub FileInUseDelay(Tx$)
        
    qq = Now + TimeSerial(0, 1, 30)
    Do
        DoEvents
        ii = FileInUse(Tx$)
        If ii = True Then Sleep (200)
    Loop Until ii = False Or qq < Now
    
    If ii = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Tx$ + vbCrLf + "Longer than 1 min 30 secs"
        Stop
    End If
End Sub

