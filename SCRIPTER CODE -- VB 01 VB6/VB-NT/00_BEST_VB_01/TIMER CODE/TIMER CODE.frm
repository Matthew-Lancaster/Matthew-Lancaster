VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   Caption         =   "TIMER CODE"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   6810
   Icon            =   "TIMER CODE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer9 
      Interval        =   1000
      Left            =   3960
      Top             =   2760
   End
   Begin VB.Timer Timer8 
      Interval        =   10000
      Left            =   3405
      Top             =   1890
   End
   Begin VB.Timer Timer7 
      Interval        =   2000
      Left            =   3828
      Top             =   972
   End
   Begin VB.Timer Timer_Mouse 
      Interval        =   100
      Left            =   2712
      Top             =   1548
   End
   Begin VB.Timer Timer6 
      Interval        =   1000
      Left            =   1848
      Top             =   1500
   End
   Begin VB.Timer Timer_ART_TRIP_WIRE 
      Interval        =   100
      Left            =   2400
      Top             =   975
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   1800
      Top             =   960
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   1200
      Top             =   960
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   2400
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1800
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1200
      Top             =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "TIMER CODE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   2970
   End
   Begin VB.Menu MNU_VBRUN 
      Caption         =   "VB RUN"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XD, OXD, OHOUR
Dim OMinModAgwa, OOLSN, TIMEDROP As Date, OMinute
Dim XTOTAG As Long, XTOTAGHWND, XD2
Dim OLSN(), NEROID, TTX As String, DIR1 As String
Dim VARCENTER, VDD
Dim OLDi, iTRIP, XHH
Dim TTDAYNOW, TTHOURNOW, TTDAYHALFNOW, TTDAYMODNOW
'PUT TIME SYNC IN FROM SHEDULAR
Dim OX1, OY1, MOUSE, XPAUSE
Dim MOUSE2, XPAUSE2
Dim PROGSTARTFLAG
'DIM AGEI

Const E = 2.7182818284
Const pi = 3.141592648
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const shrdNoMRUString = &H2
Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400
Const LF_FACESIZE = 32
Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4
Const GRADIENT_FILL_RECT_H As Long = &H0
Const GRADIENT_FILL_RECT_V  As Long = &H1


Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public cProcesses As New clsCnProc




' return the Enabled state of the screen saver

Function GetScreenSaverState() As Boolean
    Dim Result As Long
    SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0, Result, 0
    GetScreenSaverState = (Result <> 0)
End Function

' enable or disable the screen saver
'
' if second argument is true, it writes changes in user's profile
' returns True if the operation was successful, False otherwise

Function SetScreenSaverState(ByVal enabled As Boolean, _
    Optional ByVal PermanentChange As Boolean) As Boolean
    Dim fuWinIni As Long
    If PermanentChange Then
        fuWinIni = SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
    End If
    SetScreenSaverState = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, enabled, _
        ByVal 0&, fuWinIni) <> 0
End Function




Private Sub Form_Activate()
Form_Resize
End Sub

Private Sub Form_Load()

Dim OLSN(2)
VDD = -1
If App.PrevInstance = True Then End
OMinute = -2

OHOUR = -10

PROGSTARTFLAG = Now + TimeSerial(0, 4, 0)
TTHOURNOW = -5
'TIME1 = TimeSerial(4, 11, 0)
'TIME2 = TimeSerial(0, 388, 0) - TIME1
'Debug.Print Format(TIME1 - TIME2, "hh:mm:ss")
'TIME2 = TimeSerial(0, 388, 0) - TIME1
'' 04:11 - 01:54 - 02:17
'' VISIT OF 05-09-2011 4:11PM
'' 3 PEOPLE = 01:54P MY INDOOR TIME
''          = 02hh 17mm OF SLEEP AS IN TIME
''Stop
''End



TTDAYNOW = Day(Now)
TTDAYHALFNOW = Day(Now * 2)
TTHOURNOW = Hour(Now)
TTDAYMODNOW = Day(Now) Mod 3
Set fs = CreateObject("Scripting.FileSystemObject")

On Error Resume Next

If IsIDE = True Then Me.WindowState = 1


LB = "TIMER 1 - RS232 WELCOME" + vbCrLf
LB = LB + "TIMER 2 - WinAmp MP3%-VideoBar" + vbCrLf
LB = LB + "TIMER 3 - MAster BATch VB6 Compiler = EXIT SUB" + vbCrLf
LB = LB + "                - WINAMP  = EXIT SUB" + vbCrLf
LB = LB + "                - WinAmp MP3% MINI #" + vbCrLf
LB = LB + "                - VolumeBar BIG" + vbCrLf
LB = LB + "                - VolumeBar MINI" + vbCrLf
LB = LB + "TIMER 4 - CHK_CDROM_DRIVES_AND_COPY_FILES" + vbCrLf
LB = LB + "                - CHK_KAT_RECORD" + vbCrLf
LB = LB + "TIMER 5 - Webcam_Motion" + vbCrLf
LB = LB + "TIMER       ART_TRIP_WIRE_Timer" + vbCrLf
LB = LB + "-------------- KAT RECORD SMALL FILES" + vbCrLf
LB = LB + "-------------- DAY -- WINRAR ARC VB" + vbCrLf
LB = LB + "-------------- NORTON OF QUITE MODE" + vbCrLf
LB = LB + "-------------- " + vbCrLf
LB = LB + "-------------- " + vbCrLf




Label1 = LB

Me.WindowState = 1
'Me.Show
DoEvents

'Call Timer7_Timer


End Sub

Private Sub Form_Resize()
'If Me.WindowState <> 0 Then Exit Sub
'If VARCENTER = True Then Exit Sub
On Error Resume Next
Label1.Top = 0
Label1.Left = 0
Label1.AutoSize = False
Label1.AutoSize = True
DoEvents

VARCENTER = True
Me.Width = Label1.Width
Me.Height = Label1.Height + 800

Me.Top = (Screen.Height / 2 - Me.Height / 2) - 400
Me.Left = Screen.Width / 2 - Me.Width / 2

End Sub

Private Sub Label1_Click()
For R = 1 To UBound(OLSN)
    OLSN(R) = 0
Next
End Sub


Private Sub MNU_VBRUN_Click()

Dim FileSpec, TT As Long
FileSpec = App.Path + "\" + App.EXEName + ".vbp"

If IsIDE = False And Dir$(FileSpec) <> "" Then
    'TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit """ + FileSpec + """", vbMinimizedNoFocus)
    TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE  """ + FileSpec + """", vbMinimizedNoFocus)
    
    End
End If


End Sub
Function ShowTaskBar()
    TaskBarhWnd = FindWindow("Shell_traywnd", "")
    If TaskBarhWnd <> 0 Then
       Call SetWindowPos(TaskBarhWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    End If
End Function
Private Sub Timer1_Timer()
'TIMER 1 RS232 WELCOME
If OHOUR <> Hour(Now * 2) Then
        OHOUR = Hour(Now * 2)
        ShowTaskBar
    End If
Exit Sub
    
On Error Resume Next
Set F = fs.getfile("D:\VB6\VB-NT\00_Best_VB_04\RS232 LOGGER\RS232 DOOR REED DATA\LAST DOOR CLOSE.txt")
    
If XONDATE > 0 And XONDATE <> F.datelastmodified Then
    'Shell "D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS\WelCome\WelCome.exe"
End If

XONDATE = F.datelastmodified





End Sub

Private Sub Timer2_Timer()
'TIMER 1 RS232 WELCOME
'TIMER 2 WinAmp MP3%-VideoBar

'If GetComputerName = "55-88-HAPPY" Then Exit Sub
    
If FindWindow(vbNullString, "WinAmp MP3%-VideoBar.exe") > 0 Then Exit Sub

Win_Main = FindWindow("Winamp v1.x", vbNullString)
If Win_Main = 0 Then Exit Sub

Win_Video = FindWindow("Winamp Video", "Winamp Video")
If Win_Video = 0 Then Exit Sub

xt = IsWindowVisible(Win_Video)
'XT = 1 = VISIBLE = TRUE
If xt = 0 Then Exit Sub


Shell "D:\VB6\VB-NT\00_Best_VB_01\WinAmp MP3%\WinAmp MP3%-VideoBar.exe", vbNormalFocus


End Sub

Private Sub Timer3_Timer()

'TIMER 1 RS232 WELCOME
'TIMER 2 WinAmp MP3%-VideoBar
'TIMER 3 MAster BATch VB6 Compiler = EXIT SUB
'                WINAMP  = EXIT SUB
'                WinAmp MP3% MINI #
'                VolumeBar BIG
'                VolumeBar MINI
'DELAY TO LOAD

If XTOTAG > 0 Then Call cmdSetTitle_Click



'If Second(Now) < 30 Then Exit Sub

XD = Second(Now)
If XD <> OXD Then
    If XD2 > 30 Then
        XD2 = 0
        GIVE = True
    End If
    
    OXD = XD
    XD2 = XD2 + 1
End If

If GIVE = False Then Exit Sub
    
    If FindWindow(vbNullString, "CID Run Me.") = 0 Then
        Shell "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\Cid-RunMe.exe", vbNormalFocus
    End If

        
    Call FIND_CAPTURE_INFAR
    
    If FindWindow(vbNullString, "MAster BATch VB6 Compiler") > 0 Then Exit Sub
    
    Win_Main = FindWindow("Winamp v1.x", vbNullString)
    If Win_Main = 0 Then Exit Sub
    
    If FindWindow(vbNullString, "WinAmp MP3% MINI #") = 0 Then
        Shell "D:\VB6\VB-NT\00_Best_VB_01\WinAmp MP3% MINI #\WinAmp MP3% MINI.EXE", vbNormalFocus
    End If
    
    If FindWindow(vbNullString, "VolumeBar BIG") = 0 Then
        Shell "D:\VB6\VB-NT\00_Best_VB_01\Winamp - Volume BIG\VolumeBar BIG.exe", vbNormalFocus
    End If
    
    If FindWindow(vbNullString, "VolumeBar MINI") = 0 Then
        Shell "D:\VB6\VB-NT\00_Best_VB_01\Winamp - Volume MINI\VolumeBar MINI.exe", vbNormalFocus
    End If


End Sub


Sub KILL_CAPTURE()

    Dim sTitle As String
    
    TXTMHWND = FindWindow(vbNullString, "IrfanCAPTURE")
    If TXTMHWND > 0 Then
        OTSS = XTOTAG
        Var = cProcesses.Convert(TXTMHWND, OTSS, cnFromhWnd Or cnToProcessID)
        cProcesses.Process_Kill (OTSS)
    End If
    

    
'    i = TXTMHWND
'    i = GetParent(i)
'    Do While i <> 0
'        i = GetParent(i)
'        If i = 0 Then i = II: Exit Do
'        II = i
'    Loop

End Sub

Sub FIND_CAPTURE_INFAR()
    
    
    Exit Sub
'    If MOUSE2 < Now Then
'        KILL_CAPTURE
'        Exit Sub
'    End If
    
    If GetWindowTitle(XTOTAGHWND) <> "IrfanCAPTURE" Then
         XTOTAGHWND = 0
'        XTOTAG = 0
    End If
    
    If FindWindow(vbNullString, "IrfanView") > 0 And _
        FindWindow(vbNullString, "IrfanCAPTURE") > 0 Then Exit Sub
                           
            KILL_CAPTURE
                    
    
    'End If
   
    
    
    
    'If FindWinPart("IrfanCAPTURE", TF) > 0 Then Exit Sub
    If FindWindow(vbNullString, "IrfanCAPTURE") > 0 Then Exit Sub



    If XTOTAG = 0 Then
       If FindWindow(vbNullString, "IrfanView") > 0 Then
            Exit Sub
       End If
    End If

    If FindWindow(vbNullString, "IrfanCAPTURE") = 0 Then
        If XTOTAG = 0 Then
'            XTOTAG = Shell("C:\Program Files\IrfanView\i_view32.exe /title=""IrfanView_CAPTURE"" /capture=6", vbNormalFocus)
            XTOTAG = ExecCmd("C:\Program Files\IrfanView\i_view32.exe  /capture=6")
            
            'Parent
            'Set new window title
            sTitle = ("IrfanCAPTURE")
            SetWindowText XTOTAG, sTitle
        End If
    End If
    
End Sub

Private Sub cmdSetTitle_Click()
    
    Dim sTitle As String
    ' Ask user for new window title
    'sTitle = InputBox("Enter new window title:", "EliteSpy +")
    'sTitle = InputBox("Enter new window title:", "EliteSpy +")
    

    
    
    TXTMHWND = FindWindow(vbNullString, "IrfanView")
    If TXTMHWND > 0 Then
        OTSS = XTOTAG
        Var = cProcesses.Convert(OTSS, TXTMHWND, cnFromProcessID Or cnTohWnd)
    End If
    
    i = TXTMHWND
    i = GetParent(i)
    Do While i <> 0
        i = GetParent(i)
        If i = 0 Then i = ii: Exit Do
        ii = i
    Loop

    sTitle = ("IrfanView CAPTURE")
    
    'Parent
    'Set new window title
    SetWindowText i, sTitle
    ' Disable window
    EnableWindow i, 0
    XTOTAG = 0
    XTOTAGHWND = ii
End Sub

Private Sub Timer4_Timer()

'TIMER 1 - RS232 WELCOME
'TIMER 2 - WinAmp MP3%-VideoBar
'TIMER 3 - MAster BATch VB6 Compiler = EXIT SUB
'                - WINAMP  = EXIT SUB
'                - WinAmp MP3% MINI #
'                - VolumeBar BIG
'                - VolumeBar MINI
'TIMER 4 - CHK_CDROM_DRIVES_AND_COPY_FILES


Call CHK_KAT_RECORD

'If Minute(Now) Mod 5 = 0 Then
'    Call CHK_CDROM_DRIVES_AND_COPY_FILES
'
'End If

Exit Sub

If Second(Now) Mod 5 = 0 And 1 = 2 Then
    Debug.Print Str(Timer) + "Start -- CHK_M2_DRIVES_AND_MOVE_FILES"
    Call CHK_M2_DRIVES_AND_MOVE_FILES
    Debug.Print Str(Timer) + "Start -- CHK_W311_DICTOR_AND_MOVE_FILES_SHARE_SEARCH"
    Call CHK_W311_DICTOR_AND_MOVE_FILES_SHARE_SEARCH
    Debug.Print Str(Timer) + "Start -- CHK_NIKON_AND_MOVE_FILES_SHARE_SEARCH"
    Call CHK_NIKON_AND_MOVE_FILES_SHARE_SEARCH
    Debug.Print Str(Timer) + "Start -- CHK_KAT_RECORD"
    Call CHK_KAT_RECORD
    Debug.Print Str(Timer) + "Done"
End If
    
End Sub

Sub CHK_M2_DRIVES_AND_MOVE_FILES()

Exit Sub
        
        On Local Error Resume Next
        
        For RY = 1 To 2
BACKTOFORM2:
        If RY = 1 Then DIR1 = "H:\DCIM\100MSDCF"
        If RY = 2 Then DIR1 = "L:\DCIM\100MSDCF"
        If RY > 2 Then Exit Sub
        
        Err.Clear
        Set F = fs.GETDRIVE(Mid(DIR1, 1, 1) + ":")
        
        If Err.Number > 0 Then RY = RY + 1: GoTo BACKTOFORM2
        
        
        If F.ISREADY = False Then RY = RY + 1: GoTo BACKTOFORM2
        If fs.FolderExists(DIR1) = False Then RY = RY + 1: GoTo BACKTOFORM2
            
        If fs.FolderExists("M:\# W880i\VIDEO\") = True Then DRI = "M" Else DRI = "D"
        
        If Dir(DIR1 + "\*.3gp") <> "" Then
            TTX = DRI + ":\# W880i\VIDEO\" + Format(Now, "YYYY-MM-DD") + "\"
            MkDir TTX
            Call ShellFileMove(DIR1 + "\*.3gp", TTX)
        End If
                    
        If Dir(DIR1 + "\*.jpg") <> "" Then
            TTX = DRI + ":\# W880i\VIDEO\JPG\" + Format(Now, "YYYY-MM-DD") + "\"
            MkDir DRI + ":\# W880i\VIDEO\JPG\"
            MkDir TTX
            Call ShellFileMove(DIR1 + "\*.jpg", TTX)
        End If

    Next

End Sub


Sub CHK_CDROM_DRIVES_AND_COPY_FILES()



Exit Sub

'Public Const DRIVE_UNKNOWN As Long = 0
'Public Const DRIVE_NO_ROOT_DIR As Long = 1
'Public Const DRIVE_REMOVABLE As Long = 2
'Public Const DRIVE_FIXED As Long = 3
'Public Const DRIVE_REMOTE As Long = 4
'Public Const DRIVE_CDROM As Long = 5
'Public Const DRIVE_RAMDISK As Long = 6

    ReDim Preserve OLSN(2)

    On Local Error Resume Next
        
    Set dc = fs.Drives
        
    For Each D In dc
            
        'Case 0: t = "Unknown"
        'Case 1: t = "Removable"
        'Case 2: t = "Fixed"
        'Case 3: t = "Network"
        'Case 4: t = "CD-ROM"
        'Case 5: t = "RAM Disk"
        
        If D.DRIVETYPE = 4 Then
            HDD = HDD + D.DriveLetter
        End If
    Next
        
        'Exit Sub
        'If RY = 1 Then DIR1 = "F:"
        'If RY = 2 Then DIR1 = "G:"
'        DIR1 = Chr(RY + 65) + ":"
'        If RY > i Then
'            Exit Sub
'        End If
        
    For R = 1 To Len(HDD)
        
        DIR1 = Mid(HDD, R, 1) + ":"
        RY = R
        
        Set F = Nothing
        Err.Clear
        Set F = fs.GETDRIVE(Mid(DIR1, 1, 1) + ":")
        
        X = 0
        If Err.Number > 0 Then OLSN(RY) = "": X = 2
        If F.ISREADY = False Then OLSN(RY) = "": X = 2
        If fs.FolderExists(DIR1) = False Then X = 2
        If F.SerialNumber = OLSN(RY) Then X = 2
        If F.ISREADY = False Then OLSN(RY) = "": X = 2
        
        If X = 0 Then
            'NOT WTH DVD
            OLSN(RY) = F.SerialNumber
            Debug.Print Hex$(F.SerialNumber) + " -- " + F.VolumeName
                
            'TTX = "M:\Videos\00 Vid Films\VOBS\# The Mother\" + Hex$(F.SerialNumber) + "--" + F.VolumeName '+ "\"
            TTX = "D:\#0 CD DVD'S\" + Hex$(F.SerialNumber) + " -X- " + F.VolumeName
            
            'If fs.FolderExists(DIR1 + "\VIDEO_TS") = True Then
            '    MkDir TTX
            '    Clipboard.Clear
            '    Clipboard.SetText TTX
            'End If
            
            'TTX = "M:\0 00 Art\# PHOTO #\" + Hex$(F.SerialNumber) + "--" + F.VolumeName '+ "\"
            
            'If fs.FolderExists(DIR1 + "\VIDEO_TS") = False Then
            '    MkDir TTX
            '    Clipboard.Clear
            '    Clipboard.SetText TTX
            'End If
            
            'MsgBox Hex$(F.SerialNumber) + " -- " + F.VolumeName, vbMsgBoxSetForeground + vbCritical
 
            'Me.WindowState = 1
       
            'Exit Sub
            
            'DRI = "M"
            'If Dir(DIR1 + "\*.*") <> "" Then
                
            
            'TTX = DRI + ":\Videos\00 Vid Films\VOBS\The Mother\" + Hex$(F.SerialNumber) + "--" + F.VolumeName + "\"
            'M:\0 00 Art\# PHOTO #
            'TTX = DRI + ":\00 Art\# PHOTO #\" + Hex$(F.SerialNumber) + "--" + F.VolumeName + "\"
            'TTX = "M:\#0 CD DVD'S"
            
'            MkDir TTX
            
            
            'Call Form1MIX.Form_Load
            ScanPath.ListView1.ListItems.Clear
            ScanPath.cboMask.Text = "*.*"
            ScanPath.chkSubFolders = vbChecked
            ScanPath.TxtPath.Text = DIR1
            Call ScanPath.cmdScan_Click
            'MsgBox "BEGIN DISK"
            
            For WE5 = ScanPath.ListView1.ListItems.Count To 1 Step -1
                A1 = ScanPath.ListView1.ListItems.Item(WE5).SubItems(1)
                B1 = ScanPath.ListView1.ListItems.Item(WE5)
                If A1 = "" Or B1 = "" Then
                    ScanPath.ListView1.ListItems.Remove (WE5)
                End If
            Next
            FLAGNOT_DONE = False
            
            For WE5 = 1 To ScanPath.ListView1.ListItems.Count
                WE7 = ScanPath.ListView1.ListItems.Count
                
                A1 = ScanPath.ListView1.ListItems.Item(WE5).SubItems(1)
                B1 = ScanPath.ListView1.ListItems.Item(WE5)
                'If Mid(A1, 4) <> "" Then Stop
                TTXZ = TTX + Mid(A1, 3)
                
                Call DIR_ON_ANOTHER_DRIVE(TTXZ)
                
                
                Set F1 = fs.getfile(TTXZ + B1)
                FH1 = F1.Size
                Set F1 = fs.getfile(A1 + B1)
                Dim FH
                FH = F1.Size
                
                TXT_ST = Str(WE5) + " Of " + Str(WE7) + "---" + Str(FH) + "--- " + A1 + B1 + "  -- " + TTXZ
                Debug.Print TXT_ST
                'FH = FH <> fH1
'                If ScanPath.ListView1.ListItems.Count = WE5 Then
'                    Th = 1
'                End If
'
'                If Th = 1 Then Call Form1MIX.cmdEject_Click
                
                'If fs.FileExists(TTX + H1 + B1) = False Then
                '    FH = True
                'End If
                
                If FH <> FH1 Then
                    'If Th = 1 Then Call Form1MIX.cmdEject_Click
                    Call ShellFileCopy(A1 + B1, TTXZ + B1, False)
                    If fs.FileExists(TTXZ + B1) = False Then FLAGNOT_DONE = True: Debug.Print "YES"
                            
                        'NOT WITH DVD
                        'Th = 1
                    '    MsgBox "DVD COPY NOT DONE"
                    '    Exit For
                    'End If
                End If

                'If NEROID < Now And NEROID > 0 Then
                '    NEROID = 0
                '    pid = -1
                '    PidBool = cProcesses.GetEXEID(pid, "C:\Program Files\Ahead\Nero ShowTime\ShowTime.exe")
                '    Process_Kill (pid)
                'End If
                
                
                'Th = 1
            Next 'COPY
        End If 'JOB WORK DO
    Next
    If OLSN(RY) <> OOLSN Then
    OOLSN = OLSN(RY)
    If FLAGNOT_DONE = True Then
'        MsgBox "DVD COPY NOT DONE" + vbCrLf + TXT_ST
    Else
'        MsgBox "DVD WORK DONE" + vbCrLf + TXT_ST
    End If
    End If
    
    'If Th = 1 Then Call Form1MIX.cmdEject_Click
        
End Sub



Sub CHK_W311_DICTOR_AND_MOVE_FILES_SHARE_SEARCH()

'MAKE SURE WHEN SHARE ALL DRIVES SET TO NAME
'LIKE WS_550M
'AND NOT WITH DRIVE LETTER AFTER LIKE WS_550M (H)
'DR1 = "\\55-88-happy\WS_311M"
'DR2 = "\\55-88-happy\WS_550M"
        
       Exit Sub
            
        DIR1 = ""
        Set dc = fs.Drives
        For Each D In dc
            '1=REMOVABLE
            If D.DRIVETYPE = 1 Then
                DIR1 = D.DriveLetter + ":\DSS_FLDA"
                DIR2 = D.DriveLetter + ":\DSS_FLD"
                If fs.FolderExists(DIR1) = True Then Exit For
                            
            End If
        Next

        If DIR1 = "" Then Exit Sub
        
        
        For R = 1 To 5
            DIR1 = DIR2 + Chr(R + 64)
            On Error Resume Next
            FILE1 = Dir(DIR1 + "\*.WMA")
            If Err.Number > 0 Then Exit Sub
            If FILE1 <> "" Then
                If InStr(FILE1, "WS550") Then
                    Set F = fs.GETDRIVE("M:")
                    If F.ISREADY = True Then
                        Call ShellFileMove(DIR1 + "\*.WMA", "M:\05 Media\04 Olympus WS-550M\")
                    Else
                        Call ShellFileMove(DIR1 + "\*.WMA", "D:\05 Media\04 Olympus WS-550M\")
                    End If
                End If
                                
                If InStr(FILE1, "WS310") Then
                    Set F = fs.GETDRIVE("M:")
                    If F.ISREADY = True Then
                        Call ShellFileMove(DIR1 + "\*.WMA", "M:\05 Media\03 Olympus WS-311M\")
                    Else
                        Call ShellFileMove(DIR1 + "\*.WMA", "D:\05 Media\03 Olympus WS-311M\")
                    End If
                End If
            
                If InStr(FILE1, "VN6") Then
                    Set F = fs.GETDRIVE("M:")
                    If F.ISREADY = True Then
                        MkDir "M:\05 Media\05 Olympus VN\"
                        Call ShellFileMove(DIR1 + "\*.WMA", "M:\05 Media\05 Olympus VN\")
                    Else
                        MkDir "D:\05 Media\05 Olympus VN\"
                        Call ShellFileMove(DIR1 + "\*.WMA", "D:\05 Media\05 Olympus VN\")
                    End If
                End If
            
            
            End If
                            
        Next


End Sub

Sub CHK_NIKON_AND_MOVE_FILES_SHARE_SEARCH()
        
        Exit Sub
        
        
        For RY = 1 To 2
BACKTOFOR:
        On Local Error Resume Next
        If RY = 1 Then DIR1 = "H:\DCIM"
        If RY = 2 Then DIR1 = "P:\DCIM"
        If RY > 2 Then Exit Sub
        
        Set F = fs.GETDRIVE(Mid(DIR1, 1, 1) + ":")
        XXV = F.VolumeName
        'Set d = FS.GetDrive(FS.GetDriveName(drvPath))

        If Err.Number > 0 Then RY = RY + 1: GoTo BACKTOFOR
        
        If F.ISREADY = False Then RY = RY + 1: GoTo BACKTOFOR
        If fs.FolderExists(DIR1) = False Then RY = RY + 1: GoTo BACKTOFOR
        Set F = fs.GetFolder(DIR1)
        If F.Size > 0 Then
            
        Set F = fs.GETDRIVE("M:")
        If F.ISREADY = True Then DRI = "M" Else DRI = "D"
        If XXV = "DV CAM 4GB" Then
            TTX = DRI + ":\# DV CAM 4GB\DV CAM 4GB - " + Format(Now, "YYYY-MM-DD") + "\"
            MkDir DRI + ":\# DV CAM 4GB"
            MkDir TTX
            Call ShellFileMove(DIR1, TTX)
        End If
        
        If XXV = "NIKON" Then
            TTX = DRI + ":\# NIKON\" + Format(Now, "YYYY-MM-DD") + "\"
            MkDir DRI + ":\# NIKON\"
            MkDir TTX
            Call ShellFileMove(DIR1, TTX)
        End If
    End If
        
        
        ':\# NIKON\
        
        'If FS.FolderExists("M:\# NIKON\") = True Then DRI = "M" Else DRI = "D"
        
        'If Dir(DIR1 + "\*.") <> "" Then
        '    TTX = DRI + ":\# W880i\VIDEO\" + Format(Now, "YYYY-MM-DD") + "\"
        '    MkDir TTX
        '    Call ShellFileMove(DIR1 + "\*.3gp", TTX)
        'End If
            
            
        'O:\DCIM\100DSCIM
            
        'FILE1 = Dir(DIR1 + "\100DSCIM\*.JPG")
        'If Err.Number > 0 Then Exit Sub
        'If FILE1 = "" Then
                    
        'If InStr(FILE1, "WS310") Then
                    
                    
                    
        
        
        Next 'RY
        
        Exit Sub
        
        For R2 = 1 To 2
        
        If R2 = 1 Then DIR1 = "\\55-88-happy\S\DCIM"
        If R2 = 2 Then DIR1 = "\\55-88-happy\G\DCIM"
        
        If DirExists(DIR1) = True Then
            SERIALDR = NIKON_OSERIALDR
        
            For R = 1 To 10
            If Dir(Mid(DIR1, 1, 15) + "\MEMSTICK.IND", vbHidden) <> "" Then Exit Sub
                
                        On Local Error Resume Next
                        Err.Clear
                        
                        
                        O = Dir(DIR1 + "\*", vbDirectory)
                        If O = "" Then Exit For
                        
                        O2 = Dir(, vbDirectory)
                        O2 = Dir(, vbDirectory)
                        'O2 = Dir()
                
                            
                        If O2 <> "" Then
                            Set F = fs.GetFolder(DIR1 + "\" + O2)
                            XF = F.Size
                            If XF = 0 Then RmDir (DIR1 + "\" + O2): Exit For
                            If InStr(NIKON_OSERIALDR, Str(D.SerialNumber) + " " + DIR1 + O2) = 0 And XF > 0 Then
                                SERIALDR = SERIALDR + " " + Str(D.SerialNumber) + " " + DIR1 + O2
                                
                                If fs.FolderExists("M:\# NIKON\") = True Then
                                    Call ShellFileMove(DIR1 + "\" + O2 + "\*.*", "M:\# NIKON\")
                                Else
                                    Call ShellFileMove(DIR1 + "\" + O2 + "\*.*", "D:\# NIKON\")
                                End If
                            
                            End If
                        End If
            Next
        End If
    Next

NIKON_OSERIALDR = SERIALDR


End Sub

Sub CHK_KAT_RECORD()
    Exit Sub
'    If (Minute(Now) = OMinute Or Minute(Now) Mod 5 <> 0) And OMinute >= 0 Then Exit Sub
    XHA = OMinute
    
    OMinute = Minute(Now)

    'On Local Error Resume Next
    
    'If R2 = 1 Then DIR1 = "\\55-88-happy\C\05 Media\ASUS - KAT RECORD"
    'If R2 = 1 Then Call ShellFileMove2(DIR1 + "\" + "*.MP3", "M:\05 Media\KAT RECORD\KAT RECORD ASUS\", True)
       
    On Local Error Resume Next
    DIR1 = "D:\KAT MP3 RECORDER"
    XXD = Dir(DIR1 + "\*.MP3")
    
    If XXD = "" Then Exit Sub
    
    'DATE NOW TODAY IF ONE PASS DONE
    If XHA > 0 Then
        ScanPath.cboDate.ListIndex = 0
        ScanPath.DTPicker1(0) = Int(Now)
    End If
    
    ScanPath.ListView1.ListItems.Clear
    ScanPath.cboMask.Text = "*.MP3"
    ScanPath.chkSubFolders = vbChecked
    ScanPath.TxtPath.Text = DIR1
     

    ScanPath.cboSize.ListIndex = 1 'Less than
    'cboSize.ListIndex = 2 'Bigger Than
    ScanPath.cboSizeType(lCount).ListIndex = 0 'Byte
    'cboSizeType(lCount).ListIndex = 2 'Meg
'    ScanPath.cboSizeType(lCount).ListIndex = 1 'K
    ScanPath.txtSize(0) = 700
'    ScanPath.Show
'    Call ScanPath.cmdScan_Click
    


           ' If A1 + B1 <> "" And A1 + D1 <> "" Then
    For WE5 = ScanPath.ListView1.ListItems.Count - 1 To 1 Step -1
        A1 = ScanPath.ListView1.ListItems.Item(WE5).SubItems(1)
        B1 = ScanPath.ListView1.ListItems.Item(WE5)
'        Shell "Explorer.exe /select, " + A1 + B1, vbNormalFocus
        
        ScanPath.ListView1.ListItems(WE5).EnsureVisible
        ScanPath.Show
        
        'ZERO BYTE KILL
        Kill A1 + B1

        'NULL FIELD REMOVE
        If A1 + B1 = "" Then ScanPath.ListView1.ListItems.Remove (WE5)
    Next

    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32
    
    ScanPath.ListView1.ListItems.Clear
    ScanPath.cboMask.Text = "*.MP3"
    ScanPath.chkSubFolders = vbChecked
    ScanPath.TxtPath.Text = DIR1
     
     
    ScanPath.cboSize.ListIndex = 1 'Less than

    ScanPath.cboSizeType(lCount).ListIndex = 1 'K
    ScanPath.txtSize(0) = 800


'    ScanPath.Show
    Call ScanPath.cmdScan_Click
    
    ScanPath.ListView1.SortKey = 1
    ScanPath.ListView1.Sorted = False
    ScanPath.ListView1.SortKey = 2
    ScanPath.ListView1.Sorted = False
    ScanPath.ListView1.SortKey = 3
    ScanPath.ListView1.Sorted = False
    ScanPath.ListView1.SortKey = 4
    ScanPath.ListView1.Sorted = False
    ScanPath.ListView1.SortKey = 5
    ScanPath.ListView1.Sorted = False
    
    ScanPath.ListView1.SortOrder = lvwAscending
    ScanPath.ListView1.SortKey = 3
    ScanPath.ListView1.SortOrder = lvwAscending
    ScanPath.ListView1.SortKey = 3
    ScanPath.ListView1.Sorted = True
'    ScanPath.ListView1.Sorted = False
           ScanPath.Show
    ScanPath.ListView1.Refresh
'    Do
    DoEvents
'    Loop Until 1 = 2
'    Exit Sub
    
    

    
    
    For WE5 = 1 To ScanPath.ListView1.ListItems.Count - 1
        FileS1 = Val(ScanPath.ListView1.ListItems.Item(WE5).SubItems(3))
        FileS2 = Val(ScanPath.ListView1.ListItems.Item(WE5 + 1).SubItems(3))
        A1 = ScanPath.ListView1.ListItems.Item(WE5).SubItems(1)
        B1 = ScanPath.ListView1.ListItems.Item(WE5)
        D1 = ScanPath.ListView1.ListItems.Item(WE5 + 1)
        D2 = ScanPath.ListView1.ListItems.Item(WE5 + 1).SubItems(1)
        
           ' MsgBox "SAMPLE OF CRC DEL TO CHK IN THIS CODE" + vbCrLf + App.Path + "\" + App.EXEName
'        ScanPath.ListView1.ListItems(WE5).EnsureVisible
        If FileS1 > 0 And FileS2 > 0 Then
        On Error Resume Next
        Err.Clear
        If FileS1 = FileS2 Then
        If m_CRC.CalculateFile(A1 + B1) = m_CRC.CalculateFile(D2 + D1) Then
            'MsgBox "SAMPLE OF CRC DEL TO CHK IN THIS CODE" + vbCrLf + App.Path + "\" + App.EXEName
            'Stop

            If Err.Number = 0 Then
                Shell "Explorer.exe /select, " + D2, vbNormalFocus '+ D1, vbNormalFocus
          
                Kill D2 + D1
            End If
        End If
        On Error GoTo 0
        
        End If
        End If
    Next
    
    
    If FileInUse(DIR1 + "\" + XXD) = True Then Exit Sub
    If fs.DriveExists("M:") = False Then Exit Sub
    

    Set F = fs.GETDRIVE("M:")
    X = 0
    If Err.Number > 0 Then X = 2
    If F.ISREADY = False Then X = 2
    
    
    'NOT IF ELEMENTS
    If F.VolumeName <> "ELEMENTS" Then X = 2
    
    
    
    If F.ISREADY = False Then X = 2
    
'    If FS.FolderExists(DIR1) = False Then X = 2
'    If F.SerialNumber = OLSN(RY) Then X = 2

    If X = 0 Then
        Call ShellFileMove(DIR1 + "\" + XXD, "M:\05 Media\KAT RECORD\KAT RECORD ACER\", True)
    End If
    
    'KAT_OSERIALDR = SERIALDR


End Sub



Private Sub Timer5_Timer()


'TIMER 1 - RS232 WELCOME
'TIMER 2 - WinAmp MP3%-VideoBar
'TIMER 3 - MAster BATch VB6 Compiler = EXIT SUB
'                - WINAMP  = EXIT SUB
'                - WinAmp MP3% MINI #
'                - VolumeBar BIG
'                - VolumeBar MINI
'TIMER 4 - CHK_CDROM_DRIVES_AND_COPY_FILES
'                 - CHK_KAT_RECORD
'TIMER 5 - Webcam_Motion
'TIMER       ART_TRIP_WIRE_Timer



'1 MINUTE
'

Dim RTX As Double
'RT = TimeSerial(Hour(Now), Minute(Now), Second(Now))
RTX = Now - (Int(Now))
RTX = RTX * 2

MinModAgwa = Minute(RTX) 'Mod 2
'MinModAgwa2 = MinModAgwa
If OMinModAgwa <> MinModAgwa Then
'    If Now - Int(Now) < SunSet And Now - Int(Now) > SunRise Then
        OMinModAgwa = MinModAgwa
        
        'Call OutputData
        On Error Resume Next
        Di7 = Shell("D:\VB6\VB-NT\00_Best_VB_04\Webcam_Motion\WebCamMatts.exe Quite", vbHide)
        On Error GoTo 0
        Process_Low (Di7)
    
'    End If
End If




End Sub



Private Sub Timer_ART_TRIP_WIRE_Timer()
Call ONEADAY

'TIMER 1 - RS232 WELCOME
'TIMER 2 - WinAmp MP3%-VideoBar
'TIMER 3 - MAster BATch VB6 Compiler = EXIT SUB
'                - WINAMP  = EXIT SUB
'                - WinAmp MP3% MINI #
'                - VolumeBar BIG
'                - VolumeBar MINI
'TIMER 4 - CHK_CDROM_DRIVES_AND_COPY_FILES
'                 - CHK_KAT_RECORD
'TIMER 5 - Webcam_Motion
'TIMER       ART_TRIP_WIRE_Timer

'i = LastModifiedToCurrent("K:\TEMP\ART_PROG_TRIP_WIRE.txt")

Dim HH As Long, OTSS As Long

HH = FindWindow(vbNullString, "Jpeg Slides PJpeg ART.exe - Application Error")

If HH > 0 Then
    AppActivate "Jpeg Slides PJpeg ART.exe - Application Error", False
    SendKeys "{ENTER}", True
End If
HH = FindWindow(vbNullString, "Jpeg Encoder -- 2009 The One --: Jpeg Slides PJpeg ART.exe - Application Error")
If HH > 0 Then
    AppActivate "Jpeg Slides PJpeg ART.exe - Application Error", False
    SendKeys "{ENTER}", True
End If

HH = FindWindow(vbNullString, "JPEG Encoder -- Shockingly Good Photo ART -- [-:¦:-•:*''' The One '''*:•-:¦:-]")
If HH > 0 Then XHH = True
'If HH = 0 And XHH = True Then End
If HH = 0 And XHH = False Then Exit Sub


'If FS.FileExists("K:\TEMP\ART_PROG_TRIP_WIRE.txt") = False And XHH = True Then End
If fs.FileExists("K:\TEMP\ART_PROG_TRIP_WIRE.txt") = False Then Exit Sub

Set F = fs.getfile("K:\TEMP\ART_PROG_TRIP_WIRE.txt")
i = F.datelastmodified
If OLDi <> i Then iTRIP = Now + TimeSerial(0, 0, 25)
OLDi = i
'Debug.Print i, OLDi

If iTRIP > Now Then Exit Sub
XHH = False

'TT = cProcesses.Convert(HH, OTSS, cnFromhWnd Or cnToProcessID)
OTSS = -1
TT = cProcesses.GetEXEID(OTSS, "D:\VB6\VB-NT\00_Best_VB_04\JPEG-Encoder-ART\Jpeg Slides PJpeg ART.exe")

If TT = True Then cProcesses.Process_Kill (OTSS)

XX = 0
Do
    XX = XX + 1
    OTSS = -1
    TT = cProcesses.GetEXEID(OTSS, "D:\VB6\VB-NT\00_Best_VB_04\JPEG-Encoder-ART\Jpeg Slides PJpeg ART.exe")
    If TT = True Then Call Sleep(500)
    If XX > 100 Then End
    'Debug.Print XX
Loop Until TT = False

HH = FindWindow(vbNullString, "Jpeg Slides PJpeg ART.exe - Application Error")
If HH > 0 Then
    AppActivate "Jpeg Slides PJpeg ART.exe - Application Error", True
    SendKeys "{ENTER}", True
End If

Shell "D:\VB6\VB-NT\00_Best_VB_04\JPEG-Encoder-ART\Jpeg Slides PJpeg ART.exe", vbNormalFocus
XHH = False

End Sub



Sub ONEADAY()



If Minute(Now) < 30 And Minute(Now) < 30 <> TTHOURNOW Then
    TTHOURNOW = Minute(Now) < 30
    
    'Call KAT_RECORD_SMALL
    
'    Shell "D:\VB6\VB-NT\00_Best_VB_01\Drive_Detach\Drive_Detach.exe"
    'NORTON SONAR PROBLEM
    'LOW MODE TRY
    'FILTER EXCLUDE FOLDER
    
    Shell "D:\VB6\VB-NT\00_Best_VB_01\Time_Sync\Time Sync.exe"
    Shell "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\URL Logger.exe"
'    Shell "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\Cid-RunMe.exe"
    Shell "D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\ClipBoard Logger.exe"
    
    
    'CRC
    Shell "D:\VB6\VB-NT\00_Best_VB_04\INet_Content_Jpgs\INet_Content_Jpgs.exe VVV"
    
End If

If TTDAYHALFNOW <> Day(Now * 2) Then
    TTDAYHALFNOW = Day(Now * 2)
    
'    Shell "D:\VB6\VB-NT\00_Best_VB_01\Sync Files\Sync Files.exe -Quite", vbNormalFocus
'    Shell "D:\VB6\VB-NT\00_Best_VB_04\INet_Content_Jpgs\INet_Content_Jpgs.exe -Quite", vbNormalNoFocus
'    Shell "D:\VB6\VB-NT\00_Best_VB_01\WinRar_Arc_Notes_VB\WinRarArc_Notes_VB6.exe"
    
End If
If TTDAYNOW <> Day(Now) Then
    TTDAYNOW = Day(Now)
        
    Shell "D:\VB6\VB-NT\00_Best_VB_01\Sync Files\Sync Files.exe -Quite", vbNormalFocus
'    Shell "D:\VB6\VB-NT\00_Best_VB_01\WebDates_Ace\WebDates.exe", vbMinimizedNoFocus


End If

If TTDAYMODNOW <> Day(Now) Mod 3 Then
    TTDAYMODNOW = Day(Now) Mod 3
'    Shell "D:\VB6\VB-NT\00_Best_VB_01\WinRar Archive Temp Folder\WinRar Archive Temp Folder RAR.exe"

 
End If

End Sub

Private Sub Timer6_Timer()
If MOUSE > Now Then Exit Sub
'MOUSE = 0

If PROGSTARTFLAG < Now Then Exit Sub


'XX = FindWinPart("Norton Utilities", True)
XX = FindWindow("ThunderRT6FormDC", "Norton Utilities")
If XX > 0 And MOUSE < Now Then
     ShowWindow XX, SW_FORCEMINIMIZE
End If


XX = FindWinPart("TreeSize Professional", True)
If XX > 0 And MOUSE < Now Then ShowWindow XX, SW_MINIMIZE
 
'
XX = FindWindow("PROCEXPL", vbNullString)

If XX > 0 And MOUSE < Now Then ShowWindow XX, SW_FORCEMINIMIZE
 
 If TIMEDROP = 0 Then TIMEDROP = Now + TimeSerial(0, 25, 0)
If TIMEDROP < Now Then Exit Sub


XX = FindWindow(vbNullString, "On-Screen Keyboard")

If XX > 0 And MOUSE < Now Then
ShowWindow XX, SW_MINIMIZE
End If
 
 
XX = FindWindow(vbNullString, "Google Earth")

If XX > 0 And MOUSE < Now Then ShowWindow XX, SW_MINIMIZE


End Sub
Public Sub FindCursor(X, Y)

'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'FindCursor Tx, Ty
'Private Type POINTAPI
'        x As Long
'        Y As Long
'End Type

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
X = P.X ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
Y = P.Y '/ GetSystemMetrics(1) * Screen.Height

End Sub
Private Sub Timer_Mouse_Timer()

Dim X1, Y1
FindCursor X1, Y1
'dim OX1,OY1,MOUSE,XPAUSE

If OX1 <> X1 Or OY1 <> Y1 Then
    MOUSE = Now + TimeSerial(0, 0, 20)
End If

If OX1 <> X1 Or OY1 <> Y1 Then
    MOUSE2 = Now + TimeSerial(0, 10, 0)
End If

If MOUSE > Now Then
    If XPAUSE <> True Then
        XPAUSE = True
        
        
        'SUSPENDCODE
        'Debug.Print "YES"
    End If
Else
    If XPAUSE <> False Then
        XPAUSE = False
        
        
        'SUSPENDCODE
    End If
End If

If MOUSE2 > Now Then
    If XPAUSE2 <> True Then
        XPAUSE2 = True
    End If
Else
    If XPAUSE2 <> False Then
        XPAUSE2 = False
    End If
End If



'If XPAUSE = True Then Command1.BackColor = &HFF&
'If XPAUSE = 0 Then Command1.BackColor = &HFF8080

OX1 = X1
OY1 = Y1

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


Private Sub KAT_RECORD_SMALL()
'Timer7.Interval = 40000

'i = Hour(Now)
'If AGEI = i Then Exit Sub
'AGEI = i


'313
Debug.Print "---------------------------------"
Debug.Print "---------------------------------"
Debug.Print "---------------------------------"
For R = 1 To 3
If R = 1 Then FD = "M:\Kat MP3 Recorder\"
'If R = 2 Then FD = "J:\Kat MP3 Recorder\"
'If R = 3 Then FD = "I:\Kat MP3 Recorder\"
'If R = 4 Then FD = "I:\Kat MP3 Recorder ACER\"
If R = 2 Then FD = "D:\Kat MP3 Recorder\"
If R = 3 Then FD = "C:\Kat MP3 Recorder\"
On Error Resume Next

MkDir FD

If Err.Number > 0 Then Exit Sub

XD = FD + Dir(FD + "*.MP3")
Debug.Print XD
On Error Resume Next
Do While XD <> ""
    DoEvents
    FR = FreeFile
    Open FD + XD For Binary Lock Read Write As #FR
        TS = LOF(FR)
    Close #FR
    If TS <= 2048 Then Kill FD + XD
    XD = Dir
    
    If XD = OXD Then Exit Do
    OXD = XD
Loop
Next

End Sub

Private Sub Timer8_Timer()
DD = Hour(Now)
If DD = VDD Then Exit Sub

VDD = Hour(Now)
If DD < 5 Or DD >= 21 Then
i = SetScreenSaverState(True, True)
Else
i = SetScreenSaverState(False, False)
End If


End Sub

Private Sub Timer9_Timer()
Dim H As Long
H = FindWindow("GfxUI", vbNullString)
If H > 0 Then
    Call SetForegroundWindow(H)
    
    SendKeys ("{ESC}"), 0
    Timer9 = False
End If


End Sub
