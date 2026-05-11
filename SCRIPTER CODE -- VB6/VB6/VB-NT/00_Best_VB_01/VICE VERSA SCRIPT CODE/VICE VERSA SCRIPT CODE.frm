VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   Caption         =   "TIMER CODE"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   3285
   StartUpPosition =   2  'CenterScreen
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
Dim OLSN, NEROID, TTX As String, DIR1 As String
Dim VARCENTER
Dim OLDi, iTRIP, XHH


Private Sub Form_Load()

'Set FS = CreateObject("Scripting.FileSystemObject")
'Me.WindowState = 1

Load Form1MIX

Call Form1MIX.cmdEject_Click
End


End Sub

Private Sub Form_Resize()
If Me.WindowState <> 0 Then Exit Sub
If VARCENTER = True Then Exit Sub
Me.Top = Screen.Height / 2 - Me.Height / 2 + 400
Me.Left = Screen.Width / 2 - Me.Width / 2 + 400
VARCENTER = True

End Sub

Private Sub Label1_Click()
OLSN = 0
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

Private Sub Timer1_Timer()

    

Set f = FS.GetFile("D:\VB6\VB-NT\00_Best_VB_04\RS232 LOGGER\RS232 DOOR REED DATA\LAST DOOR CLOSE.txt")
    
If XONDATE > 0 And XONDATE <> f.datelastmodified Then
    Shell "D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS\WelCome\WelCome.exe"
End If

XONDATE = f.datelastmodified





End Sub

Private Sub Timer2_Timer()


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

Private Sub Timer4_Timer()

If Second(Now) Mod 2 = 0 Then
Call CHK_CDROM_DRIVES_AND_COPY_FILES
End If
If Second(Now) Mod 5 = 0 And 1 = 2 Then
    Debug.Print str(Timer) + "Start -- CHK_M2_DRIVES_AND_MOVE_FILES"
    Call CHK_M2_DRIVES_AND_MOVE_FILES
    Debug.Print str(Timer) + "Start -- CHK_W311_DICTOR_AND_MOVE_FILES_SHARE_SEARCH"
    Call CHK_W311_DICTOR_AND_MOVE_FILES_SHARE_SEARCH
    Debug.Print str(Timer) + "Start -- CHK_NIKON_AND_MOVE_FILES_SHARE_SEARCH"
    Call CHK_NIKON_AND_MOVE_FILES_SHARE_SEARCH
    Debug.Print str(Timer) + "Start -- CHK_KAT_RECORD"
    Call CHK_KAT_RECORD
    Debug.Print str(Timer) + "Done"
End If
    
End Sub

Sub CHK_M2_DRIVES_AND_MOVE_FILES()

        On Local Error Resume Next
        
        For RY = 1 To 2
BACKTOFORM2:
        If RY = 1 Then DIR1 = "H:\DCIM\100MSDCF"
        If RY = 2 Then DIR1 = "L:\DCIM\100MSDCF"
        If RY > 2 Then Exit Sub
        
        Err.Clear
        Set f = FS.GetDrive(Mid(DIR1, 1, 1) + ":")
        
        If Err.Number > 0 Then RY = RY + 1: GoTo BACKTOFORM2
        
        
        If f.IsReady = False Then RY = RY + 1: GoTo BACKTOFORM2
        If FS.FolderExists(DIR1) = False Then RY = RY + 1: GoTo BACKTOFORM2
            
        If FS.FolderExists("M:\# W880i\VIDEO\") = True Then DRI = "M" Else DRI = "D"
        
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

        On Local Error Resume Next
'Exit Sub
        For RY = 1 To 1
BACKTOFORM3:
        If RY = 1 Then DIR1 = "G:"
        'If RY = 2 Then DIR1 = "L:\DCIM\100MSDCF"
        If RY > 2 Then Exit Sub
        
        Err.Clear
        Set f = FS.GetDrive(Mid(DIR1, 1, 1) + ":")
        
        If Err.Number > 0 Then RY = RY + 1: OLSN = "": GoTo BACKTOFORM3
        
        If f.IsReady = False Then RY = RY + 1: OLSN = "": GoTo BACKTOFORM3
        If FS.FolderExists(DIR1) = False Then RY = RY + 1: GoTo BACKTOFORM3
        If f.SerialNumber = OLSN Then RY = RY + 1: GoTo BACKTOFORM3
        
        'NOT WTH DVD
        OLSN = f.SerialNumber
        Debug.Print Hex$(f.SerialNumber) + " -- " + f.VolumeName
            
            TTX = "M:\Videos\00 Vid Films\VOBS\# The Mother\" + Hex$(f.SerialNumber) + "--" + f.VolumeName '+ "\"
            
            If FS.FolderExists(DIR1 + "\VIDEO_TS") = True Then
                MkDir TTX
                Clipboard.Clear
                Clipboard.SetText TTX
            End If
            
            TTX = "M:\0 00 Art\# PHOTO #\" + Hex$(f.SerialNumber) + "--" + f.VolumeName '+ "\"
            
            If FS.FolderExists(DIR1 + "\VIDEO_TS") = False Then
                MkDir TTX
                Clipboard.Clear
                Clipboard.SetText TTX
            End If
            
            
            MsgBox Hex$(f.SerialNumber) + " -- " + f.VolumeName, vbMsgBoxSetForeground + vbCritical
        Exit Sub
            
        DRI = "M"
        'If Dir(DIR1 + "\*.*") <> "" Then
            
            
            TTX = DRI + ":\Videos\00 Vid Films\VOBS\The Mother\" + Hex$(f.SerialNumber) + "--" + f.VolumeName + "\"
'M:\0 00 Art\# PHOTO #
            TTX = DRI + ":\00 Art\# PHOTO #\" + Hex$(f.SerialNumber) + "--" + f.VolumeName + "\"
            MkDir TTX
            
            Call Form1MIX.Form_Load
            ScanPath.ListView1.ListItems.Clear
            ScanPath.cboMask.Text = "*.*"
            ScanPath.chkSubFolders = vbChecked
            ScanPath.txtPath.Text = DIR1
            Call ScanPath.cmdScan_Click
            'MsgBox "BEGIN DISK"
            
            For WE5 = 1 To ScanPath.ListView1.ListItems.Count
                WE7 = ScanPath.ListView1.ListItems.Count
                A1$ = ScanPath.ListView1.ListItems.item(WE5).SubItems(1)
                B1$ = ScanPath.ListView1.ListItems.item(WE5)
                H1 = Mid(A1, 4)
                Call DIR_ON_ANOTHER_DRIVE(TTX + H1)
                Set F1 = FS.GetFile(A1 + B1)
                Dim FH
                FH = F1.Size
                Set F1 = FS.GetFile(TTX + H1 + B1)
                
                Debug.Print str(WE5) + " Of " + str(WE7) + "---" + str(F1.Size) + "--- " + A1 + B1 + "  -- " + f.VolumeName + " -- " + Hex$(f.SerialNumber)
                FH = FH <> F1.Size
                If ScanPath.ListView1.ListItems.Count = WE5 Then
                Th = 1
                End If
                            If Th = 1 Then Call Form1MIX.cmdEject_Click
                If FS.FileExists(TTX + H1 + B1) = False Then
                    FH = True
                End If
                
                If FH Then
                                If Th = 1 Then Call Form1MIX.cmdEject_Click
                    Call ShellFileCopy(A1$ + B1, TTX + H1 + B1, False)
                    If FS.FileExists(TTX + H1 + B1) = False Then
                        
                        'NOT WITH DVD
                        'Th = 1
                        MsgBox "DVD COPY NOT DONE"
                        Exit For
                    End If
                    
                    'If NEROID < Now And NEROID > 0 Then
                    '    NEROID = 0
                    '    pid = -1
                    '    PidBool = cProcesses.GetEXEID(pid, "C:\Program Files\Ahead\Nero ShowTime\ShowTime.exe")
                    '    Process_Kill (pid)
                    'End If
                    
                    
                    'Th = 1
                End If
            Next
            If Th = 1 Then Call Form1MIX.cmdEject_Click
        
        'End If
                    
    Next

End Sub



Sub CHK_W311_DICTOR_AND_MOVE_FILES_SHARE_SEARCH()

'MAKE SURE WHEN SHARE ALL DRIVES SET TO NAME
'LIKE WS_550M
'AND NOT WITH DRIVE LETTER AFTER LIKE WS_550M (H)
'DR1 = "\\55-88-happy\WS_311M"
'DR2 = "\\55-88-happy\WS_550M"
        
        
            
        DIR1 = ""
        Set dc = FS.Drives
        For Each D In dc
            '1=REMOVABLE
            If D.DriveType = 1 Then
                DIR1 = D.DriveLetter + ":\DSS_FLDA"
                DIR2 = D.DriveLetter + ":\DSS_FLD"
                If FS.FolderExists(DIR1) = True Then Exit For
                            
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
                    Set f = FS.GetDrive("M:")
                    If f.IsReady = True Then
                        Call ShellFileMove(DIR1 + "\*.WMA", "M:\05 Media\04 Olympus WS-550M\")
                    Else
                        Call ShellFileMove(DIR1 + "\*.WMA", "D:\05 Media\04 Olympus WS-550M\")
                    End If
                End If
                                
                If InStr(FILE1, "WS310") Then
                    Set f = FS.GetDrive("M:")
                    If f.IsReady = True Then
                        Call ShellFileMove(DIR1 + "\*.WMA", "M:\05 Media\03 Olympus WS-311M\")
                    Else
                        Call ShellFileMove(DIR1 + "\*.WMA", "D:\05 Media\03 Olympus WS-311M\")
                    End If
                End If
            
                If InStr(FILE1, "VN6") Then
                    Set f = FS.GetDrive("M:")
                    If f.IsReady = True Then
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
        
        For RY = 1 To 2
BACKTOFOR:
        On Local Error Resume Next
        If RY = 1 Then DIR1 = "H:\DCIM"
        If RY = 2 Then DIR1 = "P:\DCIM"
        If RY > 2 Then Exit Sub
        
        Set f = FS.GetDrive(Mid(DIR1, 1, 1) + ":")
        XXV = f.VolumeName
        'Set d = FS.GetDrive(FS.GetDriveName(drvPath))

        If Err.Number > 0 Then RY = RY + 1: GoTo BACKTOFOR
        
        If f.IsReady = False Then RY = RY + 1: GoTo BACKTOFOR
        If FS.FolderExists(DIR1) = False Then RY = RY + 1: GoTo BACKTOFOR
        Set f = FS.GetFolder(DIR1)
        If f.Size > 0 Then
            
        Set f = FS.GetDrive("M:")
        If f.IsReady = True Then DRI = "M" Else DRI = "D"
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
                            Set f = FS.GetFolder(DIR1 + "\" + O2)
                            XF = f.Size
                            If XF = 0 Then RmDir (DIR1 + "\" + O2): Exit For
                            If InStr(NIKON_OSERIALDR, str(D.SerialNumber) + " " + DIR1 + O2) = 0 And XF > 0 Then
                                SERIALDR = SERIALDR + " " + str(D.SerialNumber) + " " + DIR1 + O2
                                
                                If FS.FolderExists("M:\# NIKON\") = True Then
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

        'On Local Error Resume Next
        
        'If R2 = 1 Then DIR1 = "\\55-88-happy\C\05 Media\ASUS - KAT RECORD"
        'If R2 = 1 Then Call ShellFileMove2(DIR1 + "\" + "*.MP3", "M:\05 Media\KAT RECORD\KAT RECORD ASUS\", True)
        
        DIR1 = "D:\05 Media\KAT RECORD"
        
        
        XXD = Dir(DIR1 + "\*.MP3")
        
        If XXD = "" Then Exit Sub
        
        If FileInUse(DIR1 + "\" + XXD) = True Then Exit Sub
            
        If FS.DriveExists("M:") = False Then Exit Sub

        Call ShellFileMove(DIR1 + "\" + XXD, "M:\05 Media\KAT RECORD\KAT RECORD ACER\", True)
        

'KAT_OSERIALDR = SERIALDR


End Sub



Private Sub Timer5_Timer()


'1 MINUTE
'
MinModAgwa = Second(Now) 'Mod 2
MinModAgwa2 = MinModAgwa
If MinModAgwa > 15 And oMinModAgwa <> Minute(Now) Then
    If Now - Int(Now) < SunSet And Now - Int(Now) > SunRise Then
        oMinModAgwa = Minute(Now)
        
        'Call OutputData
        On Error Resume Next
        DI7 = Shell("D:\VB6\VB-NT\00_Best_VB_04\Webcam_Motion\WebCamMatts.exe Quite", vbHide)
        On Error GoTo 0
        'Process_Low (Di7)
    
    End If
End If




End Sub



Private Sub Timer_ART_TRIP_WIRE_Timer()

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
If FS.FileExists("K:\TEMP\ART_PROG_TRIP_WIRE.txt") = False Then Exit Sub

Set f = FS.GetFile("K:\TEMP\ART_PROG_TRIP_WIRE.txt")
i = f.datelastmodified
If OLDi <> i Then iTRIP = Now + TimeSerial(0, 0, 25)
OLDi = i
'Debug.Print i, OLDi

If iTRIP > Now Then Exit Sub
XHH = False

'TT = cProcesses.Convert(HH, OTSS, cnFromhWnd Or cnToProcessID)
OTSS = -1
TT = cProcesses.GetEXEID(OTSS, "D:\VB6\VB-NT\00_Best_VB_04\JPEG-Encoder-ART\Jpeg Slides PJpeg ART.exe")

If TT = True Then Process_Kill (OTSS)

xx = 0
Do
    xx = xx + 1
    OTSS = -1
    TT = cProcesses.GetEXEID(OTSS, "D:\VB6\VB-NT\00_Best_VB_04\JPEG-Encoder-ART\Jpeg Slides PJpeg ART.exe")
    If TT = True Then Call Sleep(500)
    If xx > 100 Then End
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

