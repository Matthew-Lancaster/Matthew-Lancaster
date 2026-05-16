VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "DARK FORM"
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H001D1D1D&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8430
      Left            =   390
      TabIndex        =   0
      Top             =   405
      Width           =   9750
      Begin MSComctlLib.ProgressBar PBar1 
         Height          =   225
         Left            =   1380
         TabIndex        =   2
         Top             =   90
         Visible         =   0   'False
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   8250
         Left            =   270
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   412.5
         ScaleMode       =   0  'User
         ScaleWidth      =   438.75
         TabIndex        =   1
         Top             =   270
         Width           =   8775
         Begin VB.DirListBox Dir1 
            Height          =   1890
            Left            =   7095
            TabIndex        =   4
            Top             =   1410
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Timer Timer2 
            Interval        =   50
            Left            =   3105
            Top             =   2595
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   2700
            Top             =   2175
         End
      End
      Begin VB.Label LabTIME2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "TIME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9075
         TabIndex        =   5
         Top             =   1185
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label LabTIME1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         Caption         =   "TIME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9090
         TabIndex        =   3
         Top             =   870
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, OTT, TT, NOCLICK, ADATE1, CNT, TX
Dim SECCNT, SECCNTLIMIT
Public WTrue As Double
Public HWTrue As Double
Public KWTrue As Double
Public TW1 As Double
Public TW2 As Double
Public TW3 As Double


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
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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

Sub FileInUseDelay(TX$)
        
    qq = Now + TimeSerial(0, 1, 30)
    Do
        DoEvents
        II = FileInUse(TX$)
        If II = True Then Sleep (200)
    Loop Until II = False Or qq < Now
    
    If II = True Then
        MsgBox "File Stuck In Use " + vbCrLf + TX$ + vbCrLf + "Longer than 1 min 30 secs"
        Stop
    End If
End Sub



Private Sub Form_Click()
End
End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then End
Load ScanPath

Me.Height = Screen.Height
Me.Width = Screen.Width
Me.Top = 0
Me.Left = 0
Frame1.Caption = App.Title
Frame1.Height = Me.ScaleHeight
Frame1.Width = Me.ScaleWidth
Frame1.Top = Me.ScaleHeight / 2 - Frame1.Height / 2
Frame1.Left = Me.ScaleWidth / 2 - Frame1.Width / 2

'ActivateForm
'Picture1.AutoSize = False
DoEvents
'Picture1.Height = Picture1.Height - 1500
'Picture1.Width = Picture1.Width - 1500
Picture1.Top = (Me.ScaleHeight / 2 - Picture1.Height / 2)
Picture1.Left = (Me.ScaleWidth / 2 - Picture1.Width / 2)

TT = Second(Now)

x = -1
Call Picture1_Click

End Sub

Private Sub Frame1_Click()
End
End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)
End
End Sub

Private Sub Picture1_Click()

If NOCLICK = True Then Exit Sub

x = Not x
ColorCycle
If x = 0 Then
    Set Picture1.Picture = Nothing
    Set Picture1.Picture = LoadPicture(App.Path + "\01 YinYang4 BlkBG.jpg")
    
Else
    Set Picture1.Picture = Nothing
    Set Picture1.Picture = LoadPicture("D:\VB6\VB-NT\00_Best_VB_04\Webcam_Motion\VBDataNoArchive\Web_Cam Pic Always.jpg")
    SECCNT = 200
End If
Picture1.Top = (Me.ScaleHeight / 2 - Picture1.Height / 2)
Picture1.Left = (Me.ScaleWidth / 2 - Picture1.Width / 2)

Call LABTIMESUB


'OTT = -1

Call PBar1SUB



End Sub

Private Sub Timer1_Timer()

If x = 0 Then PBar1.Visible = False: LabTIME1.Visible = False: LabTIME2.Visible = False
If x = 0 Then Exit Sub

'MALE>HE WANTS NOW TO THE NEAREST MINUTE

Call PBar1SUB

SECCNT = SECCNT + 1
SECCNTLIMIT = 3
If ADATE1 < Now - TimeSerial(0, 10, 0) And SECCNT > SECCNTLIMIT Then
    If ScanPath.ListView1.ListItems.Count = 0 Then
        WEBCAM
        SECCNT = 0
        PBar1SUB
        'Exit Sub
    End If
    
    SECCNT = 0
    
    NOCLICK = True
    
    
    CNT = CNT + 1
    If CNT > ScanPath.ListView1.ListItems.Count - 1 Then CNT = 0
    
    A1 = ScanPath.ListView1.ListItems.Item(CNT).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(CNT)
        
        
    TX = A1 + B1
    
    Call LABTIMESUB
    
    FileInUseDelay (TX)
    On Error Resume Next
    Set Picture1.Picture = Nothing
    Set Picture1.Picture = LoadPicture(TX)
    If Err.Number = 0 Then TT = 0
    NOCLICK = False
    If CNT Mod 10 = 0 Then Call ColorCycle
    Exit Sub
End If


If ADATE1 < Now - TimeSerial(0, 10, 0) Then Exit Sub

If OTT <> ADATE1 And x = -1 Then
    ScanPath.ListView1.ListItems.Clear

    Sleep 1
    
    NOCLICK = True
    TX = "D:\VB6\VB-NT\00_Best_VB_04\Webcam_Motion\VBDataNoArchive\Web_Cam Pic Always.jpg"
    Call LABTIMESUB
    FileInUseDelay (TX)
    On Error Resume Next
    Set Picture1.Picture = Nothing
    Set Picture1.Picture = LoadPicture(TX)
    If Err.Number = 0 Then TT = 0
    NOCLICK = False
    Call ColorCycle

End If
OTT = ADATE1



End Sub

Sub LABTIMESUB()


LabTIME1.Top = PBar1.Height + 15 'Picture1.Top - LabTIME1.Height
LabTIME1.Left = (Picture1.Left + Picture1.Width) - LabTIME1.Width

LabTIME2.Top = LabTIME1.Top + LabTIME1.Height + 15 'Picture1.Top - LabTIME1.Height
LabTIME2.Left = (Picture1.Left + Picture1.Width) - LabTIME2.Width


If ScanPath.ListView1.ListItems.Count > 0 And TX <> "" Then
    WEBFILE = TX
Else
    WEBFILE = "D:\VB6\VB-NT\00_Best_VB_04\Webcam_Motion\VBDataNoArchive\Web_Cam Pic Always.jpg"
End If


Set Fs = CreateObject("Scripting.FileSystemObject")
Set F = Fs.getfile(WEBFILE)
ADATE1 = F.datelastmodified

LabTIME1 = Format(Now, "DD MMM YY -- HH:MM:Ss")
LabTIME2 = Format(ADATE1, "DD MMM YY -- HH:MM:Ss")


End Sub



Sub PBar1SUB()
DoEvents

PBar1.Top = 0 'Picture1.Top - PBar1.Height
PBar1.Left = Picture1.Left
PBar1.Width = Picture1.Width
PBar1.Height = 88
PBar1.Max = 60
PBar1.Min = 0

Call LABTIMESUB

'TT = Minute(Now)
TT = TT + 1

If ScanPath.ListView1.ListItems.Count > 0 Then
    PBar1.Max = SECCNTLIMIT
    If SECCNT > SECCNTLIMIT Then SECCNT = SECCNTLIMIT
    PBar1.Value = SECCNT

Else


    If TT > 60 Then
        PBar1.Max = TT
    Else
        PBar1.Max = 60
    End If
    PBar1.Value = TT

End If





If Picture1.Width < 1 Then Exit Sub

PBar1.Visible = True
LabTIME1.Visible = True
LabTIME2.Visible = True

End Sub



Private Sub ColorCycle()

TTR = 0
Do
WTrue = WTrue + TW1

If WTrue > 255 Then TW1 = -6: WTrue = WTrue + TW1
If WTrue < 1 Then TW1 = 6: WTrue = WTrue + TW1

HWTrue = HWTrue + TW2

If HWTrue > 255 Then TW2 = -7: HWTrue = HWTrue + TW2
If HWTrue < 1 Then TW2 = 7: HWTrue = HWTrue + TW2

KWTrue = KWTrue + TW3

If KWTrue > 255 Then TW3 = -8: KWTrue = KWTrue + TW3
If KWTrue < 1 Then TW3 = 8: KWTrue = KWTrue + TW3
   
TTR = TTR + 1
Loop Until TTR = 100
   
Frame1.BackColor = RGB(KWTrue / 2, HWTrue / 2, WTrue / 2) ' Set drawing color.
'FRAME1.ForeColor = RGB(255 - KWTrue, 255 - HWTrue, 255 - WTrue)
'frmPassLock.Label4.BackColor = RGB(kwtrue, hwtrue, wtrue)   ' Set drawing color.
'frmPassLock.Label4.ForeColor = RGB(255 - kwtrue, 255 - hwtrue, 255 - wtrue)
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Call ColorCycle
End Sub


Sub WEBCAM()


        DAYSBACK = 1
        If DAYSBACK > 0 Then
            XX = True
            
            
            'ScanPath.DTPicker1(0) = Now - (DaysBack - 1)
            
            Dir1.Path = "M\0 00 Art Loggers\Web_Cam\"
            R5 = Dir1.ListCount
            OLDFC = 0
            TTH = 0
            Do
                R5 = R5 - 1
                Td = Dir1.List(R5)
                
                ScanPath.TxtPath.Text = Td
                Call ScanPath.cmdScan_Click
                
                If ScanPath.ListView1.ListItems.Count <> OLDFC Then
                    TTH = TTH + 1
                End If
                OLDFC = ScanPath.ListView1.ListItems.Count
            
            
            
            Loop Until TTH >= DAYSBACK Or R5 = 0
            'Slider1.Value = 0

    End If

End Sub
