VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form WMI_CPU 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "WMI Demo - CPU Information"
   ClientHeight    =   5190
   ClientLeft      =   2355
   ClientTop       =   1500
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   795
      Top             =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   1305
      Top             =   2280
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   2550
      Left            =   0
      TabIndex        =   7
      Top             =   780
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   4498
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2055
      Top             =   135
   End
   Begin VB.TextBox txtCpu 
      Height          =   2715
      Left            =   7365
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   885
      Width           =   3705
   End
   Begin VB.ListBox lstCPU 
      Height          =   2595
      Left            =   3675
      TabIndex        =   0
      Top             =   1410
      Width           =   2580
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "END PROGS AND REBOOT"
      Top             =   570
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Cpu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   -15
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "--%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   -15
      TabIndex        =   5
      ToolTipText     =   "END PROGS"
      Top             =   270
      Width           =   495
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Information for Selected CPU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3690
      TabIndex        =   4
      Top             =   555
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4500
      TabIndex        =   3
      Top             =   840
      Width           =   585
   End
   Begin VB.Label lblCPUList 
      AutoSize        =   -1  'True
      Caption         =   "CPUs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4500
      TabIndex        =   1
      Top             =   1125
      Width           =   480
   End
End
Attribute VB_Name = "WMI_CPU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************\
'NOTES:

'YOU MUST HAVE WMI SDK INSTALLED.  YOU CAN GET IT AT
'http://msdn.microsoft.com/downloads/sdks/wmi/default.asp
'***********************************************************************
Const WM_CLOSE = &H10
Dim Oth, CPUNOW, CLOSEPROG
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Width2

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Function MoveWindow _
        Lib "user32" _
        (ByVal hwnd As Long, _
         ByVal x As Long, _
         ByVal y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim Ryu4 As RECT
Dim Ryu5 As RECT

Private asCpuPaths() As String
Private m_objCPUSet As SWbemObjectSet
Private m_objWMINameSpace As SWbemServices
'Option Explicit

    
Public Function AlwaysOnTop(ByVal hwnd As Long)  'Makes a form always on top
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Public Function NotAlwaysOnTop(ByVal hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function



Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then End

ProgressBar1.ToolTipText = App.Path + " -- " + App.EXEName
Label2.ToolTipText = App.Path + " -- " + App.EXEName

WMI_CPU.Width = ProgressBar1.Width
xxr20 = GetWindowRect(Me.hwnd, Ryu5)
Width2 = Ryu5.Right - Ryu5.Left

WMI_CPU.Height = ProgressBar1.Top + ProgressBar1.Height
WMI_CPU.Width = ProgressBar1.Width
Dim oCpu As SWbemObject 'WMI Object, in this case, local CPUs
Dim sPath As String, sCaption As String

Dim lElement As Long
ReDim asCpuPaths(0) As String


On Error GoTo ErrorHandler

'Get Default NameSpace, which will be the one for the local machine
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Set m_objWMINameSpace = GetObject("winmgmts:")
lstCPU.Clear


'Get CPU set

GoTo jump:

Set m_objCPUSet = m_objWMINameSpace.InstancesOf("Win32_Processor")
sCaption = m_objCPUSet.Count & " processor"
If m_objCPUSet.Count <> 1 Then sCaption = sCaption & "s"
sCaption = sCaption & " detected on this machine"
lblTitle.Caption = sCaption
'Populate list box with CPU names
                
                
'Set oCpu = m_objCPUSet
                
'rr$ = oCpu.Path_.DisplayName
GoTo jump:
                
For Each oCpu In m_objCPUSet
     With oCpu
        sPath = .Path_ & ""
            If sPath <> "" Then
                lstCPU.AddItem .Name
                'save path to array, so on machines with multiple CPUs,
                'each can be identified and their info loaded into text box
                
                lElement = IIf(asCpuPaths(0) = "", 0, UBound(asCpuPaths) + 1)
                ReDim Preserve asCpuPaths(lElement) As String
                asCpuPaths(lElement) = sPath
            End If
     End With
Next
jump:

'If lstCPU.ListCount <> 0 Then lstCPU.ListIndex = 0
'Call lstCPU_Click
 



CleanUp:
Set oCpu = Nothing

'Load WMI_CPU2


Exit Sub

ErrorHandler:
MsgBox "CPU Information could not be displayed due to the following error: " & Err.Description, , "WMI Demo Failed"
GoTo CleanUp
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set m_objCPUSet = Nothing
Set m_objWMINameSpace = Nothing
End Sub

Private Sub Label1_Click()

Call ENDPROGS

End Sub

Private Sub Label2_Click()
End

End Sub

Private Sub Label3_Click()
Shell "C:\Program Files\Tweak-XP Pro\reboot.exe", vbNormalFocus
Call ENDPROGS
End
End Sub

Private Sub lstCPU_Click()
'Refer to SDK documentation for more detail about each of these properties
Dim sInfoString As String
On Error Resume Next
'Set oCpu = m_objCPUSet(asCpuPaths(lstCPU.ListIndex))

'With oCpu
    'sInfoString = "Description: " & .Description & vbCrLf
    'sInfoString = sInfoString & "Processor ID: " & .ProcessorID & vbCrLf
    'sInfoString = sInfoString & "Status: " & .Status & vbCrLf
    'sInfoString = sInfoString & "Manufacturer: " & .Manufacturer & vbCrLf
    'sInfoString = sInfoString & "Availability: " & AvailabilityToString(.Availability) & vbCrLf
    'sInfoString = sInfoString & "Load Percentage: " & .LoadPercentage & vbCrLf
    'sInfoString = sInfoString & "Current Clock Speed: " & .CurrentClockSpeed & " MHz" & vbCrLf
    'sInfoString = sInfoString & "Maximum Clock Speed: " & .MaxClockSpeed & vbCrLf
    'sInfoString = sInfoString & "Level 2 Cache Size: " & .L2CacheSize & vbCrLf
    'sInfoString = sInfoString & "Level 2 Cache Speed: " & .L2CacheSpeed & vbCrLf
    'sInfoString = sInfoString & "Power Management Supported: " & .PowerManagementSupported
'End With
'txtCpu.Text = sInfoString



End Sub
'Conversions from code to string were developed
'based on information in WMI SDK documentation
Private Function AvailabilityToString(Code As Integer) As String
Dim sAns As String

Select Case Code
    Case 1, 2
        sAns = "Unknown"
    Case 3
        sAns = "Running/Full Power"
    Case 4
        sAns = "Warning"
    Case 5
        sAns = "In Test"
    Case 6
        sAns = "Not Applicable"
    Case 7
        sAns = "Power Off"
    Case 8
        sAns = "Off Line"
    Case 9
        sAns = "Off Duty"
    Case 10
        sAns = "Degraded"
    Case 11
        sAns = "Not Installed"
    Case 12
        sAns = "Install Error"
    Case 13
        sAns = "Power Save - Unknown"
    Case 14
        sAns = "Power Save - Low Power Mode"
    Case 15
        sAns = "Power Save - Standby"
    Case 16
        sAns = "Power Cycle"
    Case 17
        sAns = "Power Save - Warning"
    Case Else
        sAns = "Unknown"
End Select
AvailabilityToString = sAns

End Function


        

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
If Button = 1 Then
    EXEs = """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".VBP"""
    Shell EXEs, vbNormalFocus
    End
End If
If Button = 2 Then
    Call ENDPROGS

End If

End Sub

Private Sub Timer1_Timer()
On Local Error Resume Next
'Call lstCPU_Click


If FindWindow(vbNullString, "CID Run Me.") = 0 And IsIDE = False Then End
If FindWindow(vbNullString, "CID Run Me.") = 0 Then End

Dim oCpu As SWbemObject
Set m_objCPUSet = m_objWMINameSpace.InstancesOf("Win32_Processor")
Set oCpu = m_objCPUSet("\\MATT-555ROIDS\root\cimv2:Win32_Processor.DeviceID=""CPU0""")
Label1 = Trim(Str(oCpu.LoadPercentage) + "%")
ProgressBar1.Value = oCpu.LoadPercentage
WMI_CPU2.ProgressBar1.Value = ProgressBar1.Value
CleanUp:
Set oCpu = Nothing



If ProgressBar1.Value <= 99 Or CPUNOW = 0 Then
    
    If CPUNOW = 0 Then CPUNOW = Now + TimeSerial(1, 0, 0)

Else
CPUNOW = 0
End If

If ProgressBar1.Value >= 99 And CPUNOW < Now And CPUNOW > 0 Then
Call ENDPROGS

End If

End Sub

Private Sub Timer2_Timer()

On Local Error Resume Next

xxr = FindWindow(vbNullString, "CID Run Me.")
'ttr$ = GetWindowTitle(Xxr)



If xxr > 0 Then

xxr18 = GetWindowRect(xxr, Ryu4)
xxr18 = Ryu4.Top + Ryu4.Bottom + Ryu4.Left + Ryu4.Right

If xxr18 <> xxr19 Then

'If Ryu4.Top = 0 Then Stop

xxr19 = xxr18
'If MeRyu3.Top < 40 Then
xxr20 = GetWindowRect(Me.hwnd, Ryu5)
MoveWindow Me.hwnd, Ryu4.Left - Width2, Ryu4.Top, Width2, (Ryu4.Bottom - Ryu4.Top), True
ProgressBar1.Height = (WMI_CPU.Height - ProgressBar1.Top) + 200
If WMI_CPU.Visible = False Then WMI_CPU.Visible = True
End If
End If


xxr = FindWindow(vbNullString, "Tidal Information...")
If xxr > 0 Then T$ = "Tidal Information..."
xxr = FindWindow(vbNullString, "Extreme Lock Build: 2011")
If xxr > 0 Then T$ = "Extreme Lock Build: 2011"
If Oth <> T$ Then
Oth = T$

If T$ = "Extreme Lock Build: 2011" Then
    AlwaysOnTop Me.hwnd
Else
    NotAlwaysOnTop Me.hwnd
End If
End If

End Sub

Private Sub Timer3_Timer()
'If Second(Now) Mod 10 <> 0 Then Exit Sub

On Error Resume Next

fr1 = FreeFile
Open "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures\Signature-CPU.txt" For Output As #fr1

Print #fr1, "Cpu Usage " + Label1

Close #fr1


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




'----------------------------------------
'Try this to find our pursuit of a neat window close
'looks same as our last one


Function CloseProgram(Optional ByVal ProcessHandle As Long, Optional ByVal ThreadHandle As Long, Optional CloseThread As Boolean = True) As Boolean
  
  Dim ReturnValue As Long
  Dim ExitCode    As Long

  If CloseThread = True Then
    ReturnValue = GetExitCodeThread(ThreadHandle, ExitCode)
    If ReturnValue <> 0 Then
  '    TerminateProcess ProcessHandle, ExitCode
       TerminateThread ThreadHandle, ExitCode
      CloseProgram = True
    Else
      CloseProgram = False
    End If
    
  Else
    ReturnValue = GetExitCodeProcess(ProcessHandle, ExitCode)
    If ReturnValue <> 0 Then
   '   TerminateThread ThreadHandle, ExitCode
       TerminateProcess ProcessHandle, ExitCode
      CloseProgram = True
    Else
      CloseProgram = False
    End If
  End If
  Exit Function
  
End Function


Function GetParentTitle(ByVal Handle As Long) As String
   Dim i As Long
   Dim j As Long, TTx1 As String
   i = Handle
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   TTx1 = GetWindowTitle(i)
   GetParentTitle = TTx1

End Function


Public Sub CloseProgramParent(ByVal Caption1 As String, ByVal Caption2 As String)

Dim i As Long
If Caption1 = "" And Caption2 = "" Then Exit Sub

If Caption1 = "" Then Caption1 = vbNullString
If Caption2 = "" Then Caption2 = vbNullString

i = FindWindow(Caption1, Caption2)

If i = 0 Then CLOSEPROG = False: Exit Sub

CLOSEPROG = True

Do While i <> 0
    SendMessage i, WM_CLOSE, 0&, 0&
    i = GetParent(i)
    If i = 0 Then Exit Do
Loop

End Sub



Sub ENDPROGS()

Dim PRC As Long, EXEs As String

Call CloseProgramParent("", "CID Run Me.")
Call CloseProgramParent("", "Tidal Information...")
Call CloseProgramParent("", "Active Idle...")
Call CloseProgramParent("", "URL Logger")
Call CloseProgramParent("", "KAT RECORDER TO USB PEN DRIVE AND HDD DUAL RECORDER")

Call CloseProgramParent("", "ClipBoard Logger")
Call CloseProgramParent("", "URL Logger")

Do
    Call CloseProgramParent("Winamp v1.x", "")
    Sleep 500
Loop Until CLOSEPROG = False


Call CloseProgramParent("", "")
Call CloseProgramParent("", "")
Call CloseProgramParent("", "")
Call CloseProgramParent("", "")
Call CloseProgramParent("", "")
Call CloseProgramParent("", "")
Call CloseProgramParent("", "")



PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\Cid-RunMe.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then If PRC > 0 Then Call Process_Kill(PRC)


PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\Tidal.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)



PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Active Idle\Active Idle.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)


PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\URL Logger.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)


'KEEPS TIMES WANT PROPER CLOSE
'PRC = -1
'EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\Cid-RunMe.exe"
'Rf = cProcesses.GetEXEID(PRC, EXEs)
'If PRC > 0 Then Call Process_Kill(PRC)

'EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Set NetWork Drives\Set-NetWork-Drives.exe", vbMinimizedNoFocus

'PRC = -1
'Rf1 = cProcesses.GetEXEID(PRC, "C:\PROGRAM FILES\PROCESSEXPLORER\PROCEXP.EXE")
'If PRC > 0 Then Call Process_Kill(PRC)

'PRC = -1
'Rf1 = cProcesses.GetEXEID(PRC, sWindowsFolder + "\SYSTEM32\TASKMGR.EXE")
'If PRC > 0 Then Call Process_Kill(PRC)

'PRC = -1
'Rf = cProcesses.GetEXEID(PRC, "C:\Program Files\KatMouse\KatMouse.exe")
'If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
Rf = cProcesses.GetEXEID(PRC, "C:\Program Files\Kat MP3 Recorder\Kat MP3 Recorder.exe")
'If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
Rf = cProcesses.GetEXEID(PRC, "D:\VB6\VB-NT\00_Best_VB_01\KAT USB PEN DRIVE RECORD\KAT_TIMER DUAL COPY.exe")
If PRC > 0 Then Call Process_Kill(PRC)

'PRC = -1
'Rf = cProcesses.GetEXEID(PRC, "C:\Program Files\SuperCopier2\SuperCopier2.exe")
'If PRC > 0 Then Call Process_Kill(PRC)

'PRC = -1
'EXEs = "C:\Program Files\Olympus\DeviceDetector\DevDtct2.exe"
'Rf = cProcesses.GetEXEID(PRC, EXEs)
'If PRC > 0 Then Call Process_Kill(PRC)

'PRC = -1
'EXEs = "C:\Program Files\WordWeb\wweb32.exe"
'Rf = cProcesses.GetEXEID(PRC, EXEs)
'If PRC > 0 Then Call Process_Kill(PRC)


PRC = -1
EXEs = "C:\Program Files\Cyberdyne Systems\iCEBar.v3\iCEBar.Net.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
'If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Wmi_CPU\WMICPU2 MINI.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Wmi_CPU\WMICPU1 BIG.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\WinAmp MP3% MINI #\WinAmp MP3% MINI.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\WinAmp MP3%\WinAmp MP3%-VideoBar.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\ClipBoard Logger.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

'PRC = -1
'EXEs = "D:\VB6\VB-NT\00_Best_VB_01\VU_METER_LOGGER\VU METER LOGGER.exe"
'Rf = cProcesses.GetEXEID(PRC, EXEs)
'If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\#00_Schedule_Timer\#00_Schedule_Timer-ACER.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\#00_Schedule_Timer\#00_Schedule_Timer-ASUS.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)
  
    
PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Drive_Detach\Drive_Detach.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)



PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_04\RS232 LOGGER\RS232 LOGGER.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Drives Gig\Drives_Gig.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\VPN_AutoDial IP\VPN_Auto_Grab_IP.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\WinAmp MP3%\WinAmp MP3%.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\VPN_AutoDial IP\VPN_Auto-Dialer IP.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Weather\BBCWeather.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\WinAmp MP3% MINI #\WinAmp MP3% MINI.EXE"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\WinAmp MP3%\WinAmp MP3%.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Drive_Detach\Drive_Detach.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\URL Logger.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\VPN_AutoDial IP\VPN_Auto-Dialer IP.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)


PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Drives Gig\Drives_Gig.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Weather\BBCWeather.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\Tidal.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)



PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_01\VPN_AutoDial IP\VPN_Auto_Grab_IP.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)

PRC = -1
EXEs = "D:\VB6\VB-NT\00_Best_VB_04\RS232 LOGGER\RS232 LOGGER.exe"
Rf = cProcesses.GetEXEID(PRC, EXEs)
If PRC > 0 Then Call Process_Kill(PRC)



End Sub




