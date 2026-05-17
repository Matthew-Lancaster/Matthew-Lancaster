VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form WMI_CPU2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   2880
      Top             =   1530
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   480
      Left            =   390
      TabIndex        =   0
      Top             =   285
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   847
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1410
      Top             =   1110
   End
End
Attribute VB_Name = "WMI_CPU2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ForeG, MELeft, MEW, METOP, ENDE
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim Rect4 As RECT

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
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

Private asCpuPaths() As String
Private m_objCPUSet As SWbemObjectSet
Private m_objWMINameSpace As SWbemServices
'Option Explicit
    
    
Public Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Public Function NotAlwaysOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function


Private Sub Form_Load()
If App.PrevInstance = True Then End
End
ProgressBar1.ToolTipText = App.Path + " -- " + App.EXEName


Me.Show
METOP = 30
Me.Top = METOP
MELeft = 680
Me.Left = MELeft
ProgressBar1.Top = 0
ProgressBar1.Left = 0
ProgressBar1.Height = 140
MEW = 1950
ProgressBar1.Width = MEW
Me.Width = ProgressBar1.Width
Me.Height = ProgressBar1.Height
AlwaysOnTop (WMI_CPU2.hWnd)

Dim oCpu As SWbemObject 'WMI Object, in this case, local CPUs
Dim sPath As String, sCaption As String

Dim lElement As Long
ReDim asCpuPaths(0) As String


On Error GoTo ErrorHandler

'Get Default NameSpace, which will be the one for the local machine
'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Set m_objWMINameSpace = GetObject("winmgmts:")
'lstCPU.Clear




CleanUp:
Set oCpu = Nothing


Exit Sub

ErrorHandler:
MsgBox "CPU Information could not be displayed due to the following error: " & Err.Description, , "WMI Demo Failed"
GoTo CleanUp

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set m_objCPUSet = Nothing
Set m_objWMINameSpace = Nothing
End Sub

Private Sub ProgressBar1_Click()
ENDE = ENDE + 1

If ENDE > 2 Then End

End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
    End
End If

End Sub

Private Sub Timer1_Timer()

On Local Error Resume Next
'Call lstCPU_Click

Dim oCpu As SWbemObject
Set m_objCPUSet = m_objWMINameSpace.InstancesOf("Win32_Processor")
Set oCpu = m_objCPUSet("\\MATT-555ROIDS\root\cimv2:Win32_Processor.DeviceID=""CPU0""")
WMI_CPU2.ProgressBar1.Value = oCpu.LoadPercentage
'WMI_CPU2.ProgressBar1.Value = 50
CleanUp:
Set oCpu = Nothing


qq = GetForegroundWindow
If qq = ForeG Then Exit Sub
ForeG = qq

If FindWindow(vbNullString, "Extreme Lock Build: 2011") > 0 Then
'    Exit Sub
    Me.Top = Screen.Height - (Me.Height) - 200
    Me.Left = 10

Else
    Me.Top = METOP
    Me.Left = MELeft
End If
    
    'Me.Top = Screen.Height - (Me.Height) - 800
    'Me.Left = 10


AlwaysOnTop (WMI_CPU2.hWnd)




End Sub

Private Sub Timer2_Timer()
ENDE = 0
End Sub
