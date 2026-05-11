VERSION 5.00
Begin VB.Form WMI_CPU 
   Caption         =   "WMI Demo - CPU Information"
   ClientHeight    =   3930
   ClientLeft      =   2415
   ClientTop       =   1845
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   6240
   Begin VB.Timer Timer1 
      Interval        =   500
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
      Left            =   45
      TabIndex        =   0
      Top             =   870
      Width           =   2580
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   540
      Left            =   2745
      TabIndex        =   5
      Top             =   885
      Width           =   1365
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
      Left            =   60
      TabIndex        =   3
      Top             =   270
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
      Left            =   45
      TabIndex        =   1
      Top             =   555
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

Private asCpuPaths() As String
Private m_objCPUSet As SWbemObjectSet
Private m_objWMINameSpace As SWbemServices
'Option Explicit

Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub Form_Load()


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

Exit Sub

ErrorHandler:
MsgBox "CPU Information could not be displayed due to the following error: " & Err.Description, , "WMI Demo Failed"
GoTo CleanUp
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set m_objCPUSet = Nothing
Set m_objWMINameSpace = Nothing
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


        

Private Sub Timer1_Timer()
'Call lstCPU_Click
Dim oCpu As SWbemObject
Set m_objCPUSet = m_objWMINameSpace.InstancesOf("Win32_Processor")
Set oCpu = m_objCPUSet("\\MATT-555ROIDS\root\cimv2:Win32_Processor.DeviceID=""CPU0""")
Label1 = Str(oCpu.LoadPercentage) + "% Cpu"
End Sub
