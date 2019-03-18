VERSION 5.00
Begin VB.MDIForm MDIProcServ 
   BackColor       =   &H8000000C&
   Caption         =   "MDIProcServ"
   ClientHeight    =   8388
   ClientLeft      =   132
   ClientTop       =   780
   ClientWidth     =   11172
   LinkTopic       =   "MDIProcServ"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Exit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu Connect 
      Caption         =   "&Connect"
   End
   Begin VB.Menu Services 
      Caption         =   "&Services"
   End
   Begin VB.Menu Processes 
      Caption         =   "&Processes"
   End
   Begin VB.Menu ComputerSystem 
      Caption         =   "Comp&uter System"
   End
   Begin VB.Menu LogicalDisk 
      Caption         =   "&Logical Disk"
   End
   Begin VB.Menu OperatingSystem 
      Caption         =   "&Operating System"
   End
   Begin VB.Menu Processor 
      Caption         =   "P&rocessor"
   End
   Begin VB.Menu EventLog 
      Caption         =   "&EventLog"
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIProcServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Locator As SWbemLocator
Public WBEMServices As SWbemServices
Public WithEvents eventSink As SWbemSink
Attribute eventSink.VB_VarHelpID = -1

' Global StrConnect As String
Public StrConnect As String

'Public TimerCount

Private Sub About_Click()

'    Load frmAbout
'    frmAbout.Show

End Sub

Private Sub MDIForm_Load()
    
    Set Locator = New SWbemLocator
    Set eventSink = New SWbemSink
    
    ' disable menu items
'    Services.Enabled = False
'    Processes.Enabled = False
'    ComputerSystem.Enabled = False
'    LogicalDisk.Enabled = False
    OperatingSystem.Enabled = False
'    Processor.Enabled = False
'    EventLog.Enabled = False

    Connect_Click
    
End Sub

Private Sub Connect_Click()

'    Load dlgConnect
'    dlgConnect.Show
    
    ' StrConnect = txtConnect.Text
    StrConnect = ""
    MDIProcServ.ConnectWBEM (StrConnect)
    
    If Err.Number <> 0 Then
        Exit Sub
    End If
    
'    frmComputerSystem.Visible = False
'    frmEventLogFile.Visible = False
'    frmLogicalDisk.Visible = False
    frmOperatingSystem.Visible = False
'    frmProcesses.Visible = False
'    frmProcessor.Visible = False
'    frmServices.Visible = False
    
    MDIProcServ.Caption = "ProcServ - building - this may take a minute or two ..."
    
'    frmComputerSystem.ShowForm
'    frmEventLogFile.ShowForm
'    frmLogicalDisk.ShowForm
    frmOperatingSystem.ShowForm
'    frmProcesses.ShowForm
'    frmProcessor.ShowForm
'    frmServices.ShowForm
    
    With MDIProcServ
        .Caption = "Done building information"
'        .Services.Enabled = True
'        .Processes.Enabled = True
'        .ComputerSystem.Enabled = True
'        .LogicalDisk.Enabled = True
        .OperatingSystem.Enabled = True
'        .Processor.Enabled = True
'        .EventLog.Enabled = True
    End With
    

End Sub

Private Sub ComputerSystem_Click()

'    frmComputerSystem.ZOrder
    
End Sub

Private Sub EventLog_Click()

'    frmEventLogFile.ZOrder

End Sub

Private Sub LogicalDisk_Click()

'    frmLogicalDisk.ZOrder

End Sub

Private Sub OperatingSystem_Click()

    frmOperatingSystem.ZOrder

End Sub

Private Sub Processes_Click()

'    frmProcesses.ZOrder

End Sub

Private Sub Processor_Click()

'    frmProcessor.ZOrder

End Sub

Private Sub Services_Click()

'    frmServices.ZOrder

End Sub

Private Sub Exit_Click()

'    Finish

End Sub

Public Sub ConnectWBEM(strComputername As String)

    On Error Resume Next

    eventSink.Cancel
   
    Set WBEMServices = Locator.ConnectServer(strComputername)
    If Err.Number <> 0 Then
        MsgBox Err.Description + Err.Source + vbCrLf _
               + "Either:" + vbCrLf _
               + "    1. You are not authorized to access this machine" + vbCrLf _
               + " or 2. This is not a Windows 2000 machine" + vbCrLf _
               + " or 3. DCOM is not setup on this Windows 2000 machine"
        Exit Sub
    End If
        
End Sub


