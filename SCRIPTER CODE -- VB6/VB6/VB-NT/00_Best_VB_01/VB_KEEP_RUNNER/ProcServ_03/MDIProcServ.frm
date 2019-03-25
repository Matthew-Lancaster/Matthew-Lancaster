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
   Visible         =   0   'False
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


' -------------------------------------------------
'MENU PULLER PROJECT AND THEN REFERENCE
'PROJECT REFERENCE
' -------------------------------------------------
'WMI Extension to DS 1.0 Type Library
'Location:
'D:\VB6\VB-NT\00_Best_VB_01\ProcServ_02\wbemads.tlb
' -------------------------------------------------
'THIS DRIVER IS FOUND THE WINDOWS XP MACHINE
'OR ONLINE SEARCH FOR wbemads.tlb
'IN THE \WINDOWS\SYSTEM32\ FOLDER
'\WINDOWS\system32\wbem\wbemads.tlb
'YOU PUT IN HERE WHEN FOUND
'\WINDOWS\SYSTEM32\
'AND NONE TO DO ALL WORK NORMAL
'OR PUT IN APP FOLDER ROOT MAYBE NOT TEST YET
' -------------------------------------------------
' USE THE BATCH COMMAND FILE TO COPY IT THERE FROM HOME FOLDER OF THIS CODE
' D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\WBEMADS.TLB.BAT
' -------------------------------------------------
'[ Sunday 21:40:00 Pm_17 March 2019 ]
' -------------------------------------------------

' THESE REFERANCE REQUIRING
' Reference=*\G{E503D000-5C7F-11D2-8B74-00104B2AFB41}#1.0#0#..\ProcServ_02\wbemads.tlb#WMI Extension to DS 1.0 Type Library
' Reference=*\G{565783C6-CB41-11D1-8B02-00600806D9B6}#1.1#0#C:\Windows\SysWOW64\wbem\wbemdisp.TLB#Microsoft WMI Scripting V1.1 Library


Public Locator As SWbemLocator
Public WBEMServices As SWbemServices
Public WithEvents eventSink As SWbemSink
Attribute eventSink.VB_VarHelpID = -1

'Global StrConnect As String
Public StrConnect As String


'Sub FORM_LOAD()
'
'Me.Hide
'
'End Sub


'Private Sub Form_Activate()
'Me.Hide
'End Sub
'Private Sub MDIForm_Activate()
'Me.Hide
'End Sub



Private Sub MDIForm_Load()
    
    Set Locator = New SWbemLocator
    Set eventSink = New SWbemSink
    
    ' disable menu items
'    Services.Enabled = False
'    Processes.Enabled = False
'    ComputerSystem.Enabled = False
'    LogicalDisk.Enabled = False
'    OperatingSystem.Enabled = False
'    Processor.Enabled = False
'    EventLog.Enabled = False

'    Me.Hide

    Call Connect_Click

    
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
'    frmOperatingSystem.Visible = False
'    frmProcesses.Visible = False
'    frmProcessor.Visible = False
'    frmServices.Visible = False
    
'    MDIProcServ.Caption = "ProcServ - building - this may take a minute or two ..."
    
     frmComputerSystem.ShowForm
'    frmEventLogFile.ShowForm
'    frmLogicalDisk.ShowForm
     frmOperatingSystem.ShowForm
'    frmProcesses.ShowForm
     frmProcessor.ShowForm
'    frmServices.ShowForm
    
'    With MDIProcServ
'        .Caption = "Done building information"
'        .Services.Enabled = True
'        .Processes.Enabled = True
'        .ComputerSystem.Enabled = True
'        .LogicalDisk.Enabled = True
'        .OperatingSystem.Enabled = True
'        .Processor.Enabled = True
'        .EventLog.Enabled = True
'    End With
    

End Sub

Private Sub ComputerSystem_Click()

    frmComputerSystem.ZOrder
    
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

    frmProcessor.ZOrder

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


