VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmComputerSystem 
   Caption         =   "Computer System"
   ClientHeight    =   7905
   ClientLeft      =   -285
   ClientTop       =   15
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   11175
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid msfgComputerSystem 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   13785
      _Version        =   393216
      Rows            =   52
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmComputerSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ShowForm()
' ComputerSystem

    Dim Enumerator As SWbemObjectSet
    Dim Object As SWbemObject
    Dim Item As ListItem

    On Error Resume Next
    
    MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_ComputerSystem")
    
    If Err.Number = 91 Then
        MsgBox "Please CONNECT to Local or Remote computer first", vbOKOnly, "Connect please!"
        Exit Sub
    End If
    
    For Each Object In Enumerator
      With msfgComputerSystem
      
        .ColWidth(0) = 3000
        .ColWidth(1) = 2000
      
        .Row = 1
        .Col = 0
        .Text = "AdminPasswordStatus"
        .Col = 1
        .Text = Object.AdminPasswordStatus
        
        .Row = 2
        .Col = 0
        .Text = "AutomaticResetBootOption"
        .Col = 1
        .Text = Object.AutomaticResetBootOption
        
        .Row = 3
        .Col = 0
        .Text = "AutomaticResetCapability"
        .Col = 1
        .Text = Object.AutomaticResetCapability
        
        .Row = 4
        .Col = 0
        .Text = "BootOptionOnLimit"
        .Col = 1
        .Text = Object.BootOptionOnLimit
        
        .Row = 5
        .Col = 0
        .Text = "BootOptionOnWatchDog"
        .Col = 1
        .Text = Object.BootOptionOnWatchDog
        
        .Row = 6
        .Col = 0
        .Text = "BootROMSupported"
        .Col = 1
        .Text = Object.BootROMSupported
        
        .Row = 7
        .Col = 0
        .Text = "BootupState"
        .Col = 1
        .Text = Object.BootupState
        
        .Row = 8
        .Col = 0
        .Text = "Caption"
        .Col = 1
        .Text = Object.Caption
        
        .Row = 9
        .Col = 0
        .Text = "ChassisBootupState"
        .Col = 1
        .Text = Object.ChassisBootupState
        
        .Row = 10
        .Col = 0
        .Text = "CreationClassName"
        .Col = 1
        .Text = Object.CreationClassName
        
        .Row = 11
        .Col = 0
        .Text = "CurrentTimeZone"
        .Col = 1
        .Text = Object.CurrentTimeZone
        
        .Row = 12
        .Col = 0
        .Text = "DaylightInEffect"
        .Col = 1
        .Text = Object.DaylightInEffect
        
        .Row = 13
        .Col = 0
        .Text = "Description"
        .Col = 1
        .Text = Object.Description
        
        .Row = 14
        .Col = 0
        .Text = "Domain"
        .Col = 1
        .Text = Object.Domain
        
        .Row = 15
        .Col = 0
        .Text = "DomainRole"
        .Col = 1
        .Text = Object.DomainRole
        
        .Row = 16
        .Col = 0
        .Text = "FrontPanelResetStatus"
        .Col = 1
        .Text = Object.FrontPanelResetStatus
        
        .Row = 17
        .Col = 0
        .Text = "InfraredSupported"
        .Col = 1
        .Text = Object.InfraredSupported
        
        .Row = 18
        .Col = 0
        .Text = "InitialLoadInfo"
        .Col = 1
        .Text = Object.InitialLoadInfo
        
        .Row = 19
        .Col = 0
        .Text = "InstallDate"
        .Col = 1
        .Text = Object.InstallDate
        
        .Row = 20
        .Col = 0
        .Text = "KeyboardPasswordStatus"
        .Col = 1
        .Text = Object.KeyboardPasswordStatus
        
        .Row = 21
        .Col = 0
        .Text = "LastLoadInfo"
        .Col = 1
        .Text = Object.LastLoadInfo
        
        .Row = 22
        .Col = 0
        .Text = "Manufacturer"
        .Col = 1
        .Text = Object.Manufacturer
        
        .Row = 23
        .Col = 0
        .Text = "Model"
        .Col = 1
        .Text = Object.Model
        
        .Row = 24
        .Col = 0
        .Text = "Name"
        .Col = 1
        .Text = Object.Name
        
        .Row = 25
        .Col = 0
        .Text = "NameFormat"
        .Col = 1
        .Text = Object.NameFormat
        
        .Row = 26
        .Col = 0
        .Text = "NetworkModelServerEnabled"
        .Col = 1
        .Text = Object.NetworkModelServerEnabled
        
        .Row = 27
        .Col = 0
        .Text = "NumberOfProcessors"
        .Col = 1
        .Text = Object.NumberOfProcessors
        
        .Row = 28
        .Col = 0
        .Text = "OEMLogoBitMap"
        .Col = 1
        .Text = Object.OEMLogoBitMap
        
        .Row = 29
        .Col = 0
        .Text = "OEMStringArray"
        .Col = 1
        .Text = Object.OEMStringArray
        
        .Row = 30
        .Col = 0
        .Text = "PauseAfterReset"
        .Col = 1
        .Text = Object.PauseAfterReset
        
        .Row = 31
        .Col = 0
        .Text = "PowerManagementCapabilities"
        .Col = 1
        .Text = Object.PowerManagementCapabilities
        
        .Row = 32
        .Col = 0
        .Text = "PowerManagementSupported"
        .Col = 1
        .Text = Object.PowerManagementSupported
        
        .Row = 33
        .Col = 0
        .Text = "PowerOnPasswordStatus"
        .Col = 1
        .Text = Object.PowerOnPasswordStatus
        
        .Row = 34
        .Col = 0
        .Text = "PowerState"
        .Col = 1
        .Text = Object.PowerState
        
        .Row = 35
        .Col = 0
        .Text = "PowerSupplyState"
        .Col = 1
        .Text = Object.PowerSupplyState
        
        .Row = 36
        .Col = 0
        .Text = "PrimaryOwnerContact"
        .Col = 1
        .Text = Object.PrimaryOwnerContact
        
        .Row = 37
        .Col = 0
        .Text = "PrimaryOwnerName"
        .Col = 1
        .Text = Object.PrimaryOwnerName
        
        .Row = 38
        .Col = 0
        .Text = "ResetCapability"
        .Col = 1
        .Text = Object.ResetCapability
        
        .Row = 39
        .Col = 0
        .Text = "ResetCount"
        .Col = 1
        .Text = Object.ResetCount
        
        .Row = 40
        .Col = 0
        .Text = "ResetLimit"
        .Col = 1
        .Text = Object.ResetLimit
        
        .Row = 41
        .Col = 0
        .Text = "Roles"
        .Col = 1
        .Text = Object.Roles
        
        .Row = 42
        .Col = 0
        .Text = "Status"
        .Col = 1
        .Text = Object.Status
        
        .Row = 43
        .Col = 0
        .Text = "SupportContactDescription"
        .Col = 1
        .Text = Object.SupportContactDescription
        
        .Row = 44
        .Col = 0
        .Text = "SystemStartupDelay"
        .Col = 1
        .Text = Object.SystemStartupDelay
        
        .Row = 45
        .Col = 0
        .Text = "SystemStartupOptions"
        .Col = 1
        .Text = Object.SystemStartupOptions
        
        .Row = 46
        .Col = 0
        .Text = "SystemStartupSettings"
        .Col = 1
        .Text = Object.SystemStartupSettings
        
        .Row = 47
        .Col = 0
        .Text = "SystemType"
        .Col = 1
        .Text = Object.SystemType
        
        .Row = 48
        .Col = 0
        .Text = "ThermalState"
        .Col = 1
        .Text = Object.ThermalState
        
        .Row = 49
        .Col = 0
        .Text = "TotalPhysicalMemory"
        .Col = 1
        .Text = Object.TotalPhysicalMemory
        
        .Row = 50
        .Col = 0
        .Text = "UserName"
        .Col = 1
        .Text = Object.UserName
        
        .Row = 51
        .Col = 0
        .Text = "WakeUpType"
        .Col = 1
        .Text = Object.WakeUpType
        
      End With
    Next
    
    Me.Show

End Sub

Private Sub Form_Load()

    frmComputerSystem.Height = MDIProcServ.Height - 200
    frmComputerSystem.Width = MDIProcServ.Width - 275

End Sub


