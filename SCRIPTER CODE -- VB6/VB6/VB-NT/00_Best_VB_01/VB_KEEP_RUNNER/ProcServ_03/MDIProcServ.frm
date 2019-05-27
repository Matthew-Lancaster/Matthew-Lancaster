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
Public Manufacturer_and_Model, TotalPhysicalMemory
Public TT_1VDT, TT_2VDT

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

    Dim Form
    On Error Resume Next
    Err.Clear
'    For Each Form In Forms
    If Form1.EXIT_TRUE = True Then
        If Err.Number = 0 Then
            Unload Me: Exit Sub
        End If
    End If
'    Next


    On Error Resume Next
    Err.Clear
    Me.Hide
    If Err.Number > 0 Then
        Exit Sub
    End If
    On Error GoTo 0
    
    
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

    Me.Hide

    Call Connect_Click

    
End Sub

Private Sub Connect_Click()

    StrConnect = ""
    MDIProcServ.ConnectWBEM (StrConnect)
    
    If Err.Number <> 0 Then
        Exit Sub
    End If
    
     Call ComputerSystem_GET_INFO
     Call OperatingSystem_GET_INFO
     Call Processor_GET_INFO
     
    
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




Public Sub ComputerSystem_GET_INFO()
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
    
    ' -------------------------------------------------------------
    ' CAN'T STOP THIS FORM FROM SHOW UP
    ' WHEN EVENT KICK OFF LINE ABOVE AND RETURN HERE
    ' SO CODE EVENT ABOVE MOVE SO SMALL AND OFF ABOVE TOP OF SCREEN
    ' -------------------------------------------------------------
    ' [ Sunday 22:34:50 Pm_17 March 2019 ]
    ' -------------------------------------------------------------
    ' HIDE WORKER HERE AFTER
    ' -------------------------------------------------------------
    On Error Resume Next
'    If Form1.VAR_FORM1_EXIST = True Then
'        If Err.Number = 0 Then VAR_FORM1_EXIST = Form1.VAR_FORM1_EXIST
'    End If
'    On Error GoTo 0
'    If VAR_FORM1_EXIST = True Then
'        ' Me.Hide
'        Me.Visible = False
'        MDIProcServ.Hide
'    End If
    
    For Each Object In Enumerator
'      With msfgComputerSystem
'
'        .ColWidth(0) = 3000
'        .ColWidth(1) = 2000
'
'        .Row = 1
'        .Col = 0
'        .Text = "AdminPasswordStatus"
'        .Col = 1
'        .Text = Object.AdminPasswordStatus
'
'        .Row = 2
'        .Col = 0
'        .Text = "AutomaticResetBootOption"
'        .Col = 1
'        .Text = Object.AutomaticResetBootOption
'
'        .Row = 3
'        .Col = 0
'        .Text = "AutomaticResetCapability"
'        .Col = 1
'        .Text = Object.AutomaticResetCapability
'
''        .Row = 4
''        .Col = 0
''        .Text = "BootOptionOnLimit"
''        .Col = 1
''        .Text = Object.BootOptionOnLimit
'
''        .Row = 5
''        .Col = 0
''        .Text = "BootOptionOnWatchDog"
''        .Col = 1
''        .Text = Object.BootOptionOnWatchDog
'
'        .Row = 6
'        .Col = 0
'        .Text = "BootROMSupported"
'        .Col = 1
'        .Text = Object.BootROMSupported
'
'        .Row = 7
'        .Col = 0
'        .Text = "BootupState"
'        .Col = 1
'        .Text = Object.BootupState
'
'        .Row = 8
'        .Col = 0
'        .Text = "Caption"
'        .Col = 1
'        .Text = Object.Caption
'
'        .Row = 9
'        .Col = 0
'        .Text = "ChassisBootupState"
'        .Col = 1
'        .Text = Object.ChassisBootupState
'
'        .Row = 10
'        .Col = 0
'        .Text = "CreationClassName"
'        .Col = 1
'        .Text = Object.CreationClassName
'
'        .Row = 11
'        .Col = 0
'        .Text = "CurrentTimeZone"
'        .Col = 1
'        .Text = Object.CurrentTimeZone
'
'        .Row = 12
'        .Col = 0
'        .Text = "DaylightInEffect"
'        .Col = 1
'        .Text = Object.DaylightInEffect
'
'        .Row = 13
'        .Col = 0
'        .Text = "Description"
'        .Col = 1
'        .Text = Object.Description
'
'        .Row = 14
'        .Col = 0
'        .Text = "Domain"
'        .Col = 1
'        .Text = Object.Domain
'
'        .Row = 15
'        .Col = 0
'        .Text = "DomainRole"
'        .Col = 1
'        .Text = Object.DomainRole
'
'        .Row = 16
'        .Col = 0
'        .Text = "FrontPanelResetStatus"
'        .Col = 1
'        .Text = Object.FrontPanelResetStatus
'
'        .Row = 17
'        .Col = 0
'        .Text = "InfraredSupported"
'        .Col = 1
'        .Text = Object.InfraredSupported
'
''        .Row = 18
''        .Col = 0
''        .Text = "InitialLoadInfo"
''        .Col = 1
''        .Text = Object.InitialLoadInfo
'
''        .Row = 19
''        .Col = 0
''        .Text = "InstallDate"
''        .Col = 1
''        .Text = Object.InstallDate
'
'        .Row = 20
'        .Col = 0
'        .Text = "KeyboardPasswordStatus"
'        .Col = 1
'        .Text = Object.KeyboardPasswordStatus
'
''        .Row = 21
''        .Col = 0
''        .Text = "LastLoadInfo"
''        .Col = 1
''        .Text = Object.LastLoadInfo
'
'        .Row = 22
'        .Col = 0
'        .Text = "Manufacturer"
'        .Col = 1
'        .Text = Object.Manufacturer
'
'        .Row = 23
'        .Col = 0
'        .Text = "Model"
'        .Col = 1
'        .Text = Object.Model
'
'        .Row = 24
'        .Col = 0
'        .Text = "Name"
'        .Col = 1
'        .Text = Object.Name
'
''        .Row = 25
''        .Col = 0
''        .Text = "NameFormat"
''        .Col = 1
''        .Text = Object.NameFormat
'
''        .Row = 26
''        .Col = 0
''        .Text = "NetworkModelServerEnabled"
''        .Col = 1
''        .Text = Object.NetworkModelServerEnabled
'
'        .Row = 27
'        .Col = 0
'        .Text = "NumberOfProcessors"
'        .Col = 1
'        .Text = Object.NumberOfProcessors
'
''        .Row = 28
''        .Col = 0
''        .Text = "OEMLogoBitMap"
''        .Col = 1
''        .Text = Object.OEMLogoBitMap
'
''        .Row = 29
''        .Col = 0
''        .Text = "OEMStringArray"
''        .Col = 1
''        .Text = Object.OEMStringArray
'
'        .Row = 30
'        .Col = 0
'        .Text = "PauseAfterReset"
'        .Col = 1
'        .Text = Object.PauseAfterReset
'
''        .Row = 31
''        .Col = 0
''        .Text = "PowerManagementCapabilities"
''        .Col = 1
''        .Text = Object.PowerManagementCapabilities
'
''        .Row = 32
''        .Col = 0
''        .Text = "PowerManagementSupported"
''        .Col = 1
''        .Text = Object.PowerManagementSupported
'
'        .Row = 33
'        .Col = 0
'        .Text = "PowerOnPasswordStatus"
'        .Col = 1
'        .Text = Object.PowerOnPasswordStatus
'
'        .Row = 34
'        .Col = 0
'        .Text = "PowerState"
'        .Col = 1
'        .Text = Object.PowerState
'
'        .Row = 35
'        .Col = 0
'        .Text = "PowerSupplyState"
'        .Col = 1
'        .Text = Object.PowerSupplyState
'
''        .Row = 36
''        .Col = 0
''        .Text = "PrimaryOwnerContact"
''        .Col = 1
''        .Text = Object.PrimaryOwnerContact
'
'        .Row = 37
'        .Col = 0
'        .Text = "PrimaryOwnerName"
'        .Col = 1
'        .Text = Object.PrimaryOwnerName
'
'        .Row = 38
'        .Col = 0
'        .Text = "ResetCapability"
'        .Col = 1
'        .Text = Object.ResetCapability
'
'        .Row = 39
'        .Col = 0
'        .Text = "ResetCount"
'        .Col = 1
'        .Text = Object.ResetCount
'
'        .Row = 40
'        .Col = 0
'        .Text = "ResetLimit"
'        .Col = 1
'        .Text = Object.ResetLimit
'
''        .Row = 41
''        .Col = 0
''        .Text = "Roles"
''        .Col = 1
''        .Text = Object.Roles
'
'        .Row = 42
'        .Col = 0
'        .Text = "Status"
'        .Col = 1
'        .Text = Object.Status
'
''        .Row = 43
''        .Col = 0
''        .Text = "SupportContactDescription"
''        .Col = 1
''        .Text = Object.SupportContactDescription
'
''        .Row = 44
''        .Col = 0
''        .Text = "SystemStartupDelay"
''        .Col = 1
''        .Text = Object.SystemStartupDelay
'
''        .Row = 45
''        .Col = 0
''        .Text = "SystemStartupOptions"
''        .Col = 1
''        .Text = Object.SystemStartupOptions
'
''        .Row = 46
''        .Col = 0
''        .Text = "SystemStartupSettings"
''        .Col = 1
''        .Text = Object.SystemStartupSettings
'
'        .Row = 47
'        .Col = 0
'        .Text = "SystemType"
'        .Col = 1
'        .Text = Object.SystemType
'
'        .Row = 48
'        .Col = 0
'        .Text = "ThermalState"
'        .Col = 1
'        .Text = Object.ThermalState
'
'        .Row = 49
'        .Col = 0
'        .Text = "TotalPhysicalMemory"
'        .Col = 1
'        .Text = Object.TotalPhysicalMemory
'
'        .Row = 50
'        .Col = 0
'        .Text = "UserName"
'        .Col = 1
'        .Text = Object.UserName
'
'        .Row = 51
'        .Col = 0
'        .Text = "WakeUpType"
'        .Col = 1
'        .Text = Object.WakeUpType
'
'      End With
      
      
'              For R = 1 To 43
'            Debug.Print Str(R) + " -- " + Item.SubItems(R)
'        Next
        
'        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "Manufacturer & Model")
'        Item_2.SubItems(1) = Object.Manufacturer + " _ " + Object.Model
'        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "TotalPhysicalMemory")
'        Item_2.SubItems(1) = Format(Object.TotalPhysicalMemory / 1024 ^ 3) + " Gigabyte"
        
        Manufacturer_and_Model = Object.Manufacturer + " _ " + Object.Model
        TotalPhysicalMemory = Format(Object.TotalPhysicalMemory / 1024 ^ 3) + " Gigabyte"
        TotalPhysicalMemory = Object.TotalPhysicalMemory
    Next
    
    ' Me.Show

End Sub

Public Sub OperatingSystem_GET_INFO()
Dim TT_1, TT_2

' Operating System

    Dim Enumerator As SWbemObjectSet
    Dim Object As SWbemObject
    Dim Item As ListItem

    On Error Resume Next
    
    MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_OperatingSystem")
    
    If Err.Number = 91 Then
        MsgBox "Please CONNECT to Local or Remote computer first", vbOKOnly, "Connect please!"
        Exit Sub
    End If
    
    ' -------------------------------------------------------------
    ' CAN'T STOP THIS FORM FROM SHOW UP
    ' WHEN EVENT KICK OFF LINE ABOVE AND RETURN HERE
    ' SO CODE EVENT ABOVE MOVE SO SMALL AND OFF ABOVE TOP OF SCREEN
    ' -------------------------------------------------------------
    ' [ Sunday 22:34:50 Pm_17 March 2019 ]
    ' -------------------------------------------------------------
    ' HIDE WORKER HERE AFTER
    ' -------------------------------------------------------------
'    On Error Resume Next
'    If Form1.VAR_FORM1_EXIST = True Then
'        If Err.Number = 0 Then VAR_FORM1_EXIST = Form1.VAR_FORM1_EXIST
'    End If
'    On Error GoTo 0
'    If VAR_FORM1_EXIST = True Then
'        Me.Hide
'        MDIProcServ.Hide
'    End If
    
    On Error Resume Next
    For Each Object In Enumerator
'        With msfgOperatingSystem
'
'          .ColWidth(0) = 3000
'          .ColWidth(1) = 3000
'
'          .Row = 1
'          .Col = 0
'          .Text = "BootDevice"
'          .Col = 1
'          .Text = Object.BootDevice
'
'          .Row = 2
'          .Col = 0
'          .Text = "BuildNumber"
'          .Col = 1
'          .Text = Object.BuildNumber
'
'          .Row = 3
'          .Col = 0
'          .Text = "BuildType"
'          .Col = 1
'          .Text = Object.BuildType
'
'          .Row = 4
'          .Col = 0
'          .Text = "Caption"
'          .Col = 1
'          .Text = Object.Caption
'
'          .Row = 5
'          .Col = 0
'          .Text = "CodeSet"
'          .Col = 1
'          .Text = Object.CodeSet
'
'          .Row = 6
'          .Col = 0
'          .Text = "CountryCode"
'          .Col = 1
'          .Text = Object.CountryCode
'
'          .Row = 7
'          .Col = 0
'          .Text = "CreationClassName"
'          .Col = 1
'          .Text = Object.CreationClassName
'
'          .Row = 8
'          .Col = 0
'          .Text = "CSCreationClassName"
'          .Col = 1
'          .Text = Object.CSCreationClassName
'
'        '  Run-time error '94':
'        '  Invalid use of Null
'' -------------------------------------------
''          .Row = 9
''          .Col = 0
''          .Text = "CSDVersion"
''          .Col = 1
''          .Text = Object.CSDVersion
'
'          .Row = 10
'          .Col = 0
'          .Text = "CSName"
'          .Col = 1
'          .Text = Object.CSName
'
'          .Row = 11
'          .Col = 0
'          .Text = "CurrentTimeZone"
'          .Col = 1
'          .Text = Object.CurrentTimeZone
'
'          .Row = 12
'          .Col = 0
'          .Text = "Debug"
'          .Col = 1
'          .Text = Object.Debug
'
'          .Row = 13
'          .Col = 0
'          .Text = "Description"
'          .Col = 1
'          .Text = Object.Description
'
'          .Row = 14
'          .Col = 0
'          .Text = "Distributed"
'          .Col = 1
'          .Text = Object.Distributed
'
'          .Row = 15
'          .Col = 0
'          .Text = "ForegroundApplicationBoost"
'          .Col = 1
'          .Text = Object.ForegroundApplicationBoost
'
'          .Row = 16
'          .Col = 0
'          .Text = "FreePhysicalMemory"
'          .Col = 1
'          .Text = Object.FreePhysicalMemory
'
'          .Row = 17
'          .Col = 0
'          .Text = "FreeSpaceInPagingFiles"
'          .Col = 1
'          .Text = Object.FreeSpaceInPagingFiles
'
'          .Row = 18
'          .Col = 0
'          .Text = "FreeVirtualMemory"
'          .Col = 1
'          .Text = Object.FreeVirtualMemory
'
'          .Row = 19
'          .Col = 0
'          .Text = "InstallDate"
'          .Col = 1
'          .Text = Object.InstallDate
'
'          .Row = 20
'          .Col = 0
'          .Text = "LastBootUpTime"
'          .Col = 1
'          .Text = Object.LastBootUpTime
'
'          .Row = 21
'          .Col = 0
'          .Text = "LocalDateTime"
'          .Col = 1
'          .Text = Object.LocalDateTime
'
'          .Row = 22
'          .Col = 0
'          .Text = "Locale"
'          .Col = 1
'          .Text = Object.Locale
'
'          .Row = 23
'          .Col = 0
'          .Text = "Manufacturer"
'          .Col = 1
'          .Text = Object.Manufacturer
'
'          .Row = 24
'          .Col = 0
'          .Text = "MaxNumberOfProcesses"
'          .Col = 1
'          .Text = Object.MaxNumberOfProcesses
'
'          .Row = 25
'          .Col = 0
'          .Text = "MaxProcessMemorySize"
'          .Col = 1
'          .Text = Object.MaxProcessMemorySize
'
'          .Row = 26
'          .Col = 0
'          .Text = "Name"
'          .Col = 1
'          .Text = Object.Name
'
'          .Row = 27
'          .Col = 0
'          .Text = "NumberOfLicensedUsers"
'          .Col = 1
'          .Text = Object.NumberOfLicensedUsers
'
'          .Row = 28
'          .Col = 0
'          .Text = "NumberOfProcesses"
'          .Col = 1
'          .Text = Object.NumberOfProcesses
'
'          .Row = 29
'          .Col = 0
'          .Text = "NumberOfUsers"
'          .Col = 1
'          .Text = Object.NumberOfUsers
'
'          .Row = 30
'          .Col = 0
'          .Text = "Organization"
'          .Col = 1
'          .Text = Object.Organization
'
'          .Row = 31
'          .Col = 0
'          .Text = "OSLanguage"
'          .Col = 1
'          .Text = Object.OSLanguage
'
'          .Row = 32
'          .Col = 0
'          .Text = "OSProductSuite"
'          .Col = 1
'          .Text = Object.OSProductSuite
'
'          .Row = 33
'          .Col = 0
'          .Text = "OSType"
'          .Col = 1
'          .Text = Object.OSType
'
'        '  Run-time error '94':
'        '  Invalid use of Null
'' -------------------------------------------
''          .Row = 34
''          .Col = 0
''          .Text = "OtherTypeDescription"
''          .Col = 1
''          .Text = Object.OtherTypeDescription
'
'        '  Run-time error '94':
'        '  Invalid use of Null
'' -------------------------------------------
''          .Row = 35
''          .Col = 0
''          .Text = "PlusProductID"
''          .Col = 1
''          .Text = Object.PlusProductID
'
'        '  Run-time error '94':
'        '  Invalid use of Null
'' -------------------------------------------
''          .Row = 36
''          .Col = 0
''          .Text = "PlusVersionNumber"
''          .Col = 1
''          .Text = Object.PlusVersionNumber
'
'          .Row = 37
'          .Col = 0
'          .Text = "Primary"
'          .Col = 1
'          .Text = Object.Primary
'
'        '  Run-time error '94':
'        '  Invalid use of Null
'' -------------------------------------------
''          .Row = 38
''          .Col = 0
''          .Text = "QuantumLength"
''          .Col = 1
''          .Text = Object.QuantumLength
'
'        '  Run-time error '94':
'        '  Invalid use of Null
'' -------------------------------------------
''          .Row = 39
''          .Col = 0
''          .Text = "QuantumType"
''          .Col = 1
''          .Text = Object.QuantumType
'
'          .Row = 40
'          .Col = 0
'          .Text = "RegisteredUser"
'          .Col = 1
'          .Text = Object.RegisteredUser
'
'          .Row = 41
'          .Col = 0
'          .Text = "SerialNumber"
'          .Col = 1
'          .Text = Object.SerialNumber
'
'          .Row = 42
'          .Col = 0
'          .Text = "ServicePackMajorVersion"
'          .Col = 1
'          .Text = Object.ServicePackMajorVersion
'
'          .Row = 43
'          .Col = 0
'          .Text = "ServicePackMinorVersion"
'          .Col = 1
'          .Text = Object.ServicePackMinorVersion
'
'          .Row = 44
'          .Col = 0
'          .Text = "SizeStoredInPagingFiles"
'          .Col = 1
'          .Text = Object.SizeStoredInPagingFiles
'
'          .Row = 45
'          .Col = 0
'          .Text = "Status"
'          .Col = 1
'          .Text = Object.Status
'
'          .Row = 46
'          .Col = 0
'          .Text = "SystemDevice"
'          .Col = 1
'          .Text = Object.SystemDevice
'
'          .Row = 47
'          .Col = 0
'          .Text = "SystemDirectory"
'          .Col = 1
'          .Text = Object.SystemDirectory
'
'        '  Run-time error '94':
'        '  Invalid use of Null
'' -------------------------------------------
''          .Row = 48
''          .Col = 0
''          .Text = "TotalSwapSpaceSize"
''          .Col = 1
''          .Text = Object.TotalSwapSpaceSize
'
'          .Row = 49
'          .Col = 0
'          .Text = "TotalVirtualMemorySize"
'          .Col = 1
'          .Text = Object.TotalVirtualMemorySize
'
'          .Row = 50
'          .Col = 0
'          .Text = "TotalVisibleMemorySize"
'          .Col = 1
'          .Text = Object.TotalVisibleMemorySize
'
'          .Row = 51
'          .Col = 0
'          .Text = "Version"
'          .Col = 1
'          .Text = Object.Version
'
'          .Row = 52
'          .Col = 0
'          .Text = "WindowsDirectory"
'          .Col = 1
'          .Text = Object.WindowsDirectory
'
'        End With
    
        TT_1 = Object.LastBootUpTime
        TT_2 = Object.InstallDate
    Next

    Dim TT_1DV
    Dim TT_1TV
    Dim TT_1DR
    Dim TT_2DV
    Dim TT_2TV
    Dim TT_2DR
    
    TT_1DV = Mid(TT_1, 1, 8)
    TT_1TV = Mid(TT_1, 9, 6)
    TT_1DV = Mid(TT_1DV, 1, 4) + "-" + Mid(TT_1DV, 5, 2) + "-" + Mid(TT_1DV, 7, 2)
    TT_1TV = Mid(TT_1TV, 1, 2) + ":" + Mid(TT_1TV, 3, 2) + ":" + Mid(TT_1TV, 5, 2)
    TT_1DR = Format(DateValue(TT_1DV) + TimeValue(TT_1TV), "DD-MMM-YYYY_HH:MM:SSam/pm")
    TT_1VDT = DateValue(TT_1DV) + TimeValue(TT_1TV)
    Form1.Text_SYSTEM_START_TIME_01.Text = "SYSTEM START " + TT_1DR
    Form1.Text_SYSTEM_START_TIME_01.FontSize = 10
    TT_2DV = Mid(TT_2, 1, 8)
    TT_2TV = Mid(TT_2, 9, 6)
    TT_2DV = Mid(TT_2DV, 1, 4) + "-" + Mid(TT_2DV, 5, 2) + "-" + Mid(TT_2DV, 7, 2)
    TT_2TV = Mid(TT_2TV, 1, 2) + ":" + Mid(TT_2TV, 3, 2) + ":" + Mid(TT_2TV, 5, 2)
    TT_2DR = Format(DateValue(TT_2DV) + TimeValue(TT_2TV), "DD-MMM-YYYY_HH:MM:SSam/pm")
    TT_2VDT = DateValue(TT_2DV) + TimeValue(TT_2TV)
    Form1.Text_OS_INSTALL_DATE_01.Text = "OS INSTALL        " + TT_2DR
    Form1.Text_OS_INSTALL_DATE_01.FontSize = 10

    Form1.Text_SYSTEM_START_TIME_02 = Str(DateDiff("h", TT_1VDT, Now)) + " HOUR"
    Form1.Text_SYSTEM_START_TIME_02.FontSize = 10
    Form1.Text_OS_INSTALL_DATE_02 = Str(DateDiff("d", TT_2VDT, Now)) + " DAY"
    Form1.Text_OS_INSTALL_DATE_02.FontSize = 10

    'Text_OS_INSTALL_DATE_01
    
End Sub

Public Sub Processor_GET_INFO()
' Processor

    Dim Enumerator As SWbemObjectSet
    Dim Object As SWbemObject
    Dim Item As ListItem
    Dim Item_2 As ListItem

    On Error Resume Next
    
    MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_Processor")
    
    If Err.Number = 91 Then
        MsgBox "Please CONNECT to Local or Remote computer first", vbOKOnly, "Connect please!"
        Exit Sub
    End If
    
    For Each Object In Enumerator
'        Set Item = lvwProcessor.ListItems.Add(, Object.DeviceID, Object.DeviceID)
'        With Item
'            .SubItems(1) = Object.AddressWidth
'            .SubItems(2) = Object.Architecture
'            .SubItems(3) = Object.Availability
'            .SubItems(4) = Object.Caption
''            .SubItems(5) = Object.ConfigManagerErrorCode
''            .SubItems(6) = Object.ConfigManagerUserConfig
'            .SubItems(7) = Object.CpuStatus
'            .SubItems(8) = Object.CreationClassName
'            .SubItems(9) = Object.CurrentClockSpeed
'            .SubItems(10) = Object.CurrentVoltage
'            .SubItems(11) = Object.DataWidth
'            .SubItems(12) = Object.Description
''            .SubItems(13) = Object.ErrorCleared
''            .SubItems(14) = Object.ErrorDescription
'            .SubItems(15) = Object.ExtClock
'            .SubItems(16) = Object.Family
''            .SubItems(17) = Object.InstallDate
'            .SubItems(18) = Object.L2CacheSize
''            .SubItems(19) = Object.L2CacheSpeed
''            .SubItems(20) = Object.LastErrorCode
'            .SubItems(21) = Object.Level
'            .SubItems(22) = Object.LoadPercentage
'            .SubItems(23) = Object.Manufacturer
'            .SubItems(24) = Object.MaxClockSpeed
'            .SubItems(25) = Object.Name
''            .SubItems(26) = Object.OtherFamilyDescription
''            .SubItems(27) = Object.PNPDeviceID
''            .SubItems(28) = Object.PowerManagementCapabilities
''            .SubItems(29) = Object.PowerManagementSuported
'            .SubItems(30) = Object.ProcessorID
'            .SubItems(31) = Object.ProcessorType
''            .SubItems(32) = Object.Revision
'            .SubItems(33) = Object.Role
'            .SubItems(34) = Object.SocketDesignation
'            .SubItems(35) = Object.Status
'            .SubItems(36) = Object.StatusInfo
''            .SubItems(37) = Object.Stepping
'            .SubItems(38) = Object.SystemCreationClassName
'            .SubItems(39) = Object.SystemName
''            .SubItems(40) = Object.UniqueID
'            .SubItems(41) = Object.UpgradeMethod
'            .SubItems(42) = Object.Version
''            .SubItems(43) = Object.VoltageCaps
'        End With
'        For R = 1 To 43
'            Debug.Print Str(R) + " -- " + Item.SubItems(R)
'        Next
        
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "Name")
        Item_2.SubItems(1) = Trim(Object.Name)
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "Description")
        Item_2.SubItems(1) = Object.Description
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "Manufacturer & Model")
        Item_2.SubItems(1) = Manufacturer_and_Model
        
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "Manufacturer")
        Item_2.SubItems(1) = Object.Manufacturer
        
        Dim A_MEM
        
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "TotalPhysicalMemory")
        A_MEM = Int(TotalPhysicalMemory / 1024 ^ 3) + 1
        Item_2.SubItems(1) = Trim(Str(A_MEM)) + " GB _ " + Format(TotalPhysicalMemory / 1024 ^ 3, "0.0000") + " GB _ " + Format(MDIProcServ.TotalPhysicalMemory) + " BYTE"
        
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "AddressWidth & L2CacheSize")
        Item_2.SubItems(1) = Trim(Str(Object.AddressWidth)) + " Bit ____ " + Trim(Str(Object.L2CacheSize)) + " Kbyte"
        
        
'        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "L2CacheSize")
'        Item_2.SubItems(1) = Trim(Str(Object.L2CacheSize))
        'Item_2.SubItems(1) = Trim(Str(Object.L2CacheSize)) + "    AddressWidth_" + Trim(Str(Object.AddressWidth)) + "   SocketDesignation_" + Object.SocketDesignation

        'Item_2.SubItems(1) = Trim(Str(Object.AddressWidth)) + "     L2CacheSize" + Trim(Str(Object.L2CacheSize)) + "         Level " + Trim(Str(Object.Level)) + "        SocketDesignation " + Object.SocketDesignation
'        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "Level")
'        Item_2.SubItems(1) = Object.Level
        
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "SocketDesignation")
        Item_2.SubItems(1) = Object.SocketDesignation
        
        'Item_2.SubItems(1) = Object.Manufacturer + "   Level_" + Trim(Str(Object.Level)) + "   ProcessorID_" + Object.ProcessorID
'        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "ProcessorID")
'        Item_2.SubItems(1) = Object.ProcessorID
        
        Dim TT_1
        
        TT_1 = ""
        TT_1 = TT_1 + "Description ___ " + Item.SubItems(12) + vbCrLf
        TT_1 = TT_1 + "Name _______ " + Item.SubItems(25) + vbCrLf
        TT_1 = TT_1 + "L2CacheSize _ " + Item.SubItems(18) + " __ "
        TT_1 = TT_1 + "AddressWidth _ " + Item.SubItems(1) + " __ "
        TT_1 = TT_1 + "Level _ " + Item.SubItems(21) + vbCrLf
        TT_1 = TT_1 + "Manufacturer __ " + Item.SubItems(23) + " __ "
        TT_1 = TT_1 + "SocketDesignation _ " + Item.SubItems(34) + vbCrLf
        TT_1 = TT_1 + "ProcessorID __ " + Item.SubItems(30)
    
    Next

End Sub

