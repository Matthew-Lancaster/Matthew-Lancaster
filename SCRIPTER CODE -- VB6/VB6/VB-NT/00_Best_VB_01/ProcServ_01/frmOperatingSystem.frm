VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmOperatingSystem 
   Caption         =   "Operating System"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   11175
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid msfgOperatingSystem 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   13785
      _Version        =   393216
      Rows            =   53
      FixedCols       =   0
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmOperatingSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ShowForm()
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
    
    For Each Object In Enumerator
      With msfgOperatingSystem
      
        .ColWidth(0) = 3000
        .ColWidth(1) = 3000
      
        .Row = 1
        .Col = 0
        .Text = "BootDevice"
        .Col = 1
        .Text = Object.BootDevice
        
        .Row = 2
        .Col = 0
        .Text = "BuildNumber"
        .Col = 1
        .Text = Object.BuildNumber
        
        .Row = 3
        .Col = 0
        .Text = "BuildType"
        .Col = 1
        .Text = Object.BuildType
        
        .Row = 4
        .Col = 0
        .Text = "Caption"
        .Col = 1
        .Text = Object.Caption
        
        .Row = 5
        .Col = 0
        .Text = "CodeSet"
        .Col = 1
        .Text = Object.CodeSet
        
        .Row = 6
        .Col = 0
        .Text = "CountryCode"
        .Col = 1
        .Text = Object.CountryCode
        
        .Row = 7
        .Col = 0
        .Text = "CreationClassName"
        .Col = 1
        .Text = Object.CreationClassName
        
        .Row = 8
        .Col = 0
        .Text = "CSCreationClassName"
        .Col = 1
        .Text = Object.CSCreationClassName
        
        .Row = 9
        .Col = 0
        .Text = "CSDVersion"
        .Col = 1
        .Text = Object.CSDVersion
        
        .Row = 10
        .Col = 0
        .Text = "CSName"
        .Col = 1
        .Text = Object.CSName
        
        .Row = 11
        .Col = 0
        .Text = "CurrentTimeZone"
        .Col = 1
        .Text = Object.CurrentTimeZone
        
        .Row = 12
        .Col = 0
        .Text = "Debug"
        .Col = 1
        .Text = Object.Debug
        
        .Row = 13
        .Col = 0
        .Text = "Description"
        .Col = 1
        .Text = Object.Description
        
        .Row = 14
        .Col = 0
        .Text = "Distributed"
        .Col = 1
        .Text = Object.Distributed
        
        .Row = 15
        .Col = 0
        .Text = "ForegroundApplicationBoost"
        .Col = 1
        .Text = Object.ForegroundApplicationBoost
        
        .Row = 16
        .Col = 0
        .Text = "FreePhysicalMemory"
        .Col = 1
        .Text = Object.FreePhysicalMemory
        
        .Row = 17
        .Col = 0
        .Text = "FreeSpaceInPagingFiles"
        .Col = 1
        .Text = Object.FreeSpaceInPagingFiles
        
        .Row = 18
        .Col = 0
        .Text = "FreeVirtualMemory"
        .Col = 1
        .Text = Object.FreeVirtualMemory
        
        .Row = 19
        .Col = 0
        .Text = "InstallDate"
        .Col = 1
        .Text = Object.InstallDate
        
        .Row = 20
        .Col = 0
        .Text = "LastBootUpTime"
        .Col = 1
        .Text = Object.LastBootUpTime
        
        .Row = 21
        .Col = 0
        .Text = "LocalDateTime"
        .Col = 1
        .Text = Object.LocalDateTime
        
        .Row = 22
        .Col = 0
        .Text = "Locale"
        .Col = 1
        .Text = Object.Locale
        
        .Row = 23
        .Col = 0
        .Text = "Manufacturer"
        .Col = 1
        .Text = Object.Manufacturer
        
        .Row = 24
        .Col = 0
        .Text = "MaxNumberOfProcesses"
        .Col = 1
        .Text = Object.MaxNumberOfProcesses
        
        .Row = 25
        .Col = 0
        .Text = "MaxProcessMemorySize"
        .Col = 1
        .Text = Object.MaxProcessMemorySize
        
        .Row = 26
        .Col = 0
        .Text = "Name"
        .Col = 1
        .Text = Object.Name
        
        .Row = 27
        .Col = 0
        .Text = "NumberOfLicensedUsers"
        .Col = 1
        .Text = Object.NumberOfLicensedUsers
        
        .Row = 28
        .Col = 0
        .Text = "NumberOfProcesses"
        .Col = 1
        .Text = Object.NumberOfProcesses
        
        .Row = 29
        .Col = 0
        .Text = "NumberOfUsers"
        .Col = 1
        .Text = Object.NumberOfUsers
        
        .Row = 30
        .Col = 0
        .Text = "Organization"
        .Col = 1
        .Text = Object.Organization
        
        .Row = 31
        .Col = 0
        .Text = "OSLanguage"
        .Col = 1
        .Text = Object.OSLanguage
        
        .Row = 32
        .Col = 0
        .Text = "OSProductSuite"
        .Col = 1
        .Text = Object.OSProductSuite
        
        .Row = 33
        .Col = 0
        .Text = "OSType"
        .Col = 1
        .Text = Object.OSType
        
        .Row = 34
        .Col = 0
        .Text = "OtherTypeDescription"
        .Col = 1
        .Text = Object.OtherTypeDescription
        
        .Row = 35
        .Col = 0
        .Text = "PlusProductID"
        .Col = 1
        .Text = Object.PlusProductID
        
        .Row = 36
        .Col = 0
        .Text = "PlusVersionNumber"
        .Col = 1
        .Text = Object.PlusVersionNumber
        
        .Row = 37
        .Col = 0
        .Text = "Primary"
        .Col = 1
        .Text = Object.Primary
        
        .Row = 38
        .Col = 0
        .Text = "QuantumLength"
        .Col = 1
        .Text = Object.QuantumLength
        
        .Row = 39
        .Col = 0
        .Text = "QuantumType"
        .Col = 1
        .Text = Object.QuantumType
        
        .Row = 40
        .Col = 0
        .Text = "RegisteredUser"
        .Col = 1
        .Text = Object.RegisteredUser
        
        .Row = 41
        .Col = 0
        .Text = "SerialNumber"
        .Col = 1
        .Text = Object.SerialNumber
        
        .Row = 42
        .Col = 0
        .Text = "ServicePackMajorVersion"
        .Col = 1
        .Text = Object.ServicePackMajorVersion
        
        .Row = 43
        .Col = 0
        .Text = "ServicePackMinorVersion"
        .Col = 1
        .Text = Object.ServicePackMinorVersion
        
        .Row = 44
        .Col = 0
        .Text = "SizeStoredInPagingFiles"
        .Col = 1
        .Text = Object.SizeStoredInPagingFiles
        
        .Row = 45
        .Col = 0
        .Text = "Status"
        .Col = 1
        .Text = Object.Status
        
        .Row = 46
        .Col = 0
        .Text = "SystemDevice"
        .Col = 1
        .Text = Object.SystemDevice
        
        .Row = 47
        .Col = 0
        .Text = "SystemDirectory"
        .Col = 1
        .Text = Object.SystemDirectory
        
        .Row = 48
        .Col = 0
        .Text = "TotalSwapSpaceSize"
        .Col = 1
        .Text = Object.TotalSwapSpaceSize
        
        .Row = 49
        .Col = 0
        .Text = "TotalVirtualMemorySize"
        .Col = 1
        .Text = Object.TotalVirtualMemorySize
        
        .Row = 50
        .Col = 0
        .Text = "TotalVisibleMemorySize"
        .Col = 1
        .Text = Object.TotalVisibleMemorySize
        
        .Row = 51
        .Col = 0
        .Text = "Version"
        .Col = 1
        .Text = Object.Version
        
        .Row = 52
        .Col = 0
        .Text = "WindowsDirectory"
        .Col = 1
        .Text = Object.WindowsDirectory
        
      End With
    Next
    
    Me.Show

End Sub

Private Sub Form_Load()

    frmOperatingSystem.Height = MDIProcServ.Height - 200
    frmOperatingSystem.Width = MDIProcServ.Width - 275

End Sub

