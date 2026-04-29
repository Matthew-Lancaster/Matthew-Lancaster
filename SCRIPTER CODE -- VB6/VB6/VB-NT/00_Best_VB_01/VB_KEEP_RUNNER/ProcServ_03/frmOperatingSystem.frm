VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmOperatingSystem 
   Caption         =   "Operating System"
   ClientHeight    =   7908
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11172
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7908
   ScaleWidth      =   11172
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid msfgOperatingSystem 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19283
      _ExtentY        =   13780
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
Dim VAR_FORM1_EXIST
Public TT_1VDT, TT_2VDT
'



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
    If Form1.VAR_FORM1_EXIST = True Then
        If Err.Number = 0 Then VAR_FORM1_EXIST = Form1.VAR_FORM1_EXIST
    End If
    On Error GoTo 0
    If VAR_FORM1_EXIST = True Then
        Me.Hide
        MDIProcServ.Hide
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
          
        '  Run-time error '94':
        '  Invalid use of Null
' -------------------------------------------
'          .Row = 9
'          .Col = 0
'          .Text = "CSDVersion"
'          .Col = 1
'          .Text = Object.CSDVersion
          
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
          
        '  Run-time error '94':
        '  Invalid use of Null
' -------------------------------------------
'          .Row = 34
'          .Col = 0
'          .Text = "OtherTypeDescription"
'          .Col = 1
'          .Text = Object.OtherTypeDescription
          
        '  Run-time error '94':
        '  Invalid use of Null
' -------------------------------------------
'          .Row = 35
'          .Col = 0
'          .Text = "PlusProductID"
'          .Col = 1
'          .Text = Object.PlusProductID
          
        '  Run-time error '94':
        '  Invalid use of Null
' -------------------------------------------
'          .Row = 36
'          .Col = 0
'          .Text = "PlusVersionNumber"
'          .Col = 1
'          .Text = Object.PlusVersionNumber
          
          .Row = 37
          .Col = 0
          .Text = "Primary"
          .Col = 1
          .Text = Object.Primary
        
        '  Run-time error '94':
        '  Invalid use of Null
' -------------------------------------------
'          .Row = 38
'          .Col = 0
'          .Text = "QuantumLength"
'          .Col = 1
'          .Text = Object.QuantumLength
          
        '  Run-time error '94':
        '  Invalid use of Null
' -------------------------------------------
'          .Row = 39
'          .Col = 0
'          .Text = "QuantumType"
'          .Col = 1
'          .Text = Object.QuantumType
          
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
        
        '  Run-time error '94':
        '  Invalid use of Null
' -------------------------------------------
'          .Row = 48
'          .Col = 0
'          .Text = "TotalSwapSpaceSize"
'          .Col = 1
'          .Text = Object.TotalSwapSpaceSize
          
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
    
        TT_1 = Object.LastBootUpTime
        TT_2 = Object.InstallDate
    Next

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

Private Sub Form_Load()
    
    frmOperatingSystem.height = MDIProcServ.height - 200
    frmOperatingSystem.width = MDIProcServ.width - 275
    On Error Resume Next
        If Form1.VAR_FORM1_EXIST = True Then
            If Err.Number = 0 Then VAR_FORM1_EXIST = Form1.VAR_FORM1_EXIST
        End If
    On Error GoTo 0
    
    If VAR_FORM1_EXIST = True Then
        MDIProcServ.height = 0
        MDIProcServ.width = 0
        MDIProcServ.Top = -1000
    End If

End Sub

