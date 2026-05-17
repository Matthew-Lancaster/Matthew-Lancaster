Attribute VB_Name = "SPECIAL_FOLDER_Mod"

'24-FEB-2010
'THIS IS THE FIRST COMMONS SHARED MODULE







'Public FS
Public sSystemFolder As String, DOMAINGET
Public sTempFolder As String
Public sWindowsFolder As String
Public sMyDocsFolder



'----------------------------------------------------
'I PUT THIS HERE FOR THE SHELL EXECUTE LOADER PROGRAM
'----------------------------------------------------
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'Function GetUserName() As String
'   Dim UserName As String * 255
'   Call GetUserNameA(UserName, 255)
'   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
'End Function
'
'Function GetComputerName() As String
'   Dim UserName As String * 255
'   Call GetComputerNameA(UserName, 255)
'   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
'End Function
'----------------------------------------------------
'----------------------------------------------------



Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long





'### USAGE
'sWindowsFolder = GetSpecialfolder(36)
'sMyDocsFolder = GetSpecialfolder(5)

'Dim R As Long
'On Error Resume Next
'For R = 0 To 120
'    Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
'Next
'End


Public Enum SpecialFolderConstants
    spfldDesktop = &H0
    spfldInternet = &H1
    spfldPrograms = &H2
    spfldControls = &H3
    spfldPrinters = &H4
    spfldPersonal = &H5
    spfldFavorites = &H6
    spfldStartUp = &H7
    spfldRecent = &H8
    spfldSendTo = &H9
    spfldBitBucket = &HA
    spfldStartMenu = &HB
    spfldDesktopDirectory = &H10
    spfldDrives = &H11
    spfldNetwork = &H12
    spfldNethood = &H13
    spfldFonts = &H14
    spfldTemplates = &H15
    spfldCommonStartMenu = &H16
    spfldCommonPrograms = &H17
    spfldCommonStartup = &H18
    spfldCommonDesktopDirectory = &H19
    spfldAppData = &H1A
    spfldPrintHood = &H1B
    spfldAltStartup = &H1D                          '' DBCS
    spfldCommonAltStartup = &H1E                   '' DBCS
    spfldCommonFavorites = &H1F
    spfldInternetCache = &H20
    spfldCookies = &H21
    spfldHistory = &H22
    spfldWindows = &H24
    spfldWindowSystem = &H25
    
    '&H27& My Pictures
    '40 User Profile
    '43 Common Files
    '46 All Users Templates
    '47 Administrative Tools
    '49 Network Connections
    
End Enum



Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type



'Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
'Public Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
'Public Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
'Public Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
'Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long


Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long



'----------------------------------------------------
'I PUT THIS HERE FOR THE SHELL EXECUTE LOADER PROGRAM
'----------------------------------------------------
'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
'----------------------------------------------------
'----------------------------------------------------



'
'
'Sub Operating_System()
'' Operating System
'
'    Dim Enumerator As SWbemObjectSet
'    Dim Object As SWbemObject
'    Dim Item As ListItem
'
'    On Error Resume Next
'
'
'
'
'    MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
'    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_Processor")
'
'    If Err.Number = 91 Then
'        MsgBox "Please CONNECT to Local or Remote computer first", vbOKOnly, "Connect please!"
'        Exit Sub
'    End If
'
'    For Each Object In Enumerator
'        With msfgProcessor
'        cpudes = Object.Name
'        End With
'    Next
'
'    t1 = InStr(cpudes, " ")
'    t1 = InStr(t1 + 1, cpudes, " ")
'    t1 = InStr(t1 + 1, cpudes, " ")
'
'    cpudes = Mid$(cpudes, 1, t1) + Mid$(cpudes, InStrRev(cpudes, " ") + 1)
'
'
'    MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
'    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_OperatingSystem")
'
'    If Err.Number = 91 Then
'        MsgBox "Please CONNECT to Local or Remote computer first", vbOKOnly, "Connect please!"
'        Exit Sub
'    End If
'
'
'
'
''
''
''    For Each Object In Enumerator
''      With msfgOperatingSystem
''
''
''        TTz1 = Object.LastBootUpTime
''        TTz2 = Object.InstallDate
''        TTz2 = Mid$(TTz2, 1, 4) + "-" + Mid$(TTz2, 5, 2) + "-" + Mid$(TTz2, 7, 2) + " " + Mid$(TTz2, 9, 2) + ":" + Mid$(TTz2, 11, 2) + ":" + Mid$(TTz2, 13, 2)
''        TTz2 = DateValue(TTz2) + TimeValue(TTz2)
''        TTz3 = Mid$(Object.Name, 1, InStr(Object.Name, "|") - 1) + " - Ver. " + Object.Version + " - SP" + Mid$(Object.CSDVersion, Len(Object.CSDVersion))
''        'TTz3 = "Windows XP Pro" + " - Ver. " + Object.Version + " - SP" + Mid$(Object.CSDVersion, Len(Object.CSDVersion))
''        TTz3 = "MS Win XP Pro" + " SP" + Mid$(Object.CSDVersion, Len(Object.CSDVersion))
''
''
'''        msgbox TTz2
''
''        Exit For
''
''        .ColWidth(0) = 3000
''        .ColWidth(1) = 3000
''
''        .Row = 1
''        .Col = 0
''        .Text = "BootDevice"
''        .Col = 1
''        .Text = Object.BootDevice
''
''        .Row = 2
''        .Col = 0
''        .Text = "BuildNumber"
''        .Col = 1
''        .Text = Object.BuildNumber
''
''        .Row = 3
''        .Col = 0
''        .Text = "BuildType"
''        .Col = 1
''        .Text = Object.BuildType
''
''        .Row = 4
''        .Col = 0
''        .Text = "Caption"
''        .Col = 1
''        .Text = Object.Caption
''
''        .Row = 5
''        .Col = 0
''        .Text = "CodeSet"
''        .Col = 1
''        .Text = Object.CodeSet
''
''        .Row = 6
''        .Col = 0
''        .Text = "CountryCode"
''        .Col = 1
''        .Text = Object.CountryCode
''
''        .Row = 7
''        .Col = 0
''        .Text = "CreationClassName"
''        .Col = 1
''        .Text = Object.CreationClassName
''
''        .Row = 8
''        .Col = 0
''        .Text = "CSCreationClassName"
''        .Col = 1
''        .Text = Object.CSCreationClassName
''
''        .Row = 9
''        .Col = 0
''        .Text = "CSDVersion"
''        .Col = 1
''        .Text = Object.CSDVersion
''
''        .Row = 10
''        .Col = 0
''        .Text = "CSName"
''        .Col = 1
''        .Text = Object.CSName
''
''        .Row = 11
''        .Col = 0
''        .Text = "CurrentTimeZone"
''        .Col = 1
''        .Text = Object.CurrentTimeZone
''
''        .Row = 12
''        .Col = 0
''        .Text = "Debug"
''        .Col = 1
''        .Text = Object.Debug
''
''        .Row = 13
''        .Col = 0
''        .Text = "Description"
''        .Col = 1
''        .Text = Object.Description
''
''        .Row = 14
''        .Col = 0
''        .Text = "Distributed"
''        .Col = 1
''        .Text = Object.Distributed
''
''        .Row = 15
''        .Col = 0
''        .Text = "ForegroundApplicationBoost"
''        .Col = 1
''        .Text = Object.ForegroundApplicationBoost
''
''        .Row = 16
''        .Col = 0
''        .Text = "FreePhysicalMemory"
''        .Col = 1
''        .Text = Object.FreePhysicalMemory
''
''        .Row = 17
''        .Col = 0
''        .Text = "FreeSpaceInPagingFiles"
''        .Col = 1
''        .Text = Object.FreeSpaceInPagingFiles
''
''        .Row = 18
''        .Col = 0
''        .Text = "FreeVirtualMemory"
''        .Col = 1
''        .Text = Object.FreeVirtualMemory
''
''        .Row = 19
''        .Col = 0
''        .Text = "InstallDate"
''        .Col = 1
''        .Text = Object.InstallDate
''
''        .Row = 20
''        .Col = 0
''        .Text = "LastBootUpTime"
''        .Col = 1
''        .Text = Object.LastBootUpTime
''
''        .Row = 21
''        .Col = 0
''        .Text = "LocalDateTime"
''        .Col = 1
''        .Text = Object.LocalDateTime
''
''        .Row = 22
''        .Col = 0
''        .Text = "Locale"
''        .Col = 1
''        .Text = Object.Locale
''
''        .Row = 23
''        .Col = 0
''        .Text = "Manufacturer"
''        .Col = 1
''        .Text = Object.Manufacturer
''
''        .Row = 24
''        .Col = 0
''        .Text = "MaxNumberOfProcesses"
''        .Col = 1
''        .Text = Object.MaxNumberOfProcesses
''
''        .Row = 25
''        .Col = 0
''        .Text = "MaxProcessMemorySize"
''        .Col = 1
''        .Text = Object.MaxProcessMemorySize
''
''        .Row = 26
''        .Col = 0
''        .Text = "Name"
''        .Col = 1
''        .Text = Object.Name
''
''        .Row = 27
''        .Col = 0
''        .Text = "NumberOfLicensedUsers"
''        .Col = 1
''        .Text = Object.NumberOfLicensedUsers
''
''        .Row = 28
''        .Col = 0
''        .Text = "NumberOfProcesses"
''        .Col = 1
''        .Text = Object.NumberOfProcesses
''
''        .Row = 29
''        .Col = 0
''        .Text = "NumberOfUsers"
''        .Col = 1
''        .Text = Object.NumberOfUsers
''
''        .Row = 30
''        .Col = 0
''        .Text = "Organization"
''        .Col = 1
''        .Text = Object.Organization
''
''        .Row = 31
''        .Col = 0
''        .Text = "OSLanguage"
''        .Col = 1
''        .Text = Object.OSLanguage
''
''        .Row = 32
''        .Col = 0
''        .Text = "OSProductSuite"
''        .Col = 1
''        .Text = Object.OSProductSuite
''
''        .Row = 33
''        .Col = 0
''        .Text = "OSType"
''        .Col = 1
''        .Text = Object.OSType
''
''        .Row = 34
''        .Col = 0
''        .Text = "OtherTypeDescription"
''        .Col = 1
''        .Text = Object.OtherTypeDescription
''
''        .Row = 35
''        .Col = 0
''        .Text = "PlusProductID"
''        .Col = 1
''        .Text = Object.PlusProductID
''
''        .Row = 36
''        .Col = 0
''        .Text = "PlusVersionNumber"
''        .Col = 1
''        .Text = Object.PlusVersionNumber
''
''        .Row = 37
''        .Col = 0
''        .Text = "Primary"
''        .Col = 1
''        .Text = Object.Primary
''
''        .Row = 38
''        .Col = 0
''        .Text = "QuantumLength"
''        .Col = 1
''        .Text = Object.QuantumLength
''
''        .Row = 39
''        .Col = 0
''        .Text = "QuantumType"
''        .Col = 1
''        .Text = Object.QuantumType
''
''        .Row = 40
''        .Col = 0
''        .Text = "RegisteredUser"
''        .Col = 1
''        .Text = Object.RegisteredUser
''
''        .Row = 41
''        .Col = 0
''        .Text = "SerialNumber"
''        .Col = 1
''        .Text = Object.serialnumber
''
''        .Row = 42
''        .Col = 0
''        .Text = "ServicePackMajorVersion"
''        .Col = 1
''        .Text = Object.ServicePackMajorVersion
''
''        .Row = 43
''        .Col = 0
''        .Text = "ServicePackMinorVersion"
''        .Col = 1
''        .Text = Object.ServicePackMinorVersion
''
''        .Row = 44
''        .Col = 0
''        .Text = "SizeStoredInPagingFiles"
''        .Col = 1
''        .Text = Object.SizeStoredInPagingFiles
''
''        .Row = 45
''        .Col = 0
''        .Text = "Status"
''        .Col = 1
''        .Text = Object.Status
''
''        .Row = 46
''        .Col = 0
''        .Text = "SystemDevice"
''        .Col = 1
''        .Text = Object.SystemDevice
''
''        .Row = 47
''        .Col = 0
''        .Text = "SystemDirectory"
''        .Col = 1
''        .Text = Object.SystemDirectory
''
''        .Row = 48
''        .Col = 0
''        .Text = "TotalSwapSpaceSize"
''        .Col = 1
''        .Text = Object.TotalSwapSpaceSize
''
''        .Row = 49
''        .Col = 0
''        .Text = "TotalVirtualMemorySize"
''        .Col = 1
''        .Text = Object.TotalVirtualMemorySize
''
''        .Row = 50
''        .Col = 0
''        .Text = "TotalVisibleMemorySize"
''        .Col = 1
''        .Text = Object.TotalVisibleMemorySize
''
''        .Row = 51
''        .Col = 0
''        .Text = "Version"
''        .Col = 1
''        .Text = Object.Version
''
''        .Row = 52
''        .Col = 0
''        .Text = "WindowsDirectory"
''        .Col = 1
''        .Text = Object.WindowsDirectory
''      End With
''
''    Exit For
''    Next
''
'
'    'MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
'    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_ComputerSystem")
'
'    If Err.Number = 91 Then
'        MsgBox "Please CONNECT to Local or Remote computer first", vbOKOnly, "Connect please!"
'        Exit Sub
'    End If
'
''    For Each Object In Enumerator
''      With msfgComputerSystem
''
''
''        TTz4 = Trim(Object.Manufacturer + " " + Object.Model) + vbCrLf + cpudes + "  "
''        Exit For
''
''            .Row = 22
''        .Col = 0
''        .Text = "Manufacturer"
''        .Col = 1
''        .Text = Object.Manufacturer
''
''        .Row = 23
''        .Col = 0
''        .Text = "Model"
''        .Col = 1
''        .Text = Object.Model
''
''
''
''    End With
''    Exit For
''    Next
'
''TTz5 = ""
''    'Processor Information
''Set obj = GetObject("winmgmts:").InstancesOf("Win32_Processor")
''    For Each obj2 In obj
''        'text3.Text = text3.Text & obj2.Caption & ENTER
''        'text3.Text = text3.Text & "Speed: " & obj2.currentclockspeed & " Mhz" & ENTER
''        TTz5 = Format(Val(obj2.CurrentClockSpeed) / 1000, "0.0") + "GHz "
''    Next
''
'
''WRONG FOR AGES
''Set obj = GetObject("winmgmts:").InstancesOf("Win32_PhysicalMemory")
'Set obj = GetObject("winmgmts:").InstancesOf("Win32_ComputerSystem")
'
'Dim i As String
'For Each obj2 In obj
'
'        'WRONG FOR AGES
'        'TTz5 = TTz5 + Trim(Trim(str(obj2.Capacity / 1024 ^ 3))) + " Gb Ram"
''        TTz5 = TTz5 + Trim(Format(obj2.TotalPhysicalMemory / 1024 ^ 3, "0.00")) + "G Ram"
'        DOMAINGET = obj2.DOMAIN
'
'
'    Exit For
'Next
'
'        'msgbox TTz4 + " " + cpudes
'
'
'
'
'
'
'
'End Sub
'
'




'MAIN_Module1   ---- HOLD FIRST
Sub MAIN2()

'EMPTY COMMONS
End

End Sub


Sub GET_USUAL_SPECIALS()
'### USAGE
sWindowsFolder = GetSpecialfolder(36)
sMyDocsFolder = GetSpecialfolder(5) 'MY DOCUMENTS

End Sub

Sub SHOW_DEBUG_SPECIALS()

Dim R As Long
On Error Resume Next
For R = 0 To 120
    Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
Next

End

End Sub

Public Function GetSpecialFolder_Show_Script_Debug(CSIDL As Long) As String

Dim R As Long
On Error Resume Next
For R = 0 To 120
    Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
Next
End
    
End Function

Public Function GetSpecialfolder(CSIDL As Long) As String
    '##############################################################################################
    'Returns the Path to a "Special" Folder (i.e. Internet History)
    '##############################################################################################
    
    Dim IDL As ITEMIDLIST
    Dim lResult As Long
    Dim sPath As String
    
    lResult = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If lResult = 0 Then
        sPath = Space$(512)
        lResult = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        
        GetSpecialfolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End Function





'2012-03-24
'RECENT ENTRY CHECK COMPATABLE

'Function GetWindowClass(ByVal hWnd As Long) As String
'    Dim Ret As Long, sText As String
'    sText = Space(255)
'    Ret = GetClassName(hWnd, sText, 255)
'    sText = Left$(sText, Ret)
'   GetWindowClass = sText
'End Function


'Function GetWindowTitle(ByVal hWnd As Long) As String
'   Dim l As Long
'   Dim S As String
'   l = GetWindowTextLength(hWnd)
'   S = Space(l + 1)
'   GetWindowText hWnd, S, l + 1
'   GetWindowTitle = Left$(S, l)
'End Function



Public Function GetWindowState(ByVal lngHwnd As Long) As Integer
    If IsWindow(lngHwnd) = 1 Then
        GetWindowState = -1
    If IsIconic(lngHwnd) <> 0 Then
        GetWindowState = vbMinimized
    ElseIf IsZoomed(lngHwnd) <> 0 Then
        GetWindowState = vbMaximized
    End If
    End If
End Function


