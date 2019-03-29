Attribute VB_Name = "modErrorManager"
'////////////////////////////////////////////////////////
'///          Error Handling & Management Module
'///                (modErrorManager.bas)
'///_____________________________________________________
'/// Error management and handling routines. Error codes
'/// enumeration are supported here.
'/// Each Class/Control/Module must have a RaiseErr method
'/// to call this global routines.
'///_____________________________________________________
'/// Last modification  : Ago/10/2000
'/// Last modified by   : Leontti R.
'/// Modification reason: Commented
'/// Project: RamoSoft Code Foundation
'/// Author: Leontti A. Ramos M. (leontti@leontti.net)
'/// RamoSoft de Mexico S.A. de C.V.
'////////////////////////////////////////////////////////
Option Explicit

#Const EXTERNAL_PROVIDER = 1
#Const ERROR_TRAP_ENABLED = 0
#Const LOG_ENABLED_SYSTEM = 1

#If ERROR_TRAP_ENABLED Then
Public Enum ErrorTrapAction
    ERR_RESUMENEXT = 0
    ERR_RESUME = 1
    ERR_QUIT = 2
    ERR_EXITSUB = 3
End Enum
#End If
Public Enum RSErrorCode
    ' Generic errors
    ecBaseError = 1000
    ecInvalidConnectInfo
    ecInvalidValue
    ' ODBC Errors
    ecODBCDriversInitFailed = 1030
    ecODBCMemoryAllocFailed
    ecODBCConnectionFailed
    ecODBCLogOutFailed
    ecODBCDriversUnloadFailed
    ecODBCFreingEnvFailed
    ' DataSet related errors
    ecDataSetEmpty = 1039
    ecODBC_SQL_Error = 1050
    ecInvalidDataSourceType
    ecErrorImportingData
    ' Data stream related
    ecCryptWrongPassword = 1060
    ecCryptAcquirefailed
    ecCryptCreateHashfailed
    ecCryptHashDatafailed
    ecCryptDeriveKeyfailed
    ecCryptEncryptFailed
    ecCompressedDataCorrupted
    ' Object related errors
    ecInvalidObjType = 1070
    ecFormAlreadyLoaded
    ecObjectDoesNotExist
    ecInvalidPropertyName
    ecInvalidDeviceContext
    ecInvalidPropertyUsage
    ecInvalidStartUpModule
    ' System related
    ecTooManyTimers = 1080
    ecCantCreateTimer
    ecNoPrintersDefined
    ecFileDoesNotExist
    ecInvalidFileFormat
    ' Parsing and variable related
    ecSymbolAlreadyDefined = 1090
    ecSymbolNotDefined
    ecEmptyStackCall
    ' SubClass error codes
    ecBaseWindowProc = 13080 ' WindowProc
    ecCantSubclass           ' Can't subclass window
    ecAlreadyAttached        ' Message already handled by another class
    ecInvalidWindow          ' Invalid window
    ecNoExternalWindow       ' Can't modify external window
End Enum

Private m_sLogPath As String
'//////////////////////////////////////////////////
'--------------- Disk Free Space ------------------
Private Declare Function GetFreeSpace Lib "KERNEL32.dll" (ByVal wFlags As Integer) As Long
Private Declare Function GetDiskFreeSpace Lib "KERNEL32.dll" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
'--------------- Registry Access ------------------
Private Const KEY_QUERY_VALUE = &H1 ' Registry Key open mode
Private Const HKEY_LOCAL_MACHINE = &H80000002 ' The Registry section we'll be visiting
Private Const HKEY_DYN_DATA = &H80000006
Private Const RK_Processor = "HARDWARE\DESCRIPTION\System\CentralProcessor\0" ' Root to the processor information
Private Const RK_Performance = "PerfStats\StatData" ' Root to performance statistics
Private Const RK_WIN32_OS = "SOFTWARE\Microsoft\Windows\CurrentVersion" ' Root to OS information on Win machines
Private Const RK_WIN32_OS_NT = "SOFTWARE\Microsoft\Windows NT\CurrentVersion" ' Root to OS information on NT machines
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
'-------------- System Information ----------------
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type
'------------ User & WorkStation Info -------------
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'------------ IP Address Information --------------
Private Type HOSTENT
    hName      As Long
    hAliases   As Long
    hAddrType  As Integer
    hLen       As Integer
    hAddrList  As Long
End Type
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
'--------- Version information constants -----------
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0
'--------- System information constants -----------
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private Const PROCESSOR_INTEL_386 = 386
Private Const PROCESSOR_INTEL_486 = 486
Private Const PROCESSOR_INTEL_PENTIUM = 586
Private Const PROCESSOR_MIPS_R4000 = 4000
Private Const PROCESSOR_ALPHA_21064 = 21064
Private Declare Sub GlobalMemoryStatus Lib "KERNEL32.dll" (lpBuffer As MEMORYSTATUS)
Private Declare Sub GetSystemInfo Lib "KERNEL32.dll" (lpSystemInfo As SYSTEM_INFO)

Public Function GetSysInfo() As String
    GetSysInfo = prvGetUserInfo & vbCrLf
    GetSysInfo = GetSysInfo & prvGetWSInfo & vbCrLf
    GetSysInfo = GetSysInfo & prvGetOSInfo & vbCrLf
    GetSysInfo = GetSysInfo & prvGetCPUInfo & vbCrLf
    GetSysInfo = GetSysInfo & prvGetMemoryStatus & vbCrLf
    GetSysInfo = GetSysInfo & prvGetDiskSpace
End Function

Public Function prvGetLocalIP() As String
    'Resolves the LrHost-name (or current machine if balnk) to an IP address
    Dim LsHostName As String * 256
    Dim LnHostIP As Long
    Dim LrHost As HOSTENT
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim LnIdx As Integer
    Dim LsIPAddress As String

    Call GetComputerName(LsHostName, 255)
    LnHostIP = gethostbyname(LsHostName)
    If (LnHostIP <> 0) Then
        CopyMemory LrHost, LnHostIP, Len(LrHost)
        CopyMemory dwIPAddr, LrHost.hAddrList, 4
        ReDim tmpIPAddr(1 To LrHost.hLen)
        CopyMemory tmpIPAddr(1), dwIPAddr, LrHost.hLen
        For LnIdx = 1 To LrHost.hLen
            LsIPAddress = LsIPAddress & tmpIPAddr(LnIdx) & "."
        Next
        prvGetLocalIP = Mid$(LsIPAddress, 1, Len(LsIPAddress) - 1)
    End If
End Function


#If ERROR_TRAP_ENABLED Then
Public Function ErrorTrap(sModuleName As String) As ErrorTrapAction
    Dim LnErrorCD As Long
    Dim LsErrorSrc As String
    Dim LsErrorDescr As String
    Dim LsCaption As String
    Dim LsLogID As String
    Dim LnLastCursor As Long
    ' Retrieve Error information as soon as posible
    ' to avoid Err object to be cleared
    LnErrorCD = Err.Number
    LsErrorSrc = Err.Source
    LsErrorDescr = Err.Description
    ' Ok, now you can safely clear Err object to avoid
    ' Some system confusion with then previous error
    Err.Clear
    ' Obtain a unique log identifier to report pourposes
    LsLogID = prvLogError(LnErrorCD, LsErrorDescr, LsErrorSrc)
    ' Stores old cursos state, sets a default one
    LnLastCursor = Screen.MousePointer
    Screen.MousePointer = vbDefault
    ' Build a descriptive form caption
    With App
        LsCaption = .EXEName & " v" & .Major & "." & .Minor & " (Rev. " & .Revision & ") "
    End With
    ' Fill in form values
    On Error GoTo ERROR_HANDLER
    With frmErrorTrap
        .TimeStamp = Now
        .SerialNumber = "Log ID: " & LsLogID
        .SystemInfo = GetSysInfo
        .ErrorNumber = LnErrorCD
        ' Description must be very informative
        .ErrorDescription = Format(Now, "Mmmm dd yyyy, h:mm") & " Hrs" & vbCrLf & _
            LsErrorDescr & vbCrLf & String(50, "-") & vbCrLf & .SystemInfo
        .ErrorSource = LsErrorSrc
        .Caption = LsCaption
        .ModuleName = sModuleName
        .Show vbModal
        ErrorTrap = .UserSelection
    End With
    Screen.MousePointer = LnLastCursor
ERROR_HANDLER:
    Unload frmErrorTrap
End Function
#End If

Private Function prvGetOSInfo() As String
    Dim LrOSVersion As OSVERSIONINFO
    
    With LrOSVersion
        .dwOSVersionInfoSize = 148
        GetVersionEx LrOSVersion
        If (.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS) Then
            If (.dwMinorVersion = 0) Then
                prvGetOSInfo = "Microsoft Windows 95"
            Else
                prvGetOSInfo = "Microsoft Windows 98"
            End If
        ElseIf (.dwPlatformID = VER_PLATFORM_WIN32_NT) Then
            If (.dwMajorVersion = 4) Then
                prvGetOSInfo = "Microsoft Windows NT"
            Else
                prvGetOSInfo = "Microsoft Windows 2000"
            End If
        End If
        prvGetOSInfo = "Op. System:  " & prvGetOSInfo & " v" & .dwMajorVersion & "." & _
            Format(.dwMinorVersion, "00") & "." & .dwBuildNumber & _
            Left$(.szCSDVersion, (InStr(1, .szCSDVersion, Chr(0)) - 1))
    End With
End Function

Private Function prvGetUserInfo() As String
    Dim LsBuffer As String
    Dim LnRetCD As Long
    
    LsBuffer = String(255, 0)
    LnRetCD = GetUserName(LsBuffer, 255)
    LsBuffer = Left$(LsBuffer, (InStr(1, LsBuffer, Chr(0)) - 1))
    prvGetUserInfo = "Logged User: " & IIf((Len(LsBuffer) = 0), "[No Logged]", "lsbuffer")
End Function

Private Function prvGetWSInfo() As String
    Dim LsBuffer As String * 256
    Dim LnRetCD As Long
    
    LnRetCD = GetComputerName(LsBuffer, 255)
    LsBuffer = Left$(LsBuffer, (InStr(1, LsBuffer, Chr(0)) - 1))
    prvGetWSInfo = "Workstation: " & RTrim$(LsBuffer) & " (IP " & prvGetLocalIP & ")"
End Function

Public Sub WriteLog(sText As String)
    On Error Resume Next
    Dim LsPath As String
    Dim LhFile As Integer
    Dim LsData As String
    
    ' Creates the log file path
    If (Len(m_sLogPath) = 0) Then
        LsPath = App.Path
    Else
        LsPath = m_sLogPath
    End If
    If (Right$(LsPath, 1) <> "\") Then LsPath = LsPath & "\"
    LsPath = LsPath & App.EXEName & "Log" & Format(Year(Now), "00") & "\" & Format(Month(Now), "00") & "\"
    ' Esure directory exists
    If (Not DirExist(LsPath)) Then
        MkDirNested LsPath
    End If
    LsPath = LsPath & Format(Day(Now), "00.log")
    ' Build log string
    If (Right$(sText, 2) = vbCrLf) Then
        sText = Left$(sText, (Len(sText) - 2))
    End If
    LsData = "[" & Format(Now, "YYYY-MM-DD ") & Format(Now, "H:MM.SS") & "] " & sText
    'Debug.Print LsData
    ' Open log file to write in
    LhFile = FreeFile
    ' If the file was just created, write down a descriptive header
    If (Dir(LsPath) = "") Then
        Open LsPath For Append As LhFile
        With App
            Dim LsDescr As String
            Print #LhFile, String(50, "=")
            #If EXTERNAL_PROVIDER Then
            LsDescr = .EXEName
            #Else
            LsDescr = IIf((.ProductName = ""), ("RamoSoft " & .EXEName), .ProductName)
            #End If
            LsDescr = LsDescr & " v" & .Major & "." & .Minor & IIf((.Revision > 0), (" (Rev. " & .Revision & ")"), "")
            Print #LhFile, LsDescr
            Print #LhFile, "Aplication log file for " & Format(Now, "Mmm dd, yyyy")
            Print #LhFile, String(50, "-")
            Print #LhFile, GetSysInfo
            Print #LhFile, String(50, "=")
        End With
    Else
        Open LsPath For Append As LhFile
    End If
    Print #LhFile, LsData
    Close #LhFile
End Sub

Private Function prvGetErrorUniqueID(lErrNum As Long) As String
    Dim LsErrorNum As String
    Dim LsVersion As String
    Dim LsDate As String
    Dim LsTime As String
    Dim LnLen As Integer
    
    On Error Resume Next
    LsErrorNum = Format(lErrNum, "00000000")
    LsVersion = Format(App.Major, "00") & Format(App.Minor, "000") & Format(App.Revision, "00")
    LsDate = Format(Now, "yyyymmdd")
    LsTime = Format(Now, "hmmss")
    If (Len(LsTime) < 6) Then LsTime = ("0" & LsTime)
    prvGetErrorUniqueID = ("E" & LsErrorNum & "-" & LsVersion & "-" & LsDate & LsTime)
End Function


Public Function prvGetMemoryStatus() As String
    Dim LrMemoryStatus As MEMORYSTATUS
    Dim LnFree As Long
    Dim LnTotal As Long
    Dim LnPercent As Single
    Dim LsBuffer As String

    With LrMemoryStatus
        .dwLength = 32
        GlobalMemoryStatus LrMemoryStatus
        LnFree = (.dwAvailPhys / 1024)
        LnTotal = (.dwTotalPhys / 1024)
    End With
    LnPercent = ((LnFree / LnTotal) * 100)
    LsBuffer = (Format(LnFree, "#,###,###,##0") & " / " & Format(LnTotal, "#,###,###,##0") & " Kb")
    LsBuffer = LsBuffer & " (" & Format(LnPercent, "##0.0#%") & " Free)"
    prvGetMemoryStatus = "RAM Memory:  " & LsBuffer
End Function


Public Sub SetLogPath(sPath As String)
    m_sLogPath = sPath
    If (Right$(m_sLogPath, 1) <> "\") Then m_sLogPath = m_sLogPath & "\"
End Sub

Public Function DirExist(sPath As String) As Boolean
    DirExist = (Dir(sPath, vbDirectory) <> "")
End Function

Private Function prvGetDiskSpace(Optional sDrive As String = "C") As String
    Dim LnSectorsPerCluster As Long
    Dim LnBytesPerSector As Long
    Dim LnFreeClustersCount As Long
    Dim LnTotalClustersCount As Long
    Dim LnRetCD As Long
    Dim LnFree As Long
    Dim LnTotal As Double
    Dim LnPercent As Single
    
    sDrive = Left$(Trim$(sDrive), 1) & ":\"    ' Ensure path is at the root.
    LnRetCD = GetDiskFreeSpace(sDrive, LnSectorsPerCluster, LnBytesPerSector, LnFreeClustersCount, LnTotalClustersCount)
    If (LnRetCD = 0) Then
        prvGetDiskSpace = "Disk Space: [Unknown]"
    Else
        LnFree = ((LnSectorsPerCluster * LnBytesPerSector * LnFreeClustersCount) / 1024)
        LnTotal = ((LnSectorsPerCluster * LnBytesPerSector * LnTotalClustersCount) / 1024)
        LnPercent = (LnFree / LnTotal)
        prvGetDiskSpace = "Disk Space:  " & (Format(LnFree, "#,###,###,##0") & " / " & _
            Format(LnTotal, "#,###,###,##0") & " Kb (" & Format(LnPercent, "##0.0#%") & " Free)")
    End If
End Function

Public Function prvGetCPUInfo() As String
    Dim LrSystemInfo As SYSTEM_INFO
    Dim LsProcessorName As String
    Dim LsBuffer As String
    Dim LsTemp As String

    GetSystemInfo LrSystemInfo
    
    LsTemp = prvGetKeyValue(HKEY_LOCAL_MACHINE, RK_Processor, "Identifier")
    If (Len(LsTemp) = 0) Then
        Select Case LrSystemInfo.dwProcessorType
            Case PROCESSOR_INTEL_386
                LsTemp = "Intel 386"
            Case PROCESSOR_INTEL_486
                LsTemp = "Intel 486"
            Case PROCESSOR_INTEL_PENTIUM
                LsTemp = "Intel Pentium"
            Case PROCESSOR_MIPS_R4000
                LsTemp = "MIPS R4000"
            Case PROCESSOR_ALPHA_21064
                LsTemp = "DEC Alpha 21064"
            Case Else
                LsTemp = "[Unknown]"
        End Select
    End If
    LsBuffer = LsTemp
    LsTemp = Trim$(prvGetKeyValue(HKEY_LOCAL_MACHINE, RK_Processor, "~MHZ"))
    If (Len(LsTemp) > 0) Then
        LsBuffer = LsBuffer & " " & LsTemp & " MHz"
    End If
    LsTemp = prvGetKeyValue(HKEY_LOCAL_MACHINE, RK_Processor, "VendorIdentifier")
    If (Len(LsTemp) > 0) Then
        LsBuffer = LsBuffer & " (" & LsTemp & ") "
    End If
    With LrSystemInfo
        If (.dwNumberOrfProcessors > 1) Then
            LsTemp = " " & .dwNumberOrfProcessors & " processors (#" & .dwActiveProcessorMask & " active)"
            LsBuffer = (LsBuffer & LsTemp)
        End If
    End With
    prvGetCPUInfo = "Processors:  " & LsBuffer
    
    'LrSystemInfo.dwProcessorTypE
    ''LrSystemInfo.dwActiveProcessorMask
    
End Function

Private Function prvGetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'==========================================================================================================='
' Returns a specified key value from the registry                                                           '
'==========================================================================================================='
    Dim LnKey As Long
    Dim LsBuffer As String
    Dim LnKeySize As Long
    Dim LnKeyType As Long
    Dim LnIdx As Integer
    
    ' Prepares data buffer
    LsBuffer = String(1024, 0)
    LnKeySize = 1024
    ' Open the registry key. Any value other than zero means something went wrong
    If (RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_QUERY_VALUE, LnKey) = 0) Then
        ' Retrieve the registry value, any value other than zero means something went wrong
        If (RegQueryValueEx(LnKey, SubKeyRef, 0, LnKeyType, LsBuffer, LnKeySize) = 0) Then
            ' Extract the useful string from the garble
            If (Asc(Mid(LsBuffer, LnKeySize, 1)) = 0) Then
                LsBuffer = Left(LsBuffer, LnKeySize - 1)
            Else
                LsBuffer = Left(LsBuffer, LnKeySize)
            End If
            ' If the returned value is a dword we need to format the value to something meaningful
            If (LnKeyType = 4) Then
                For LnIdx = Len(LsBuffer) To 1 Step -1
                    prvGetKeyValue = prvGetKeyValue + Hex(Asc(Mid(LsBuffer, LnIdx, 1)))
                Next
                prvGetKeyValue = Format("&H" & prvGetKeyValue)
            Else
                prvGetKeyValue = LsBuffer
            End If
        End If
    End If
    RegCloseKey LnKey
End Function

Public Function MkDirNested(sFullPath As String) As Boolean
    On Error GoTo ErrHandler
    Dim LnNextSlash As Integer
    Dim LnStartPos As Integer
    Dim LsCurDir As String
    ' Set first char
    LnStartPos = 1
    ' validates path syntax
    If (Right$(sFullPath, 1) <> "\") Then sFullPath = (sFullPath & "\")
    Do
        LnNextSlash = InStr(LnStartPos, sFullPath, "\")
        If (LnNextSlash >= LnStartPos) Then
            LsCurDir = Left$(sFullPath, LnNextSlash)
            If (Not DirExist(LsCurDir)) Then
               ' Create the dir
               MkDir LsCurDir
            End If
             ' Check if it's the last char and exit if true
             LnStartPos = (LnNextSlash + 1)
             If (LnStartPos >= Len(sFullPath)) Then Exit Do
        End If
    Loop
    MkDirNested = True
    Exit Function
ErrHandler:
    MkDirNested = False
End Function

Public Function GetLogPath() As String
    GetLogPath = m_sLogPath
End Function

Public Sub RaiseError(ByVal lErrNum As RSErrorCode, sModuleName As String, _
    Optional sRoutineName As String, Optional sDescription As String, _
    Optional lLineNumber As Long)
    Dim LsSource As String
    
    LsSource = App.EXEName & "." & sModuleName
    If Len(sRoutineName) Then LsSource = LsSource & "." & sRoutineName
    If (lLineNumber <> 0) Then LsSource = LsSource & "." & CStr(lLineNumber)
    If (lErrNum > ecBaseError) And (Len(sDescription) = 0) Then _
        sDescription = prvGetErrorDescription(lErrNum)
    #If LOG_ENABLED_SYSTEM Then
    Trace sDescription, LsSource, lErrNum, vbLogEventTypeError
    #Else
    prvLogError lErrNum, sDescription, LsSource
    #End If
    Err.Clear
    On Error GoTo 0
    If Len(sDescription) Then
        Err.Raise lErrNum, LsSource, sDescription
    Else
        Err.Raise lErrNum, LsSource
    End If
End Sub

Private Function prvGetErrorDescription(lErrCD As RSErrorCode) As String
    Select Case lErrCD
        Case ecInvalidConnectInfo
            prvGetErrorDescription = "Invalid connection info."
        Case ecInvalidValue
            prvGetErrorDescription = "Invalid value."
        Case ecODBCDriversInitFailed
            prvGetErrorDescription = "Unable to initialize ODBC API drivers!."
        Case ecODBCMemoryAllocFailed
            prvGetErrorDescription = "Could not allocate memory for connection\statement Handle!."
        Case ecODBCFreingEnvFailed
            prvGetErrorDescription = "Error Freeing Environment From ODBC Drivers."
        Case ecODBCConnectionFailed
            prvGetErrorDescription = "Could not establish connection to ODBC driver!."
        Case ecODBCLogOutFailed
            prvGetErrorDescription "Error logging out of data source!."
        Case ecODBCDriversUnloadFailed
            prvGetErrorDescription = "Error unloading ODBC drivers!."
        Case ecDataSetEmpty
            prvGetErrorDescription = "No data returned."
        Case ecODBC_SQL_Error
            prvGetErrorDescription = "SQL driver error."
        Case ecInvalidDataSourceType
            prvGetErrorDescription = "Invalid Data Source Type."
        Case ecErrorImportingData
            prvGetErrorDescription = "Error importing data."
        Case ecCryptWrongPassword
            prvGetErrorDescription = "Invalid decryption password was passed."
        Case ecInvalidObjType
            prvGetErrorDescription = "Invalid Object Type."
        Case ecFormAlreadyLoaded
            prvGetErrorDescription = "Form already loaded."
        Case ecObjectDoesNotExist
            prvGetErrorDescription = "Object does not exist."
        Case ecInvalidPropertyName
            prvGetErrorDescription = "Invalid property name."
        Case ecInvalidDeviceContext
            prvGetErrorDescription = "Invalid Device Context."
        Case ecInvalidPropertyUsage
            prvGetErrorDescription = "Invalid property usage at this time."
        Case ecInvalidStartUpModule
            prvGetErrorDescription = "Invalid StartUp module."
        Case ecTooManyTimers
            prvGetErrorDescription = "No more than 10 timers allowed per class."
        Case ecCantCreateTimer
            prvGetErrorDescription = "Can't create system timer."
        Case ecNoPrintersDefined
            prvGetErrorDescription = "No printers defined."
        Case ecFileDoesNotExist
            prvGetErrorDescription = "File does not exist."
        Case ecInvalidFileFormat
            prvGetErrorDescription = "Invalid file format or signature."
        Case ecSymbolAlreadyDefined
            prvGetErrorDescription = "Symbol already defined."
        Case ecSymbolNotDefined
            prvGetErrorDescription = "Symbol not defined."
        Case ecEmptyStackCall
            prvGetErrorDescription = "Invalid Empty Stack Call."
        Case ecCryptAcquirefailed
            prvGetErrorDescription = "Function CryptoAcquireContext failed"
        Case ecCryptCreateHashfailed
            prvGetErrorDescription = "Function CryptCreateHash failed"
        Case ecCryptHashDatafailed
            prvGetErrorDescription = "Function CryptHashData failed"
        Case ecCryptDeriveKeyfailed
            prvGetErrorDescription = "Function CryptDeriveKey failed"
        Case ecCryptEncryptFailed
            prvGetErrorDescription = "Function CryptEncrypt failed"
        Case ecCompressedDataCorrupted
            prvGetErrorDescription = "Data might be corrupted (Invalid Checksum)."
        Case ecBaseWindowProc
            prvGetErrorDescription = "Base Window Procedure (SubClass)."
        Case ecCantSubclass           ' Can't subclass window
            prvGetErrorDescription = "Can't subclass window"
        Case ecAlreadyAttached        ' Message already handled by another class
            prvGetErrorDescription = "Message already handled by another class"
        Case ecInvalidWindow          ' Invalid window
            prvGetErrorDescription = "Invalid window"
        Case ecNoExternalWindow       ' Can't modify external window
            prvGetErrorDescription = "Can't modify external window"
        Case Else
            prvGetErrorDescription = "Unknown error."
    End Select
End Function

Public Sub Trace(sMessage As String, Optional sSource As String, _
    Optional lEventCD As Long, Optional iEventType _
    As LogEventTypeConstants = vbLogEventTypeInformation)
    Dim LsMsg As String
    Dim LsHdr As String
    Dim LsSrc As String
    
    Select Case iEventType
        Case vbLogEventTypeError
            LsHdr = "Error ["
        Case vbLogEventTypeInformation
            LsHdr = "Info ["
        Case vbLogEventTypeWarning
            LsHdr = "Warning ["
        Case Else
            LsHdr = "Event ["
    End Select
    LsHdr = LsHdr & Format(Now, "mm-dd-yyyy h:mm] ")
    If (lEventCD <> 0) Then
        LsMsg = "[CD 0x" & Hex(lEventCD) & "] "
    End If
    LsMsg = LsMsg & sMessage
    If (Len(sSource) = 0) Then
        LsSrc = "Unknown Source"
    Else
        LsSrc = "Source: " & sSource
    End If
    App.LogEvent LsHdr & LsMsg & ", " & LsSrc, iEventType
    WriteLog String(50, "-")
    WriteLog LsHdr
    WriteLog LsMsg
    WriteLog sSource
    WriteLog String(50, "-")
End Sub

Private Function prvLogError(lErrorNum As Long, sDescription As String, sSource As String) As String
    Dim LsFile As String
    Dim LnFileHandler As Integer

    On Error GoTo prvLogError_EH
    LsFile = App.Path & "\" & App.EXEName & ".err"
    prvLogError = prvGetErrorUniqueID(lErrorNum)
    LnFileHandler = FreeFile
    Open LsFile For Append As #LnFileHandler
    Print #LnFileHandler, String(60, "=")
    Print #LnFileHandler, "Event ID: " & prvLogError
    Print #LnFileHandler, String(60, "-")
    Print #LnFileHandler, "Error number " & CStr(lErrorNum) & ", ocurred on " & Format(Now, "Mmmm dd yyyy, h:mm") & " Hrs"
    Print #LnFileHandler, App.EXEName & " version " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    Print #LnFileHandler, "Source: " & sSource
    Print #LnFileHandler, "Description: " & sDescription
    Print #LnFileHandler, String(60, "-")
    Print #LnFileHandler, GetSysInfo
    Print #LnFileHandler, String(60, "_")
prvLogError_EH:
    Close #LnFileHandler
End Function



'http://stackoverflow.com/questions/15765804/check-os-and-processor-is-32-bit-or-64-bit

'AddressWidth
'
'On a 32-bit operating system, the value is 32 and on a 64-bit operating system it is 64.
'Relevant VB6 code is:

'HEAVY TASK - SWITCH TO PROBLEM CAN HAPPEN

Public Function GetOsBitness() As String
    Dim ProcessorSet As Object
    Dim CPU As Object

    Set ProcessorSet = GetObject("Winmgmts:"). _
        ExecQuery("SELECT * FROM Win32_Processor")
    For Each CPU In ProcessorSet
        GetOsBitness = CStr(CPU.AddressWidth)
    Next
End Function

'Is processor 32-bit or 64-bit?
'
'For processor one can check DataWidth property:
'
'DataWidth
'
'On a 32-bit processor, the value is 32 and on a 64-bit processor it is 64.
'Relevant VB6 code is:

'HEAVY TASK - SWITCH TO PROBLEM CAN HAPPEN

Public Function GetCpuBitness() As String
    Dim ProcessorSet As Object
    Dim CPU As Object

    Set ProcessorSet = GetObject("Winmgmts:"). _
        ExecQuery("SELECT * FROM Win32_Processor")
    For Each CPU In ProcessorSet
        GetCpuBitness = CStr(CPU.DataWidth)
    Next
End Function


'------------------------------------------------


Public Function GetOsArchitecture()
    If IsAtLeastVista Then
        GetOsArchitecture = GetVistaOsArchitecture
    Else
        GetOsArchitecture = GetXpOsArchitecture
    End If
End Function

Private Function IsAtLeastVista() As Boolean
    IsAtLeastVista = GetOsVersion >= "6.0"
End Function

Private Function GetOsVersion() As String
    Dim OperatingSystemSet As Object
    Dim OS As Object

    Set OperatingSystemSet = GetObject("winmgmts:{impersonationLevel=impersonate}"). _
                                    InstancesOf("Win32_OperatingSystem")
    For Each OS In OperatingSystemSet
        GetOsVersion = Left$(Trim$(OS.Version), 3)
    Next
End Function

Private Function GetVistaOsArchitecture() As String
    Dim OperatingSystemSet As Object
    Dim OS As Object

    Set OperatingSystemSet = GetObject("Winmgmts:"). _
        ExecQuery("SELECT * FROM Win32_OperatingSystem")
    For Each OS In OperatingSystemSet
        GetVistaOsArchitecture = Left$(Trim$(OS.OSArchitecture), 2)
    Next
End Function

Private Function GetXpOsArchitecture() As String
    Dim ComputerSystemSet As Object
    Dim Computer As Object
    Dim SystemType As String

    Set ComputerSystemSet = GetObject("Winmgmts:"). _
        ExecQuery("SELECT * FROM Win32_ComputerSystem")
    For Each Computer In ComputerSystemSet
        SystemType = UCase$(Left$(Trim$(Computer.SystemType), 3))
    Next

    GetXpOsArchitecture = IIf(SystemType = "X86", "32", "64")
End Function

