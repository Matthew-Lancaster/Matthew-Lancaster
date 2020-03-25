VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCnProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#####################################################################################
'#  Process Info, Traversal & Conversion (PID/EXE/WinTitle/Class/hWnd) Class - NT4 Friendly (clsCnProc.cls)
'#      By: Nick Campbeln
'#
'#      Revision History:
'#          1.0 (Aug 28, 2002):
'#              Initial Release
'#
'#      Copyright © 2002 Nick Campbeln (opensource@nick.campbeln.com)
'#          This source code is provided 'as-is', without any express or implied warranty. In no event will the author(s) be held liable for any damages arising from the use of this source code. Permission is granted to anyone to use this source code for any purpose, including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:
'#          1. The origin of this source code must not be misrepresented; you must not claim that you wrote the original source code. If you use this source code in a product, an acknowledgment in the product documentation would be appreciated but is not required.
'#          2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original source code.
'#          3. This notice may not be removed or altered from any source distribution.
'#              (NOTE: This license is borrowed from zLib.)
'#
'#  NOTE: WinNT4 support requires that "PSAPI.dll" be present in the target WinNT4 system's path (current directory, system directory, etc). This DLL is not installed by default on WinNT4, so be advised. It is provided in the PSC.com zip as 'PSAPI-dll', simply rename to 'PSAPI.dll'.
'#
'#  Please remember to vote on PSC.com if you like this code!
'#  Code URL: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=38425&lngWId=1
'#####################################################################################
'# Future Features: Add ModuleID & ParentProcessID as class properties and Convert() options?



'#####################################################################################
'# Private sub/function/type/const/etc definitions required by the class
'#####################################################################################
'#########################################################
'# General definitions
'#########################################################
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Const MAX_PATH = 260

Private sEXEName As String
Private hProcess As Long
Private lProcessID As Long
Private bIsNT4 As Boolean
Private bIsOpen As Boolean


'#########################################################
'# Functions/Consts/Types used for GetVersionEx()
'#########################################################
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'Private Const VER_PLATFORM_WIN32s = 0
'Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long           '#### NT: Build Number, 9x: High-Order has Major/Minor ver, Low-Order has build
    PlatformID As Long
    szCSDVersion As String * 128    '#### NT: ie- "Service Pack 3", 9x: 'arbitrary additional information'
End Type


'#########################################################
'# Win32 (non-NT4) specific definitions
'#########################################################
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

    '#### Required private data members
Private uProcess As PROCESSENTRY32
Private hSnapShot As Long


'#########################################################
'# Windows NT4 specific definitions
'#    NOTE: Remember to distribute the PSAPI.DLL file.
'#########################################################
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
'Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
'Private Const SYNCHRONIZE = &H100000
'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
'Private Const PROCESS_ALL_ACCESS = &H1F0FFF

    '#### Required private data members
Private Modules(1 To 200) As Long
Private lProcessIDs() As Long
Private lcbNeeded As Long
Private lRetVal As Long
Private i As Long


'#########################################################
'# Convert() releated definitions
'#########################################################
    '#### Functions/Consts used for GetHWnd() (part of Convert())
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

    '#### eNum used with Convert()
Public Enum cnProcessConv
    cnFromClass = 1
    cnFromEXE = 2
    cnFromhWnd = 4
    cnFromProcessID = 8
    cnFromTitle = 16
    cnToClass = 32
    cnToEXE = 64
    cnTohWnd = 128
    cnToProcessID = 256
    cnToTitle = 512
End Enum



'#####################################################################################
'# Class Functions
'#####################################################################################
'#########################################################
'# Class constructor to init private variables
'#########################################################
Private Sub Class_Initialize()
    Dim oOSInfo As OSVERSIONINFO

        '#### Init the class vars
    sEXEName = ""
    lProcessID = 0
    bIsOpen = False

        '#### Determine the value of bIsNT4
    With oOSInfo
        .OSVSize = Len(oOSInfo)
        .szCSDVersion = Space(128)
        lRetVal = GetVersionEx(oOSInfo)
        bIsNT4 = (.PlatformID = VER_PLATFORM_WIN32_NT And .dwVerMajor = 4)
    End With
End Sub


'#########################################################
'# Class destructor to kill private variables
'#########################################################
Private Sub Class_Terminate()
        '#### If we're running under a system that supports CreateToolhelp32Snapshot()
    If (Not bIsNT4) Then
            '#### Close the hSnapShot handle
        Call CloseHandle(hSnapShot)
    End If

        '#### Set bIsOpen
    bIsOpen = False
End Sub


'#########################################################
'# Get/Let Properties
'#########################################################
Public Property Get Class() As String
        '#### If conversion is not successful, reset the return value
    If (Not Convert(lProcessID, Class, cnFromProcessID Or cnToClass)) Then
        Class = ""
    End If
End Property
Public Property Get EXE() As String
    EXE = sEXEName
End Property
Public Property Get hWnd() As Long
        '#### If conversion is not successful, reset the return value
    If (Not Convert(lProcessID, hWnd, cnFromProcessID Or cnTohWnd)) Then
        hWnd = -1
    End If
End Property
Public Property Get ProcessID() As Long
    ProcessID = lProcessID
End Property
Public Property Get Title() As String
        '#### If conversion is not successful, reset the return value
    If (Not Convert(lProcessID, Title, cnFromProcessID Or cnToTitle)) Then
        Title = ""
    End If
End Property



'#####################################################################################
'# Public subs/functions
'#####################################################################################
'#########################################################
'# Opens the processes and moves to the first ProcessID
'#########################################################
Public Function OpenProcesses() As Boolean
        '#### If we're running under WinNT4
    If (bIsNT4) Then
        Dim lcb As Long

            '#### Init the API vars
        lcbNeeded = 96
        lcb = 8

            '#### While lcbNeeded is bigger then lcb, loop to find the correct size of lProcessIDs()
        Do While (lcb <= lcbNeeded)
                '#### Increase the size of lcb by 2, dividing the result by 4 to convert from bytes to processes
            lcb = lcb * 2
            ReDim lProcessIDs(lcb / 4)

                '#### If the return value is 0, error out
            If (EnumProcesses(lProcessIDs(1), lcb, lcbNeeded) = 0) Then
                GoTo OpenProcesses_Error
            End If
        Loop

            '#### Init i, set bIsOpen and move to the first process, returning the result of the NextProcess() call to the caller
        i = 1
        bIsOpen = True
        OpenProcesses = NextProcess()

        '#### Else we're running under a system that supports CreateToolhelp32Snapshot()
    Else
            '#### Setup hSnapShot, begin to setup the uProcess UDT
        hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
        uProcess.dwSize = Len(uProcess)

            '#### If hSnapShot was successfully setup
        If (hSnapShot <> 0) Then
                '#### Find the first hProcess and set the return value
            hProcess = ProcessFirst(hSnapShot, uProcess)
            OpenProcesses = (hProcess <> 0)

                '#### If a valid hProcess was found, set sEXEName and lProcessID
            If (OpenProcesses) Then
                sEXEName = TrimNull(uProcess.szExeFile)
                lProcessID = uProcess.th32ProcessID
            End If

            '#### Else hSnapShot was not successfully setup, so error out
        Else
            GoTo OpenProcesses_Error
        End If
    End If

        '#### Set bIsOpen
    bIsOpen = OpenProcesses
    Exit Function

OpenProcesses_Error:
    OpenProcesses = False
    bIsOpen = False
End Function


'#########################################################
'# Moves to the next lProcessID, setting sEXEName and lProcessID
'#########################################################
Public Function NextProcess() As Boolean
        '#### If there is currently process info open
    If (bIsOpen) Then
            '#### If we're running under WinNT4
        If (bIsNT4) Then
                '#### If we are still within the bounds of lProcessIDs()
            If (i <= (lcbNeeded / 4)) Then
                    '#### Setup hProcess and set the return value
                hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcessIDs(i))
                NextProcess = True

                    '#### If hProcess returned a valid value
                If (hProcess <> 0) Then
                        '#### Set lProcessID
                    lProcessID = lProcessIDs(i)

                        '#### If we are able to retrieve the module handels for the found hProcess
                    If (EnumProcessModules(hProcess, Modules(1), 200, 0) <> 0) Then
                            '#### Init the sEXEName buffer, retrieve the module name then trim it based on lRetVal
                        sEXEName = Space(MAX_PATH)
                        lRetVal = GetModuleFileNameExA(hProcess, Modules(1), sEXEName, 500)
                        sEXEName = Left(sEXEName, lRetVal)
                    End If

                    '#### Else hProcess did not return a valid value, so reset lProcessID and sEXEName
                Else
                    lProcessID = 0
                    sEXEName = ""
                End If

                    '#### Close hProcess and inc i for the next call
                Call CloseHandle(hProcess)
                i = i + 1

                '#### Else we're outside the bounds of lProcessIDs(), so return false
            Else
                NextProcess = False
            End If

            '#### Else we're running under a system that supports CreateToolhelp32Snapshot()
        Else
                '#### If the current hProcess is valid
            If (hProcess <> 0) Then
                    '#### Move hProcess to the next hProcess and set the return value
                hProcess = ProcessNext(hSnapShot, uProcess)
                NextProcess = (hProcess <> 0)

                    '#### If a valid hProcess was found, set sEXEName and lProcessID
                If (NextProcess) Then
                    sEXEName = uProcess.szExeFile
                    lProcessID = uProcess.th32ProcessID
                End If

                '#### Else the current hProcess is invalid, so return false
            Else
                NextProcess = False
            End If
        End If

        '#### Else we're not currently open, so return false
    Else
        NextProcess = False
    End If
End Function


'#########################################################
'# Properly closes the processes
'#########################################################
Public Sub CloseProcesses()
        '#### Reset the class
    Call Class_Terminate
    Call Class_Initialize
End Sub


'#########################################################
'# Converts information about a process from one form to another (hWnd, ProcessID, and RunningEXE)
'#########################################################
Public Function Convert(ByRef InputData, ByRef OutputData, ByVal eConversionType As cnProcessConv) As Boolean
    Dim sClass As String
    Dim sTitle As String
    Dim sEXE As String
    Dim lhWnd As Long
    Dim lProcessID As Long

        '#### Init the vars to invalid values for an hWnd/ProcessID
    lhWnd = -1
    lProcessID = -1

        '#### If we are converting from a Class name
    If (eConversionType And cnFromClass) Then
        sClass = InputData

        '#### Else if we are converting from an EXE name
    ElseIf (eConversionType And cnFromEXE) Then
        sEXE = InputData

            '#### If we're not converting from EXE to EXE then we'll need the lProcessID
        If (Not CBool(eConversionType And cnToEXE)) Then
            If (Not GetEXEID(lProcessID, sEXE)) Then GoTo Convert_Error
        End If

        '#### Else if we are converting from an hWnd
    ElseIf (eConversionType And cnFromhWnd) Then
        If (IsNumeric(InputData)) Then
            lhWnd = InputData
        Else
            GoTo Convert_Error
        End If

        '#### Else if we are converting from a ProcessID
    ElseIf (eConversionType And cnFromProcessID) Then
        If (IsNumeric(InputData)) Then
            lProcessID = InputData
        Else
            GoTo Convert_Error
        End If

        '#### Else if we are converting from a Title
    ElseIf (eConversionType And cnFromTitle) Then
        sTitle = InputData

        '#### Else we are missing a valid 'from'
    Else
        GoTo Convert_Error
    End If

        '#### If we are converting to a Class name
    If (eConversionType And cnToClass) Then
            '#### If sClass is not already set, determine its value now then set the OutputData
        If (Len(sClass) = 0) Then
            If (Not GetHWnd(lProcessID, lhWnd, sTitle, sClass)) Then GoTo Convert_Error
        End If
        OutputData = sClass

        '#### Else if we are converting to an EXE name
    ElseIf (eConversionType And cnToEXE) Then
            '#### If sEXE is not already set, determine its value now then set the OutputData
        If (Len(sEXE) = 0) Then
            If (Not GetEXEID(lProcessID, sEXE)) Then GoTo Convert_Error
        End If
        OutputData = sEXE

        '#### Else if we are converting to an hWnd
    ElseIf (eConversionType And cnTohWnd) Then
            '#### If lhWnd is not already set, determine its value now then set the OutputData
        If (lhWnd = -1) Then
            If (Not GetHWnd(lProcessID, lhWnd, sTitle, sClass)) Then GoTo Convert_Error
        End If
        OutputData = lhWnd

        '#### Else if we are converting to a ProcessID
    ElseIf (eConversionType And cnToProcessID) Then
            '#### If lProcessID is not already set, determine its value now then set the OutputData
        If (lProcessID = -1) Then
            If (Not GetHWnd(lProcessID, lhWnd, sTitle, sClass)) Then GoTo Convert_Error
        End If
        OutputData = lProcessID

        '#### Else if we are converting to a Title
    ElseIf (eConversionType And cnToTitle) Then
            '#### If sTitle is not already set, determine its value now then set the OutputData
        If (Len(sTitle) = 0) Then
            If (Not GetHWnd(lProcessID, lhWnd, sTitle, sClass)) Then GoTo Convert_Error
        End If
        OutputData = sTitle

        '#### Else we are missing a valid 'to'
    Else
        GoTo Convert_Error
    End If

        '#### If we got here, all is well
    Convert = True
    Exit Function

    '#### If we end up here some sort of error occurred, so return false
Convert_Error:
    Convert = False
End Function



'#####################################################################################
'# Private subs/functions
'#####################################################################################
'#########################################################
'# Returns a windows EXE and ProcessID from the passed EXE or ProcessID
'#########################################################
Private Function GetEXEID(ByRef lRunningID As Long, ByRef sRunningEXE As String) As Boolean
        '#### Default the return value
    GetEXEID = False

        '#### If we're able to successfully open the processes
    If (OpenProcesses()) Then
            '#### Get the name of the EXE
        sRunningEXE = UCase(GetFileName(sRunningEXE))

            '#### Do while we still have processes
        Do
                '#### If the ProcessIDs match
            If (lProcessID = lRunningID) Then
               'lRunningID = lProcessID
                sRunningEXE = TrimNull(sEXEName)
                GetEXEID = True
                GoTo GetEXEID_Close

                '#### Else if the EXE names match
            ElseIf (InStr(1, UCase(sEXEName), sRunningEXE, vbBinaryCompare) > 0) Then
                lRunningID = lProcessID
                sRunningEXE = TrimNull(sEXEName)
                GetEXEID = True
                GoTo GetEXEID_Close
            End If
        Loop While (NextProcess())

            '#### If we make it here, the lRunningID/sRunningEXE was not found
        lRunningID = -1
        sRunningEXE = ""

        '#### Else the processes were not successfully opened
    Else
        lRunningID = -1
        sRunningEXE = ""
    End If

GetEXEID_Close:
        '#### Close the class
    Call CloseProcesses
End Function


'#########################################################
'# Returns a windows hWnd, ProcessID, Title and Class from the passed hWnd, ProcessID, Title or Class
'#########################################################
Private Function GetHWnd(ByRef hProcessID As Long, ByRef hWnd As Long, ByRef sTitle As String, ByRef sClass As String) As Boolean
    Dim sWinTitle As String
    Dim sClassName As String
    Dim hWndChild As Long
    Dim hWndChildProcessID As Long

        '#### Get the Desktop handle while getting the first child under the Desktop and default the return value
    hWndChild = GetWindow(GetDesktopWindow(), GW_CHILD)
    GetHWnd = False

        '#### While we still have hWndChild(en) to look at
    Do While (hWndChild <> 0)
            '#### Get the ThreadProcessID of the window
        Call GetWindowThreadProcessId(hWndChild, hWndChildProcessID)

            '#### Get the current window's title
        sWinTitle = Space$(255)
        sWinTitle = Left(sWinTitle, GetWindowText(hWndChild, sWinTitle, 255))

            '#### Get the current window's class
        sClassName = Space$(255)
        sClassName = Left(sClassName, GetClassName(hWndChild, sClassName, 255))

            '#### If we have a match with the hProcessID or hWnd, return the values
        If (hWndChildProcessID = hProcessID Or hWndChild = hWnd) Then
            hProcessID = hWndChildProcessID
            hWnd = hWndChild
            sTitle = sWinTitle
            sClass = sClassName
            GetHWnd = True
            Exit Do

            '#### Else if sWinTitle has a value that is like sTitle, return the values
        ElseIf (Len(sWinTitle) > 0 And Len(sTitle) > 0) Then
            If (sWinTitle Like sTitle) Then
                hProcessID = hWndChildProcessID
                hWnd = hWndChild
                sTitle = sWinTitle
                sClass = sClassName
                GetHWnd = True
                Exit Do
            End If

            '#### Else if sClassName has a value that is like sClass, return the values
        ElseIf (Len(sClassName) > 0 And Len(sClass) > 0) Then
            If (sClassName Like sClass) Then
                hProcessID = hWndChildProcessID
                hWnd = hWndChild
                sTitle = sWinTitle
                sClass = sClassName
                GetHWnd = True
                Exit Do
            End If
        End If

            '#### We've not yet found a match, so get the next hWndChild
        hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
    Loop
End Function


'#########################################################
'# Peals off the last filename element from the passed sPath and returns it to the caller
'#########################################################
Private Function GetFileName(ByVal sPath As String) As String
    Dim iLastSpace As Integer
    Dim iLastSlash As Integer

        '#### If sPath has a value, process it
    If (Len(sPath) > 0) Then
            '#### Deteremine the index of the last "\" and the following " " (if there is one)
        iLastSlash = InStrRev(sPath, "\", -1)
        iLastSpace = InStr(iLastSlash + 1, sPath, " ")

            '#### If a space was found in sPath, make sure we only remove the EXEs name and not any arguments
        If (iLastSpace > 0) Then
            GetFileName = Mid(sPath, iLastSlash + 1, iLastSpace - iLastSlash - 1)

            '#### Else there were no arguments, so peal off the name
        Else
            GetFileName = Mid(sPath, iLastSlash + 1)
        End If
    End If
End Function


'#########################################################
'# Trims the passed sString up to vbNull
'#########################################################
Private Function TrimNull(ByVal sString As String) As String
    Dim lIndex As Long

        '#### Default the return value and determine if there is a vbNull in sString
    TrimNull = sString
    lIndex = InStr(1, TrimNull, Chr(0), vbBinaryCompare)

        '#### If a vbNull was present, trim up to it and return
    If (lIndex > 0) Then TrimNull = Left(TrimNull, lIndex - 1)
End Function
