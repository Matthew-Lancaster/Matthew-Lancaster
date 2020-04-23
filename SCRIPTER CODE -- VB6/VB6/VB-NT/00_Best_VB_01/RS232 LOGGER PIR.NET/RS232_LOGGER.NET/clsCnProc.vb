Option Strict Off
Option Explicit On
Friend Class clsCnProc
	
	' Example
	' Form or Module
	' Public cProcesses As New clsCnProc
	
	' VAR = cProcesses.Convert(CPHwnd, Otss, cnFromProcessID Or cnTohWnd)
	' VAR = cProcesses.Convert(CPHwnd, Otss, cnFromhWnd Or cnToProcessID)
	
	' SEEMS TO BE ERROR USING HWND TO EXE IT WANTS PROCESSID
	' HAVE TO RUN TWO ROUTINE
	' VAR = cProcesses.Convert(CPHwnd, Otss, cnFromhWnd Or cnToEXE)
	
	' SEE EXAMPLE
	' VAR = cProcesses.Convert(CPHwnd, Otss, cnFromhWnd Or cnToProcessID)
	' VAR = cProcesses.Convert(CPHwnd, Otss, cnFromProcessID Or cnToEXE)
	' VAR = cProcesses.Convert(CPHwnd, Otss, cnFromProcessID Or cnTohWnd)
	
	
	' VAR = GetEXEID(ByRef lRunningID As Long, ByRef sRunningEXE As String) As Boolean
	
	
	'PID HAS TO BE -1
	'VAR = cProcesses.GetEXEID(-1, "C:\Program Files\KatMouse\Kat.EXE ")"
	'GetFileName -- Modified to get whole exe name not trim to space
	
	'VAR = cProcesses.Process_Kill(PID)
	
	
	'#### eNum used with Convert()
	'Public Enum cnProcessConv
	'    cnFromClass = 1
	'    cnFromEXE = 2
	'    cnFromhWnd = 4
	'    cnFromProcessID = 8
	'    cnFromTitle = 16
	'    cnToClass = 32
	'    cnToEXE = 64
	'    cnTohWnd = 128
	'    cnToProcessID = 256
	'    cnToTitle = 512
	'End Enum
	
	
	
	
	
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
	Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Integer) As Integer
	Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Integer) As Integer
	Private Declare Function GetWindowLong Lib "user32"  Alias "GetWindowLongA"(ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
	' PRIVATE Constants for GetWindowLong API declaration
	Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Integer, ByVal bInheritHandle As Integer, ByVal dwProcId As Integer) As Integer
	Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Integer, ByRef lphModule As Integer, ByVal cb As Integer, ByRef cbNeeded As Integer) As Integer
	Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Integer, ByVal hModule As Integer, ByVal ModuleName As String, ByVal nSize As Integer) As Integer
	Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Integer, ByVal cb As Integer, ByRef cbNeeded As Integer) As Integer
	
	Private Declare Function GetModuleFileNameEx Lib "psapi.dll"  Alias "GetModuleFileNameExA"(ByVal hProcess As Integer, ByVal hModule As Integer, ByVal lpFileName As String, ByVal nSize As Integer) As Integer
	
	Private Const GWL_STYLE As Short = (-16)
	Private Const GWL_EXSTYLE As Short = (-20)
	
	' PRIVATE Constants for window styles
	Private Const WS_BORDER As Integer = &H800000
	Private Const WS_CAPTION As Integer = &HC00000
	Private Const WS_CHILD As Integer = &H40000000
	Private Const WS_CLIPCHILDREN As Integer = &H2000000
	Private Const WS_CLIPSIBLINGS As Integer = &H4000000
	Private Const WS_DLGFRAME As Integer = &H400000
	Private Const WS_GROUP As Integer = &H20000
	Private Const WS_HSCROLL As Integer = &H100000
	Private Const WS_MAXIMIZEBOX As Integer = &H10000
	Private Const WS_MINIMIZEBOX As Integer = &H20000
	Private Const WS_SYSMENU As Integer = &H80000
	Private Const WS_POPUP As Integer = &H80000000
	Private Const WS_POPUPWINDOW As Boolean = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
	Private Const WS_TABSTOP As Integer = &H10000
	Private Const WS_THICKFRAME As Integer = &H40000
	Private Const WS_SIZEBOX As Integer = WS_THICKFRAME
	Private Const WS_VISIBLE As Integer = &H10000000
	Private Const WS_VSCROLL As Integer = &H200000
	Private Const WM_CLOSE As Integer = &H10
	
	' Private constants for extended window styles
	Private Const WS_EX_CLIENTEDGE As Integer = &H200
	Private Const WS_EX_STATICEDGE As Integer = &H20000
	
	Private Const DELETE As Integer = &H10000
	Private Const READ_CONTROL As Integer = &H20000
	Private Const WRITE_DAC As Integer = &H40000
	Private Const WRITE_OWNER As Integer = &H80000
	Private Const SYNCHRONIZE As Integer = &H100000
	Private Const STANDARD_RIGHTS_REQUIRED As Integer = &HF0000
	Private Const STANDARD_RIGHTS_READ As Integer = READ_CONTROL
	Private Const STANDARD_RIGHTS_WRITE As Integer = READ_CONTROL
	Private Const STANDARD_RIGHTS_EXECUTE As Integer = READ_CONTROL
	Private Const STANDARD_RIGHTS_ALL As Integer = &H1F0000
	Private Const SPECIFIC_RIGHTS_ALL As Integer = &HFFFF
	Private Const GENERIC_READ As Integer = &H80000000
	Private Const GENERIC_WRITE As Integer = &H40000000
	Private Const GENERIC_EXECUTE As Integer = &H20000000
	Private Const GENERIC_ALL As Integer = &H10000000
	
	Private Const PROCESS_TERMINATE As Integer = &H1
	Private Const PROCESS_CREATE_THREAD As Integer = &H2
	Private Const PROCESS_SET_SESSIONID As Integer = &H4
	Private Const PROCESS_VM_OPERATION As Integer = &H8
	Private Const PROCESS_VM_READ As Integer = &H10
	Private Const PROCESS_VM_WRITE As Integer = &H20
	Private Const PROCESS_DUP_HANDLE As Integer = &H40
	Private Const PROCESS_CREATE_PROCESS As Integer = &H80
	Private Const PROCESS_SET_QUOTA As Integer = &H100
	Private Const PROCESS_SET_INFORMATION As Integer = &H200
	Private Const PROCESS_QUERY_INFORMATION As Integer = &H400
	Private Const PROCESS_ALL_ACCESS As Integer = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
	
	Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Integer, ByVal uExitCode As Integer) As Integer
	
	
	Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Integer)
	Private Const MAX_PATH As Short = 260
	
	Private sEXEName As String
	Private hProcess As Integer
	Private lProcessID As Integer
	Private bIsNT4 As Boolean
	Private bIsOpen As Boolean
	
	
	Private Declare Function FindWindowDLL Lib "user32"  Alias "FindWindowA"(ByVal lpClassName As Integer, ByVal lpWindowName As Integer) As Integer
	
	'#########################################################
	'# Functions/Consts/Types used for GetVersionEx()
	'#########################################################
	'UPGRADE_WARNING: Structure OSVERSIONINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetVersionEx Lib "kernel32"  Alias "GetVersionExA"(ByRef lpVersionInformation As OSVERSIONINFO) As Integer
	'Private Const VER_PLATFORM_WIN32s = 0
	'Private Const VER_PLATFORM_WIN32_WINDOWS = 1
	Private Const VER_PLATFORM_WIN32_NT As Short = 2
	Private Structure OSVERSIONINFO
		Dim OSVSize As Integer
		Dim dwVerMajor As Integer
		Dim dwVerMinor As Integer
		Dim dwBuildNumber As Integer '#### NT: Build Number, 9x: High-Order has Major/Minor ver, Low-Order has build
		Dim PlatformID As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(128),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=128)> Public szCSDVersion() As Char '#### NT: ie- "Service Pack 3", 9x: 'arbitrary additional information'
	End Structure
	
	
	'#########################################################
	'# Win32 (non-NT4) specific definitions
	'#########################################################
	Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Integer, ByRef lProcessID As Integer) As Integer
	'UPGRADE_WARNING: Structure PROCESSENTRY32 may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function ProcessFirst Lib "kernel32"  Alias "Process32First"(ByVal hSnapShot As Integer, ByRef uProcess As PROCESSENTRY32) As Integer
	'UPGRADE_WARNING: Structure PROCESSENTRY32 may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function ProcessNext Lib "kernel32"  Alias "Process32Next"(ByVal hSnapShot As Integer, ByRef uProcess As PROCESSENTRY32) As Integer
	Private Const TH32CS_SNAPPROCESS As Integer = 2
	Private Structure PROCESSENTRY32
		Dim dwSize As Integer
		Dim cntUsage As Integer
		Dim th32ProcessID As Integer
		Dim th32DefaultHeapID As Integer
		Dim th32ModuleID As Integer
		Dim cntThreads As Integer
		Dim th32ParentProcessID As Integer
		Dim pcPriClassBase As Integer
		Dim dwFlags As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(MAX_PATH),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=MAX_PATH)> Public szExeFile() As Char
	End Structure
	
	'#### Required private data members
	Private uProcess As PROCESSENTRY32
	Private hSnapShot As Integer
	
	
	'#########################################################
	'# Windows NT4 specific definitions
	'#    NOTE: Remember to distribute the PSAPI.DLL file.
	'#########################################################
	'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
	'Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
	'Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
	'Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
	
	'Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
	'Private Const SYNCHRONIZE = &H100000
	'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
	'Private Const PROCESS_ALL_ACCESS = &H1F0FFF
	
	'#### Required private data members
	'UPGRADE_WARNING: Lower bound of array Modules was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Private Modules(200) As Integer
	Private lProcessIDs() As Integer
	Private lcbNeeded As Integer
	Private lRetVal As Integer
	Private i As Integer
	
	
	'#########################################################
	'# Convert() releated definitions
	'#########################################################
	'#### Functions/Consts used for GetHWnd() (part of Convert())
	Private Declare Function GetDesktopWindow Lib "user32" () As Integer
	Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
	Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Integer, ByRef lpdwProcessId As Integer) As Integer
	Private Declare Function GetWindowText Lib "user32"  Alias "GetWindowTextA"(ByVal hWnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer
	Private Declare Function GetClassName Lib "user32"  Alias "GetClassNameA"(ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
	Private Const GW_HWNDNEXT As Short = 2
	Private Const GW_CHILD As Short = 5
	
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
	
	Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
	'Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
	'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
	'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
	Private Declare Function GetWindowTextLength Lib "user32"  Alias "GetWindowTextLengthA"(ByVal hWnd As Integer) As Integer
	
	Private Const GW_OWNER As Short = 4
	Private Const WS_EX_TOOLWINDOW As Integer = &H80
	Private Const WS_EX_APPWINDOW As Integer = &H40000
	Private Const GA_ROOT As Short = 2
	
	Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Integer, ByVal gaFlags As Integer) As Integer
	Private Declare Function FindWindow Lib "user32"  Alias "FindWindowA"(ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
	
	
	
	
	Private Function GetWindowTitle(ByVal hWnd As Integer) As String
		Dim L As Integer
		Dim S As String
		L = GetWindowTextLength(hWnd)
		S = Space(L + 1)
		GetWindowText(hWnd, S, L + 1)
		GetWindowTitle = Left(S, L)
	End Function
	
	'#####################################################################################
	'# Class Functions
	'#####################################################################################
	'#########################################################
	'# Class constructor to init private variables
	'#########################################################
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
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
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'#########################################################
	'# Class destructor to kill private variables
	'#########################################################
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'#### If we're running under a system that supports CreateToolhelp32Snapshot()
		If (Not bIsNT4) Then
			'#### Close the hSnapShot handle
			Call CloseHandle(hSnapShot)
		End If
		
		'#### Set bIsOpen
		bIsOpen = False
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	'#########################################################
	'# Get/Let Properties
	'#########################################################
	'UPGRADE_NOTE: Class was upgraded to Class_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public ReadOnly Property Class_Renamed() As String
		Get
			'#### If conversion is not successful, reset the return value
			If (Not Convert(lProcessID, Class_Renamed, cnProcessConv.cnFromProcessID Or cnProcessConv.cnToClass)) Then
				Class_Renamed = ""
			End If
		End Get
	End Property
	Public ReadOnly Property EXE() As String
		Get
			EXE = sEXEName
		End Get
	End Property
	Public ReadOnly Property hWnd() As Integer
		Get
			'#### If conversion is not successful, reset the return value
			If (Not Convert(lProcessID, hWnd, cnProcessConv.cnFromProcessID Or cnProcessConv.cnTohWnd)) Then
				hWnd = -1
			End If
		End Get
	End Property
	Public ReadOnly Property ProcessID() As Integer
		Get
			ProcessID = lProcessID
		End Get
	End Property
	Public ReadOnly Property title() As String
		Get
			'#### If conversion is not successful, reset the return value
			If (Not Convert(lProcessID, title, cnProcessConv.cnFromProcessID Or cnProcessConv.cnToTitle)) Then
				title = ""
			End If
		End Get
	End Property
	
	
	
	'#####################################################################################
	'# Public subs/functions
	'#####################################################################################
	'#########################################################
	'# Opens the processes and moves to the first ProcessID
	'#########################################################
	Public Function OpenProcesses() As Boolean
		'#### If we're running under WinNT4
		Dim lcb As Integer
		If (bIsNT4) Then
			
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
			hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
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
	
	'#####################################################################################
	'# Public subs/functions
	'#####################################################################################
	'#########################################################
	'# Opens the processes and moves to the first ProcessID
	'# WANT THE FULL EXE PATH
	'#########################################################
	Public Function OpenProcessesEXEPATH() As Boolean
		'#### If we're running under WinNT4
		Dim sEXEUsage As Object
		Dim lcb As Integer
		If (bIsNT4) Then
			
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
			OpenProcessesEXEPATH = NextProcess()
			
			'#### Else we're running under a system that supports CreateToolhelp32Snapshot()
		Else
			'#### Setup hSnapShot, begin to setup the uProcess UDT
			hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
			uProcess.dwSize = Len(uProcess)
			
			'#### If hSnapShot was successfully setup
			If (hSnapShot <> 0) Then
				'#### Find the first hProcess and set the return value
				hProcess = ProcessFirst(hSnapShot, uProcess)
				OpenProcessesEXEPATH = (hProcess <> 0)
				
				'#### If a valid hProcess was found, set sEXEName and lProcessID
				If (OpenProcesses) Then
					sEXEName = TrimNull(uProcess.szExeFile)
					'UPGRADE_WARNING: Couldn't resolve default property of object sEXEUsage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sEXEUsage = uProcess.cntUsage
					'UPGRADE_WARNING: Couldn't resolve default property of object sEXEUsage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If sEXEUsage > 0 Then Stop
					
					lProcessID = uProcess.th32ProcessID
				End If
				
				'#### Else hSnapShot was not successfully setup, so error out
			Else
				GoTo OpenProcesses_Error
			End If
		End If
		
		'#### Set bIsOpen
		bIsOpen = OpenProcessesEXEPATH
		Exit Function
		
OpenProcesses_Error: 
		OpenProcessesEXEPATH = False
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
		Call Class_Terminate_Renamed()
		Call Class_Initialize_Renamed()
	End Sub
	
	
	Public Function Get_PID_From_HWND(ByRef InputData As Object, ByRef OutputData As Object) As Boolean
		
		Dim sClass As String
		Dim sTitle As String
		Dim sEXE As String
		Dim lHwnd As Integer
		Dim lProcessID As Integer
		
		'#### Init the vars to invalid values for an hWnd/ProcessID
		lHwnd = -1
		lProcessID = -1
		
		'#### Else if we are converting from an hWnd
		If (IsNumeric(InputData)) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object InputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lHwnd = InputData
		Else
			GoTo Convert_Error
		End If
		
		'#### Else if we are converting to a ProcessID
		'#### If lProcessID is not already set, determine its value now then set the OutputData
		If (lProcessID = -1) Then
			If (Not GetHwnd(lProcessID, lHwnd, sTitle, sClass)) Then GoTo Convert_Error
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object OutputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		OutputData = lProcessID
		
		'#### If we got here, all is well
		Get_PID_From_HWND = True
		Exit Function
		
		
		'#### If we end up here some sort of error occurred, so return false
Convert_Error: 
		Get_PID_From_HWND = False
	End Function
	
	
	'#########################################################
	'# Converts information about a process from one form to another (hWnd, ProcessID, and RunningEXE)
	'#########################################################
	Public Function Convert(ByRef InputData As Object, ByRef OutputData As Object, ByVal eConversionType As cnProcessConv) As Boolean
		Dim sClass As String
		Dim sTitle As String
		Dim sEXE As String
		Dim lHwnd As Integer
		Dim lProcessID As Integer
		
		'#### Init the vars to invalid values for an hWnd/ProcessID
		lHwnd = -1
		lProcessID = -1
		
		'#### If we are converting from a Class name
		If (eConversionType And cnProcessConv.cnFromClass) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object InputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sClass = InputData
			
			'#### Else if we are converting from an EXE name
		ElseIf (eConversionType And cnProcessConv.cnFromEXE) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object InputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sEXE = InputData
			
			'#### If we're not converting from EXE to EXE then we'll need the lProcessID
			If (Not CBool(eConversionType And cnProcessConv.cnToEXE)) Then
				If (Not GetEXEID(lProcessID, sEXE)) Then GoTo Convert_Error
			End If
			
			'#### Else if we are converting from an hWnd
		ElseIf (eConversionType And cnProcessConv.cnFromhWnd) Then 
			If (IsNumeric(InputData)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object InputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lHwnd = InputData
			Else
				GoTo Convert_Error
			End If
			
			'#### Else if we are converting from a ProcessID
		ElseIf (eConversionType And cnProcessConv.cnFromProcessID) Then 
			If (IsNumeric(InputData)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object InputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lProcessID = InputData
			Else
				GoTo Convert_Error
			End If
			
			'#### Else if we are converting from a Title
		ElseIf (eConversionType And cnProcessConv.cnFromTitle) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object InputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sTitle = InputData
			
			'#### Else we are missing a valid 'from'
		Else
			GoTo Convert_Error
		End If
		
		'#### If we are converting to a Class name
		If (eConversionType And cnProcessConv.cnToClass) Then
			'#### If sClass is not already set, determine its value now then set the OutputData
			If (Len(sClass) = 0) Then
				If (Not GetHwnd(lProcessID, lHwnd, sTitle, sClass)) Then GoTo Convert_Error
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object OutputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutputData = sClass
			
			'SEEMS TO BE ERROR USING HWND IT WANTS PROCESSID
			'HAVE TO RUN TWO ROUTINE
			'#### Else if we are converting to an EXE name
		ElseIf (eConversionType And cnProcessConv.cnToEXE) Then 
			'#### If sEXE is not already set, determine its value now then set the OutputData
			If (Len(sEXE) = 0) Then
				
				If (Not GetEXEID(lProcessID, sEXE)) Then GoTo Convert_Error
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object OutputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutputData = sEXE
			
			'#### Else if we are converting to an hWnd
		ElseIf (eConversionType And cnProcessConv.cnTohWnd) Then 
			'#### If lhWnd is not already set, determine its value now then set the OutputData
			If (lHwnd = -1) Then
				If (Not GetHwnd(lProcessID, lHwnd, sTitle, sClass)) Then GoTo Convert_Error
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object OutputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutputData = lHwnd
			
			'#### Else if we are converting to a ProcessID
		ElseIf (eConversionType And cnProcessConv.cnToProcessID) Then 
			'#### If lProcessID is not already set, determine its value now then set the OutputData
			If (lProcessID = -1) Then
				If (Not GetHwnd(lProcessID, lHwnd, sTitle, sClass)) Then GoTo Convert_Error
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object OutputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutputData = lProcessID
			
			'#### Else if we are converting to a Title
		ElseIf (eConversionType And cnProcessConv.cnToTitle) Then 
			'#### If sTitle is not already set, determine its value now then set the OutputData
			If (Len(sTitle) = 0) Then
				If (Not GetHwnd(lProcessID, lHwnd, sTitle, sClass)) Then GoTo Convert_Error
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object OutputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
	
	
	'#########################################################
	'# Converts information about a process from one form to another (hWnd, ProcessID, and RunningEXE)
	'#########################################################
	Public Function Convert_2(ByRef InputData As Object, ByRef OutputData As Object, ByVal eConversionType As cnProcessConv) As Boolean
		Dim sClass As String
		Dim sTitle As String
		Dim sEXE As String
		Dim lHwnd As Integer
		Dim lProcessID As Integer
		
		'#### Init the vars to invalid values for an hWnd/ProcessID
		lHwnd = -1
		lProcessID = -1
		
		'#### If we are converting from a Class name
		If (eConversionType And cnProcessConv.cnFromClass) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object InputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sClass = InputData
			
			'#### Else if we are converting from an EXE name
		ElseIf (eConversionType And cnProcessConv.cnFromEXE) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object InputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sEXE = InputData
			
			'#### If we're not converting from EXE to EXE then we'll need the lProcessID
			If (Not CBool(eConversionType And cnProcessConv.cnToEXE)) Then
				If (Not GetEXEID(lProcessID, sEXE)) Then GoTo Convert_Error
			End If
			
			'#### Else if we are converting from an hWnd
		ElseIf (eConversionType And cnProcessConv.cnFromhWnd) Then 
			If (IsNumeric(InputData)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object InputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lHwnd = InputData
			Else
				GoTo Convert_Error
			End If
			
			'#### Else if we are converting from a ProcessID
		ElseIf (eConversionType And cnProcessConv.cnFromProcessID) Then 
			If (IsNumeric(InputData)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object InputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lProcessID = InputData
			Else
				GoTo Convert_Error
			End If
			
			'#### Else if we are converting from a Title
		ElseIf (eConversionType And cnProcessConv.cnFromTitle) Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object InputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sTitle = InputData
			
			'#### Else we are missing a valid 'from'
		Else
			GoTo Convert_Error
		End If
		
		'#### If we are converting to a Class name
		If (eConversionType And cnProcessConv.cnToClass) Then
			'#### If sClass is not already set, determine its value now then set the OutputData
			If (Len(sClass) = 0) Then
				If (Not GetHwnd(lProcessID, lHwnd, sTitle, sClass)) Then GoTo Convert_Error
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object OutputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutputData = sClass
			
			'SEEMS TO BE ERROR USING HWND IT WANTS PROCESSID
			'HAVE TO RUN TWO ROUTINE
			'#### Else if we are converting to an EXE name
		ElseIf (eConversionType And cnProcessConv.cnToEXE) Then 
			'#### If sEXE is not already set, determine its value now then set the OutputData
			If (Len(sEXE) = 0) Then
				If (Not GetEXEID(lProcessID, sEXE)) Then GoTo Convert_Error
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object OutputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutputData = sEXE
			
			'#### Else if we are converting to an hWnd
		ElseIf (eConversionType And cnProcessConv.cnTohWnd) Then 
			'#### If lhWnd is not already set, determine its value now then set the OutputData
			If (lHwnd = -1) Then
				' If (Not GetHWndFromProcess(lProcessID, lHwnd, sTitle, sClass)) Then GoTo Convert_Error
				If (Not GetHwnd_From_PID(lProcessID)) Then GoTo Convert_Error
				
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object OutputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutputData = lHwnd
			
			'#### Else if we are converting to a ProcessID
		ElseIf (eConversionType And cnProcessConv.cnToProcessID) Then 
			'#### If lProcessID is not already set, determine its value now then set the OutputData
			If (lProcessID = -1) Then
				If (Not GetHwnd(lProcessID, lHwnd, sTitle, sClass)) Then GoTo Convert_Error
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object OutputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutputData = lProcessID
			
			'#### Else if we are converting to a Title
		ElseIf (eConversionType And cnProcessConv.cnToTitle) Then 
			'#### If sTitle is not already set, determine its value now then set the OutputData
			If (Len(sTitle) = 0) Then
				If (Not GetHwnd(lProcessID, lHwnd, sTitle, sClass)) Then GoTo Convert_Error
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object OutputData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutputData = sTitle
			
			'#### Else we are missing a valid 'to'
		Else
			GoTo Convert_Error
		End If
		
		'#### If we got here, all is well
		Convert_2 = True
		Exit Function
		
		'#### If we end up here some sort of error occurred, so return false
Convert_Error: 
		Convert_2 = False
	End Function
	
	
	
	'#####################################################################################
	'# Private subs/functions
	'#####################################################################################
	'#########################################################
	'# Returns a windows EXE and ProcessID from the passed EXE or ProcessID
	'# USE WILD CARD TO COUNT HOW MANY EXE RUNNING
	'#########################################################
	Public Function GetEXEID_WILDCARD(ByRef lRunningID As Integer) As Boolean
		'#### Default the return value
		Dim sRunningEXE_ENTRY As Object
		Dim CountEXE As Object
		GetEXEID_WILDCARD = False
		'sRunningEXE_ENTRY = sRunningEXE
		'#### If we're able to successfully open the processes
		If (OpenProcesses()) Then
			'#### Get the name of the EXE
			'sRunningEXE = UCase(GetFileName(sRunningEXE))
			
			'#### Do while we still have processes
			Do 
				'#### If the ProcessIDs match
				'            If (lProcessID = lRunningID) Then
				'               'lRunningID = lProcessID
				'                sRunningEXE = TrimNull(sEXEName)
				'                GetEXEID = True
				'                GoTo GetEXEID_Close
				'
				'                '#### Else if the EXE names match
				'            ElseIf (InStr(1, UCase(sEXEName), sRunningEXE, vbBinaryCompare) > 0) Then
				'                lRunningID = lProcessID
				'                sRunningEXE = TrimNull(sEXEName)
				'                GetEXEID = True
				'                GoTo GetEXEID_Close
				'            End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object CountEXE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CountEXE = CountEXE + 1
			Loop While (NextProcess())
			
			'#### If we make it here, the lRunningID/sRunningEXE was not found
			'lRunningID = -1
			'sRunningEXE = ""
			
			'UPGRADE_WARNING: Couldn't resolve default property of object CountEXE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lRunningID = CountEXE
			
			
			'#### Else the processes were not successfully opened
		Else
			lRunningID = -1
			'sRunningEXE = ""
		End If
		
GetEXEID_Close: 
		'#### Close the class
		Call CloseProcesses()
		
		'If sRunningEXE = "" Then sRunningEXE = sRunningEXE_ENTRY
		
	End Function
	
	
	'#####################################################################################
	'# Private subs/functions
	'#####################################################################################
	'#########################################################
	'# Returns a windows EXE and ProcessID from the passed EXE or ProcessID
	'# USE WILD CARD TO COUNT HOW MANY EXE RUNNING
	'# AND CREATE A SCRIPT OF PROSESSES
	'#########################################################
	Public Function GetEXEID_WILDCARD_PROCESS_SCRIPT(ByRef lRunningID As Integer, ByRef PROCESS_SCRIPT As String) As Boolean
		'#### Default the return value
		Dim sRunningEXE_ENTRY As Object
		Dim CountEXE As Object
		GetEXEID_WILDCARD_PROCESS_SCRIPT = False
		PROCESS_SCRIPT = ""
		'sRunningEXE_ENTRY = sRunningEXE
		'#### If we're able to successfully open the processes
		If (OpenProcessesEXEPATH()) Then
			'#### Get the name of the EXE
			'        sRunningEXE = UCase(GetFileName(sRunningEXE))
			
			
			'#### Do while we still have processes
			Do 
				'#### If the ProcessIDs match
				'            If (lProcessID = lRunningID) Then
				'               'lRunningID = lProcessID
				'                sRunningEXE = TrimNull(sEXEName)
				PROCESS_SCRIPT = PROCESS_SCRIPT & Str(lProcessID) & "-" & TrimNull(sEXEName) & vbCrLf
				'                GetEXEID = True
				'                GoTo GetEXEID_Close
				'
				'                '#### Else if the EXE names match
				'            ElseIf (InStr(1, UCase(sEXEName), sRunningEXE, vbBinaryCompare) > 0) Then
				'                lRunningID = lProcessID
				'                sRunningEXE = TrimNull(sEXEName)
				'                GetEXEID = True
				'                GoTo GetEXEID_Close
				'            End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object CountEXE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CountEXE = CountEXE + 1
			Loop While (NextProcess())
			
			'#### If we make it here, the lRunningID/sRunningEXE was not found
			'lRunningID = -1
			'sRunningEXE = ""
			
			'UPGRADE_WARNING: Couldn't resolve default property of object CountEXE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lRunningID = CountEXE
			
			
			'#### Else the processes were not successfully opened
		Else
			lRunningID = -1
			'sRunningEXE = ""
		End If
		
GetEXEID_Close: 
		'#### Close the class
		Call CloseProcesses()
		
		'If sRunningEXE = "" Then sRunningEXE = sRunningEXE_ENTRY
		
	End Function
	
	
	
	'#####################################################################################
	'# Private subs/functions
	'#####################################################################################
	'#########################################################
	'# Returns a windows EXE and ProcessID from the passed EXE or ProcessID
	'#########################################################
	Public Function GetEXEID(ByRef lRunningID As Integer, ByRef sRunningEXE As String) As Boolean
		'#### Default the return value
		Dim sRunningEXE_ENTRY As Object
		GetEXEID = False
		'UPGRADE_WARNING: Couldn't resolve default property of object sRunningEXE_ENTRY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sRunningEXE_ENTRY = sRunningEXE
		'#### If we're able to successfully open the processes
		If (OpenProcesses()) Then
			'#### Get the name of the EXE
			If sRunningEXE <> "" Then
				sRunningEXE = UCase(GetFileName(sRunningEXE))
			End If
			
			'#### Do while we still have processes
			Do 
				'#### If the ProcessIDs match
				If (lProcessID = lRunningID) Then
					'lRunningID = lProcessID
					sRunningEXE = TrimNull(sEXEName)
					GetEXEID = True
					GoTo GetEXEID_Close
					
					'MATT.LAN@BTINTERNET.COM -- FIX BUG
					'MIST NOT CHECK EXE IF VOID
					'#### Else if the EXE names match
				ElseIf sRunningEXE <> "" And (InStr(1, UCase(sEXEName), sRunningEXE, CompareMethod.Binary) > 0) Then 
					lRunningID = lProcessID
					sRunningEXE = TrimNull(sEXEName)
					GetEXEID = True
					' Exit Do
					GoTo GetEXEID_Close
				End If
			Loop While (NextProcess())
			
			'#### If we make it here, the lRunningID/sRunningEXE was not found
			lRunningID = -1
			sRunningEXE = ""
			
			'#### Else the processes were not successfully opened
		Else
			lRunningID = -1
			'sRunningEXE = ""
		End If
		
GetEXEID_Close: 
		'#### Close the class
		Call CloseProcesses()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object sRunningEXE_ENTRY. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If sRunningEXE = "" Then sRunningEXE = sRunningEXE_ENTRY
		
	End Function
	
	' ---------------------------------------------------------
	' CREDIT DUE BUT DIDN'T DELCARE ANYTHING
	' ----
	' getting hwnd from process-VBForums
	' http://www.vbforums.com/showthread.php?561413-getting-hwnd-from-process
	' ----
	' [ Tuesday 11:49:00 Am_23 October 2018 ]
	' ---------------------------------------------------------
	Public Function GetHwnd_From_PID(ByVal PID As Integer) As Integer
		' Return the window handle for an instance handle.
		Dim lHwnd As Integer
		Dim test_pid As Integer
		Dim Thread_ID As Integer
		Dim lExStyle As Integer
		Dim bNoOwner As Boolean
		'Private Const GW_OWNER = 4
		'Private Const WS_EX_TOOLWINDOW = &H80
		'Private Const WS_EX_APPWINDOW = &H40000
		'Private Const GA_ROOT = 2
		'Private Declare Function GetAncestor Lib "user32.dll" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
		
		' Get the first window handle.
		lHwnd = FindWindow(vbNullString, vbNullString)
		' Loop until we find the target or we run out
		' of windows.
		Do While lHwnd <> 0
			' check if window is visible or not
			If IsWindowVisible(lHwnd) Then
				' See if this window has a parent. If not,
				' it is a top-level window.
				If GetParent(lHwnd) = 0 Then
					bNoOwner = (GetWindow(lHwnd, GW_OWNER) = 0)
					lExStyle = GetWindowLong(lHwnd, GWL_EXSTYLE)
					
					If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
						' This is a top-level window. See if
						' it has the target instance handle.
						Thread_ID = GetWindowThreadProcessId(lHwnd, test_pid)
						' Debug.Print test_pid, Thread_ID
						If test_pid = PID Then
							lHwnd = GetAncestor(lHwnd, GA_ROOT)
							'check if a top window
							
							'If IsTopWIndow(lHwnd) Then
							GetHwnd_From_PID = lHwnd
							'End If
							Exit Function
						End If
					End If
				End If
			End If
			' Examine the next window.
			lHwnd = GetWindow(lHwnd, GW_HWNDNEXT)
		Loop 
	End Function
	
	'----
	'[RESOLVED] Get exe name from Process ID fails Help-VBForums
	'http://www.vbforums.com/showthread.php?763427-RESOLVED-Get-exe-name-from-Process-ID-fails-Help
	'----
	Public Function GetEXE_Path_From_ProcessID(ByRef PID As Integer) As String
		Dim Process As Object
		Dim lzFile As String
		On Error GoTo GetFileErr
		
		'Scan process and find pid then return the path and exe name
		'UPGRADE_WARNING: Couldn't resolve default property of object GetObject().ExecQuery. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each Process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
			'UPGRADE_WARNING: Couldn't resolve default property of object Process.ProcessID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Debug.Print(CInt(Process.ProcessID))
			'UPGRADE_WARNING: Couldn't resolve default property of object Process.ProcessID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (PID = CInt(Process.ProcessID)) Then
				'Return exe path
				'UPGRADE_WARNING: Couldn't resolve default property of object Process.ExecutablePath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				lzFile = Process.ExecutablePath
				Exit For
			End If
		Next Process
		
		GetEXE_Path_From_ProcessID = lzFile
		Exit Function
GetFileErr: 
		GetEXE_Path_From_ProcessID = vbNullString
	End Function
	
	
	'#########################################################
	'# Returns a windows hWnd, ProcessID, Title and Class from the passed hWnd, ProcessID, Title or Class
	'#########################################################
	Private Function GetHWndFromProcess(ByRef hProcessID As Integer, ByRef hWnd As Integer, ByRef sTitle As String, ByRef sClass As String) As Boolean
		Dim sWinTitle As String
		Dim sClassName As String
		Dim hWndChild As Integer
		Dim hWndChildProcessID As Integer
		
		'#### Get the Desktop handle while getting the first child under the Desktop and default the return value
		hWndChild = GetWindow(GetDesktopWindow(), GW_CHILD)
		GetHWndFromProcess = False
		
		'#### While we still have hWndChild(en) to look at
		Do While (hWndChild <> 0)
			'#### Get the ThreadProcessID of the window
			Call GetWindowThreadProcessId(hWndChild, hWndChildProcessID)
			
			'#### Get the current window's title
			sWinTitle = Space(255)
			sWinTitle = Left(sWinTitle, GetWindowText(hWndChild, sWinTitle, 255))
			
			'#### Get the current window's class
			sClassName = Space(255)
			sClassName = Left(sClassName, GetClassName(hWndChild, sClassName, 255))
			
			'#### If we have a match with the hProcessID or hWnd, return the values
			If (hWndChildProcessID = hProcessID) Then
				hProcessID = hWndChildProcessID
				hWnd = hWndChild
				sTitle = sWinTitle
				sClass = sClassName
				GetHWndFromProcess = True
				'Exit Do
			End If
			
			'#### We've not yet found a match, so get the next hWndChild
			hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
		Loop 
	End Function
	
	
	
	
	'#########################################################
	'# Returns a windows hWnd, ProcessID, Title and Class from the passed hWnd, ProcessID, Title or Class
	'#########################################################
	Private Function GetHwnd(ByRef hProcessID As Integer, ByRef hWnd As Integer, ByRef sTitle As String, ByRef sClass As String) As Boolean
		Dim sWinTitle As String
		Dim sClassName As String
		Dim hWndChild As Integer
		Dim hWndChildProcessID As Integer
		
		'#### Get the Desktop handle while getting the first child under the Desktop and default the return value
		hWndChild = GetWindow(GetDesktopWindow(), GW_CHILD)
		GetHwnd = False
		
		'#### While we still have hWndChild(en) to look at
		Do While (hWndChild <> 0)
			'#### Get the ThreadProcessID of the window
			Call GetWindowThreadProcessId(hWndChild, hWndChildProcessID)
			
			'#### Get the current window's title
			sWinTitle = Space(255)
			sWinTitle = Left(sWinTitle, GetWindowText(hWndChild, sWinTitle, 255))
			
			'#### Get the current window's class
			sClassName = Space(255)
			sClassName = Left(sClassName, GetClassName(hWndChild, sClassName, 255))
			
			'#### If we have a match with the hProcessID or hWnd, return the values
			If (hWndChildProcessID = hProcessID Or hWndChild = hWnd) Then
				hProcessID = hWndChildProcessID
				hWnd = hWndChild
				sTitle = sWinTitle
				sClass = sClassName
				GetHwnd = True
				Exit Do
				
				'#### Else if sWinTitle has a value that is like sTitle, return the values
			ElseIf (Len(sWinTitle) > 0 And Len(sTitle) > 0) Then 
				If (sWinTitle Like sTitle) Then
					hProcessID = hWndChildProcessID
					hWnd = hWndChild
					sTitle = sWinTitle
					sClass = sClassName
					GetHwnd = True
					Exit Do
				End If
				
				'#### Else if sClassName has a value that is like sClass, return the values
			ElseIf (Len(sClassName) > 0 And Len(sClass) > 0) Then 
				If (sClassName Like sClass) Then
					hProcessID = hWndChildProcessID
					hWnd = hWndChild
					sTitle = sWinTitle
					sClass = sClassName
					GetHwnd = True
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
		Dim iLastSpace As Short
		Dim iLastSlash As Short
		
		'#### If sPath has a value, process it
		If (Len(sPath) > 0) Then
			'#### Deteremine the index of the last "\" and the following " " (if there is one)
			iLastSlash = InStrRev(sPath, "\", -1)
			iLastSpace = InStr(iLastSlash + 1, sPath, " ")
			
			'#### If a space was found in sPath, make sure we only remove the EXEs name and not any arguments
			If (iLastSpace > 0) Then
				'GetFileName = Mid(sPath, iLastSlash + 1, iLastSpace - iLastSlash - 1)
				
				'#### Else there were no arguments, so peal off the name
				
				'ABORT THIS CAUSING ERROR TRIM AWAY PART FILE EXE NAME
				GetFileName = Mid(sPath, iLastSlash + 1)
			Else
				GetFileName = Mid(sPath, iLastSlash + 1)
			End If
		End If
	End Function
	
	
	'#########################################################
	'# Trims the passed sString up to vbNull
	'#########################################################
	Private Function TrimNull(ByVal sString As String) As String
		Dim lIndex As Integer
		
		'#### Default the return value and determine if there is a vbNull in sString
		TrimNull = sString
		lIndex = InStr(1, TrimNull, Chr(0), CompareMethod.Binary)
		
		'#### If a vbNull was present, trim up to it and return
		If (lIndex > 0) Then TrimNull = Left(TrimNull, lIndex - 1)
	End Function
	
	
	
	
	
	
	Public Function Process_Kill(ByRef P_ID As Integer) As Integer
		'// Kill the wanted process
		
		'EXAMPLE ---- VAR = cProcesses.Process_Kill(PID)
		
		'MODIFIED TO ADD 201204182300 HOUR ---
		
		
		Dim hProcess As Integer
		Dim lExitCode As Integer
		Dim XI As Integer
		
		hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID)
		
		XI = TerminateProcess(hProcess, lExitCode) = False
		
		Call CloseHandle(hProcess)
		
	End Function
	
	
	
	
	
	Public Function FindHighestHandle() As Integer
		
		Dim test_pid, test_hwnd, test_thread_id As Integer
		
		Dim cText As String
		Dim wText As String
		Dim HighTest_hWnd As Integer
		
		'Find the first window
		test_hwnd = FindWindowDLL(0, 0)
		
		Do While test_hwnd <> 0
			
			wText = GetWindowTitle(test_hwnd)
			
			If test_hwnd > HighTest_hWnd Then
				HighTest_hWnd = test_hwnd
				wText = GetWindowTitle(test_hwnd)
			End If
			'retrieve the next window
			test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
			
		Loop 
		
		FindHighestHandle = HighTest_hWnd
		
		
		
		
		'Date 06-10-08
		'"[System Process]"
		'1593707806   -- "Program Manager"
		'This SysUptime 61 Days 0h 3m 54s 532 mil
		'Last SysUptime 29 Days 3h 1m 49s 203 mil
		
		
		
		
	End Function
	
	
	
	
	
	Public Function GetFileFromHwnd(ByRef lngHwnd As Object) As String
		
		'-----------------------------------------------------------
		'REPAIR DONE NOW GET FULL PATH OF EXE BY USE THE OPENPROCESS
		'COPY PASTE THE DECLARE AND DONE
		'Thu 16 February 2017 03:11:44----------
		'-----------------------------------------------------------
		'-----------------------------------------------------------
		
		'------------------------------------------------------------------
		'------------------------------------------------------------------
		'TEMPORARY AS FULL EXE WAS NOT WORKING FOR A BIT
		'------------------------------------------------------------------
		'MsgBox getfilefromhwnd(Me.hwnd)
		'Var = cProcesses.Convert(lngHwnd, vfile, cnFromhWnd Or cnToEXE)
		'Var = cProcesses.Convert(lngHwnd, PID, cnFromhWnd Or cnToProcessID)
		'Var = cProcesses.Convert(PID, vfile, cnFromProcessID Or cnToEXE)
		'GetFileFromHwnd = vfile
		'Exit Function
		'------------------------------------------------------------------
		
		'------------------------------------------------------------------
		Dim bla, lngProcess, hProcess, C As Integer
		Dim strFile As String
		Dim X As Object
		
		strFile = New String(Chr(0), 256)
		'UPGRADE_WARNING: Couldn't resolve default property of object lngHwnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		X = GetWindowThreadProcessId(lngHwnd, lngProcess)
		hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lngProcess)
		'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		X = EnumProcessModules(hProcess, bla, 4, C)
		C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
		GetFileFromHwnd = Left(strFile, C)
		'------------------------------------------------------------------
		
	End Function
End Class