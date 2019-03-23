VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRunAs"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'RunAs Class
'21 March 2001
'Version 1.0.2
'Author Darran Jones www.dalsoft.co.uk
'Support at Support@dalsoft.co.uk
'Suggestions to RunAsSuggestions@dalsoft.co.uk
'Please read readme.txt before using this class
'Please note by using this class you agree _
'to the disclaimer in readme.txt

Option Explicit
'Run as Constants
Private Const LOGON_WITH_PROFILE = &H1&
Private Const LOGON_NETCREDENTIALS_ONLY = &H2&
Private Const CREATE_DEFAULT_ERROR_MODE = &H4000000
Private Const CREATE_NEW_CONSOLE = &H10&
Private Const CREATE_NEW_PROCESS_GROUP = &H200&
Private Const CREATE_SEPARATE_WOW_VDM = &H800&
Private Const CREATE_SUSPENDED = &H4&
Private Const CREATE_UNICODE_ENVIRONMENT = &H400&
Private Const ABOVE_NORMAL_PRIORITY_CLASS = &H8000&
Private Const BELOW_NORMAL_PRIORITY_CLASS = &H4000&
Private Const HIGH_PRIORITY_CLASS = &H80&
Private Const IDLE_PRIORITY_CLASS = &H40&
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const REALTIME_PRIORITY_CLASS = &H100&
'Process type lib
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
'start up type lib
Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
'API call for Run As
Private Declare Function CreateProcessWithLogonW Lib "Advapi32" (ByVal lpUsername As Long, ByVal lpDomain As Long, ByVal lpPassword As Long, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInfo As PROCESS_INFORMATION) As Long
'API call for closing a handle
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'local variable(s) to hold property value(s)
'and for use in internal methods
Private mvarsUserName As String 'local
Private mvarsPassword As String 'local
Private mvarsCommand As String 'local
Private mvarsDomain As String 'local

Public Property Let sUserName(ByVal vData As String)
    'Set the local variable to the passed in value of the let.
    mvarsUserName = vData
End Property

Public Property Get sUserName() As String
    'Set the get value to the local variable value
    sUserName = mvarsUserName
End Property

Public Property Let sPassword(ByVal vData As String)
    mvarsPassword = vData
End Property

Public Property Get sPassword() As String
    sPassword = mvarsPassword
End Property

Public Property Let sCommand(ByVal vData As String)
    mvarsCommand = vData
End Property

Public Property Get sCommand() As String
   sCommand = mvarsCommand
End Property

Public Property Let sDomain(ByVal vData As String)
    mvarsDomain = vData
End Property

Public Property Get sDomain() As String
   sDomain = mvarsDomain
End Property

Public Sub RunAs()
'Run As method
Dim lpUsername As String, lpDomain As String, lpPassword As String, lpApplicationName As String
Dim lpCommandLine As String, lpCurrentDirectory As String
Dim StartInfo As STARTUPINFO, ProcessInfo As PROCESS_INFORMATION
         
         lpUsername = mvarsUserName
         lpDomain = mvarsDomain
         lpPassword = mvarsPassword
         lpApplicationName = vbNullString
         lpCommandLine = mvarsCommand
         lpCurrentDirectory = vbNullString 'use standard directory
         StartInfo.cb = LenB(StartInfo) 'initialize structure
         StartInfo.dwFlags = 0&
         CreateProcessWithLogonW StrPtr(lpUsername), StrPtr(lpDomain), StrPtr(lpPassword), LOGON_WITH_PROFILE, StrPtr(lpApplicationName), StrPtr(lpCommandLine), CREATE_DEFAULT_ERROR_MODE Or CREATE_NEW_CONSOLE Or CREATE_NEW_PROCESS_GROUP, ByVal 0&, StrPtr(lpCurrentDirectory), StartInfo, ProcessInfo
         CloseHandle ProcessInfo.hThread 'close the handle to the main thread, since we don't use it
         CloseHandle ProcessInfo.hProcess 'close the handle to the process, since we don't use it
         'note that closing the handles of the main thread and the process do not terminate the process
End Sub

Private Sub Class_Initialize()
    'You could put default Values for your network here
    sUserName = ""
    sPassword = ""
    sCommand = ""
    sDomain = ""
End Sub

