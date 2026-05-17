Attribute VB_Name = "suspendprocess"
'//////////////////////////////////////////////////////////////////////////////'
' This code were explicitly developed for PSC(Planet Source Code) Users,
' as Open Source Project. This code are property of their author.
' The code is provided "as is" WITHOUT any warranty.

' You may use any of this code in you're own application(s).

' Code by Janevski Blagoj
' Please vote for me on planet-source-code.com
' e-mail: blagoj_bl@yahoo.com for comments or help
' (c) XbXan 2006
'//////////////////////////////////////////////////////////////////////////////'


Option Explicit
'To suspend or resume thread
Private Declare Function SuspendThread Lib "kernel32" (ByVal hthread As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hthread As Long) As Long
'To open a thread
Private Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
'THREAD_SUSPEND_RESUME can be used instead of THREAD_ALL_ACCESS for suspending and resuming a thread
Private Const THREAD_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF
'Current running process PID or ID
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

' ToolHelp 32-bit VB implementation
'--------------------------------------------
'Note: These are not all the functions from ToolHelp.

Private Const MAX_PATH = 260
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4

' Describe a process
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

' Describe a thread
Private Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type

'Functions
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal dwProcessId As Long) As Long
Private Declare Function Thread32First Lib "kernel32" (ByVal hObject As Long, p As THREADENTRY32) As Boolean
Private Declare Function Thread32Next Lib "kernel32" (ByVal hObject As Long, p As THREADENTRY32) As Boolean

' This function is not from ToolHelp but you need it to destroy a snapshot
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function SuspendResumeProcess(ByVal procid As Long, ByVal suspendresume As Boolean) As Boolean
Dim hsnapshot As Long
Dim htthread As Long
Dim pthread As Boolean
Dim pt As THREADENTRY32

SuspendResumeProcess = False
'Take the snapshot
hsnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0)

'The dwSize parameter must be set to Len(THREADENTRY32)
'(in c++ sizeof(THREADENTRY32)) in order
'Thread32First to work
pt.dwSize = Len(pt)

pthread = Thread32First(hsnapshot, pt)

While pthread
    'if thread's owner process's PID or ID is found then
    'suspend it or resume it depending on suspendresume value
    If pt.th32OwnerProcessID = procid Then
        htthread = OpenThread(THREAD_ALL_ACCESS, 0, pt.th32ThreadID)
        If htthread <> 0 Then
            If suspendresume Then SuspendThread (htthread) Else ResumeThread (htthread)
            CloseHandle htthread
            SuspendResumeProcess = True
        End If
    End If
    'keep searching
    pthread = Thread32Next(hsnapshot, pt)
Wend
'destroy the snapshot
CloseHandle hsnapshot
End Function

