Attribute VB_Name = "modProcess"
Option Explicit

Public Const MAX_PATH As Integer = 260

' the process type
Public Type PROCESSENTRY32
    dwSize As Long                      ' Struct size
    cntUsage As Long
    th32ProcessID As Long               ' process ID
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long                  ' thread number
    th32ParentProcessID As Long
    pcPriClassBase As Long              ' process priority
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type

' the processes
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Public Const TH32CS_SNAPPROCESS As Long = 2&

' Close an opened handle
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' List the processes running on the system
Public Sub mpListProcess(ByRef tabID() As Long)
    Dim hSnapshot   As Long
    Dim uProcess    As PROCESSENTRY32
    Dim T           As Long
    Dim Nb          As Long
    ReDim tabID(0) As Long
    ' Snaphshot of the processes
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = 0 Then Exit Sub
    uProcess.dwSize = Len(uProcess)
    ' the first one
    T = ProcessFirst(hSnapshot, uProcess)
    Nb = 0
    ' for each process
    Do While T
        ' save the process
        ReDim Preserve tabID(Nb) As Long
        tabID(Nb) = uProcess.th32ProcessID
        Nb = Nb + 1
        ' take the next one
        T = ProcessNext(hSnapshot, uProcess)
    Loop
    CloseHandle hSnapshot
End Sub

' Get the process name with his process ID
Public Function mpGetProcessName(ByVal ProcessID As Long) As String
    Dim hSnapshot As Long
    Dim uProcess As PROCESSENTRY32
    Dim T As Long
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = 0 Then Exit Function
    uProcess.dwSize = Len(uProcess)
    T = ProcessFirst(hSnapshot, uProcess)
    Do While T
        ' If that is the one we're looking for
        If uProcess.th32ProcessID = ProcessID Then
            mpGetProcessName = uProcess.szexeFile
            mpGetProcessName = Left(mpGetProcessName, InStr(mpGetProcessName, vbNullChar) - 1)
            Exit Function
        End If
        T = ProcessNext(hSnapshot, uProcess)
    Loop
End Function
