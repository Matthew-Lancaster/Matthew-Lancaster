Attribute VB_Name = "Modulare"
Option Explicit

'============================================
'Author: Joe Wong (Come from China)
'ICQ NO:7601450
'PLS VOTE FOR ME AND GIVE ME COMMENTS,thanks!
'============================================

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPheaplist = &H1
Public Const TH32CS_SNAPthread = &H4
Public Const TH32CS_SNAPmodule = &H8
Public Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule
Public Const MAX_PATH As Integer = 260

'define PROCESSENTRY32 structure
Public Type PROCESSENTRY32
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

'Hardware scanner
'================
Public isClient As Boolean
Public isClienta As Boolean
Public strUserName As String
Public strPassword As String
Public klientoID As Integer
Public webUserName As String
Public webPassword As String
Public oDeviceType() As Variant
Public oDeviceCaption() As Variant
Public oDeviceParam() As Variant
Public oDeviceInterf() As Variant
Public eilute As Integer
Public isHardware As Boolean

'Advanced port scanner
'=====================
Public Const CL As String = "CGILog.ini"
Public RSContinue As Long

Function KeepGoing(Index As Integer)
   frmMain.S1(Index).Close
   If RSContinue < Val(frmMain.P2) Then
      frmMain.S1(Index).Connect , RSContinue
      RSContinue = RSContinue + 1
   Else
      Unload frmMain.S1(Index)
   End If
End Function


