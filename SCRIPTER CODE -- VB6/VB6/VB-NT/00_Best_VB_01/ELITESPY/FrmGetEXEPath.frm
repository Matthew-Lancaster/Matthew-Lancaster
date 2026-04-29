VERSION 5.00
Begin VB.Form FrmGetEXEPath 
   Caption         =   "Form1"
   ClientHeight    =   5736
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9972
   LinkTopic       =   "Form1"
   ScaleHeight     =   5736
   ScaleWidth      =   9972
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Left            =   264
      TabIndex        =   1
      Top             =   60
      Width           =   2028
   End
   Begin VB.ListBox List1 
      Height          =   4272
      Left            =   252
      TabIndex        =   0
      Top             =   384
      Width           =   8328
   End
End
Attribute VB_Name = "FrmGetEXEPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----
'[RESOLVED] Get Full path of running Exe-VBForums
'http://www.vbforums.com/showthread.php?593457-RESOLVED-Get-Full-path-of-running-Exe
'----

Option Explicit

Private Const TH32CS_SNAPPROCESS = &H2
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Private Const PROCESS_VM_READ As Long = (&H10)
Private Const MAX_PATH As Integer = &H104

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

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
    ByVal lFlags As Long, _
    ByVal lProcessID As Long _
    ) As Long
     
Private Declare Function Process32First Lib "kernel32" ( _
    ByVal hSnapShot As Long, _
    uProcess As PROCESSENTRY32 _
    ) As Long
     
Private Declare Function Process32Next Lib "kernel32" ( _
    ByVal hSnapShot As Long, _
    uProcess As PROCESSENTRY32 _
    ) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" ( _
     ByVal dwDesiredAccess As Long, _
     ByVal bInheritHandle As Boolean, _
     ByVal dwProcessId As Long _
     ) As Long

Private Declare Function EnumProcessModules Lib "psapi.dll" ( _
     ByVal hProcess As Long, _
     ByRef lphModule As Long, _
     ByVal cb As Long, _
     ByRef lpcbNeeded As Long) As Long

Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" ( _
     ByVal hProcess As Long, _
     ByVal hModule As Long, _
     ByVal lpFileName As String, _
     ByVal nSize As Long) As Long

Private Declare Sub CloseHandle Lib "kernel32" ( _
    ByVal hPass As Long _
    )

Private Sub Command1_Click()
    Dim PE As PROCESSENTRY32
    Dim hSnap As Long
    Dim Result As Boolean
    Dim hProcess As Long
    Dim FileName As String * MAX_PATH
    
    PE.dwSize = Len(PE)
    List1.Clear
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    Result = Process32First(hSnap, PE)
    Do While Result
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, False, PE.th32ProcessID)
            If Not EnumProcessModules(hProcess, 0, 0, 0) = 0 Then
                GetModuleFileNameEx hProcess, 0, FileName, MAX_PATH
                List1.AddItem FileName
            End If
            CloseHandle hProcess
        Result = Process32Next(hSnap, PE)
    Loop
    CloseHandle hSnap
End Sub

