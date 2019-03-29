Attribute VB_Name = "APIModule"
Option Explicit

''''''''''''''''''''''RegisterServiceProcess''''''''''''''''''''''''''''''''''''''''
'The RegisterServiceProcess function registers or _
 unregisters a service process. A service process _
 continues to run after the user logs off.
Public Declare Function RegisterServiceProcess Lib _
    "kernel32.dll" (ByVal dwProcessId As Long, _
     ByVal dwType As Long) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''GetCurrentProcessID'''''''''''''''''''
'The GetCurrentProcessId function returns the process _
 identifier of the calling process.
Public Declare Function GetCurrentProcessId Lib "kernel32" _
() As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const SND_FILENAME = &H20000     '  name is a file name
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const SPIF_UPDATEINIFILE = &H1

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_CURRENT_USER = &H80000001
Private Const REG_SZ = 1                         ' Unicode nul terminated string

Public Sub AddInRun()
Dim retVal As Long, KeyHandle As Long
Dim Value As String
    Value = App.Path & "\" & App.EXEName & ".EXE"
    retVal = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", KeyHandle)
    retVal = RegSetValueEx(KeyHandle, "MyApp", 0&, REG_SZ, ByVal Value, 255&)
    retVal = RegCloseKey(KeyHandle)
End Sub

