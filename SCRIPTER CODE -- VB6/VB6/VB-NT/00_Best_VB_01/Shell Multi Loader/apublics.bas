Attribute VB_Name = "APublics"
Public FontSizez, SetTrueToLoadLast, LoadFolder, LoadFolderFile, FixIS
Public Capp1, Capp2, Capp3, QQ$

Public A1$, B1$, C1$, OIP$, OIP2$, InfoRapid

Public A4$(), B4$(), C4$()

Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

