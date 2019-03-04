Attribute VB_Name = "Mod_1_VAR_COMMON"
Public cProcesses As New clsCnProc

Public Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

