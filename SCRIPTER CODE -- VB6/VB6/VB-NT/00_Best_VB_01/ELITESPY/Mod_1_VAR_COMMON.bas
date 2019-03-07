Attribute VB_Name = "Mod_1_VAR_COMMON"
Public cProcesses As New clsCnProc

Public Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Public VAR_ARRAY
Public BLOCK_RUN_1()
Public BLOCK_RUN_2()
Public BLOCK_RUN_3()

