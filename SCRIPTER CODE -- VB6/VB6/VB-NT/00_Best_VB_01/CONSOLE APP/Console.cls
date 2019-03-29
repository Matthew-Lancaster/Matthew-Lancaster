VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const FOREGROUND_BLUE = &H1
Private Const FOREGROUND_GREEN = &H2
Private Const FOREGROUND_RED = &H4
Private Const BACKGROUND_BLUE = &H10
Private Const BACKGROUND_GREEN = &H20
Private Const BACKGROUND_RED = &H40
Private Const BACKGROUND_INTENSITY = &H80&
Private Const BACKGROUND_SEARCH = &H20&
Private Const FOREGROUND_INTENSITY = &H8&
Private Const FOREGROUND_SEARCH = (&H10&)
Private Const ENABLE_LINE_INPUT = &H2&
Private Const ENABLE_ECHO_INPUT = &H4&
Private Const ENABLE_MOUSE_INPUT = &H10&
Private Const ENABLE_PROCESSED_INPUT = &H1&
Private Const ENABLE_WINDOW_INPUT = &H8&
Private Const ENABLE_PROCESSED_OUTPUT = &H1&
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_ERROR_HANDLE = -12&
Private Const INVALID_HANDLE_VALUE = -1&

Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Loaded As Boolean
Private hConsoleOut As Long, hConsoleIn As Long, hConsoleErr As Long



Sub StartUp(cTitle As String)
'Create console
If Loaded Then
 Shutdown
End If
If AllocConsole() Then
hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
If hConsoleOut = INVALID_HANDLE_VALUE Then MsgBox "Unable to get STDOUT"
hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
If hConsoleOut = INVALID_HANDLE_VALUE Then MsgBox "Unable to get STDIN"
Else
MsgBox "Couldn't allocate console"
End If
'Set the caption of the console window
SetConsoleTitle cTitle
'Set the background color of the text in the console
SetConsoleTextAttribute hConsoleOut, FOREGROUND_GREEN Or FOREGROUND_INTENSITY

Loaded = True
End Sub

Sub Shutdown()
'Delete console
CloseHandle hConsoleOut
CloseHandle hConsoleIn
FreeConsole
Loaded = False
End Sub

Sub ConsoleWriteLine(sInput As String)
ConsoleWrite sInput + vbCrLf
End Sub

Sub ConsoleWrite(sInput As String)
If Loaded = True Then
Dim cWritten As Long
WriteConsole hConsoleOut, ByVal sInput, Len(sInput), cWritten, ByVal 0&
End If
End Sub

Function ConsoleReadLine() As String
If Loaded = True Then
Dim ZeroPos As Long
'Create a buffer
ConsoleReadLine = String(10, 0)
'Read the input
ReadConsole hConsoleIn, ConsoleReadLine, Len(ConsoleReadLine), vbNull, vbNull
'Strip off trailing vbCrLf and Chr$(0)'s
ZeroPos = InStr(ConsoleReadLine, Chr$(0))
If ZeroPos > 0 Then ConsoleReadLine = Left$(ConsoleReadLine, ZeroPos - 3)
End If
End Function

Function IsLoaded() As Boolean
IsLoaded = Loaded
End Function

Sub SetTitle(NewTitle As String)
SetConsoleTitle NewTitle
End Sub



