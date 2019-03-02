Attribute VB_Name = "modMainCopyFileClipboard"
Option Explicit

' ************************************************************************************************************
' * Main declaration, subroutines and functions
' ************************************************************************************************************

' ************************************************************************************************************
' * Copyright ©2016, Frank Donckers
' * All source code in this file is licensed under a modified BSD license.
' * This means you may use the code in your own projects IF you provide attribution.
' ************************************************************************************************************

'API:
'----
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long

'Public constant variables:
'--------------------------
Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

'Public variables:
'-----------------
Public sReturn              As String
Public i                    As Integer
Private sContentsFile       As String

Sub Main()
    'sCommand =C:\temp\TestFile.txt"
    If Command$ <> "" Then
        sContentsFile = Command$
        If Dir(sContentsFile, vbNormal) = "" Then
            MsgBox "Invalid filename passed."
        End If
    End If
    If sContentsFile = "" Then
        frmCopyFile2ClipBoard.Show 1
    Else
        sReturn = File_To_ClipBoard(sContentsFile)
        MsgBox sReturn
    End If
    End
End Sub


'Copy string to ClipBoard
Function File_To_ClipBoard(sFileName As String) As String
    Dim iFileNum        As Integer
    Dim hGlobalMemory   As Long
    Dim lpGlobalMemory  As Long
    Dim hClipMemory     As Long
    Dim X               As Long
    Dim sFileContents   As String
    Dim sErr            As String
    
    'Clear returnvalue.
    File_To_ClipBoard = ""
    
    On Error GoTo FileNotRead
    
    'Check existence of file.
    If Dir(sFileName, vbNormal) = "" Then
        sErr = "File not found."
        GoTo ExitClean
    End If
    'Get free fileport.
    iFileNum = FreeFile
    'Read file to a string.
    Open sFileName For Binary As #iFileNum
        sFileContents = Space$(LOF(iFileNum))
        Get #iFileNum, , sFileContents
    Close #iFileNum
    
FileNotRead:
    'Return error in case file could not be read.
    If Trim$(sFileContents) = "" Then
        sErr = "File can not be read."
        GoTo ExitClean
    End If
    'Allocate movable global memory.
    hGlobalMemory = GlobalAlloc(GHND, Len(sFileContents) + 1)
    'Lock the block to get a far pointer to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    'Copy the string to this global memory.
    lpGlobalMemory = lstrcpy(lpGlobalMemory, sFileContents)
    'Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        sErr = "Can not unlock ClipBoard"
        GoTo ExitHere
    End If
    'Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        sErr = "Can not open ClipBoard."
        GoTo ExitHere
    End If
    'Clear the Clipboard.
    X = EmptyClipboard()
    'Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)
    'Check ClipBoard value
    If hClipMemory = 0 Then
        sErr = "Can not copy to ClipBoard."
        GoTo ExitClean
    End If
    
ExitHere:
    'Close the ClipBoard.
    If CloseClipboard() = 0 Then
        sErr = "Can not close ClipBoard."
    End If
    
ExitClean:
    'Return success in case of no error.
    If sErr = "" Then
        sErr = "Copy to ClipBoard is executed."
    Else
        sErr = "Error: " & sErr
    End If
    'Return value.
    File_To_ClipBoard = sErr
End Function

' Returns the path of a filename
Public Function GetPathForFile(sPath As String, Optional bStripPath As Boolean = True) As String
    Dim iPos As Integer
    
    GetPathForFile = ""
    If sPath = "" Then Exit Function
    sPath = Trim$(sPath)
    iPos = InStrRev(sPath, "\", , vbTextCompare)
    If iPos > 0 Then
        If bStripPath Then
            GetPathForFile = Trim$(Left$(sPath, iPos - 1))
        Else
            GetPathForFile = Trim$(Left$(sPath, iPos))
        End If
    End If
End Function


