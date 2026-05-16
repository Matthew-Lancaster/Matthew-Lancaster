Attribute VB_Name = "Module1"
Dim FS, F, TS, FR1, FR2, iVal, iVal2, ReturnValue, QQ, ii, TF1, Flag2, TESTMODE
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)



Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Option Explicit

 Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
 Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
 Const INFINITE = -1&

 
 
 Private Function shellAndWait(ByVal fileName As String) As Long
     Dim executionStatus As Long
     Dim processHandle As Long
     Dim ReturnValue As Long
     'Execute the application/file
     'executionStatus = Shell(fileName, vbNormalFocus)
     executionStatus = Shell(fileName, vbHide)

     'Get the Process Handle
     processHandle = OpenProcess(&H100000, True, executionStatus)

     'Wait till the application is finished
     ReturnValue = WaitForSingleObject(processHandle, INFINITE)

     'Send the Return Value Back
     shellAndWait = ReturnValue

 End Function

 Private Sub Command1_Click()
     'launch the application and wait
     Dim fileName As String
     Dim retrunValue As Long

     fileName = "EXCEL.EXE"
     retrunValue = shellAndWait(fileName)

     If retrunValue = 0 Then
         MsgBox "Application executed and was finished"
     End If
 End Sub



Sub Main()

If App.PrevInstance = True Then End


Set FS = CreateObject("Scripting.FileSystemObject")

If IsIDE = True Then
    
    On Error Resume Next
    
    Set F = FS.Getfile("C:\Program Files\Microsoft Visual Studio\VB98\Link.exe")
    If F.datelastmodified = DateSerial(1999, 6, 2) Then
        FS.CopyFile "C:\Program Files\Microsoft Visual Studio\VB98\Link.exe", "C:\Program Files\Microsoft Visual Studio\VB98\LinkBAK.exe"
    End If
    If Err.Number > 0 Then MsgBox Err.Description: Stop


    
    FS.CopyFile "C:\Program Files\Microsoft Visual Studio\VB98\Linker2.exe", "C:\Program Files\Microsoft Visual Studio\VB98\Link.exe"
    If Err.Number > 0 Then MsgBox Err.Description
    
    '-------------
    TESTMODE = True
'    End
    '-------------

End If








On Error Resume Next
TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Linkcounter1.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Input Lock Read Write As #FR1
    Line Input #FR1, iVal
Close #1
On Error GoTo 0

If TESTMODE = False Then iVal2 = Val(iVal) + 1

TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Linkcounter.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Output Lock Read Write As #FR1
    Print #FR1, iVal2
Close #1

If TESTMODE = False Then
TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\LastLink.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Output Lock Read Write As #FR1
    Print #FR1, Now
Close #1

TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\LastLinksDays48Hours.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Append Lock Read Write As #FR1
    Print #FR1, Now
Close #1

End If



Dim HALO, Lethal, Flag
Flag = 0: HALO = 0: Flag2 = 0
TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\LastLinksDaysOUT.txt"
FileInUseDelay (TS)
FR2 = FreeFile
Open TS For Output Lock Read Write As #FR2

TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\LastLinksDays48Hours.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Input As #FR1
    Do
        Line Input #FR1, HALO
        If DateValue(HALO) + TimeValue(HALO) > Now - 30 Then Flag2 = 1
        If DateValue(HALO) + TimeValue(HALO) > Now - 2 Then Flag = 1
        If Flag2 = 1 Then Print #FR2, HALO
        If Flag = 1 Then Lethal = Lethal + 1
    
    Loop Until EOF(FR1)
Close #FR1, FR2

Kill "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\LastLinksDays48Hours.txt"
Name "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\LastLinksDaysOUT.txt" As "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\LastLinksDays48Hours.txt"


TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\LastLink48HourCount.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Output Lock Read Write As #FR1
    Print #FR1, Lethal + 1
Close #1






'------------------------------


Dim HALO2, TUFF
Flag = 0: HALO = 0: Flag2 = 0: Lethal = 0

'NEW CREATE
TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Line--OBJECTS-48Hr-OUT.txt"
FileInUseDelay (TS)
FR2 = FreeFile

Open TS For Output Lock Read Write As #FR2
TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Line--OBJECTS-48Hr.txt"

On Error Resume Next '--- FILE INPUT MIGHT BE ZERO
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Input Lock Read Write As #FR1
    Do
        Line Input #FR1, HALO
        
        If Trim(HALO) = "" Then
            If Flag2 = 1 Then Print #FR2, ""
            Do
                If EOF(FR1) Then Exit Do
                Line Input #FR1, HALO
                If EOF(FR1) Then Exit Do
            Loop Until HALO <> ""
            HALO2 = Mid$(HALO, 5, 19)
            'If HALO2 <> "" Then Debug.Print DateValue(HALO2) + TimeValue(HALO2)
        End If
        
        If HALO2 <> "" And Flag = 0 Then
            If DateValue(HALO2) + TimeValue(HALO2) > Now - 30 And Flag2 = 0 Then
                Flag2 = 1
                Print #FR2, ""
            End If
            If DateValue(HALO2) + TimeValue(HALO2) > Now - 2 Then
                Flag = 1
            End If
        End If
            
        If Flag2 = 1 And HALO <> "" Then
            Print #FR2, HALO
        End If
        
        If Flag = 1 And HALO <> "" Then
            TUFF = 0
            Do
                TUFF = InStr(TUFF + 1, HALO, ".OBJ""")
                If TUFF > 0 Then Lethal = Lethal + 1
            Loop Until TUFF = 0
        End If
    
    Loop Until EOF(FR1)
Close #FR1, FR2
On Error GoTo 0 '--- FILE INPUT MIGHT BE ZERO

Kill "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Line--OBJECTS-48Hr.txt"
Name "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Line--OBJECTS-48Hr-OUT.txt" As "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Line--OBJECTS-48Hr.txt"




TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Line--OBJECTS-48Hr-COUNT.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Output Lock Read Write As #FR1
    Print #FR1, Lethal + 1
Close #1












'Shell "C:\Program Files\Microsoft Visual Studio\VB98\LinkBAK " + Command$

If IsIDE = False Then ReturnValue = shellAndWait("C:\Program Files\Microsoft Visual Studio\VB98\LinkBAK " + Command$)
If IsIDE = True Then Exit Sub






TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Line--OBJECTS-48Hr.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Append Lock Read Write As #FR1
    Print #FR1, Format(Now, "DDD DD-MM-YYYY HH:MM:SS") + " - " + Command$
Close #1

'IF NEED DATE GO BACK FURTHER USE THIS
TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Line--OBJECTS-TOTAL.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Append Lock Read Write As #FR1
    Print #FR1, Format(Now, "DDD DD-MM-YYYY HH:MM:SS") + " - " + Command$
Close #1


TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Lines.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Append Lock Read Write As #FR1
    Print #FR1, Format(Now, "DDD DD-MM-YYYY HH:MM:SS") + " - " + Command$
Close #1

'--------------------- EXIT CODES
'---------------------


TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Lines.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Append Lock Read Write As #FR1
    Print #FR1, Format(Now, "DDD DD-MM-YYYY HH:MM:SS") + " - Exit Code =" + Str(ReturnValue) + vbCrLf
Close #1


TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Line--OBJECTS-48Hr.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Append Lock Read Write As #FR1
    Print #FR1, Format(Now, "DDD DD-MM-YYYY HH:MM:SS") + " - Exit Code =" + Str(ReturnValue) + vbCrLf
Close #1

'IF NEED DATE GO BACK FURTHER USE THIS
TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Line--OBJECTS-TOTAL.txt"
FileInUseDelay (TS)
FR1 = FreeFile
Open TS For Append Lock Read Write As #FR1
    Print #FR1, Format(Now, "DDD DD-MM-YYYY HH:MM:SS") + " - Exit Code =" + Str(ReturnValue) + vbCrLf
Close #1


If IsIDE = False Then ExitProcess (ReturnValue)

'End

End Sub



'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************



Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function


Sub FileInUseDelay(TX)
    QQ = Now + TimeSerial(0, 5, 0)
    Do
        DoEvents
        ii = FileInUse(TX)
        If ii = True Then Sleep (200)
    Loop Until ii = False Or QQ < Now
    
    If ii = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + TX + vbCrLf + "Longer than 5 Min"
        Stop
    End If
End Sub
