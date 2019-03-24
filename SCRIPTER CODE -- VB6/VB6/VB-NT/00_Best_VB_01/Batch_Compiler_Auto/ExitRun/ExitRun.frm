VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ExitRun"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1815
      Top             =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim ee
Private Sub Form_Load()

If IsIDE = False Then
    'Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" /r """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
    'End
End If

'MsgBox "00 x+str(Now)"
'ee = Now + TimeSerial(0, 0, 2)
End Sub

Private Sub Timer1_Timer()

Set fs = CreateObject("Scripting.FileSystemObject")

tt = FindWindow(vbNullString, "Batch VB6 Compiler")
If tt > 0 Then Exit Sub
'If ee > Now Then Exit Sub

On Error GoTo eh

A1 = "D:\VB6\VB-NT\00_Best_VB_01\Batch_Compiler_Auto\BatchCompiler-2.exe"
A2 = "D:\VB6\VB-NT\00_Best_VB_01\Batch_Compiler_Auto\BatchCompiler.exe"
Set F1 = fs.getfile(A1)
Set F2 = fs.getfile(A2)

If F1.DATELASTMODIFIED = F2.DATELASTMODIFIED Then End

If F1.DATELASTMODIFIED > F2.DATELASTMODIFIED Then
    fs.CopyFile A1, A2
    Call ShellEnd
End If

If F1.DATELASTMODIFIED < F2.DATELASTMODIFIED Then
    fs.CopyFile A2, A1
    Call ShellEnd
End If

End

eh:
Exit Sub

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

Sub ShellEnd()
    Shell "D:\VB6\VB-NT\00_Best_VB_01\Batch_Compiler_Auto\BatchCompiler.exe", vbNormalFocus
    End
End Sub

