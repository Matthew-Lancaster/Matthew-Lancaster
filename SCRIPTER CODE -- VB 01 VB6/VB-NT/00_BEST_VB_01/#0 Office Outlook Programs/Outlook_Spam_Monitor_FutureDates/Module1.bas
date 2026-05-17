Attribute VB_Name = "Module1"
Public Timer4
Public Rrr
Public Rrr2
Public Title$
Public t1$, t2$, t3$, t4$, t5$, t6$, t7$, t8$

Public MyItemsFrm4
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long



'***********************************************
'# Check, whether we are in the IDE
Public Function IsIDE() As Boolean
'  If IsIDE Then

  Debug.Assert Not TestIDE(IsIDE)
End Function


Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function

