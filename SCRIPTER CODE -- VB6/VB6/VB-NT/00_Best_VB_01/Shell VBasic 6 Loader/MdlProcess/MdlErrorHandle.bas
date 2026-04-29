Attribute VB_Name = "mdlErrorHandle"
Option Explicit
Public Const ERROR_MOD_NOT_FOUND As Long = 126
Public Sub Err_Dll(ErrorNum As Long, ErrorDesc As String, Source As String, SubOrFunction As String)
    File_WriteTo "ERROR: " & ErrorNum & " at " & Source & "\" & SubOrFunction & " >>> " & ErrorDesc
End Sub
Public Sub Err_Vb(ErrorNum As Long, ErrorDesc As String, Source As String, SubOrFunction As String)
    File_WriteTo "VB ERROR: " & ErrorNum & " at " & Source & "\" & SubOrFunction & " >>> " & ErrorDesc
End Sub
Public Sub File_WriteTo(Text As String)
    '// Allways use this in programs:
 '   Open App.Path & "\PROGRAM.LOG" For Append As #1
 '       Print #1, Text
 '   Close #1
End Sub

