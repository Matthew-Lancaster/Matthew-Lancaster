Attribute VB_Name = "modDivers"
' Delete the extension of a file
Public Function GetFileNameWithoutExtension(FileName As String) As String
    Dim lDotPos As Long
    lDotPos = InStrRev(FileName, ".")
    If lDotPos = 0 Then
        GetFileNameWithoutExtension = FileName
    Else
        GetFileNameWithoutExtension = Left$(FileName, lDotPos - 1)
    End If
End Function
