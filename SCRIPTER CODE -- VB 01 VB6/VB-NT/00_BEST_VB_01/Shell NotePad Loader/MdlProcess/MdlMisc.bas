Attribute VB_Name = "mdlMisc"
Option Explicit
Const sLocation As String = "mdlMisc"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long


Public Function GetParentIndex(TreeView As TreeView, ParentImage As Integer) As Long
    '// Just to get the roots
    Dim i As Long
    Dim pCount As Long
    With TreeView
        For i = 1 To .SelectedItem.Index
            If .Nodes(i).Image = ParentImage Then
                pCount = pCount + 1
            End If
        Next i
    End With
    GetParentIndex = pCount
End Function


Public Function FileName_Parse(ByVal Path As String) As String
    '// Takes a full file specification and returns the path
    Dim A
    For A = Len(Path) To 1 Step -1
        If Mid$(Path, A, 1) = "\" Or Mid$(Path, A, 1) = "/" Then
            FileName_Parse = Mid$(Path, A + 1)
            Exit Function
        End If
    Next A
    FileName_Parse = Path
End Function

Public Function Path_Parse(Path As String) As String
    '// Takes a full file specification and returns the path
    Dim A
    For A = Len(Path) To 1 Step -1
        If Mid$(Path, A, 1) = "\" Or Mid$(Path, A, 1) = "/" Then
            '// Add the correct path separator for the input
            If Mid$(Path, A, 1) = "\" Then
                Path_Parse = LCase$(Left$(Path, A - 1) & "\")
            Else
                Path_Parse = LCase$(Left$(Path, A - 1) & "/")
            End If
            Exit Function
        End If
    Next A
End Function
Function File_Start(FileName As String, Action As String) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    File_Start = ShellExecute(Scr_hDC, Action, FileName, "", Left(FileName, 3), 1)
End Function

