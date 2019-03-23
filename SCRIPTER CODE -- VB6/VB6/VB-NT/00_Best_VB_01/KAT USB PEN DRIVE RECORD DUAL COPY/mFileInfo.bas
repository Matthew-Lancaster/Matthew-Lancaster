Attribute VB_Name = "mFileInfo"
Public FS




Public Function GetFileSize(Filename) As String
    On Error GoTo Gfserror
    Dim TempStr As String
    TempStr = FileLen(Filename)


    If TempStr >= "1024" Then
        'KB
        TempStr = CCur(TempStr / 1024) & "KB"
    Else


        If TempStr >= "1048576" Then
            'MB
            TempStr = CCur(TempStr / (1024 * 1024)) & "KB"
        Else
            TempStr = CCur(TempStr) & "B"
        End If
    End If
    GetFileSize = TempStr
    Exit Function
Gfserror:
    GetFileSize = "0B"
    Resume
End Function


Public Function GetAttrib(Filename) As String
    On Error GoTo GAError
    Dim TempStr As String
    TempStr = GetAttr(Filename)


    If TempStr = "64" Then
        TempStr = "Alias"
    End If


    If TempStr = "32" Then
        TempStr = "Archive"
    End If


    If TempStr = "16" Then
        TempStr = "Directory"
    End If


    If TempStr = "2" Then
        TempStr = "Hidden"
    End If


    If TempStr = "0" Then
        TempStr = "Normal"
    End If


    If TempStr = "1" Then
        TempStr = "ReadOnly"
    End If


    If TempStr = "4" Then
        TempStr = "System"
    End If


    If TempStr = "8" Then
        TempStr = "Volume"
    End If
    GetAttrib = TempStr
    Exit Function
GAError:
    GetAttrib = "Unknown"
    Resume
End Function


Public Sub SetHidden(Filename As String)
    On Error Resume Next
    SetAttr Filename, vbHidden
End Sub


Public Sub SetReadOnly(Filename As String)
    On Error Resume Next
    SetAttr Filename, vbReadOnly
End Sub


Public Sub SetSystem(Filename As String)
    On Error Resume Next
    SetAttr Filename, vbSystem
End Sub


Public Sub SetNormal(Filename As String)
    On Error Resume Next
    SetAttr Filename, vbNormal
End Sub


Public Function GetFileExtension(Filename As String)
    On Error Resume Next
    Dim TempStr As String
    TempStr = Right(Filename, 2)


    If Left(TempStr, 1) = "." Then
        GetFileExtension = Right(Filename, 1)
        Exit Function
    Else
        TempStr = Right(Filename, 3)


        If Left(TempStr, 1) = "." Then
            GetFileExtension = Right(Filename, 2)
            Exit Function
        Else
            TempStr = Right(Filename, 4)


            If Left(TempStr, 1) = "." Then
                GetFileExtension = Right(Filename, 3)
                Exit Function
            Else
                TempStr = Right(Filename, 5)


                If Left(TempStr, 1) = "." Then
                    GetFileExtension = Right(Filename, 4)
                    Exit Function
                Else
                    GetFileExtension = "Unknown"
                End If
            End If
        End If
    End If
    
End Function


Public Function GetFileDate(Filename As String) As String
    On Error Resume Next
    GetFileDate = FileDateTime(Filename)
End Function

Function FileDateTime(Filespec)

    Dim Ldate As Date
    Set FS = CreateObject("Scripting.FileSystemObject")

    Ldate = 0
    If FS.FileExists(Filespec) Then
        Set F = FS.GetFile(Filespec)
        Ldate = F.DateLastModified
    End If
FileDateTime = Ldate
End Function

Function FileSize(Filespec)

    Dim Ldate As Date

    FileSize = 0
    If FS.FileExists(Filespec) Then
        Set F = FS.GetFile(Filespec)
        FileSize = F.Size
    End If

End Function

Public Sub DeleteFile(Filename As String)
    On Error GoTo DelError
    Kill Filename
    Exit Sub
DelError:
    MsgBox "Error deleting File"
    Resume
End Sub


Public Sub CopyFile(Source As String, Destination As String)
    On Error GoTo CopyError
    FileCopy Source, Destination
    Exit Sub
CopyError:
    MsgBox "Error copying File"
    Resume
End Sub


Public Sub MoveFile(Source As String, Destination As String)
    On Error GoTo MoveError
    FileCopy Source, Destination
    Kill Source
    Exit Sub
MoveError:
    MsgBox "Error moving File"
    Resume
End Sub


Public Sub MakeDIR(Path As String)
    On Error GoTo DIRError
    MkDir Path
    Exit Sub
DIRError:
    MsgBox "Error creating Directory"
    Resume
End Sub


Public Sub RemoveDIR(Path As String)
    On Error GoTo DIRError2
    RmDir Path
    Exit Sub
DIRError2:
    MsgBox "Error removing Directory"
    Resume
End Sub


Public Sub CloseAllFiles()
    On Error Resume Next
    Reset
End Sub

