Imports System.IO
Imports System.IO.File
Imports System.Runtime.InteropServices

Public Class Form1

    <StructLayout(LayoutKind.Sequential, Pack:=1)> _
    Public Structure SHQUERYRBINFO
        Dim cbSize As Int32
        Dim i64Size As Long
        Dim i64NumItems As Long
    End Structure

    <DllImport("shell32.dll")> _
    Private Shared Function SHQueryRecycleBin(ByVal pszRootPath As String, ByRef ptSHQueryRBInfo As SHQUERYRBINFO) As Int32
    End Function

    Private Enum RecycleFlags As Int32
        SHERB_NOCONFIRMATION = &H1
        SHERB_NOPROGRESSUI = &H2
        SHERB_NOSOUND = &H4
    End Enum

    <DllImport("shell32.dll")> _
    Private Shared Function SHEmptyRecycleBin(ByVal hwnd As IntPtr, ByVal pszRootPath As String, ByVal dwFlags As RecycleFlags) As Int32
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFile.Select(0, 0)
        MakeDummyFile()
    End Sub

    Private Sub MakeDummyFile()
        Dim file_name As String = Application.StartupPath
        If file_name.EndsWith("\bin\Debug") Then
            file_name = file_name.Substring(0, file_name.LastIndexOf("\bin\Debug"))
        End If
        If Not file_name.EndsWith("\") Then
            file_name &= "\"
        End If
        file_name &= "test.txt"

        txtFile.Text = file_name

        If Not Exists(file_name) Then
            Using sw As StreamWriter = CreateText(file_name)
                sw.WriteLine("This is a test file for moving to the wastebasket.")
                sw.Close()
            End Using
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim opt_recycle As FileIO.RecycleOption
        If radDeletePermanently.Checked Then
            opt_recycle = FileIO.RecycleOption.DeletePermanently
        Else
            opt_recycle = FileIO.RecycleOption.SendToRecycleBin
        End If

        Dim opt_confirm As FileIO.UIOption
        If chkConfirmDelete.Checked Then
            opt_confirm = FileIO.UIOption.AllDialogs
        Else
            opt_confirm = FileIO.UIOption.OnlyErrorDialogs
        End If

        Try
            My.Computer.FileSystem.DeleteFile( _
                txtFile.Text, opt_confirm, opt_recycle)
        Catch ex As Exception
            MessageBox.Show("Error deleting file" & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        MakeDummyFile()
    End Sub

    Private Sub btnEmpty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEmpty.Click
        Dim options As RecycleFlags = 0

        If Not chkProgress.Checked Then options = options Or RecycleFlags.SHERB_NOPROGRESSUI
        If Not chkPlaySound.Checked Then options = options Or RecycleFlags.SHERB_NOSOUND
        If Not chkConfirmEmpty.Checked Then options = options Or RecycleFlags.SHERB_NOCONFIRMATION

        Try
            SHEmptyRecycleBin(Nothing, Nothing, options)
        Catch ex As Exception
            MessageBox.Show("Error emptying wastebasket" & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub tmrCheckBin_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrCheckBin.Tick
        Static had_error As Boolean = False
        If had_error Then Exit Sub

        Try
            Dim info As SHQUERYRBINFO
            info.cbSize = Len(info)
            SHQueryRecycleBin(Nothing, info)
            lblNumItems.Text = info.i64NumItems & " items (" & info.i64Size & " bytes)"
        Catch ex As Exception
            had_error = True
            MessageBox.Show("Error getting wastebasket information" & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
