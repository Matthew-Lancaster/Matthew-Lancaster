Imports System.Runtime.InteropServices

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'MsgBox(ListLoggedOnUsers("1-ASUS-X5DIJ"))
        ' STRING_1 = ListLoggedOnUsers("7-ASUS-GL522VW")
        STRING_1 = ListLoggedOnUsers("8-MSI-GP62M-7RD")
        If STRING_1 <> "" Then
            MsgBox(STRING_1)
        End If
        'MsgBox(ListLoggedOnUsers("7-ASUS-GL522VW"))
        'MsgBox(ListLoggedOnUsers("7-ASUS-GL522VW"))
        'MsgBox(ListLoggedOnUsers("7-ASUS-GL522VW"))
        End
    End Sub
End Class

Module Network

        Public Const NERR_SUCCESS As Integer = 0&
        Public Const MAX_PREFERRED_LENGTH As Integer = -1
        Public Const ERROR_MORE_DATA As Integer = 234&

        Public Structure WKSTA_USER_INFO_0
            <MarshalAs(UnmanagedType.LPTStr)> Public wkui0_username As String
        End Structure

        Public Structure WKSTA_USER_INFO_1
            <MarshalAs(UnmanagedType.LPTStr)> Public wkui1_username As String
            <MarshalAs(UnmanagedType.LPTStr)> Public wkui1_logon_domain As String
            <MarshalAs(UnmanagedType.LPTStr)> Public wkui1_oth_domains As String
            <MarshalAs(UnmanagedType.LPTStr)> Public wkui1_logon_server As String
        End Structure

        Public Declare Auto Function NetWkstaUserEnum Lib "netapi32" ( _
         ByVal servername As String, ByVal level As Integer, _
         ByRef buffer As IntPtr, ByVal prefMaxLen As Integer, _
         ByRef entriesRead As Integer, ByRef totalEntries As Integer, _
         ByRef resumeHandle As Integer) As Integer

        Public Declare Auto Function NetApiBufferFree Lib "netapi32" (ByVal buffer As IntPtr) As Long

        Public Function ListLoggedOnUsers(ByVal ComputerName As String) As String

            Dim server As String = "\\" & ComputerName
            Dim bufptr As IntPtr

            Dim dwEntriesread As Integer
            Dim dwTotalentries As Integer
            Dim dwResumehandle As Integer
            Dim nStatus As Integer

        Dim nStructSize As Integer = Marshal.SizeOf(GetType(WKSTA_USER_INFO_1))
            Dim Cnt As Integer

            Dim wui1 As WKSTA_USER_INFO_1

            Do
            nStatus = NetWkstaUserEnum(server, 1, bufptr, MAX_PREFERRED_LENGTH, dwEntriesread, dwTotalentries, dwResumehandle)
                If nStatus = NERR_SUCCESS Or nStatus = ERROR_MORE_DATA Then
                    If dwEntriesread > 0 Then
                        For Cnt = 0 To dwEntriesread - 1
                            Dim ptr As New IntPtr(bufptr.ToInt64() + Cnt * nStructSize)
                        wui1 = DirectCast(Marshal.PtrToStructure(ptr, GetType(WKSTA_USER_INFO_1)), WKSTA_USER_INFO_1)
                            Console.WriteLine(wui1.wkui1_username)
                        Next
                    End If
                Else
                    Select Case nStatus
                        Case 5
                            MsgBox(String.Concat("Access denied to computer ", server, "."), MsgBoxStyle.Exclamation)
                        Case Else
                            MsgBox(String.Concat("An error occurred retrieving the logged on users for computer ", server))
                    End Select
                End If
            Loop While nStatus = ERROR_MORE_DATA
            NetApiBufferFree(bufptr)

        End Function

    End Module


