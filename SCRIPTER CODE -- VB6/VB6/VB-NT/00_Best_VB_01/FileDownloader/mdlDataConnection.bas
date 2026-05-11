Attribute VB_Name = "mdlDataConnection"
'
'                                                                         The KPD-Team, March 13th, 2002
'                                                                          http://www.allapi.net/
'                                                                          Version: 1.5
'
'Notes for this module file:
'    Most methods of this module may not be used by your program. They are only
'    needed for the DataConnection class file.
'    All methods in this file, with the exception of the StartWinsock and the EndWinSock
'    methods,  are subject to change. We strongly suggest *NOT* to use them in your projects
'    (and frankly, we don't see any reason why you should use them).
'
'    Before an application can use the FileDownloaded class, it *MUST* call the
'    StartWinsock method from this module!
'    If you don't do this, every call to the WinSock API will fail!
'
'    If you find any bugs in this file or you want to suggest a new feature,
'    don 't hesitate to contact us at KPDTeam@allapi.net
'
'    We strongly suggest that you don't modify this module. Thatway, if a new
'    version of this file is released, you simply have to replace the classe file and/or
'    module file with the new file(s) and you're done.
'    If you make your own modifications, chances are you may introduce some new
'    bugs in this file and you cannot update this file as easy as normal.
'    Instead of adding your own methods, let us know about your request and we'll
'    certainly consider it.

'Copyright Notice And disclaimer:
'    The FileDownloader class file is created by The KPD-Team, 2001. You're free
'    to use this class file in your own projects, however you may not claim that
'    you have written this class file.
'    You must leave this copyright notice in the file if you intend to distribute
'    the source code of your program.
'    This class file is provided "as is", with no guarantee of completeness or accuracy
'    and without warranty of any kind, express or implied. It is intended solely
'    to provide general guidance on matters of interest for the personal use of
'    the user of this program, who accepts full responsibility for its use.
'    For more information about this, please contact us at KPDTeam@allapi.net.

Option Explicit
Private Const WSA_DESCRIPTIONLEN = 256
Private Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Private Const WSA_SYS_STATUS_LEN = 128
Private Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1
Private Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
Private Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
Public SubclassWindow As Long
Public ConnectionCount As Long
Public UpperBound As Long
Private Connections() As DataConnection
Private WSAStartedUp As Boolean     'Flag to keep track of whether winsock WSAStartup wascalled
'Public Function StartWinsock(sDescription As String) As Boolean
'  Used to initialize the WinSock API for this process
'  Before trying to connect to a remote/proxy server, you *MUST*
'  call this function!
'Parameters: String, used to return a description of the WinSock API
'                   (may be a vbNullString)
'Return Value: Boolean that indicates success or failure
Public Function StartWinsock(sDescription As String) As Boolean
    Dim StartupData As WSADataType
    If Not WSAStartedUp Then
        If Not WSAStartup(&H101, StartupData) Then
            WSAStartedUp = True
            sDescription = StartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function
'Public Sub EndWinsock()
'  Used to clean up the WinSock API for this process
'  Before ending your application, you must call this function
'Parameters: NONE
'Return Value: NONE
Public Sub EndWinsock()
    If WSAIsBlocking() Then
        Call WSACancelBlockingCall
    End If
    Call WSACleanup
    WSAStartedUp = False
End Sub
'Public Function DataProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'  Used to process messages sent to the Subclass Window
'  Applications should not use this function
'Parameters: Window Procedure parameters
'Return Value: 0
Public Function DataProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Index As Long
    Index = FindConnection(wParam)
    If Index >= 0 Then
        Connections(Index).ProcessMessage lParam
    End If
End Function
'Public Sub AddConnection(Connection As DataConnection)
'  Used to add a new DataConnection object to the connection pool
'  Applications should not use this function
'Parameters: New DataConnection
'Return Value: NONE
Public Sub AddConnection(Connection As DataConnection)
    Dim Index As Long
    Index = FindConnectionObject(Connection)
    If Index = -1 Then
        Index = FindFreePlace()
        If Index = -1 Then
            ReDim Preserve Connections(0 To ConnectionCount) As DataConnection
            Set Connections(ConnectionCount) = Connection
            UpperBound = UpperBound + 1
        Else
            Set Connections(Index) = Connection
        End If
        ConnectionCount = ConnectionCount + 1
    End If
End Sub
'Public Sub RemoveConnection(Connection As DataConnection)
'  Used to remove a new DataConnection object from the connection pool
'  Applications should not use this function
'Parameters: DataConnection to remove
'Return Value: NONE
Public Sub RemoveConnection(Connection As DataConnection)
    Dim Index As Long
    Index = FindConnectionObject(Connection)
    If Index >= 0 Then
        Set Connections(Index) = Nothing
        ConnectionCount = ConnectionCount - 1
    End If
End Sub
'Public Function FindConnection(FindSocket As Long) As Long
'  Used to find the index of a DataConnection object in the connection pool
'  Applications should not use this function
'Parameters: Socket handle of the dataconnection to find
'Return Value: Index of DataConnection
Public Function FindConnection(FindSocket As Long) As Long
    Dim Cnt As Long
    FindConnection = -1
    For Cnt = 0 To UpperBound - 1
        If Not (Connections(Cnt) Is Nothing) Then
            If Connections(Cnt).SocketHandle = FindSocket Then
                FindConnection = Cnt
                Exit For
            End If
        End If
    Next Cnt
End Function
'Public Function FindConnectionObject(FindObject As DataConnection) As Long
'  Used to find the index of a DataConnection object in the connection pool
'  Applications should not use this function
'Parameters: Dataconnection to find
'Return Value: Index of DataConnection
Public Function FindConnectionObject(FindObject As DataConnection) As Long
    Dim Cnt As Long
    FindConnectionObject = -1
    For Cnt = 0 To UpperBound - 1
        If Not (Connections(Cnt) Is Nothing) Then
            If Connections(Cnt) Is FindObject Then
                FindConnectionObject = Cnt
                Exit For
            End If
        End If
    Next Cnt
End Function
'Public Function FindFreePlace() As Long
'  Used to find the index of an empty place in the connection pool array
'  Applications should not use this function
'Parameters: NONE
'Return Value: New index
Public Function FindFreePlace() As Long
    Dim Cnt As Long
    FindFreePlace = -1
    For Cnt = 0 To UpperBound - 1
        If (Connections(Cnt) Is Nothing) Then
            FindFreePlace = Cnt
            Exit For
        End If
    Next Cnt
End Function
'Public Function ReadNextLine(ByRef sInput As String) As String
'  Returns the first line of a string and removes it from that string
'  Applications may use this function, but this function may not be included
'  in future releases of this module
'Parameters: Input string
'Return Value: First line of the input string
Public Function ReadNextLine(ByRef sInput As String) As String
    Dim ZeroPos As Long
    ZeroPos = InStr(1, sInput, vbCrLf)
    If ZeroPos > 0 Then
        ReadNextLine = Left$(sInput, ZeroPos - 1)
        sInput = Right$(sInput, Len(sInput) - Len(ReadNextLine) - 2)
    Else
        ReadNextLine = sInput
        sInput = ""
    End If
End Function

