Attribute VB_Name = "modIp"
Option Explicit

Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS As Long = 0
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD As Long = 1
Public Const SOCKET_ERROR As Long = -1

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type

Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
(ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long

Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Declare Function gethostname Lib "WSOCK32.DLL" _
(ByVal szHost As String, ByVal dwHostLen As Long) As Long

Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
(ByVal szHost As String) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Public Function GetIPAddress() As String
    Dim sHostName As String * 256
    Dim lpHost As Long
    Dim HOST As HOSTENT
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim i As Integer
    Dim sIPAddr As String
    
    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If
    
    'gethostname returns the name of the local host into the buffer specified by the name
    ' parameter.The HOST  'name is returned as a null-terminated string. The 'form of the host name is
    ' dependent on the Windows 'Sockets provider - it can be a simple host name, or 'it can be a fully
    ' qualified domain name. However, it 'is guaranteed that the name returned will be successfully
    ' parsed by gethostbyname and WSAAsyncGetHostByName.
    
    ' In actual application, if no local host name has been 'configured, gethostname must succeed
    ' and return a token'host name that gethostbyname or WSAAsyncGetHostByName 'can resolve.
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPAddress = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
        " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    ' gethostbyname returns a pointer to a HOSTENT structure - a structure allocated by Windows
    ' Sockets.The HOSTENT  'structure contains the results of a successful search 'for the host specified
    ' in the name parameter.
    ' The application must never attempt to modify this 'structure or to free any of its components.
    ' Furthermore, 'only one copy of this structure is allocated per thread,
    ' so the application should copy any information it needs 'before issuing any other Windows
    ' Sockets function calls. 'gethostbyname function cannot resolve IP address strings'passed to it.
    ' Such a request is treated exactly as if an 'unknown host name were passed. Use inet_addr to
    ' convert 'an IP address string the string to an actual IP address, 'then use another function,
    ' gethostbyaddr, to obtain the 'contents of the HOSTENT structure.
    sHostName = Trim$(sHostName)
    lpHost = gethostbyname(sHostName)
    
    If lpHost = 0 Then
        GetIPAddress = ""
        MsgBox "Windows Sockets are not responding. " & _
        "Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    'to extract the returned IP address, we have to copy the HOST structure and its members
    CopyMemory HOST, lpHost, Len(HOST)
    CopyMemory dwIPAddr, HOST.hAddrList, 4
    
    'create an array to hold the result
    ReDim tmpIPAddr(1 To HOST.hLen)
    CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
    
    'and with the array, build the actual address, appending a period between members
    For i = 1 To HOST.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    
    'the routine adds a period to the end of the string, so remove it here
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup
End Function
Public Function GetIPHostName() As String
    Dim sHostName As String * 256
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
        " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup
End Function

Public Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function

Public Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function

Public Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
End Sub

Public Function SocketsInitialize() As Boolean
    Dim WSAD As WSADATA
    Dim sLoByte As String
    Dim sHiByte As String
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        MsgBox "The 32-bit Windows Socket is not responding."
        SocketsInitialize = False
        Exit Function
    End If
    
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
               CStr(MIN_SOCKETS_REQD) & " supported sockets."
        SocketsInitialize = False
        Exit Function
    End If
    
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
       (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
       HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
             sHiByte = CStr(HiByte(WSAD.wVersion))
             sLoByte = CStr(LoByte(WSAD.wVersion))
             MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
                   " is not supported by 32-bit Windows Sockets."
             SocketsInitialize = False
             Exit Function
    End If
    
    'must be OK, so lets do it
    SocketsInitialize = True
End Function

