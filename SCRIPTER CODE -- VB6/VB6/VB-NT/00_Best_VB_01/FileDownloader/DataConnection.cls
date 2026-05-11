VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'                                                                         The KPD-Team, March 13th, 2002
'                                                                          http://www.allapi.net/
'                                                                          Version: 1.5
'
'Notes for this class file:
'    If you use this class to retrieve files larger than 2.1Gb, you're bound to
'    have problems with it. This class makes extensive use of the Long data type
'    and therefore the maximum filesize that can be downloaded with this class
'    is 2,147,483,647 bytes.
'
'    If you find any bugs in this file or you want to suggest a new feature,
'    don 't hesitate to contact us at KPDTeam@allapi.net
'
'    We strongly suggest that you don't modify this class module. Thatway, if a new
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
'
' New in Version 1.1:
'  - Bug fixes !!
'  - Added new bugs
'  - RealReceivedBytes property
'  - MaxDownload property
'  - FailureToString method
'  - FetchURLString method
'  - QuerySent event
'
'New in Version 1.2:
' - Support for password protected sites
' - UseHttpAuthorization property
' - HttpUser property
' - HttpPass property
'
'New in Version 1.3:
' - ObjectMoved Event
' - Removed some bugs

Option Explicit
'<WinSock definitions>
Private Const SD_BOTH = &H2
Private Const SD_SEND = &H1
Private Const SD_RECEIVE = &H0
Private Const FD_SETSIZE = 64
Private Const hostent_size = 16
Private Const IPPROTO_TCP = 6
Private Const INADDR_NONE = &HFFFF
Private Const INVALID_SOCKET = -1
Private Const SOCKET_ERROR = -1
Private Const SOCK_STREAM = 1
Private Const AF_INET = 2
Private Const PF_INET = 2
Private Const sockaddr_size = 16
Private Const SOL_SOCKET = &HFFFF&
Private Const SO_LINGER = &H80&
Private Const FD_READ = &H1&
Private Const FD_WRITE = &H2&
Private Const FD_CONNECT = &H10&
Private Const FD_CLOSE = &H20&
Private Const Base64Alphabet As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Private Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Private Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
Private Declare Function Connect Lib "wsock32.dll" Alias "connect" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Private Declare Function getsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
Private Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
Private Declare Function recv Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function Send Lib "wsock32.dll" Alias "send" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function ShutDown Lib "wsock32.dll" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
Private Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private saZero As sockaddr
'</WinSock definitions>
Private Const GWL_WNDPROC = (-4)
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3&
Private Const INVALID_HANDLE_VALUE = -1&
Private Const CREATE_ALWAYS = 2
Private Const FILE_BEGIN = 0
Private Const OPEN_ALWAYS = 4
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
'The Connected event is raised when a connection has been established with either
'the proxy or the host.
Event Connected(ID As Long)
'The Disconnected event is raised when the file has been successfully downloaded
'and the connection to the host or proxy has been closed.
Event Disconnected(ID As Long)
'The DownloadFailed event is raised when a download fails. This can be for various
'reasons; the Why parameter indicates the source of the problem:
'   � dfConnectionLost
'      The connection with the proxy or host has been closed, before downloading
'      the entire file.
'   � dfCancelledByUser
'      The download has been aborted by the user.
'   � dfCannotConnect
'      Unable to connect to the host.
'   � dfCannotConnectToProxy
'      Unable to connect to the proxy or unable to send data to it.
'   � dfCannotCreateFile
'      Unable to create the file, specified in the Filename property.
'   � dfCannotCreateWindow
'      Unable to create a necessary window.
'   � dfCannotSendData
'      Unable to send data to the host or proxy
'   � dfHTTPError
'      The server didn't reply with a 200 or 206 reply code. For more information
'      about the server reply code, take a look at the HTTPReply property.
'   � dfProxyAuthMethodNotAccepted
'      The selected authentication method was not accepted by the (SOCKS5) proxy.
'   � dfProxyRequestRejectedOrFailed
'      The connection request has failed or the proxy server has rejected it.
'   � dfUserPasswordNotAccepted
'      The user/password combination was not accepted by the proxy.
'   . dfInvalidURL
'      The specified URL is invalid
Event DownloadFailed(Why As DOWNLOAD_FAILURE, ID As Long)
'Event BytesReceived(lByteCount As Long, ID As Long)
'The BytesReceived event is raised whenever new data has arrived. The number
'of arrived bytes is specified by lByteCount.
Event BytesReceived(lByteCount As Long, ID As Long)
'The StreamBytes event is only raised when the DownloadType has been set to
'dtStream. The byte array is the array that contains the arrived bytes and
'lByteCount is the number of bytes contained in this array.
Event StreamBytes(lByteCount As Long, bBytes() As Byte, ID As Long)
'The QuerySent event is raised when the HTTP query has been successfully sent
Event QuerySent(ID As Long)
'The ObjectMoved event is raised when the server replies with a 301 or a 302
'message. This means the object, pointed to by the URL, has either temporarily
'or permanently moved.
'If AutoRedirect is set to True, this event will not be raised and the class
'will automatically start downloading from the new URL.
Event ObjectMoved(sNewUrl As String, ID As Long)
Public Enum DOWNLOAD_TYPE
    dtToFile = 0
    dtToBuffer = 1
    dtStream = 2
End Enum
Public Enum PROXY_TYPE
    ptNoProxy = 0
    ptStandardProxy = 1
    ptSOCKS4Proxy = 2
    ptSOCKS5Proxy = 3
End Enum
Public Enum DOWNLOAD_FAILURE
    dfConnectionLost = 0
    dfCancelledByUser = 1
    dfCannotCreateWindow = 2
    dfCannotSendData = 3
    dfHTTPError = 4
    dfCannotConnect = 5
    dfCannotCreateFile = 6
    dfCannotConnectToProxy = 7
    dfProxyRequestRejectedOrFailed = 8
    dfProxyAuthMethodNotAccepted = 9
    dfUserPasswordNotAccepted = 10
    dfInvalidURL = 11
End Enum
Public Enum CONNECTION_TYPE
    ctHTTP = 0
    'ctFTP = 1 'Perhaps supported in future versions
End Enum
Public Enum PROXY_AUTHENTICATION
    paNone = 0
    paUserPass = 2
End Enum
Private m_DownloadType As DOWNLOAD_TYPE
Private m_ProxyType As PROXY_TYPE
Private m_Filename As String
Private m_UseTempFile As Boolean
Private m_ProxyHost As String
Private m_ProxyPort As Long
Private m_URL As String
Private m_PacketSize As Long
Private m_ID As Long
Private m_Host As String
Private m_Port As Long
Private m_AllowCache As Boolean
Private m_IsConnected As Boolean
Private m_ReceivedBytes As Long
Private m_DownloadLength As Long
Private m_ConnectionType As CONNECTION_TYPE
Private m_SocksUsername As String
Private m_SocksPassword As String
Private m_SocksUseRemoteDNS As Boolean
Private m_ProxyAuthentication As PROXY_AUTHENTICATION
Private m_HTTPReply As Long
Private m_UseResume As Boolean
Private m_ResumeFrom As Long
Private m_AutoRedirect As Boolean
Private m_MaxDownload As Long
Private m_HttpUser As String
Private m_HttpPass As String
Private m_UseHttpAuthorization As Boolean
Private hPrevProc As Long
Private theSocket As Long
Private ParseHeader As Boolean
Private Header As String
Private DataBuffer() As Byte
Private DataInBuffer As Long
Private DataBufferSize As Long
Private FileHandle As Long
Private SocksAuthenticated As Boolean
Private SocksStep As Byte
Private UserCancelled As Boolean
Private NewLocation As String
Private Moving As Boolean
Private ProcessedCloseMessage As Boolean
Private Sub Class_Initialize()
    'Initialize to default values
    m_DownloadType = dtToFile
    m_ProxyType = ptNoProxy
    m_Filename = ""
    m_UseTempFile = True
    m_ProxyHost = "localhost" 'this computer
    m_ProxyPort = 80 'standard HTTP port
    m_Port = 80 'standard HTTP port
    m_URL = ""
    m_PacketSize = 1024
    m_ID = 0
    m_ProxyAuthentication = paNone
    m_SocksUseRemoteDNS = False
    m_SocksUsername = ""
    m_ConnectionType = ctHTTP
    m_DownloadLength = -1
    m_UseResume = False
    m_AutoRedirect = True
    m_MaxDownload = -1
    m_HttpUser = ""
    m_HttpPass = ""
    m_UseHttpAuthorization = False
    If SubclassWindow = 0 Then
        ' Create an invisible window (needed for the asynchronous WinSock functions)
        SubclassWindow = CreateWindowEx(0, "STATIC", "KPD_WINSOCK_SUBCLASS_LABEL", 0, 0, 0, 0, 0, 0, ByVal 0&, 0, ByVal 0&)
        'Uhoh... CretaeWindowEx failed for some reason
        If SubclassWindow = 0 Then
            RaiseEvent DownloadFailed(dfCannotCreateWindow, ID)
            Exit Sub
        End If
        'Subclass that window
        hPrevProc = SetWindowLong(SubclassWindow, GWL_WNDPROC, AddressOf DataProc)
    End If
    'Add this DataConnection class to the connection pool
    AddConnection Me
End Sub
Private Sub Class_Terminate()
    'close the socket, it here is one
    closesocket theSocket
    'close the file handle, if there is one
    CloseHandle FileHandle
    'clear the download buffer
    ClearBuffer
    'remove this class from the connection pool
    RemoveConnection Me
    'if every connection is destroyed...
    If ConnectionCount = 0 Then
        '... destroy the subclass window too
        SetWindowLong SubclassWindow, GWL_WNDPROC, hPrevProc
        DestroyWindow SubclassWindow
        SubclassWindow = 0
    End If
End Sub
'Public Sub FetchURL()
'  Start downloading data from the specified URL.
'Parameters: NONE
'Return Value: NONE
Public Sub FetchURL()
    ProcessedCloseMessage = False
    'Header must be parsed
    ParseHeader = True
    Header = ""
    'We're not yet SOCKS authenticated
    SocksAuthenticated = False
    m_ReceivedBytes = 0
    m_DownloadLength = -1
    SocksStep = 0
    m_HTTPReply = 0
    UserCancelled = False
    ClearBuffer
    'If ResumeFrom is set to byte 0, turn off UseResume
    If ResumeFrom = 0 Then UseResume = False
    'If we're downloading to a file...
    If DownloadType = dtToFile And Not Moving Then
        '... open/create it
        If m_UseTempFile Then Filename = GetRandomFilename
        If UseResume Then
            FileHandle = CreateFile(Filename, GENERIC_WRITE, 0, ByVal 0&, OPEN_ALWAYS, 0, ByVal 0&)
        Else
            FileHandle = CreateFile(Filename, GENERIC_WRITE, 0, ByVal 0&, CREATE_ALWAYS, 0, ByVal 0&)
        End If
        If FileHandle = INVALID_HANDLE_VALUE Then
            RaiseEvent DownloadFailed(dfCannotCreateFile, ID)
            Exit Sub
        End If
    End If
    If theSocket <> INVALID_SOCKET Then closesocket theSocket: DoEvents
    'Connect to the remote server
    If m_ProxyType = ptNoProxy Then
        theSocket = ConnectSock(Hostname, Port, 0, SubclassWindow)
    Else 'Catch all proxys
        theSocket = ConnectSock(ProxyHostname, ProxyPort, 0, SubclassWindow)
    End If
    If theSocket = INVALID_SOCKET Then
        CloseHandle FileHandle
        If m_ProxyType = ptNoProxy Then
            RaiseEvent DownloadFailed(dfCannotConnect, ID)
        Else
            RaiseEvent DownloadFailed(dfCannotConnectToProxy, ID)
        End If
        Exit Sub
    Else
        m_IsConnected = True
        If Not (Moving) Then RaiseEvent Connected(ID)
    End If
End Sub
'Public Sub FetchURLString(ByVal sURL As String)
'  Start downloading data from the specified URL. The URL must contain
'  the protocol ('http://') and the hostname. Optionally, it can contain the
'  port number. If no port is specified, the Port property will be set to
'  the default port 80.
'  Previous URL or Port properties settings will be ignored.
'  When an invalid URL is specified, the DownloadFailed event will
'  be raised with the Why parameter set to dfInvalidURL
'Parameters: URL to download from
'Return Value: NONE
Public Sub FetchURLString(ByVal sURL As String)
    Dim ZeroPos As Long, sHost As String, lPort As Long
    Dim sTURL As String
    sTURL = sURL
    sURL = LCase$(sURL) & "/"
    ZeroPos = InStr(1, sURL, "://")
    If ZeroPos > 0 Then
        Select Case Left$(sURL, ZeroPos - 1)
            Case "http"
                ZeroPos = InStr(8, sURL, "/")
                If ZeroPos = 0 Then
                    ZeroPos = Len(sURL)
                End If
                sURL = Mid$(sURL, 8, ZeroPos - 8)
                ZeroPos = InStr(1, sURL, ":")
                If ZeroPos > 0 Then
                    lPort = Val(Right$(sURL, Len(sURL) - ZeroPos))
                    If lPort <= 0 Then
                        RaiseEvent DownloadFailed(dfInvalidURL, ID)
                        Exit Sub
                    End If
                    sURL = Left$(sURL, ZeroPos - 1)
                Else
                    lPort = 80
                End If
                If sURL <> "" Then
                    URL = sTURL
                    Hostname = sURL
                    Port = lPort
                Else
                    RaiseEvent DownloadFailed(dfInvalidURL, ID)
                End If
            Case Else
                RaiseEvent DownloadFailed(dfInvalidURL, ID)
                Exit Sub
        End Select
    Else
        RaiseEvent DownloadFailed(dfInvalidURL, ID)
        Exit Sub
    End If
    FetchURL
End Sub
'Public Function GetBuffer(bBytes() As Byte) As Long
'  Returns a byte array with the data of the buffer. This byte array starts from
'  index 0
'Parameters: Byte array that receives the data of the buffer.
'Return Value: Number of bytes in buffer
Public Function GetBuffer(bBytes() As Byte) As Long
    If DataInBuffer > 0 Then
        ReDim Preserve DataBuffer(0 To DataInBuffer - 1) As Byte
        DataBufferSize = DataInBuffer
        bBytes() = DataBuffer()
        GetBuffer = DataInBuffer
    End If
End Function
'Public Function GetBufferAsString() As String
'  Returns a string with the data of the buffer.
'Parameters: NONE
'Return Value: the data
Public Function GetBufferAsString() As String
    If DataInBuffer > 0 Then
        GetBufferAsString = Left$(StrConv(DataBuffer(), vbUnicode), DataInBuffer)
    End If
End Function
'Public Function GetBufferSize() As Long
'  Returns the number of bytes in the buffer.
'Parameters: NONE
'Return Value: the number of bytes in the buffer
Public Function GetBufferSize() As Long
    GetBufferSize = DataInBuffer
End Function
'Public Sub ClearBuffer()
'  The ClearBuffer method clears the buffer that stores data when the DownloadType
'  is set to dtToBuffer. It may be useful to clear this buffer after downloading
'  a large file, to avoid unnecessary memory usage.
'Parameters: NONE
'Return Value: NONE
Public Sub ClearBuffer()
    ReDim Preserve DataBuffer(0 To 0) As Byte
    DataBufferSize = 1
    DataInBuffer = 0
End Sub
'Public Sub Disconnect()
'  Closes the socket and also closes the file handle if the DownloadType is set
'  to dtToFile.
'Parameters: NONE
'Return Value: NONE
Public Sub Disconnect()
    If IsConnected Then
        UserCancelled = True
        Moving = False
        ShutDown theSocket, SD_BOTH
        ProcessMessage FD_CLOSE
        closesocket theSocket
        theSocket = 0
        m_IsConnected = False
    End If
End Sub
'DownloadType [DOWNLOAD_TYPE][Get/Let][Default: dtToFile]
'  Specifies what should be done when data arrives:
'    � dtToFile
'       All incoming data is written to a file. The filename is specified
'       by the Filename property and/or the UseTempFile property.
'    � dtToBuffer
'       All incoming data is written to a buffer. The client may query
'       the contents of the buffer with the GetBuffer or GetBufferAsString
'       function.
'    � dtStream
'       Whenever new data is available, the event StreamBytes is raised.
'       One of the parameters of this event is a byte buffer with the new
'       data.
Public Property Get DownloadType() As DOWNLOAD_TYPE
    DownloadType = m_DownloadType
End Property
Public Property Let DownloadType(NewType As DOWNLOAD_TYPE)
    m_DownloadType = NewType
End Property
'Filename [String][Get/Let][Default: Empty String]
'  Specifies the filename where the data should be saved to. This
'   property is ignored when the DownloadType property is not set to
'   dtToFile or when UseTempFile is set to True.
Public Property Get Filename() As String
    Filename = m_Filename
End Property
Public Property Let Filename(NewFile As String)
    m_Filename = NewFile
End Property
'UseTempFile [Boolean][Get/Let][Default: True]
'  Specifies whether or not to generate a temporary file. If this option is set
'   to True, the class will create a temporary file in the Windows temporary
'   directory and adjust the Filename property accordingly. The client application
'   is responsible for deleting the temporary files after downloading them.
Public Property Get UseTempFile() As Boolean
    UseTempFile = m_UseTempFile
End Property
Public Property Let UseTempFile(NewTemp As Boolean)
    m_UseTempFile = NewTemp
End Property
'ProxyType [PROXY_TYPE][Get/Let][Default: ptNoProxy]
'  Specifies the type of proxy to use.
'    � ptNoProxy
'       No proxy will be used when connecting to the hostname.
'    � ptStandardProxy
'       A 'normal' proxy will be used when connecting to the host.
'    � ptSOCKS4Proxy
'       A SOCKS4 proxy will be used when connecting to the host.
'    � ptSOCKS5Proxy
'       A SOCKS5 proxy will be used when connecting to the host.
Public Property Get ProxyType() As PROXY_TYPE
    ProxyType = m_ProxyType
End Property
Public Property Let ProxyType(NewType As PROXY_TYPE)
    m_ProxyType = NewType
End Property
'ProxyHostname [String][Get/Let][Default: 'localhost']
'  Specifies the hostname or IP address of the proxy server.
Public Property Get ProxyHostname() As String
    ProxyHostname = m_ProxyHost
End Property
Public Property Let ProxyHostname(NewHost As String)
    m_ProxyHost = NewHost
End Property
'ProxyPort [Long][Get/Let][Default: 80]
'  Specifies the port where the class should connect on when it's connecting
'   to the proxy server. If you're using a SOCKS proxy, this value should
'   probably be set to 1080.
Public Property Get ProxyPort() As Long
    ProxyPort = m_ProxyPort
End Property
Public Property Let ProxyPort(NewPort As Long)
    m_ProxyPort = NewPort
End Property
'Hostname [String][Get/Let][Default: Empty String]
'  Specifies the hostname or the IP address the Download Class should
'   connect to. For instance, if you want to download the file
'   'http://www.allapi.net/downloads/myfile.ext', the Hostname property
'   should be set to 'www.allapi.net'.
Public Property Get Hostname() As String
    Hostname = m_Host
End Property
Public Property Let Hostname(NewHost As String)
    m_Host = NewHost
End Property
'Port [Long][Get/Let][Default: 80]
'  Specifies the remote port that will be used to connect to the server.
Public Property Get Port() As Long
    Port = m_Port
End Property
Public Property Let Port(NewPort As Long)
    m_Port = NewPort
End Property
'URL [String][Get/Let][Default: Empty String]
'  Specifies the URL of the data to download.
'   Note that this URL can change during the download if you enabled the
'   AutoRedirect property.
Public Property Get URL() As String
    URL = m_URL
End Property
Public Property Let URL(NewURL As String)
    m_URL = NewURL
End Property
'HttpUser [String][Get/Let][Default: Empty String]
'  Specifies the Username to use when HTTP Authentication is required
Public Property Get HttpUser() As String
    HttpUser = m_HttpUser
End Property
Public Property Let HttpUser(NewUser As String)
    m_HttpUser = NewUser
End Property
'HttpPass [String][Get/Let][Default: Empty String]
'  Specifies the Password to use when HTTP Authentication is required
Public Property Get HttpPass() As String
    HttpPass = m_HttpPass
End Property
Public Property Let HttpPass(NewPass As String)
    m_HttpPass = NewPass
End Property
'UseHttpAuthentication [Boolean][Get/Let][Default: False]
'  Specifies whether to send the HttpUser/HttpPass along with
'  the HTTP request.
Public Property Get UseHttpAuthorization() As Boolean
    UseHttpAuthorization = m_UseHttpAuthorization
End Property
Public Property Let UseHttpAuthorization(NewAuth As Boolean)
    m_UseHttpAuthorization = NewAuth
End Property
'PacketSize [Long][Get/Let][Default: 1024]
'  Specifies the preferred size of retrieving new data. This value has
'   to be positive.
Public Property Get PacketSize() As Long
    PacketSize = m_PacketSize
End Property
Public Property Let PacketSize(NewSize As Long)
    If NewSize > 0 Then m_PacketSize = NewSize
End Property
'ConnectionType [CONNECTION_TYPE][Get/Let][Default: ctHTTP]
'  Specifies the protocol to be used to download the requested file.
'  Currently, only ctHTTP is supported. Future versions may include
'  ctFTP or other protocols.
Public Property Get ConnectionType() As CONNECTION_TYPE
    ConnectionType = m_ConnectionType
End Property
Public Property Let ConnectionType(NewType As CONNECTION_TYPE)
    m_ConnectionType = NewType
End Property
'ID [Long][Get/Let][Default: 0]
'  Specifies a value that stores any extra data needed for your program.
'  Unlike other properties, the value of the ID property isn't used by
'  the Download Class. You can use this property to identify objects.
Public Property Get ID() As Long
    ID = m_ID
End Property
Public Property Let ID(NewID As Long)
    m_ID = NewID
End Property
'AllowCache [Boolean][Get/Let][Default: False]
'  Specifies whether proxies should return a cached version of the
'  requested file if one is available. If set to False, no cached
'  versions are allowed.
Public Property Get AllowCache() As Boolean
    AllowCache = m_AllowCache
End Property
Public Property Let AllowCache(NewAllow As Boolean)
    m_AllowCache = NewAllow
End Property
'ProxyUsername [String][Get/Let][Default: Empty String]
'  Specifies the username to be used when authenticating on a SOCKS proxy.
Public Property Get ProxyUsername() As String
    ProxyUsername = m_SocksUsername
End Property
Public Property Let ProxyUsername(NewUser As String)
    m_SocksUsername = NewUser
End Property
'ProxyPassword [String][Get/Let][Default: Empty String]
'  Specifies the password that should be used when logging in on a SOCKS5
'  proxy server.
Public Property Get ProxyPassword() As String
    ProxyPassword = m_SocksPassword
End Property
Public Property Let ProxyPassword(NewPass As String)
    m_SocksPassword = NewPass
End Property
'ProxyUseRemoteDNS [Boolean][Get/Let][Default: False]
'  Specifies whether domain names should be resolved locally or on the proxy
'  server. If this value is set to True, domain names are resolved on the
'  proxy server, in the other case, domain names are resolved on the local
'  machine. (Warning: not all SOCKS4 proxy servers support Remote DNS!)
Public Property Get ProxyUseRemoteDNS() As Boolean
    ProxyUseRemoteDNS = m_SocksUseRemoteDNS
End Property
Public Property Let ProxyUseRemoteDNS(NewDNS As Boolean)
    m_SocksUseRemoteDNS = NewDNS
End Property
'ProxyAuthentication [PROXY_AUTHENTICATION][Get/Let][Default: paNone]
'  Specifies the authentication scheme that should be used when connecting
'  to a SOCKS5 proxy server.
'   � paNone
'      No authentication is used.
'   � paUserPass
'      Use user/password authentication.
Public Property Get ProxyAuthentication() As PROXY_AUTHENTICATION
    ProxyAuthentication = m_ProxyAuthentication
End Property
Public Property Let ProxyAuthentication(NewAuthentication As PROXY_AUTHENTICATION)
    m_ProxyAuthentication = NewAuthentication
End Property
'IsConnected [Boolean][Get][Default: N/A]
'  Returns True if the class is connected to a server or False in
'  the other case.
Public Property Get IsConnected() As Boolean
    IsConnected = m_IsConnected
End Property
'ReceivedBytes [Long][Get][Default: 0]
'  Specifies the number of received bytes from the host. If you're using
'  the resume function, this value will start from the ResumeFrom value.
Public Property Get ReceivedBytes() As Long
    If MaxDownload >= 0 Then
        ReceivedBytes = IIf(m_ReceivedBytes > MaxDownload, MaxDownload, m_ReceivedBytes)
    Else
        ReceivedBytes = m_ReceivedBytes
    End If
End Property
'RealReceivedBytes [Long][Get][Default: 0]
'  Specifies the number of received bytes from the host.
Public Property Get RealReceivedBytes() As Long
    If UseResume Then
        RealReceivedBytes = ReceivedBytes - ResumeFrom
    Else
        RealReceivedBytes = ReceivedBytes
    End If
End Property
'MaxDownload [Long][Get/Let][Default: -1]
'  Specifies the maximum bytes to download. If this number is reached, the connection
'  will be automatically closed.
Public Property Get MaxDownload() As Long
    MaxDownload = m_MaxDownload
End Property
Public Property Let MaxDownload(NewMax As Long)
    m_MaxDownload = NewMax
End Property
'DownloadLength [Long][Get][Default: -1]
'  Returns the size of the file that is currently being downloaded,
'  or -1 if the size is unknown (this is often the case with scripts).
Public Property Get DownloadLength() As Long
    DownloadLength = m_DownloadLength
End Property
'HTTPReply [Long][Get][Default: N/A]
'  Returns the status code of the current HTTP connection. Usually
'  this is either '200' (OK) or '206' (Partial Data, when you're
'  resuming a file download) or '404' (file not found). For a full
'  list of possible status codes, take a look at the HTTP protocol
'  at http://www.w3c.org.
'  If this property returns '999', the downloaded HTTP header was
'  invalid.
Public Property Get HTTPReply() As Long
    HTTPReply = m_HTTPReply
End Property
'UseResume [Boolean][Get/Let][Default: False]
'  Specifies whether or not to resume a download.
Public Property Get UseResume() As Boolean
    UseResume = m_UseResume
End Property
Public Property Let UseResume(NewResume As Boolean)
    m_UseResume = NewResume
End Property
'ResumeFrom [Long][Get/Let][Default: 0]
'  Specifies the position in the remote file where to start downloading from.
'  The first byte in a file is byte 0, not byte 1. Hence, if you downloaded
'  205 bytes from a remote file and you want to resume the download, you should
'  set this value to '205'.
Public Property Get ResumeFrom() As Long
    ResumeFrom = m_ResumeFrom
End Property
Public Property Let ResumeFrom(NewResume As Long)
    If NewResume < 0 Then NewResume = 0
    m_ResumeFrom = NewResume
End Property
'AutoRedirect [Boolean][Get/Let][Default: True]
'  If the server returns a 301 or 302 status code (which means
'  the requested resource has been, respectively, Permantently
'  and Temporarily moved), the class automatically changes the
'  URL property and restarts the download. (Do note that, if it
'  restarts the download, the Connected and Disconnected events
'  will not be raised)
Public Property Get AutoRedirect() As Boolean
    AutoRedirect = m_AutoRedirect
End Property
Public Property Let AutoRedirect(NewRedirect As Boolean)
    m_AutoRedirect = NewRedirect
End Property
'SocketHandle [Long][Get][Default: N/A]
'  Returns the handle to the socket. In most cases, this handle is
'  not needed by your application.
Public Property Get SocketHandle() As Long
    SocketHandle = theSocket
End Property
'Public Sub ProcessMessage(Message As Long)
'  Processes a WinSock message, sent to the subclass window.
'  This function may *NOT* be used by your application!
'  It is only needed by the DataConnection module
'Parameters: Message to be processed
'Return Value: NONE
Public Sub ProcessMessage(Message As Long)
    Dim bSocks() As Byte, RemoteAddress As Long
    Dim HttpQuery As String
    Dim x As Long, bBytes() As Byte
    Select Case Message
        Case FD_CONNECT
            'Connect event
        Case FD_WRITE
            If Not m_IsConnected Then Exit Sub
            'If we're not yet authenticated and we're using a SOCKS4 proxy then...
            If Not (SocksAuthenticated) And (ProxyType = ptSOCKS4Proxy) Then
                'For more information about this, please read the SOCKS4 protocol
                '(http://www.socks.nec.com/protocol/socks4.protocol and
                ' http://www.socks.nec.com/protocol/socks4a.protocol)
                If ProxyUseRemoteDNS Then
                    ReDim bSocks(0 To Len(ProxyUsername) + Len(Hostname) + 9) As Byte
                Else
                    ReDim bSocks(0 To Len(ProxyUsername) + 8) As Byte
                    RemoteAddress = GetHostByNameAlias(Hostname)
                End If
                bSocks(0) = 4 'SOCKS version 4
                bSocks(1) = 1 'CONNECT
                bSocks(2) = Int(m_Port / 256)
                bSocks(3) = m_Port Mod 256
                If ProxyUseRemoteDNS Then
                    'bSocks(4 - 6) = 0
                    bSocks(7) = 1 'not zero
                Else
                    CopyMemory bSocks(4), RemoteAddress, 4
                End If
                CopyMemory bSocks(8), ByVal ProxyUsername, Len(ProxyUsername)
                If ProxyUseRemoteDNS Then
                    CopyMemory bSocks(9 + Len(ProxyUsername)), ByVal Hostname, Len(Hostname)
                End If
                If SendData(theSocket, bSocks()) = SOCKET_ERROR Then
                    RaiseEvent DownloadFailed(dfCannotConnectToProxy, ID)
                End If
            'If we're not yet authenticated and we're using a SOCKS5 proxy then...
            ElseIf Not (SocksAuthenticated) And (ProxyType = ptSOCKS5Proxy) Then
                'For more information about this, please read the SOCKS5 RFCs
                ' (RFC1928 and RFC1929)
                If SocksStep = 0 Then
                    ReDim bSocks(0 To IIf(ProxyAuthentication = paNone, 2, 3)) As Byte
                    bSocks(0) = 5 'SOCKS version 5
                    bSocks(1) = IIf(ProxyAuthentication = paNone, 1, 2) 'n method(s) listed
                    bSocks(2) = ProxyAuthentication 'SOCKS version 5
                    'bSocks(3) = 0
                ElseIf (SocksStep = 1) And (ProxyAuthentication = paUserPass) Then
                    ReDim bSocks(0 To 2 + Len(ProxyUsername) + Len(ProxyPassword)) As Byte
                    bSocks(0) = 1 'Subnegotiation protocol version 1
                    bSocks(1) = Len(ProxyUsername)
                    CopyMemory bSocks(2), ByVal ProxyUsername, Len(ProxyUsername)
                    bSocks(2 + Len(ProxyUsername)) = Len(ProxyPassword)
                    CopyMemory bSocks(3 + Len(ProxyUsername)), ByVal ProxyPassword, Len(ProxyPassword)
                ElseIf ((SocksStep = 1) And (ProxyAuthentication = paNone)) Or ((SocksStep = 2) And (ProxyAuthentication = paUserPass)) Then
                    If ProxyUseRemoteDNS Then
                        ReDim bSocks(0 To 6 + Len(Hostname)) As Byte
                        bSocks(0) = 5 'SOCKS version 5
                        bSocks(1) = 1 'CONNECT
                        bSocks(2) = 0 'Reserved
                        bSocks(3) = 3 'Address = Domain Name
                        bSocks(4) = Len(Hostname)
                        CopyMemory bSocks(5), ByVal Hostname, Len(Hostname)
                        bSocks(5 + Len(Hostname)) = Int(m_Port / 256)
                        bSocks(6 + Len(Hostname)) = m_Port Mod 256
                    Else
                        ReDim bSocks(0 To 9) As Byte
                        bSocks(0) = 5 'SOCKS version 5
                        bSocks(1) = 1 'CONNECT
                        bSocks(2) = 0 'Reserved
                        bSocks(3) = 1 'Address = IP address
                        RemoteAddress = GetHostByNameAlias(Hostname)
                        CopyMemory bSocks(4), RemoteAddress, 4
                        bSocks(8) = Int(m_Port / 256)
                        bSocks(9) = m_Port Mod 256
                    End If
                End If
                'Send the data
                If SendData(theSocket, bSocks()) = SOCKET_ERROR Then
                    RaiseEvent DownloadFailed(dfCannotConnectToProxy, ID)
                End If
            'If we're not using a proxy, or if we're already authenticated
            Else
                'If UseResume Then
                    'HttpQuery = "GET " + URL + " HTTP/1.1" + vbCrLf
                'Else
                    HttpQuery = "GET " + URL + " HTTP/1.0" + vbCrLf
                'End If
                HttpQuery = HttpQuery + "Accept: */*" + vbCrLf + "User-Agent: AllAPI.net Download Class" + vbCrLf + "Host: " + Hostname + vbCrLf
                'Resume file, if UseResume is enabled
                If UseResume Then
                    HttpQuery = HttpQuery + "Range: bytes=" + CStr(ResumeFrom) + "-" + vbCrLf
                End If
                If ProxyType = ptStandardProxy Then
                    HttpQuery = HttpQuery + "Proxy-Connection: Keep-Alive" + vbCrLf
                End If
                'Disable cache, if AllowCache is false
                If Not (AllowCache) Then
                    HttpQuery = HttpQuery + "Pragma: no-cache" + vbCrLf + "Cache-Control: no-cache" + vbCrLf
                End If
                'Check for HTTP Authentication
                If UseHttpAuthorization And Not (HttpPass = "" And HttpUser = "") Then
                    HttpQuery = HttpQuery + "Authorization: Basic " + ToBase64(HttpUser & ":" & HttpPass) + vbCrLf
                End If
                'Terminate the request with carriage return and line feed
                HttpQuery = HttpQuery + vbCrLf
                'Send the request
                If SendData(theSocket, HttpQuery) = SOCKET_ERROR Then
                    m_IsConnected = False
                    closesocket theSocket
                    RaiseEvent DownloadFailed(dfCannotSendData, ID)
                Else
                    RaiseEvent QuerySent(ID)
                End If
            End If
        Case FD_READ
            If Not m_IsConnected Then Exit Sub
            ReDim bBytes(0 To PacketSize - 1) As Byte
            'If we're not yet authenticated, it's a reply from the proxy server
            If (Not (SocksAuthenticated)) And ((ProxyType = ptSOCKS4Proxy) Or (ProxyType = ptSOCKS5Proxy)) Then
                'Let's see what the proxy has to say...
                x = recv(theSocket, bBytes(0), PacketSize, 0)
                If x > 0 Then
                    'For more info, read the SOCKS protocols
                    ReDim Preserve bBytes(0 To x - 1) As Byte
                    If ProxyType = ptSOCKS4Proxy Then
                        If bBytes(1) <> 90 Then
                            RaiseEvent DownloadFailed(dfProxyRequestRejectedOrFailed, ID)
                            closesocket theSocket
                            Exit Sub
                        End If
                        SocksAuthenticated = True
                    Else
                        If SocksStep = 0 Then
                            If bBytes(1) = 255 Then
                                RaiseEvent DownloadFailed(dfProxyAuthMethodNotAccepted, ID)
                                closesocket theSocket
                                Exit Sub
                            End If
                            ProxyAuthentication = bBytes(1)
                        ElseIf (SocksStep = 1) And (ProxyAuthentication = paUserPass) Then
                            If bBytes(1) <> 0 Then
                                RaiseEvent DownloadFailed(dfUserPasswordNotAccepted, ID)
                                closesocket theSocket
                                Exit Sub
                            End If
                        ElseIf ((SocksStep = 1) And (ProxyAuthentication = paNone)) Or ((SocksStep = 2) And (ProxyAuthentication = paUserPass)) Then
                            If bBytes(1) <> 0 Then
                                RaiseEvent DownloadFailed(dfProxyRequestRejectedOrFailed, ID)
                                closesocket theSocket
                                Exit Sub
                            End If
                            SocksAuthenticated = True
                        End If
                        SocksStep = SocksStep + 1
                    End If
                    ProcessMessage FD_WRITE
                End If
            'we received a message from our remote server
            Else
                Do
                    'read the bytes
                    x = recv(theSocket, bBytes(0), PacketSize, 0)
                    If x > 0 Then
                        'If the header hasn't been parsed yet, ...
                        If ParseHeader Then
                            '... parse it
                            Dim CrLfPos As Long, TempHead As String
                            'Convert the received bytes to a String
                            TempHead = Header & StrConv(bBytes, vbUnicode)
                            'Find the header terminator (vbCrLf & vbCrLf)
                            CrLfPos = InStr(1, TempHead, vbCrLf & vbCrLf, vbBinaryCompare)
                            'If we found the header terminator...
                            If CrLfPos > 0 Then
                                '... then extract and process the header
                                Header = Header & Left$(TempHead, CrLfPos + 3 - Len(Header))
                                'we don't have to parse the header anymore
                                ParseHeader = False
                                'process the header
                                ProcessHeader
                                x = x - CrLfPos - 3 - Len(Header)
                                'If we have already received some bytes of the URL, pass them to the user
                                If x > 0 And Not Moving Then
                                    Dim bTempBytes() As Byte
                                    ReDim bTempBytes(0 To x - 1)
                                    CopyMemory bTempBytes(0), bBytes(CrLfPos + 3), x
                                    SignalDataArrival bTempBytes(), x
                                End If
                            Else
                                'Header terminator not found... :(
                                Header = Header & TempHead
                            End If
                        Else
                            'We have received some bytes... pass them to the user
                            If x <> PacketSize Then ReDim Preserve bBytes(0 To x - 1)
                            SignalDataArrival bBytes(), x
                        End If
                    End If
                Loop Until x <> PacketSize
            End If
        Case FD_CLOSE
            If ProcessedCloseMessage Then Exit Sub
            'shut the socket down
            ShutDown theSocket, SD_SEND
            'before closing the socket, read from it
            ProcessMessage FD_READ
            'If we're not in the process of being redirected...
            If Not (Moving) Then
                '... then close the file handle
                CloseHandle FileHandle
                m_IsConnected = False
                ProcessedCloseMessage = True
                closesocket theSocket
                theSocket = INVALID_SOCKET
                'If we have received the entire file, then raise a Disconnected event
                If ((DownloadLength = -1) Or (DownloadLength = ReceivedBytes) Or ((ReceivedBytes >= MaxDownload) And (MaxDownload >= 0))) And Not UserCancelled Then
                    RaiseEvent Disconnected(ID)
                'If not, raise a DownloadFailed event
                Else
                    If UserCancelled Then
                        RaiseEvent DownloadFailed(dfCancelledByUser, ID)
                    Else
                        RaiseEvent DownloadFailed(dfConnectionLost, ID)
                    End If
                End If
            Else
                closesocket theSocket
            End If
    End Select
End Sub
'Private Sub SignalDataArrival(bBytes() As Byte, lBytes As Long)
'  Used internally to raise a BytesReceived (and a StreamBytes) event
'Parameters: Received bytes
'Return Value: NONE
Private Sub SignalDataArrival(bBytes() As Byte, lBytes As Long)
    Dim bCloseIt As Boolean
    If MaxDownload >= 0 Then
        If ReceivedBytes + lBytes >= MaxDownload Then
            lBytes = lBytes - (ReceivedBytes + lBytes - MaxDownload)
            If lBytes <= 0 Then
                ShutDown theSocket, SD_BOTH
                ProcessMessage FD_CLOSE
                closesocket theSocket
                m_IsConnected = False
                Exit Sub
            End If
            ReDim Preserve bBytes(0 To lBytes - 1) As Byte
        End If
    End If
    m_ReceivedBytes = m_ReceivedBytes + lBytes
    If DownloadType = dtToFile Then
        'Write buffer to file
        Dim Ret As Long
        WriteFile FileHandle, bBytes(0), lBytes, Ret, ByVal 0&
    ElseIf DownloadType = dtStream Then
        RaiseEvent StreamBytes(lBytes, bBytes(), ID)
    ElseIf DownloadType = dtToBuffer Then
        If DataInBuffer + lBytes > DataBufferSize Then
            'Increase the Buffer Size with the number of bytes plus 10% of the buffer
            DataBufferSize = DataBufferSize + lBytes + Int((DataBufferSize / 100) * 10)
            ReDim Preserve DataBuffer(0 To DataBufferSize - 1) As Byte
        End If
        CopyMemory DataBuffer(DataInBuffer), bBytes(0), lBytes
        DataInBuffer = DataInBuffer + lBytes
    End If
    RaiseEvent BytesReceived(ReceivedBytes, ID)
    If bCloseIt Then
        ShutDown theSocket, SD_BOTH
        ProcessMessage FD_CLOSE
        closesocket theSocket
        m_IsConnected = False
    End If
End Sub
'Private Sub ProcessHeader()
'  Used internally to process the header of the HTTP request
'Parameters: NONE
'Return Value: NONE
Private Sub ProcessHeader()
    Dim lPos1 As Long, lPos2 As Long
    Dim sLine As String, ColPos As Long
    Dim sFirstLine As String
    Moving = False
    sFirstLine = ReadNextLine(Header)
    'Process the HTTP header fields
    '(we only need Content-Length and Location)
    While LenB(Header) > 0
        sLine = ReadNextLine(Header)
        ColPos = InStr(1, sLine, ":", vbBinaryCompare)
        If ColPos > 0 Then
            Select Case LCase$(Left$(sLine, ColPos - 1))
                Case "content-length" 'Content-Length: 12345
                    m_DownloadLength = Val(Right$(sLine, Len(sLine) - 15))
                    If UseResume Then m_DownloadLength = m_DownloadLength + ResumeFrom
                Case "location" 'Location: www.newplace.com
                    NewLocation = Trim$(Right$(sLine, Len(sLine) - 9))
            End Select
        End If
    Wend
    'Process the first line of the reply
    lPos1 = InStr(1, sFirstLine, " ", vbBinaryCompare)
    lPos2 = InStr(lPos1 + 1, sFirstLine, " ", vbBinaryCompare)
    If lPos2 - lPos1 > 1 Then
        m_HTTPReply = Val(Mid$(sFirstLine, lPos1 + 1, lPos2 - lPos1 - 1))
        ' "Partial data" reply
        If m_HTTPReply = 206 Then
            m_ReceivedBytes = ResumeFrom
            If DownloadType = dtToFile Then
                If GetFileSize(FileHandle, ByVal 0&) >= ResumeFrom Then SetFilePointer FileHandle, ResumeFrom, ByVal 0&, FILE_BEGIN
            End If
        ' "OK" reply
        ElseIf m_HTTPReply = 200 Then
            UseResume = False
            If DownloadType = dtToFile Then SetFilePointer FileHandle, 0, ByVal 0&, FILE_BEGIN
        ' "Moved to" reply
        ElseIf (m_HTTPReply = 301 Or m_HTTPReply = 302) Then
            If AutoRedirect Then
                If LenB(NewLocation) > 0 Then
                    Moving = True
                    URL = NewLocation
                    FetchURL
                Else
                    RaiseEvent DownloadFailed(dfHTTPError, ID)
                End If
            ElseIf LenB(NewLocation) > 0 Then
                RaiseEvent ObjectMoved(NewLocation, ID)
            Else
                RaiseEvent DownloadFailed(dfHTTPError, ID)
            End If
        Else
            RaiseEvent DownloadFailed(dfHTTPError, ID)
        End If
    Else
        m_HTTPReply = 999
        RaiseEvent DownloadFailed(dfHTTPError, ID)
    End If
End Sub
'Private Function GetRndChar() As String
'  Used internally to generate a random Alpha Numeric character
'Parameters: NONE
'Return Value: Random Alpha Numeric character
Private Function GetRndChar() As String
    Randomize
    If Int(Rnd * 3) = 0 Then 'Numbers
        GetRndChar = Chr$(Int(Rnd * 10) + 48)
    Else
        GetRndChar = Chr$(Int(Rnd * 26) + 65)
        If Int(Rnd * 2) = 0 Then
            GetRndChar = LCase$(GetRndChar)
        End If
    End If
End Function
'Private Function GetRandomFilename() As String
'  Used internally to generate a random file name
'Parameters: NONE
'Return Value: Random filename
Private Function GetRandomFilename() As String
    Dim Ret As Long, TempFile As String
    'Get the Windows Temporary directory
    GetRandomFilename = Space(256)
    Ret = GetTempPath(Len(GetRandomFilename), GetRandomFilename)
    If Ret <= 0 Then Exit Function
    GetRandomFilename = Left$(GetRandomFilename, Ret)
    If Right$(GetRandomFilename, 1) <> "\" Then GetRandomFilename = GetRandomFilename + "\"
    'Find a temporary filename
    Do
        For Ret = 1 To 8
            TempFile = TempFile + GetRndChar
        Next Ret
        TempFile = TempFile + ".tmp"
        'Make sure the file doesn't already exist
        Ret = CreateFile(GetRandomFilename + TempFile, 0, 0, ByVal 0&, OPEN_EXISTING, 0, ByVal 0&)
        If Ret = INVALID_HANDLE_VALUE Then Exit Do
        CloseHandle Ret
    Loop
    GetRandomFilename = GetRandomFilename + TempFile
End Function
'Public Function FailureToString(Why As DOWNLOAD_FAILURE) As String
'  Used to convert a DOWNLOAD_FAILURE number to a String
'Parameters: The DOWNLOAD_FAILURE you wish to convert
'Return Value: String description of the failure
Public Function FailureToString(Why As DOWNLOAD_FAILURE) As String
    Select Case Why
        Case dfConnectionLost
            FailureToString = "The connection with the proxy or host has been closed, before downloading the entire file."
        Case dfCancelledByUser
            FailureToString = "The download has been aborted by the user."
        Case dfCannotConnect
            FailureToString = "Unable to connect to the host."
        Case dfCannotConnectToProxy
            FailureToString = "Unable to connect to the proxy or unable to send data to it."
        Case dfCannotCreateFile
            FailureToString = "Unable to create the file, specified in the Filename property."
        Case dfCannotCreateWindow
            FailureToString = "Unable to create a necessary window."
        Case dfCannotSendData
            FailureToString = "Unable to send data to the host or proxy"
        Case dfHTTPError
            FailureToString = "The server didn't reply with a 200 or 206 reply code."
        Case dfProxyAuthMethodNotAccepted
            FailureToString = "The selected authentication method was not accepted by the proxy."
        Case dfProxyRequestRejectedOrFailed
            FailureToString = "The connection request has failed or the proxy server has rejected it."
        Case dfUserPasswordNotAccepted
            FailureToString = "The user/password combination was not accepted by the proxy."
        Case Else
            FailureToString = "Unknown Error!"
    End Select
End Function
'Base64 functions
Public Function TripletToQuartet(bIn() As Byte, lStart As Long) As String
    Dim c As Byte
    c = ShiftRight(bIn(lStart), 2) And &H3F
    TripletToQuartet = TripletToQuartet & Mid(Base64Alphabet, c + 1, 1)
    c = ShiftLeft(bIn(lStart) And &H3, 4) Or (ShiftRight(bIn(lStart + 1), 4) And &HF)
    TripletToQuartet = TripletToQuartet & Mid(Base64Alphabet, c + 1, 1)
    c = ShiftLeft(bIn(lStart + 1) And &HF, 2) Or (ShiftRight(bIn(lStart + 2), 6) And &H3)
    TripletToQuartet = TripletToQuartet & Mid(Base64Alphabet, c + 1, 1)
    c = bIn(lStart + 2) And &H3F
    TripletToQuartet = TripletToQuartet & Mid(Base64Alphabet, c + 1, 1)
End Function
Public Function ToBase64(sNormString As String) As String
    Dim Cnt As Long, bBytes() As Byte
    bBytes = StrConv(sNormString, vbFromUnicode)
    For Cnt = 0 To Len(sNormString) - 3 Step 3
        ToBase64 = ToBase64 & TripletToQuartet(bBytes, Cnt)
    Next Cnt
    Select Case Len(sNormString) Mod 3
        Case 1
            ReDim Preserve bBytes(0 To UBound(bBytes) + 2)
            ToBase64 = ToBase64 & Left$(TripletToQuartet(bBytes, Len(sNormString) - 1), 2) & "=="
        Case 2
            ReDim Preserve bBytes(0 To UBound(bBytes) + 3)
            ToBase64 = ToBase64 & Left$(TripletToQuartet(bBytes, Len(sNormString)), 3) & "="
    End Select
End Function
Public Function ShiftLeft(ByVal bShift As Byte, lPlaces As Integer) As Byte
    If lPlaces >= 0 Then ShiftLeft = bShift * (2 ^ lPlaces) Mod 256
End Function
Public Function ShiftRight(ByVal bShift As Byte, lPlaces As Integer) As Byte
    If lPlaces >= 0 Then ShiftRight = bShift \ (2 ^ lPlaces)
End Function
'WinSock functions
'  Based on the mdlWinSock file
Private Function ConnectSock(ByVal Host As String, ByVal Port As Long, retIpPort As String, ByVal HWndToMsg As Long) As Long
    Dim s As Long, SelectOps As Long
    Dim sockin As sockaddr
    sockin = saZero
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_SOCKET Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    sockin.sin_addr = GetHostByNameAlias(Host)
    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    retIpPort = GetAscIP(sockin.sin_addr) & ":" & ntohs(sockin.sin_port)
    s = socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    If s < 0 Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If SetSockLinger(s, 1, 0) = SOCKET_ERROR Then
        If s > 0 Then
            Call closesocket(s)
        End If
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If Connect(s, sockin, sockaddr_size) <> 0 Then
        If s > 0 Then
            Call closesocket(s)
        End If
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
    If WSAAsyncSelect(s, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
        If s > 0 Then
            Call closesocket(s)
        End If
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    ConnectSock = s
End Function
Private Function SetSockLinger(ByVal SockNum As Long, ByVal OnOff As Integer, ByVal LingerTime As Integer) As Long
    Dim Linger As LingerType
    Linger.l_onoff = OnOff
    Linger.l_linger = LingerTime
    If setsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
        SetSockLinger = SOCKET_ERROR
    Else
        If getsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
            SetSockLinger = SOCKET_ERROR
        End If
    End If
End Function
Private Function GetAscIP(ByVal inn As Long) As String
    On Error Resume Next
    Dim lpStr As Long, nStr As Long
    Dim retString As String
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        GetAscIP = "255.255.255.255"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    GetAscIP = retString
    If Err Then GetAscIP = "255.255.255.255"
End Function
Private Function GetHostByNameAlias(ByVal Hostname As String) As Long
    On Error Resume Next
    'Return IP address as a long, in network byte order
    Dim phe As Long     ' pointer to host information entry
    Dim heDestHost As HostEnt 'hostent structure
    Dim addrList As Long
    Dim retIP As Long
    'first check to see if what we have been passed is a valid IP
    retIP = inet_addr(Hostname)
    If retIP = INADDR_NONE Then
        'it wasn't an IP, so do a DNS lookup
        phe = gethostbyname(Hostname)
        If phe <> 0 Then
            'Pointer is non-null, so copy in hostent structure
            MemCopy heDestHost, ByVal phe, hostent_size
            'Now get first pointer in address list
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
        Else
            'its not a valid address
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
    If Err Then GetHostByNameAlias = INADDR_NONE
End Function
Private Function SendData(ByVal s As Long, vMessage As Variant) As Long
    Dim TheMsg() As Byte, sTemp As String
    TheMsg = ""
    Select Case VarType(vMessage)
        Case vbArray + vbByte  'byte array
            sTemp = vMessage
            TheMsg = sTemp
        Case vbString      'string, if we recieve a string, its assumed we are linemode
            sTemp = StrConv(vMessage, vbFromUnicode)
        Case Else
            sTemp = CStr(vMessage)
            sTemp = StrConv(vMessage, vbFromUnicode)
    End Select
    TheMsg = sTemp
    If UBound(TheMsg) > -1 Then
        SendData = Send(s, TheMsg(0), (UBound(TheMsg) - LBound(TheMsg) + 1), 0)
    End If
End Function

