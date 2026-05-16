Attribute VB_Name = "modWinsockAPI"
Option Explicit

Public Const INADDR_NONE = &HFFFF
Public Const SOCKET_ERROR = -1
Public Const SOCKET_TIMEOUT = -2

' /*
' * All Windows Sockets error constants are biased by WSABASEERR from
' * the "normal"
' */

Private Const WSABASEERR = 10000

' /*
' * Windows Sockets definitions of regular Microsoft C error constants
' */

Public Const WSAEINTR = (WSABASEERR + 4)
Public Const WSAEBADF = (WSABASEERR + 9)
Public Const WSAEACCES = (WSABASEERR + 13)
Public Const WSAEFAULT = (WSABASEERR + 14)
Public Const WSAEINVAL = (WSABASEERR + 22)
Public Const WSAEMFILE = (WSABASEERR + 24)

' /*
' * Windows Sockets definitions of regular Berkeley error constants
' */

Public Const WSAEWOULDBLOCK = (WSABASEERR + 35)
Public Const WSAEINPROGRESS = (WSABASEERR + 36)
Public Const WSAEALREADY = (WSABASEERR + 37)
Public Const WSAENOTSOCK = (WSABASEERR + 38)
Public Const WSAEDESTADDRREQ = (WSABASEERR + 39)
Public Const WSAEMSGSIZE = (WSABASEERR + 40)
Public Const WSAEPROTOTYPE = (WSABASEERR + 41)
Public Const WSAENOPROTOOPT = (WSABASEERR + 42)
Public Const WSAEPROTONOSUPPORT = (WSABASEERR + 43)
Public Const WSAESOCKTNOSUPPORT = (WSABASEERR + 44)
Public Const WSAEOPNOTSUPP = (WSABASEERR + 45)
Public Const WSAEPFNOSUPPORT = (WSABASEERR + 46)
Public Const WSAEAFNOSUPPORT = (WSABASEERR + 47)
Public Const WSAEADDRINUSE = (WSABASEERR + 48)
Public Const WSAEADDRNOTAVAIL = (WSABASEERR + 49)
Public Const WSAENETDOWN = (WSABASEERR + 50)
Public Const WSAENETUNREACH = (WSABASEERR + 51)
Public Const WSAENETRESET = (WSABASEERR + 52)
Public Const WSAECONNABORTED = (WSABASEERR + 53)
Public Const WSAECONNRESET = (WSABASEERR + 54)
Public Const WSAENOBUFS = (WSABASEERR + 55)
Public Const WSAEISCONN = (WSABASEERR + 56)
Public Const WSAENOTCONN = (WSABASEERR + 57)
Public Const WSAESHUTDOWN = (WSABASEERR + 58)
Public Const WSAETOOMANYREFS = (WSABASEERR + 59)
Public Const WSAETIMEDOUT = (WSABASEERR + 60)
Public Const WSAECONNREFUSED = (WSABASEERR + 61)
Public Const WSAELOOP = (WSABASEERR + 62)
Public Const WSAENAMETOOLONG = (WSABASEERR + 63)
Public Const WSAEHOSTDOWN = (WSABASEERR + 64)
Public Const WSAEHOSTUNREACH = (WSABASEERR + 65)
Public Const WSAENOTEMPTY = (WSABASEERR + 66)
Public Const WSAEPROCLIM = (WSABASEERR + 67)
Public Const WSAEUSERS = (WSABASEERR + 68)
Public Const WSAEDQUOT = (WSABASEERR + 69)
Public Const WSAESTALE = (WSABASEERR + 70)
Public Const WSAEREMOTE = (WSABASEERR + 71)

' /*
' * Extended Windows Sockets error constant definitions
' */

Public Const WSASYSNOTREADY = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)
Public Const WSANOTINITIALISED = (WSABASEERR + 93)
Public Const WSAEDISCON = (WSABASEERR + 101)
Public Const WSAENOMORE = (WSABASEERR + 102)
Public Const WSAECANCELLED = (WSABASEERR + 103)
Public Const WSAEINVALIDPROCTABLE = (WSABASEERR + 104)
Public Const WSAEINVALIDPROVIDER = (WSABASEERR + 105)
Public Const WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)
Public Const WSASYSCALLFAILURE = (WSABASEERR + 107)
Public Const WSASERVICE_NOT_FOUND = (WSABASEERR + 108)
Public Const WSATYPE_NOT_FOUND = (WSABASEERR + 109)
Public Const WSA_E_NO_MORE = (WSABASEERR + 110)
Public Const WSA_E_CANCELLED = (WSABASEERR + 111)
Public Const WSAEREFUSED = (WSABASEERR + 112)

Public Const WSAHOST_NOT_FOUND = 11001
Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129
Public Const WSATRY_AGAIN = 11002
Public Const WSANO_RECOVERY = 11003
Public Const WSANO_DATA = 11004

Public Const SOL_SOCKET = 65535
Public Const SO_PROTOCOL_INFO = &H2004

Public Const FD_SETSIZE = 64

Public Type WSAData
      wVersion       As Integer
      wHighVersion   As Integer
      szDescription  As String * WSADESCRIPTION_LEN
      szSystemStatus As String * WSASYS_STATUS_LEN
      iMaxSockets    As Integer
      iMaxUdpDg      As Integer
      lpVendorInfo   As Long
End Type

Private Type HOSTENT
      hName     As Long
      hAliases  As Long
      hAddrType As Integer
      hLength   As Integer
      hAddrList As Long
End Type

Private Type servent
      s_name    As Long
      s_aliases As Long
      s_port    As Integer
      s_proto   As Long
End Type

Private Type protoent
      p_name    As String 'Official name of the protocol
      p_aliases As Long   'Null-terminated array of alternate names
      p_proto   As Long   'Protocol number, in host byte order
End Type

Private Type sockaddr_in
      sin_family       As Integer
      sin_port         As Integer
      sin_addr         As Long
      sin_zero(1 To 8) As Byte
End Type

Public Type timeval
      tv_sec  As Long   'seconds
      tv_usec As Long   'and microseconds
End Type

Public Type fd_set
      fd_count                  As Long '// how many are SET?
      fd_array(1 To FD_SETSIZE) As Long '// an array of SOCKETs
End Type

Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Private Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Private Declare Function gethostname Lib "ws2_32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
Private Declare Function getservbyname Lib "ws2_32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
Private Declare Function getprotobynumber Lib "ws2_32.dll" (ByVal proto As Long) As Long
Private Declare Function getprotobyname Lib "ws2_32.dll" (ByVal proto_name As String) As Long
Private Declare Function getservbyport Lib "ws2_32.dll" (ByVal Port As Integer, ByVal proto As Long) As Long
Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Private Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Private Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Private Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Private Declare Function Connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As sockaddr_in, ByVal namelen As Long) As Long
Private Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Private Declare Function getpeername Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Private Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Private Declare Function vbselect Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByRef readfds As Any, ByRef writefds As Any, ByRef exceptfds As Any, ByRef TimeOut As Long) As Long
Private Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, from As sockaddr_in, fromlen As Long) As Long
Private Declare Function send Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, to_addr As sockaddr_in, ByVal tolen As Long) As Long
Private Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As sockaddr_in, ByRef addrlen As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function getsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Any, ByRef optlen As Long) As Long

' /*
' * Address families.
' */

Public Enum AddressFamily
   AF_UNSPEC = 0      '/* unspecified */
' /*
' * Although  AF_UNSPEC  is  defined for backwards compatibility, using
' * AF_UNSPEC for the "af" parameter when creating a socket is STRONGLY
' * DISCOURAGED.    The  interpretation  of  the  "protocol"  parameter
' * depends  on the actual address family chosen.  As environments grow
' * to  include  more  and  more  address families that use overlapping
' * protocol  values  there  is  more  and  more  chance of choosing an
' * undesired address family when AF_UNSPEC is used.
' */
   AF_UNIX = 1           '/* local to host (pipes, portals) */
   AF_INET = 2        '/* internetwork: UDP, TCP, etc. */
   AF_IMPLINK = 3     '/* arpanet imp addresses */
   AF_PUP = 4         '/* pup protocols: e.g. BSP */
   AF_CHAOS = 5       '/* mit CHAOS protocols */
   AF_NS = 6          '/* XEROX NS protocols */
   AF_IPX = AF_NS     '/* IPX protocols: IPX, SPX, etc. */
   AF_ISO = 7         '/* ISO protocols */
   AF_OSI = AF_ISO    '/* OSI is ISO */
   AF_ECMA = 8        '/* european computer manufacturers */
   AF_DATAKIT = 9     '/* datakit protocols */
   AF_CCITT = 10      '/* CCITT protocols, X.25 etc */
   AF_SNA = 11        '/* IBM SNA */
   AF_DECnet = 12     '/* DECnet */
   AF_DLI = 13        '/* Direct data link interface */
   AF_LAT = 14        '/* LAT */
   AF_HYLINK = 15     '/* NSC Hyperchannel */
   AF_APPLETALK = 16  '/* AppleTalk */
   AF_NETBIOS = 17    '/* NetBios-style addresses */
   AF_VOICEVIEW = 18  '/* VoiceView */
   AF_FIREFOX = 19    '/* Protocols from Firefox */
   AF_UNKNOWN1 = 20   '/* Somebody is using this! */
   AF_BAN = 21        '/* Banyan */
   AF_ATM = 22        '/* Native ATM Services */
   AF_INET6 = 23      '/* Internetwork Version 6 */
   AF_CLUSTER = 24    '/* Microsoft Wolfpack */
   AF_12844 = 25      '/* IEEE 1284.4 WG AF */
   AF_MAX = 26
End Enum

' Socket types

Public Enum SocketType
   SOCK_STREAM = 1    ' /* stream socket */
   SOCK_DGRAM = 2     ' /* datagram socket */
   SOCK_RAW = 3       ' /* raw-protocol interface */
   SOCK_RDM = 4       ' /* reliably-delivered message */
   SOCK_SEQPACKET = 5 ' /* sequenced packet stream */
End Enum

' /*
' * Protocols
' */

Public Enum SocketProtocol
   IPPROTO_IP = 0             '/* dummy for IP */
   IPPROTO_ICMP = 1           '/* control message protocol */
   IPPROTO_IGMP = 2           '/* internet group management protocol */
   IPPROTO_GGP = 3            '/* gateway^2 (deprecated) */
   IPPROTO_TCP = 6            '/* tcp */
   IPPROTO_PUP = 12           '/* pup */
   IPPROTO_UDP = 17           '/* user datagram protocol */
   IPPROTO_IDP = 22           '/* xns idp */
   IPPROTO_ND = 77            '/* UNOFFICIAL net disk proto */
   IPPROTO_RAW = 255          '/* raw IP packet */
   IPPROTO_MAX = 256
End Enum

' Maximum queue length specifiable by listen.
Public Const SOMAXCONN = &H7FFFFFFF

Public Enum WinsockVersion
   SOCKET_VERSION_11 = &H101
   SOCKET_VERSION_22 = &H202
End Enum

Public Enum IPEndPointFields
   LOCAL_HOST
   LOCAL_HOST_IP
   LOCAL_PORT
   REMOTE_HOST
   REMOTE_HOST_IP
   REMOTE_PORT
End Enum

Private Const MAX_PROTOCOL_CHAIN    As Long = 6

Private Type WSAPROTOCOLCHAIN
   ChainLen                         As Long
   ChainEntries(MAX_PROTOCOL_CHAIN) As Long
End Type

Private Type Guid
   Data1         As Long
   Data2         As Integer
   Data3         As Integer
   Data4(0 To 7) As Byte
End Type

Private Const WSAPROTOCOL_LEN As Long = 256

Private Type WSAPROTOCOL_INFO
   dwServiceFlags1    As Long
   dwServiceFlags2    As Long
   dwServiceFlags3    As Long
   dwServiceFlags4    As Long
   dwProviderFlags    As Long
   ProviderId         As Guid
   dwCatalogEntryId   As Long
   ProtocolChain      As WSAPROTOCOLCHAIN
   iVersion           As Long
   iAddressFamily     As Long
   iMaxSockAddr       As Long
   iMinSockAddr       As Long
   iSocketType        As Long
   iProtocol          As Long
   iProtocolMaxOffset As Long
   iNetworkByteOrder  As Long
   iSecurityScheme    As Long
   dwMessageSize      As Long
   dwProviderReserved As Long
   szProtocol         As String * WSAPROTOCOL_LEN
End Type
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const DNS_RECURSION As Byte = 1

' /* DNS Lookup
Private Type DNS_HEADER
   qryID As Integer
   options As Byte
   response As Byte
   qdcount As Integer
   ancount As Integer
   nscount As Integer
   arcount As Integer
End Type

' get dns
Public Const MAX_HOSTNAME_LEN = 128
Public Const MAX_DOMAIN_NAME_LEN = 128
Public Const MAX_SCOPE_ID_LEN = 256
Public Const ERROR_SUCCESS As Long = 0

Private Type IP_ADDRESS_STRING
   IpAddr(0 To 15)     As Byte
End Type

Private Type IP_MASK_STRING
   IpMask(0 To 15)     As Byte
End Type

Private Type IP_ADDR_STRING
   dwNext               As Long
   IpAddress(0 To 15)   As Byte
   IpMask(0 To 15)      As Byte
   dwContext            As Long
End Type

Private Type FIXED_INFO
   HostName(0 To (MAX_HOSTNAME_LEN + 3))      As Byte
   DomainName(0 To (MAX_DOMAIN_NAME_LEN + 3)) As Byte
   CurrentDnsServer        As Long
   DnsServerList           As IP_ADDR_STRING
   NodeType                As Long
   ScopeId(0 To (MAX_SCOPE_ID_LEN + 3))       As Byte
   EnableRouting           As Long
   EnableProxy             As Long
   EnableDns               As Long
End Type

Public Type MX_RECORD
   Server                  As String
   Pref                    As Integer
End Type
Public Type MX_INFO
   Best                    As String
   Domain                  As String
   List()                  As MX_RECORD
   Count                   As Long
End Type

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767
Private Declare Function GetNetworkParams Lib "iphlpapi.dll" (pFixedInfo As Any, pOutBufLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' The function takes a Double containing a value in the 
' range of an unsigned Long and returns a Long that you 
' can pass to an API that requires an unsigned Long
Public Function UnsignedToLong(Value As Double) As Long
      If Value < 0 Or Value >= OFFSET_4 Then
         Error 6 ' Overflow
      End If

      If Value <= MAXINT_4 Then
         UnsignedToLong = Value
      Else
         UnsignedToLong = Value - OFFSET_4
      End If
End Function

' The function takes an unsigned Long from an API and 
' converts it to a Double for display or arithmetic purposes
Public Function LongToUnsigned(Value As Long) As Double
      If Value < 0 Then
         LongToUnsigned = Value + OFFSET_4
      Else
         LongToUnsigned = Value
      End If
End Function

' The function takes a Long containing a value in the range 
' of an unsigned Integer and returns an Integer that you 
' can pass to an API that requires an unsigned Integer
Public Function UnsignedToInteger(Value As Long) As Integer
      If Value < 0 Or Value >= OFFSET_2 Then
         Error 6 ' Overflow
      End If
      If Value <= MAXINT_2 Then
         UnsignedToInteger = Value
      Else
         UnsignedToInteger = Value - OFFSET_2
      End If
End Function

' The function takes an unsigned Integer from and API and 
' converts it to a Long for display or arithmetic purposes
Public Function IntegerToUnsigned(Value As Integer) As Long
      If Value < 0 Then
         IntegerToUnsigned = Value + OFFSET_2
      Else
         IntegerToUnsigned = Value
      End If
End Function

Public Function StringFromPointer(ByVal lPointer As Long) As String
Dim strTemp As String
Dim lRetVal As Long

      ' prepare the strTemp buffer
      strTemp = String$(lstrlen(ByVal lPointer), 0)
      ' copy the string into the strTemp buffer
      lRetVal = lstrcpy(ByVal strTemp, ByVal lPointer)
      ' return a string
      If lRetVal Then
         StringFromPointer = strTemp
      End If
End Function

Public Function GetAddressLong(ByVal strHostName As String) As Long
Dim lngPtrToHOSTENT As Long    ' pointer to HOSTENT structure returned by the gethostbyname function
Dim udtHostent      As HOSTENT ' structure which stores all the host info
Dim lngPtrToIP      As Long    ' pointer to the IP address list
Dim lngAddress      As Long

      lngAddress = inet_addr(strHostName)

      If lngAddress = INADDR_NONE Then

         lngPtrToHOSTENT = gethostbyname(strHostName)

         If lngPtrToHOSTENT <> 0 Then
            ' The gethostbyname function has found the address

            ' Copy retrieved data to udtHostent structure
            RtlMoveMemory udtHostent, lngPtrToHOSTENT, LenB(udtHostent)
 
            ' Now udtHostent.hAddrList member contains
            ' an array of IP addresses

            ' Get a pointer to the first address
            RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
            ' Get the address
            RtlMoveMemory lngAddress, lngPtrToIP, udtHostent.hLength
         Else
            lngAddress = INADDR_NONE
         End If
      End If
      GetAddressLong = lngAddress
End Function

Public Function GetErrorDescription(ByVal lngErrorCode As Long) As String
Dim strDesc As String
      Select Case lngErrorCode
         Case WSAEACCES
            strDesc = "Permission denied."
         Case WSAEADDRINUSE
            strDesc = "Address already in use."
         Case WSAEADDRNOTAVAIL
            strDesc = "Cannot assign requested address."
         Case WSAEAFNOSUPPORT
            strDesc = "Address family not supported by protocol family."
         Case WSAEALREADY
            strDesc = "Operation already in progress."
         Case WSAECONNABORTED
            strDesc = "Software caused connection abort."
         Case WSAECONNREFUSED
            strDesc = "Connection refused."
         Case WSAECONNRESET
            strDesc = "Connection reset by peer."
         Case WSAEDESTADDRREQ
            strDesc = "Destination address required."
         Case WSAEFAULT
            strDesc = "Bad address."
         Case WSAEHOSTDOWN
            strDesc = "Host is down."
         Case WSAEHOSTUNREACH
            strDesc = "No route to host."
         Case WSAEINPROGRESS
            strDesc = "Operation now in progress."
         Case WSAEINTR
            strDesc = "Interrupted function call."
         Case WSAEINVAL
            strDesc = "Invalid argument."
         Case WSAEISCONN
            strDesc = "Socket is already connected."
         Case WSAEMFILE
            strDesc = "Too many open files."
         Case WSAEMSGSIZE
            strDesc = "Message too long."
         Case WSAENETDOWN
            strDesc = "Network is down."
         Case WSAENETRESET
            strDesc = "Network dropped connection on reset."
         Case WSAENETUNREACH
            strDesc = "Network is unreachable."
         Case WSAENOBUFS
            strDesc = "No buffer space available."
         Case WSAENOPROTOOPT
            strDesc = "Bad protocol option."
         Case WSAENOTCONN
            strDesc = "Socket is not connected."
         Case WSAENOTSOCK
            strDesc = "Socket operation on nonsocket."
         Case WSAEOPNOTSUPP
            strDesc = "Operation not supported."
         Case WSAEPFNOSUPPORT
            strDesc = "Protocol family not supported."
         Case WSAEPROCLIM
            strDesc = "Too many processes."
         Case WSAEPROTONOSUPPORT
            strDesc = "Protocol not supported."
         Case WSAEPROTOTYPE
            strDesc = "Protocol wrong type for socket."
         Case WSAESHUTDOWN
            strDesc = "Cannot send after socket shutdown."
         Case WSAESOCKTNOSUPPORT
            strDesc = "Socket type not supported."
         Case WSAETIMEDOUT
            strDesc = "Connection timed out."
         Case WSATYPE_NOT_FOUND
            strDesc = "Class type not found."
         Case WSAEWOULDBLOCK
            strDesc = "Resource temporarily unavailable."
         Case WSAHOST_NOT_FOUND
            strDesc = "Host not found."
         Case WSANOTINITIALISED
            strDesc = "Successful WSAStartup not yet performed."
         Case WSANO_DATA
            strDesc = "Valid name, no data record of requested type."
         Case WSANO_RECOVERY
            strDesc = "This is a nonrecoverable error."
         Case WSASYSCALLFAILURE
            strDesc = "System call failure."
         Case WSASYSNOTREADY
            strDesc = "Network subsystem is unavailable."
         Case WSATRY_AGAIN
            strDesc = "Nonauthoritative host not found."
         Case WSAVERNOTSUPPORTED
            strDesc = "Winsock.dll version out of range."
         Case WSAEDISCON
            strDesc = "Graceful shutdown in progress."
         Case Else
            strDesc = "Unknown error."
      End Select
      GetErrorDescription = strDesc
End Function

' ********************************************************************************
' Author    :Oleg Gdalevich
' Purpose   :Initializes the Winsock service
' Returns   :Zero if successful. Otherwise, it returns error code.
' '          The description of the error can be retrieved with
' '          the GetErrorDescription function.
' Arguments :Version - Winsock's version: 1.1 or 2.2
' ********************************************************************************
Public Function InitializeWinsock(Version As WinsockVersion) As Long
Dim udtWinsockData  As WSAData
Dim lngRetValue     As Long

      ' start up winsock service
      lngRetValue = WSAStartup(Version, udtWinsockData)
      ' assign returned value
      InitializeWinsock = lngRetValue
End Function

' ********************************************************************************
' Author    :Darren Lawrence
' Date/Time :03-July-2005
' Purpose   :Determing the Sockets Address Family
' Returns   :AddressFamily if successful. Otherwise, it returns INADDR_NONE error code.
' Arguments :lngSocket - The Socket Handle
' ********************************************************************************
Public Function GetAddressFamily(lngSocket As Long) As AddressFamily
Dim lngAdrFamily     As Long
Dim lngBufferSize    As Long
Dim lngRetValue      As Long
Dim udtProtocolInfo  As WSAPROTOCOL_INFO

      lngBufferSize = LenB(udtProtocolInfo)
      lngRetValue = getsockopt(lngSocket, SOL_SOCKET, SO_PROTOCOL_INFO, udtProtocolInfo, lngBufferSize)
      
      If lngRetValue = 0 Then
         GetAddressFamily = udtProtocolInfo.iAddressFamily
      Else
         GetAddressFamily = INADDR_NONE
      End If
End Function

' ********************************************************************************
' Author    :Oleg Gdalevich
' Purpose   :Creates a new socket
' Returns   :The socket handle if successful, otherwise - INVALID_SOCKET
' Arguments :AdrFamily   - The Address Family
' '          SckType     - The Type of socket to Connect
' '          SckProtocol - Protocol to use
' ********************************************************************************
Public Function vbSocket(ByVal AdrFamily As AddressFamily, ByVal SckType As SocketType, ByVal SckProtocol As SocketProtocol) As Long
Dim lngRetValue As Long

      On Error GoTo vbSocket_Err_Handler

      ' Call the socket Winsock API function in order create a new socket
      lngRetValue = socket(AdrFamily, SckType, SckProtocol)
      ' Assign returned value
      vbSocket = lngRetValue

EXIT_LABEL:
      Exit Function

vbSocket_Err_Handler:
      vbSocket = SOCKET_ERROR
      
End Function

' ********************************************************************************
' Author    :Oleg Gdalevich
' Date/Time :05-Oct-2001
' Purpose   :Establishes connection to the remote host.
' Return    :If no error occurs, returns zero. Otherwise, it returns SOCKET_ERROR.
' Arguments :lngSocket     - the socket to establish the connection for.
' '          strRemoteHost - name or IP address of the host to connect to
' '          intRemotePort - the port number to connect to
' ********************************************************************************
Public Function vbConnect(ByVal lngSocket As Long, ByVal strRemoteHost As String, ByVal intRemotePort As Integer) As Long
Dim udtSocketAddress As sockaddr_in
Dim lngReturnValue   As Long
Dim lngAddress       As Long
Dim lngAdrFamily     As Long

      On Error GoTo ERROR_HANDLER

      vbConnect = SOCKET_ERROR

      ' Check the socket handle
      If Not lngSocket > 0 Then
         ' Failure
         Exit Function
      End If

      ' Check the remote host address argument
      If Len(strRemoteHost) = 0 Then
         ' Failure
         Exit Function
      End If

      ' Check the port number
      If Not intRemotePort > 0 Then
         ' Failure
         Exit Function
      End If

      ' Prepare the sockaddr_in structure to pass to the  connect Winsock API function
      lngAdrFamily = GetAddressFamily(lngSocket)
      If lngAdrFamily = INADDR_NONE Then
         ' Failure
         Exit Function
      End If

      ' the strRemoteHost may contain the host name  or IP address
      ' GetAddressLong returns a valid value anyway
      lngAddress = GetAddressLong(strRemoteHost)
      If lngAddress = INADDR_NONE Then
         ' Failure
         Exit Function
      End If

      With udtSocketAddress
         .sin_addr = lngAddress
         ' convert the port number to the network byte ordering
         .sin_port = htons(UnsignedToInteger(CLng(intRemotePort)))
         .sin_family = lngAdrFamily
      End With
      
      vbConnect = Connect(lngSocket, udtSocketAddress, LenB(udtSocketAddress))

EXIT_LABEL:
      Exit Function

ERROR_HANDLER:

      vbConnect = SOCKET_ERROR
End Function

' ********************************************************************************
' Author    :Oleg Gdalevich
' Date/Time :02-Nov-2001
' Purpose   :Binds the socket to the local address
' Return    :If no error occurs, returns zero. Otherwise, it returns SOCKET_ERROR.
' Arguments :lngSocket    - the socket to bind
' '          strLocalHost - name or IP address of the local host to bind to
' '          lngLocalPort - the port number to bind to
' ********************************************************************************
Public Function vbBind(ByVal lngSocket As Long, ByVal strLocalHost As String, ByVal lngLocalPort As Long) As Long
Dim udtSocketAddress As sockaddr_in
Dim lngReturnValue   As Long
Dim lngAddress       As Long
Dim lngAdrFamily     As Long

      On Error GoTo ERROR_HANDLER

      vbBind = SOCKET_ERROR
      ' Check the socket handle
      If Not lngSocket > 0 Then
         ' Failure
         Exit Function
      End If
      ' Check the local host address argument
      If Len(strLocalHost) = 0 Then
         ' Failure
         Exit Function
      End If
      ' Check the port number
      If Not lngLocalPort > 0 Then
         ' Failure
         Exit Function
      End If
      
      ' Prepare the sockaddr_in structure to pass to the
      ' bind Winsock API function

      ' The sin_family member of the structure needs
      ' the address family value that we can retieve
      ' with CProtocol class

      lngAdrFamily = GetAddressFamily(lngSocket)
      If lngAdrFamily = INADDR_NONE Then
         Exit Function
      End If

      With udtSocketAddress
         .sin_addr = lngAddress
         ' convert the port number to the network byte ordering
         .sin_port = htons(UnsignedToInteger(lngLocalPort))
         .sin_family = lngAdrFamily
      End With
      vbBind = bind(lngSocket, udtSocketAddress, LenB(udtSocketAddress))

EXIT_LABEL:
      Exit Function

ERROR_HANDLER:
      vbBind = SOCKET_ERROR
End Function

' ********************************************************************************
' Author    :Oleg Gdalevich
' Date/Time :21.10.01
' Purpose   :Retrieves IP address or host name or port number of
' '          an end-point of the connection established
' '          on the socket - lngSocket
'
' Return    :If no errors occures, the function returns the value
' '          requested by the EndPointField argument.
' '          Otherwise, it returns the value of SOCKET_ERROR
'
' Arguments :
' '          lngSocket -  socket's handle. The socket must be connected to the remote host.
' '          EndPointField - specifies the value to return:
' '                                                      LOCAL_HOST
' '                                                      LOCAL_HOST_IP
' '                                                      LOCAL_PORT
' '                                                      REMOTE_HOST
' '                                                      REMOTE_HOST_IP
' '                                                      REMOTE_PORT
' ********************************************************************************
Public Function GetIPEndPointField(ByVal lngSocket As Long, _
      ByVal EndPointField As IPEndPointFields) As Variant
      
Dim udtSocketAddress    As sockaddr_in
Dim lngReturnValue      As Long
Dim lngPtrToAddress     As Long

      On Error GoTo ERROR_HANDLER

      Select Case EndPointField
         Case LOCAL_HOST, LOCAL_HOST_IP, LOCAL_PORT
            ' If the info of a local end-point of the connection is
            ' requested, call the getsockname Winsock API function
            lngReturnValue = getsockname(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
         Case REMOTE_HOST, REMOTE_HOST_IP, REMOTE_PORT
            ' If the info of a remote end-point of the connection is
            ' requested, call the getpeername Winsock API function
            lngReturnValue = getpeername(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
      End Select

      If lngReturnValue = 0 Then
         ' If no errors were occurred, the getsockname or getpeername
         ' function returns 0.
         Select Case EndPointField
            Case LOCAL_PORT, REMOTE_PORT
            
               ' If the port number is requested, retrieve that value
               ' from the sin_port member of the udtSocketAddress
               ' structure, and change the byte order of that value from
               ' the network to host byte order.
               GetIPEndPointField = IntegerToUnsigned(ntohs(udtSocketAddress.sin_port))
            
            Case LOCAL_HOST_IP, REMOTE_HOST_IP

               ' The host address is stored in the sin_addr member of the structure
               ' as 4-byte value.

               ' To get an IP address of the host:

               ' Get pointer to the string that contains the IP address
               lngPtrToAddress = inet_ntoa(udtSocketAddress.sin_addr)
 
               ' Retrieve that string by the pointer
               GetIPEndPointField = StringFromPointer(lngPtrToAddress)

            Case LOCAL_HOST, REMOTE_HOST

               ' The same procedure as for an IP address.
               ' But here is the GetHostNameByAddress function call
               ' to retrieve host name by IP address.
               lngPtrToAddress = inet_ntoa(udtSocketAddress.sin_addr)
               GetIPEndPointField = GetHostNameByAddress(StringFromPointer(lngPtrToAddress))

         End Select
      Else
         GetIPEndPointField = SOCKET_ERROR
      End If

EXIT_LABEL:
      Exit Function

ERROR_HANDLER:
      GetIPEndPointField = SOCKET_ERROR
    
End Function

Private Function GetHostNameByAddress(strIpAddress As String) As String
Dim lngInetAdr As Long
Dim lngPtrHostEnt As Long
Dim lngPtrHostName As Long
Dim strHostName As String
Dim udtHostent As HOSTENT

      strIpAddress = Trim$(strIpAddress)

      ' Valid IP address contains at least 7 characters
      If Len(strIpAddress) > 6 Then
         ' Convert the IP address string to Long
         lngInetAdr = inet_addr(strIpAddress)

         ' ## Retrieve host name
         
         ' Get the pointer to the HostEnt structure
         lngPtrHostEnt = gethostbyaddr(lngInetAdr, 4, AF_INET)
         ' Copy data into the HostEnt structure
         RtlMoveMemory udtHostent, ByVal lngPtrHostEnt, LenB(udtHostent)
         ' Prepare the buffer to receive a string
         strHostName = String(256, 0)
         ' Copy the host name into the strHostName variable
         RtlMoveMemory ByVal strHostName, ByVal udtHostent.hName, 256
         ' Cut received string by first chr(0) character
         strHostName = Left(strHostName, InStr(1, strHostName, Chr(0)) - 1)
         ' Return the found value
         GetHostNameByAddress = strHostName
      End If
End Function

' ********************************************************************************
' Author    :Darren Lawrence
' Date/Time :03-July-2005
' Purpose   :Sends raw byte array data to the remote host with connected socket
' Returns   :Number of bytes sent to the remote host
' Arguments :lngSocket    - the socket connected to the remote host
' '          strData      - data to send
' '          SocketBuffer - The Socket Header
' ********************************************************************************
Public Function vbSendByteArrayTo(ByVal lngSocket As Long, strData() As Byte, SocketBuffer As sockaddr_in) As Long
Dim lngBytesSent    As Long

      If IsConnected(lngSocket) And UBound(strData) > 0 Then
         ' Call the send Winsock API function in order to send data
         lngBytesSent = sendto(lngSocket, strData(0), UBound(strData) + 1, 0&, SocketBuffer, Len(SocketBuffer))
         vbSendByteArrayTo = lngBytesSent
      Else
         vbSendByteArrayTo = SOCKET_ERROR
      End If
End Function

' ********************************************************************************
' Author    :Darren Lawrence
' Date/Time :03-July-2005
' Purpose   :Retrieves data from the Winsock buffer.
' Returns   :Number of bytes received.
' Arguments :lngSocket         - the socket connected to the remote host
' '          strData           - data array to receive into
' '          SocketBuffer      - The Socket Header
' '          MAX_BUFFER_LENGTH - Number of Bytes maximum to receive
' ********************************************************************************
Public Function vbRecvByteArrayFrom(ByVal lngSocket As Long, strData() As Byte, SocketBuffer As sockaddr_in, Optional TimeOut As Long = 10, Optional MAX_BUFFER_LENGTH = 8192) As Long
Dim lngBytesReceived As Long
Dim strTempBuffer    As String
Dim WaitUntil        As Double
Dim bOK              As Boolean

      ' Check the socket for readabilty with
      ' the IsDataAvailable function
      WaitUntil = DateAdd("s", TimeOut, Now()) ' Add the TimeOut value to the current time
      ' '                                        So we can timeout at that time
      bOK = False
      Do

         If IsDataAvailable(lngSocket) Then
            ' Call the recv Winsock API function in order to read data from the buffer
            lngBytesReceived = recvfrom(lngSocket, strData(0), MAX_BUFFER_LENGTH, 0&, SocketBuffer, Len(SocketBuffer))
            ' If lngBytesReceived is equal to 0 or -1, we have nothing to do with that
            vbRecvByteArrayFrom = lngBytesReceived
            bOK = True
            Exit Function
         Else
            ' iCounter = iCounter + 1
            ' Something wrong with the socket.
            ' Maybe the lngSocket argument is not a valid socket handle at all
            ' Or not waited long enough
            vbRecvByteArrayFrom = SOCKET_ERROR
            ' Exit Function
         End If
         
      Loop Until WaitUntil < Now() Or bOK ' Loop until the Time started is Less than the Current time
      ' If we get this far then waiting for receive timed out
      vbRecvByteArrayFrom = SOCKET_TIMEOUT
End Function

' ********************************************************************************
' Author    :Oleg Gdalevich
' Date/Time :27-Nov-2001
' Purpose   :Sends data to the remote host with connected socket
' Returns   :Number of bytes sent to the remote host
' Arguments :lngSocket    - the socket connected to the remote host
' '          strData      - data to send
' ********************************************************************************
Public Function vbSendString(ByVal lngSocket As Long, strData As String) As Long
Dim arrBuffer()     As Byte
Dim lngBytesSent    As Long
Dim lngBufferLength As Long

      lngBufferLength = Len(strData)
      If IsConnected(lngSocket) And lngBufferLength > 0 Then
         ' Convert the data string to a byte array
         arrBuffer() = StrConv(strData, vbFromUnicode)
         ' Call the send Winsock API function in order to send data
         lngBytesSent = send(lngSocket, arrBuffer(0), lngBufferLength, 0&)
         vbSendString = lngBytesSent
      Else
         vbSendString = SOCKET_ERROR
      End If
End Function

' ********************************************************************************
' Author    :Oleg Gdalevich
' Date/Time :27-Nov-2001
' Purpose   :Retrieves data from the Winsock buffer.
' Returns   :Number of bytes received.
' Arguments :lngSocket    - the socket connected to the remote host
' '          strBuffer    - buffer to read data to
' ********************************************************************************
Public Function vbRecv(ByVal lngSocket As Long, strBuffer As String, Optional TimeOut As Long = 10, Optional MAX_BUFFER_LENGTH = 8192) As Long
Dim lngBytesReceived                   As Long
Dim strTempBuffer                      As String
Dim WaitUntil                          As Double
Dim bOK                                As Boolean

      ' Check the socket for readabilty with
      ' the IsDataAvailable function
      WaitUntil = DateAdd("s", TimeOut, Now()) ' Add the TimeOut value to the current time
      ' '                                        So we can timeout at that time
      bOK = False
      
      Do
      
         ' Check the socket for readabilty with
         ' the IsDataAvailable function
         If IsDataAvailable(lngSocket) Then
            ReDim arrBuffer(MAX_BUFFER_LENGTH) As Byte

            ' Call the recv Winsock API function in order to read data from the buffer
            lngBytesReceived = recv(lngSocket, arrBuffer(0), MAX_BUFFER_LENGTH, 0&)
            If lngBytesReceived > 0 Then
               ' If we have received some data, convert it to the Unicode
               ' string that is suitable for the Visual Basic String data type
               strTempBuffer = StrConv(arrBuffer, vbUnicode)
               ' Remove unused bytes
               strBuffer = Left$(strTempBuffer, lngBytesReceived)
            End If
            ' If lngBytesReceived is equal to 0 or -1, we have nothing to do with that
            bOK = True
            vbRecv = lngBytesReceived
            Exit Function
         Else
            ' Something wrong with the socket.
            ' Maybe the lngSocket argument is not a valid socket handle at all
            vbRecv = SOCKET_ERROR
         End If
      Loop Until WaitUntil < Now() Or bOK ' Loop until the Time started is Less than the Current time
      ' If we get this far then waiting for receive timed out
      vbRecv = SOCKET_TIMEOUT
End Function

' ********************************************************************************
' Author    :Oleg Gdalevich
' Date/Time :15-Dec-2001
' Purpose   :Turns a socket into a listening state.
' Return    :If no error occurs, returns zero. Otherwise, it returns SOCKET_ERROR.
' '          Arguments :lngSocketHandle - the socket to turn into a listening state.
' ********************************************************************************
Public Function vbListen(ByVal lngSocketHandle As Long) As Long

Dim lngRetValue As Long

      lngRetValue = listen(lngSocketHandle, SOMAXCONN)

      vbListen = lngRetValue

      ' We have nothing to do as vbListen function returns the same value
      ' as the listen Winsock API function.

End Function

' ********************************************************************************
' Author    :Oleg Gdalevich
' Date/Time :15-Dec-2001
' Purpose   :Accepts a connection request, and creates a new socket.
' Return    :If no error occurs, returns the new socket's handle. Otherwise, it returns INVALID_SOCKET.
' Arguments :lngSocketHandle - the listening socket.
' ********************************************************************************
Public Function vbAccept(ByVal lngSocketHandle As Long) As Long
Dim lngRetValue         As Long
Dim udtSocketAddress    As sockaddr_in
Dim lngBufferSize       As Long
      
      ' Calculate the buffer size
      lngBufferSize = LenB(udtSocketAddress)
      ' Call the accept Winsock API function in order to create a new socket
      lngRetValue = accept(lngSocketHandle, udtSocketAddress, lngBufferSize)
      vbAccept = lngRetValue
End Function

Public Function IsConnected(ByVal lngSocket As Long) As Boolean
Dim udtRead_fd      As fd_set
Dim udtWrite_fd     As fd_set
Dim udtError_fd     As fd_set
Dim lngSocketCount  As Long
      
      udtWrite_fd.fd_count = 1
      udtWrite_fd.fd_array(1) = lngSocket
      lngSocketCount = vbselect(0&, udtRead_fd, udtWrite_fd, udtError_fd, 0&)
      IsConnected = CBool(lngSocketCount)
End Function

Public Function IsDataAvailable(ByVal lngSocket As Long) As Boolean
Dim udtRead_fd As fd_set
Dim udtWrite_fd As fd_set
Dim udtError_fd As fd_set
Dim lngSocketCount As Long

      udtRead_fd.fd_count = 1
      udtRead_fd.fd_array(1) = lngSocket
      lngSocketCount = vbselect(0&, udtRead_fd, udtWrite_fd, udtError_fd, 0&)
      IsDataAvailable = CBool(lngSocketCount)
End Function

Public Function GetDNSAddresses() As String()
Dim i As Long
Dim lServers As Long
Dim sPrefServer As String
Dim asDNSServers() As String

      ' pass an empty string and string array
      ' to the function. Return value is the
      ' number of DNS servers found
      Call GetDNSServers(asDNSServers())
      ' Return all servers found
      If UBound(asDNSServers()) > 0 Then
         For i = 0 To UBound(asDNSServers()) - 1
            ReDim Preserve asDNSServers(UBound(asDNSServers) + 1) As String
            asDNSServers(UBound(asDNSServers)) = asDNSServers(i)
         Next
      End If
End Function
' Return the Primary DNS Address
Public Function GetPrimaryDNS() As String
Dim asDNSServers() As String

      
      GetPrimaryDNS = GetDNSServers(asDNSServers())
End Function

' Return the Prefered DNS Server and a List if any
' in the ServersFound Array
Private Function GetDNSServers(ByRef ServersFound() As String) As String
Dim buff()        As Byte
Dim cbRequired    As Long
Dim nStructSize   As Long
Dim ptr           As Long
Dim tFixedInfo    As FIXED_INFO
Dim tIPData       As IP_ADDR_STRING
Dim i             As Long

      ReDim ServersFound(0) As String

      nStructSize = LenB(tIPData)

      ' call the api first to determine the
      ' size required for the values to be returned
      Call GetNetworkParams(ByVal 0&, cbRequired)

      If cbRequired > 0 Then

         ReDim buff(0 To cbRequired - 1) As Byte

         If GetNetworkParams(buff(0), cbRequired) = ERROR_SUCCESS Then

            ptr = VarPtr(buff(0))
            CopyMemory tFixedInfo, ByVal ptr, Len(tFixedInfo)
            With tFixedInfo
               ' identify the current dns server
               CopyMemory tIPData, ByVal VarPtr(.CurrentDnsServer) + 4, nStructSize
               GetDNSServers = TrimNull(StrConv(tIPData.IpAddress, vbUnicode))
               ' obtain a pointer to the
               ' DnsServerList array
               ptr = VarPtr(.DnsServerList)
               ' the IP_ADDR_STRING dwNext member indicates
               ' that more than one DNS server may be listed,
               ' so a loop as needed
               Do While (ptr <> 0)
                  ' copy each into an IP_ADDR_STRING type
                  CopyMemory tIPData, ByVal ptr, nStructSize

                  With tIPData
                     ' extract the server address and
                     ' cast to the array
                     ReDim Preserve ServersFound(0 To i) As String
                     ServersFound(i) = TrimNull(StrConv(tIPData.IpAddress, vbUnicode))
                     ptr = .dwNext
                  End With
                  i = i + 1
               Loop
            End With
         End If
      End If
End Function
' returns IP as long, in network byte order
Public Function GetHostByNameAlias(ByVal HostName$) As Long
Dim phe As Long
Dim heDestHost As HOSTENT
Dim addrList As Long
Dim retIP As Long

      retIP = inet_addr(HostName$)
      If retIP = INADDR_NONE Then
         phe = gethostbyname(HostName$)
         If phe <> 0 Then
            CopyMemory heDestHost, ByVal phe, Len(heDestHost)
            CopyMemory addrList, ByVal heDestHost.hAddrList, 4
            CopyMemory retIP, ByVal addrList, heDestHost.hLength
         Else
            retIP = INADDR_NONE
         End If
      End If
      GetHostByNameAlias = retIP
End Function

' returns your local machines name
Public Function GetLocalHostName() As String
Dim sName As String
   
      sName = String(256, 0)
      Call gethostname(sName, 256)
      If InStr(sName, Chr(0)) Then
         sName = Left(sName, InStr(sName, Chr(0)) - 1)
      End If
      GetLocalHostName = sName
End Function

Public Function GetHostByAddress(ByVal addr As Long) As String
Dim phe As Long
Dim ret As Long
Dim heDestHost As HOSTENT
Dim HostName As String

      phe = gethostbyaddr(addr, 4, AddressFamily.AF_INET)
      If phe Then
         CopyMemory heDestHost, ByVal phe, Len(heDestHost)
         HostName = String(256, 0)
         CopyMemory ByVal HostName, ByVal heDestHost.hName, 256
         GetHostByAddress = Left(HostName, InStr(HostName, Chr(0)) - 1)
      Else
         GetHostByAddress = "Unknown"
      End If
End Function

Public Function GetAscIP(ByVal inn As Long) As String
Dim nStr As Long
Dim lpStr As Long
Dim retString As String

      retString = String(32, 0)
      lpStr = inet_ntoa(inn)
      If lpStr Then
         nStr = lstrlen(lpStr)
         If nStr > 32 Then nStr = 32
         CopyMemory ByVal retString, ByVal lpStr, nStr
         retString = Left(retString, nStr)
         GetAscIP = retString
      Else
         GetAscIP = "255.255.255.255"
      End If
End Function

Private Function TrimNull(item As String)
Dim pos As Integer

      pos = InStr(item, Chr$(0))
      If pos Then
         TrimNull = Left$(item, pos - 1)
      Else
         TrimNull = item
      End If
End Function
' Takes sDomain and converts it to the QNAME-type string, returns that.
' QNAME is how a DNS server expects the string.
Public Function MakeQName(sDomain As String) As String
Dim iQCount As Integer      ' Character count (between dots)
Dim iNdx As Integer         ' Index into sDomain string
Dim iCount As Integer       ' Total chars in sDomain string
Dim sQName As String        ' QNAME string
Dim sDotName As String      ' Temp string for chars between dots
Dim sChar As String         ' Single char from sDomain string

      iNdx = 1
      iQCount = 0
      iCount = Len(sDomain)
      ' While we haven't hit end-of-string
      While (iNdx <= iCount)
         ' Read a single char from our domain
         sChar = Mid(sDomain, iNdx, 1)
         ' If the char is a dot, then put our character count and the part of the string
         If (sChar = ".") Then
            sQName = sQName & Chr(iQCount) & sDotName
            iQCount = 0
            sDotName = ""
         Else
            sDotName = sDotName + sChar
            iQCount = iQCount + 1
         End If
         iNdx = iNdx + 1
      Wend

      sQName = sQName & Chr(iQCount) & sDotName

      MakeQName = sQName
End Function

' Parse the server name out of the MX record, returns it in variable sName, iNdx is also
' modified to point to the end of the parsed structure.
Private Sub ParseName(dnsReply() As Byte, iNdx As Integer, ByRef sName As String)
Dim iCompress As Integer  ' Compression index (index into original buffer)
Dim iChCount  As Integer  ' Character count (number of chars to read from buffer)

      ' While we didn't encounter a null char (end-of-string specifier)
      While (dnsReply(iNdx) <> 0)
         ' Read the next character in the stream (length specifier)
         iChCount = dnsReply(iNdx)
         ' If our length specifier is 192 (0xc0) we have a compressed string
         If (iChCount = 192) Then
            ' Read the location of the rest of the string (offset into buffer)
            iCompress = dnsReply(iNdx + 1)
            ' Call ourself again, this time with the offset of the compressed string
            ParseName dnsReply(), iCompress, sName
            ' Step over the compression indicator and compression index
            iNdx = iNdx + 2
            ' After a compressed string, we are done
            Exit Sub
         End If

         ' Move to next char
         iNdx = iNdx + 1
         ' While we should still be reading chars
         While (iChCount)
            ' add the char to our string
            sName = sName & Chr(dnsReply(iNdx))
            iChCount = iChCount - 1
            iNdx = iNdx + 1
         Wend
         ' If the next char isn't null then the string continues, so add the dot
         If (dnsReply(iNdx) <> 0) Then
            sName = sName & "."
         End If
      Wend
End Sub

' Parses the buffer returned by the DNS server, returns the best MX server (lowest preference
' number), iNdx is modified to point to current buffer position (should be the end of buffer
' by the end, unless a record other that MX is found)
Private Function GetMXName(dnsReply() As Byte, iNdx As Integer, iAnCount As Integer) As MX_INFO
Dim iChCount  As Integer      ' Character counter
Dim sTemp     As String       ' Holds original query string
Dim iBestPref As Integer      ' Holds the "best" preference number (lowest)
Dim sBestMX   As String       ' Holds the "best" MX record (the one with the lowest preference)
Dim sName     As String
Dim iPref     As Integer
Dim TempMX    As MX_INFO

      ' init
      iBestPref = -1

      ParseName dnsReply(), iNdx, sTemp
      ' Step over null
      iNdx = iNdx + 2
      TempMX.Domain = sTemp
      ' Step over 6 bytes (not sure what the 6 bytes are, but all other
      ' documentation shows steping over these 6 bytes)
      iNdx = iNdx + 6
      On Error Resume Next
      ReDim Preserve TempMX.List(iAnCount - 1) As MX_RECORD
      TempMX.Count = iAnCount
      While (iAnCount)
         ' Check to make sure we received an MX record
         If (dnsReply(iNdx) = 15) Then

            sName = ""

            ' Step over the last half of the integer that specifies the record type (1 byte)
            ' Step over the RR Type, RR Class, TTL (3 integers - 6 bytes)
            iNdx = iNdx + 1 + 6

            ' Read the MX data length specifier
            ' (not needed, hence why it's commented out)
            ' MemCopy iMXLen, dnsReply(iNdx), 2
            ' iMXLen = ntohs(iMXLen)

            ' Step over the MX data length specifier (1 integer - 2 bytes)
            iNdx = iNdx + 2

            CopyMemory iPref, dnsReply(iNdx), 2
            iPref = ntohs(iPref)
            ' Step over the MX preference value (1 integer - 2 bytes)
            iNdx = iNdx + 2

            ' Have to step through the byte-stream, looking for 0xc0 or 192 (compression char)
            ParseName dnsReply(), iNdx, sName

            If (iBestPref = -1 Or iPref < iBestPref) Then
               iBestPref = iPref
               sBestMX = sName
               TempMX.Best = sName
            End If

            ' Step over 3 useless bytes
            iNdx = iNdx + 3

            TempMX.List(iAnCount - 1).Server = sName
            TempMX.List(iAnCount - 1).Pref = iPref
         Else

            Exit Function
         End If
         iAnCount = iAnCount - 1
      Wend

      GetMXName = TempMX

End Function

Public Function GetMXRecord(sDomainToFind As String, sDNSAddress As String) As MX_INFO
Dim SocketBuffer As sockaddr_in
Dim lngDNS As Long
Dim lngRetValue As Long
Dim lngSocket As Long

Dim dnsHead As DNS_HEADER
Dim DNSserver() As String
Dim DNSpreferred As String
Dim dnsQuery() As Byte ' The Initial Query is In this
Dim sQName As String

Dim dnsQueryNdx As Integer
Dim iTemp As Integer
Dim iNdx As Integer
Dim dnsReply(2048) As Byte ' Server Reply is Stored in this
Dim iAnCount As Integer ' Reply Counter
Dim MXRecord As MX_INFO
           
      GetMXRecord = MXRecord ' Set to a Empty record incase of errors
           
      ' Create the Socket
      lngSocket = vbSocket(AF_INET, SOCK_DGRAM, IPPROTO_IP)
      If lngSocket = SOCKET_ERROR Then
         Exit Function
      End If
      
      ' Turn DNS IP into a Long Address
      lngDNS = GetHostByNameAlias(sDNSAddress)
      If lngDNS = SOCKET_ERROR Then
         Exit Function
      End If
      
      ' Setup the connnection parameters
      SocketBuffer.sin_family = AF_INET
      SocketBuffer.sin_port = htons(53)
      SocketBuffer.sin_addr = lngDNS
      ' SocketBuffer.sin_zero = String$(8, 0)
      
      ' Set the DNS parameters
      dnsHead.qryID = htons(&H11DF)
      dnsHead.options = DNS_RECURSION
      dnsHead.qdcount = htons(1)
      dnsHead.ancount = 0
      dnsHead.nscount = 0
      dnsHead.arcount = 0
      
      dnsQueryNdx = 0
      
      ReDim dnsQuery(4000) As Byte
      
      ' Setup the dns structure to send the query in
      
      ' First goes the DNS header information
      CopyMemory dnsQuery(dnsQueryNdx), dnsHead, 12
      dnsQueryNdx = dnsQueryNdx + 12
      
      ' Then the domain name (as a QNAME)
      sQName = MakeQName(CStr(sDomainToFind))
      
      iNdx = 0
      While (iNdx < Len(sQName))
         dnsQuery(dnsQueryNdx + iNdx) = Asc(Mid(sQName, iNdx + 1, 1))
         iNdx = iNdx + 1
      Wend
      
      dnsQueryNdx = dnsQueryNdx + Len(sQName)
      
      ' Null terminate the string
      dnsQuery(dnsQueryNdx) = &H0
      dnsQueryNdx = dnsQueryNdx + 1
      ' The type of query (15 means MX query)
      iTemp = htons(15)
      CopyMemory dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
      dnsQueryNdx = dnsQueryNdx + Len(iTemp)
      ' The class of query (1 means INET)
      iTemp = htons(1)
      CopyMemory dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
      dnsQueryNdx = dnsQueryNdx + Len(iTemp)
      
      ReDim Preserve dnsQuery(dnsQueryNdx - 1)
      ' Send the query to the DNS server
      
      If modWinsockAPI.vbSendByteArrayTo(lngSocket, dnsQuery, SocketBuffer) = SOCKET_ERROR Then
         ' If the vbSendByteArrayTo function has returned a value of
         ' SOCKET_ERROR, Then exit
         Exit Function
      End If
      ' Wait for answer from the DNS server
      If modWinsockAPI.vbRecvByteArrayFrom(lngSocket, dnsReply, SocketBuffer, 10, 2048) = SOCKET_ERROR Then
         ' If the vbRecvByteArrayFrom function has returned a value of
         ' SOCKET_ERROR, Then exit
         Exit Function
      End If
      
      ' Get the number of answers
      CopyMemory iAnCount, dnsReply(6), 2
      iAnCount = ntohs(iAnCount)
      ' Parse the answer buffer
      
      MXRecord = GetMXName(dnsReply(), 12, iAnCount)
      
      Call closesocket(lngSocket)
      
      GetMXRecord = MXRecord

End Function

