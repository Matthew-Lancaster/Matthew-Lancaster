
            FileDownloader Class File Documentation
            ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                                           The KPD-Team, November 24th, 2001
                                           http://www.allapi.net/
                                           Version 1.2
                                           [this documentation is out of date]
                                           [consult the source code for up to date information]


   The same (and more) documentation can be found in the source files
   of the DataConnection class. 


   Properties:
   ~~~~~~~~~~~

    - AllowCache [Boolean][Get/Let][Default: False]
     Specifies whether proxies should return a cached version of the
     requested file if one is available. If set to False, no cached
     versions are allowed.

    - AutoRedirect [Boolean][Get/Let][Default: True]
     If the server returns a 301 or 302 status code (which means
     the requested resource has been, respectively, Permantently
     and Temporarily moved), the class automatically changes the
     URL property and restarts the download. (Do note that, if it
     restarts the download, the Connected and Disconnected events
     will not be raised)

    - ConnectionType [CONNECTION_TYPE][Get/Let][Default: ctHTTP]
     Specifies the protocol to be used to download the requested file.
     Currently, only ctHTTP is supported. Future versions may include
     ctFTP or other protocols.

    - DownloadLength [Long][Get][Default: -1]
     Returns the size of the file that is currently being downloaded,
     or -1 if the size is unknown (this is often the case with scripts).

    - DownloadType [DOWNLOAD_TYPE][Get/Let][Default: dtToFile]
     Specifies what should be done when data arrives:
      · dtToFile
       All incoming data is written to a file. The filename is specified
       by the Filename property and/or the UseTempFile property.
      · dtToBuffer
       All incoming data is written to a buffer. The client may query
       the contents of the buffer with the GetBuffer or GetBufferAsString
       function.
      · dtStream
       Whenever new data is available, the event StreamBytes is raised.
       One of the parameters of this event is a byte buffer with the new
       data.

    - Filename [String][Get/Let][Default: Empty String]
     Specifies the filename where the data should be saved to. This
     property is ignored when the DownloadType property is not set to
     dtToFile or when UseTempFile is set to True.

    - Hostname [String][Get/Let][Default: Empty String]
     Specifies the hostname or the IP address the Download Class should
     connect to. For instance, if you want to download the file
     ‘http://www.allapi.net/downloads/myfile.ext’, the Hostname property
     should be set to ‘www.allapi.net’.

    - HTTPReply [Long][Get][Default: N/A]
     Returns the status code of the current HTTP connection. Usually
     this is either ‘200’ (OK) or ‘206’ (Partial Data, when you’re
     resuming a file download) or ‘404’ (file not found). For a full
     list of possible status codes, take a look at the HTTP protocol
     at http://www.w3c.org.
     If this property returns ‘999’, the downloaded HTTP header was
     invalid.

    - ID [Long][Get/Let][Default: 0]
     Specifies a value that stores any extra data needed for your program.
     Unlike other properties, the value of the ID property isn't used by
     the Download Class. You can use this property to identify objects. 

    - IsConnected [Boolean][Get][Default: N/A]
     Returns True if the class is connected to a server or False in
     the other case.

    - MaxDownload [Long][Get/Let][Default: -1]
     Specifies the maximum bytes to download. If this number is reached,
     the connection will be automatically closed.

    - PacketSize [Long][Get/Let][Default: 1024]
     Specifies the preferred size of retrieving new data. This value has
     to be positive.

    - Port [Long][Get/Let][Default: 80]
     Specifies the remote port that will be used to connect to the server.

    - ProxyAuthentication [PROXY_AUTHENTICATION][Get/Let][Default: paNone]
     Specifies the authentication scheme that should be used when connecting
     to a SOCKS5 proxy server.
      · paNone
       No authentication is used.
      · paUserPass
       Use user/password authentication.

    - ProxyHostname [String][Get/Let][Default: ‘localhost’]
     Specifies the hostname or IP address of the proxy server.

    - ProxyPassword [String][Get/Let][Default: Empty String]
     Specifies the password that should be used when logging in on a SOCKS5
     proxy server.

    - ProxyPort [Long][Get/Let][Default: 80]
     Specifies the port where the class should connect on when it’s connecting
     to the proxy server. If you’re using a SOCKS proxy, this value should
     probably be set to 1080.

    - ProxyType [PROXY_TYPE][Get/Let][Default: ptNoProxy]
     Specifies the type of proxy to use.
      · ptNoProxy
       No proxy will be used when connecting to the hostname.
      · ptStandardProxy
       A ‘normal’ proxy will be used when connecting to the host.
      · ptSOCKS4Proxy
       A SOCKS4 proxy will be used when connecting to the host.
      · ptSOCKS5Proxy
       A SOCKS5 proxy will be used when connecting to the host.

    - ProxyUseRemoteDNS [Boolean][Get/Let][Default: False]
     Specifies whether domain names should be resolved locally or on the proxy
     server. If this value is set to True, domain names are resolved on the
     proxy server, in the other case, domain names are resolved on the local
     machine. (Warning: not all SOCKS4 proxy servers support Remote DNS!)

    - ProxyUsername [String][Get/Let][Default: Empty String]
     Specifies the username to be used when authenticating on a SOCKS proxy.

    - RealReceivedBytes [Long][Get][Default: 0]
     Specifies the number of received bytes from the host.

    - ReceivedBytes [Long][Get][Default: 0]
     Specifies the number of received bytes from the host. If you’re using
     the resume function, this value will start from the ResumeFrom value.

    - ResumeFrom [Long][Get/Let][Default: 0]
     Specifies the position in the remote file where to start downloading from.
     The first byte in a file is byte 0, not byte 1. Hence, if you downloaded
     205 bytes from a remote file and you want to resume the download, you should
     set this value to ‘205’.

    - URL [String][Get/Let][Default: Empty String]
     Specifies the URL of the data to download.
     Note that this URL can change during the download if you enabled the
     AutoRedirect property.

    - UseResume [Boolean][Get/Let][Default: False]
     Specifies whether or not to resume a download.

    - UseTempFile [Boolean][Get/Let][Default: True]
     Specifies whether or not to generate a temporary file. If this option is set
     to True, the class will create a temporary file in the Windows temporary
     directory and adjust the Filename property accordingly. The client application
     is responsible for deleting the temporary files after downloading them.


   Methods:
   ~~~~~~~~

    - Public Sub ClearBuffer()
     The ClearBuffer method clears the buffer that stores data when the DownloadType
     is set to dtToBuffer. It may be useful to clear this buffer after downloading
     a large file, to avoid unnecessary memory usage.
      Parameters: NONE
      Return Value: NONE

    - Public Sub Disconnect()
     Closes the socket and also closes the file handle if the DownloadType is set
     to dtToFile.
      Parameters: NONE
      Return Value: NONE

    - Public Function FailureToString(Why As DOWNLOAD_FAILURE) As String
     Used to convert a DOWNLOAD_FAILURE number to a String
      Parameters: The DOWNLOAD_FAILURE you wish to convert
      Return Value: String description of the failure

    - Public Sub FetchURL()
     Start downloading data from the specified URL.
      Parameters: NONE
      Return Value: NONE

    - Public Sub FetchURLString(ByVal sURL As String)
     Start downloading data from the specified URL. The URL must contain
     the protocol ('http://') and the hostname. Optionally, it can contain the
     port number. If no port is specified, the Port property will be set to
     the default port 80.
     Previous URL or Port properties settings will be ignored.
     When an invalid URL is specified, the DownloadFailed event will
     be raised with the Why parameter set to dfInvalidURL
      Parameters: URL to download from
      Return Value: NONE

    - Public Function GetBuffer(bBytes() As Byte) As Long
     Returns a byte array with the data of the buffer. This byte array starts from
     index 0.
      Parameters: Byte array that receives the data of the buffer.
      Return Value: Number of bytes in buffer

    - Public Function GetBufferAsString() As String
     Returns a string with the data of the buffer.
      Parameters: NONE
      Return Value: the data

    - Public Function GetBufferSize() As Long
     Returns the number of bytes in the buffer.
      Parameters: NONE
      Return Value: the number of bytes in the buffer


   Events:
   ~~~~~~~

    - Event Connected(ID As Long)
     The Connected event is raised when a connection has been established with either
     the proxy or the host.

    - Event Disconnected(ID As Long)
     The Disconnected event is raised when the file has been successfully downloaded
     and the connection to the host or proxy has been closed.

    - Event DownloadFailed(Why As DOWNLOAD_FAILURE, ID As Long)
     The DownloadFailed event is raised when a download fails. This can be for various
     reasons; the Why parameter indicates the source of the problem:
     · dfConnectionLost
      The connection with the proxy or host has been closed, before downloading
      the entire file.
     · dfCancelledByUser
      The download has been aborted by the user.
     · dfCannotConnect
      Unable to connect to the host.
     · dfCannotConnectToProxy
      Unable to connect to the proxy or unable to send data to it.
     · dfCannotCreateFile
      Unable to create the file, specified in the Filename property.
     · dfCannotCreateWindow
      Unable to create a necessary window.
     · dfCannotSendData
      Unable to send data to the host or proxy
     · dfCannotStartWinSock
      Unable to start WinSock. This means that every call to the WinSock API
      will fail.
     · dfHTTPError
      The server didn’t reply with a 200 or 206 reply code. For more information
      about the server reply code, take a look at the HTTPReply property.
     · dfProxyAuthMethodNotAccepted
      The selected authentication method was not accepted by the (SOCKS5) proxy.
     · dfProxyRequestRejectedOrFailed
      The connection request has failed or the proxy server has rejected it.
     · dfUserPasswordNotAccepted
      The user/password combination was not accepted by the proxy.
     . dfInvalidURL
      The specified URL is invalid

    - Event BytesReceived(lByteCount As Long, ID As Long)
     The BytesReceived event is raised whenever new data has arrived. The number
     of arrived bytes is specified by lByteCount.

    - Event StreamBytes(lByteCount As Long, bBytes() As Byte, ID As Long)
     The StreamBytes event is only raised when the DownloadType has been set to
     dtStream. The byte array is the array that contains the arrived bytes and
     lByteCount is the number of bytes contained in this array.

    - Event QuerySent(ID As Long)
     The QuerySent event is raised when the HTTP query has been successfully sent.


   How to use this class in your project:
   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    Simply add the DataConnection.cls and mdlDataConnection.mdl files to your
    project and you can use the DataConnection class in your program.
    The sample form explains how to use the class file. I provides access to
    most of the DataConnection functionality.


   Final notes:
   ~~~~~~~~~~~~

    If you use this class to retrieve files larger than 2.1Gb, you're bound to
    have problems with it. This class makes extensive use of the Long data type
    and therefore the maximum filesize that can be downloaded with this class
    is 2,147,483,647 bytes.

    Before an application can use the FileDownloaded class, it *MUST* call the
    StartWinsock method from this module!
    If you don't do this, every call to the WinSock API will fail!

    If you find any bugs in this file or you want to suggest a new feature,
    don 't hesitate to contact us at KPDTeam@allapi.net

    We strongly suggest that you don't modify this class module. Thatway, if a new
    version of this file is released, you simply have to replace the classe file and/or
    module file with the new file(s) and you're done.
    If you make your own modifications, chances are you may introduce some new
    bugs in this file and you cannot update this file as easy as normal.
    Instead of adding your own methods, let us know about your request and we'll
    certainly consider it.


   Copyright Notice and disclaimer:
   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    The FileDownloader class file is created by The KPD-Team, 2001. You're free
    to use this class file in your own projects, however you may not claim that
    you have written this class file.
    You must leave this copyright notice in the file if you intend to distribute
    the source code of your program.
    This class file is provided "as is", with no guarantee of completeness or accuracy
    and without warranty of any kind, express or implied. It is intended solely
    to provide general guidance on matters of interest for the personal use of
    the user of this program, who accepts full responsibility for its use.
    For more information about this, please contact us at KPDTeam@allapi.net.

