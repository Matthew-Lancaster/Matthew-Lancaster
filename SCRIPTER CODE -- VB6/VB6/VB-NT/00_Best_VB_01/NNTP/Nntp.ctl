VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.UserControl NNTP 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1395
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Nntp.ctx":0000
   ScaleHeight     =   420
   ScaleWidth      =   1395
   ToolboxBitmap   =   "Nntp.ctx":07B9
   Begin VB.Timer Tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "NNTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NNTP-Control (c)2000-2001 Uwe Keller
'Homepage: www.uwekeller.de
'Email: visualbasic@uwekeller.de
'Developed in accordance with RFC977 / RFC1036 /RFC2980
'comments with *** are developer comments, please ignore

'***set header fields (msgid msgno subject author lines bytes xref reference) after GotoNext/GotoLast or retrieving article/header
'***Newsgroup-Descriptions "LIST NEWSGROUPS [wildmat]"

Option Explicit
Option Compare Text


'internal variables
Dim SendUsername As Boolean 'flag whether sending login name required
Dim SendPassword As Boolean 'flag whether sending login password required
Dim Content As TNNTPData 'kind of received data
Dim Data As New TStrList 'place for all incoming data (newsgroups, headers, articles, bodies...)
Dim Holdover As String 'spooling incomplete lines
Dim LastReceive As Single 'time of last receival (required for timeout detection)
Dim MsgIdCnt As Long 'internal Message-Id counter for posting
Dim KeepAlive As Boolean 'flag whether KeepAlive is currently in use
Dim CmdStartTime As Single 'time of current command start
Dim CmdStartBytes As Long 'bytecounter before executing command
Dim CmdEndTime As Single 'time of current command end
Dim CmdEndBytes As Long 'bytecounter after executing command

'server
Dim Host1 As String 'servername or IP-address
Dim Hostinfo1 As String 'info got during login
Dim Username1 As String 'username if required
Dim Password1 As String 'password if required

'newsgroup
Dim Newsgroup1 As String 'name of newsgroup
Dim ArticleCount1 As Long 'estimated number of articles in selected newsgroup
Dim FirstArticle1 As Long 'first article in selected newsgroup
Dim LastArticle1 As Long 'last article in selected newsgroup
Dim Posting1 As Boolean 'flag whether posting is allowed

'article
Dim MsgNo1 As Long 'current message number
Dim Subject1 As String 'subject line of current article
Dim Author1 As String 'email and name of articles author
Dim MsgId1 As String 'current message id
Dim References1 As String 'references
Dim Bytes1 As Long 'size of article in bytes
Dim Lines1 As Long 'size of article in lines
Dim XRef1 As String 'cross-reference data
Dim Newsreader1 As String 'newsreader
Dim Text1 As String 'place for message text (article / header / body)
Dim Message1 As String 'message text without attachments

'other
Public State1 As TNNTPState 'current state of control
Dim Send1 As String 'place for posting article (header+text1)
Dim Received1 As Long 'total amount of received data in bytes
Dim Timeout1 As Single 'number of seconds until timeout occurs
Dim Refresh1 As Single 'number of seconds until refresh will be sent
Dim Debugger1 As Boolean 'flag whether in debug mode
Dim DataIndex1 As Long 'pointer to current data line (newsgroup/header)
Dim DataSize1 As Long 'size of received data
Dim DateTime1 As Date 'date/time of newsgroup or article

'control states
Public Enum TNNTPState
    NNTP_STAT_DISCONNECTED = 0 'disconnected
    NNTP_STAT_DISCONNECTING = 1 'disconnecting
    NNTP_STAT_CONNECTED = 2 'connected (idle)
    NNTP_STAT_CONNECTING = 3 'connecting
    NNTP_STAT_SELECTING = 4 'selecting newsgroup
    NNTP_STAT_GETLIST = 5 'receiving newsgroups
    NNTP_STAT_GETXOVER = 6 'receiving headers
    NNTP_STAT_GETSTAT = 7 'receiving article state
    NNTP_STAT_GETARTICLE = 8 'receiving article
    NNTP_STAT_GETHEADER = 9 'receiving header
    NNTP_STAT_GETBODY = 10 'receiving body
    NNTP_STAT_GOTOLAST = 11 'seeking previous article
    NNTP_STAT_GOTONEXT = 12 'seeking next article
    NNTP_STAT_POSTING = 13 'posting
End Enum

'return codes
Public Enum TNNTPRc
    NNTP_RC_NONE = 0
    NNTP_RC_NOTCONNECTED = 1
    NNTP_RC_HOSTNOTFOUND = 2
    NNTP_RC_BUSY = 3
    NNTP_RC_SOCKERR = 4
    NNTP_RC_CANCELED = 5
    NNTP_RC_TIMEOUT = 6
    NNTP_RC_POSTING = 200
    NNTP_RC_NOPOSTING = 201
    NNTP_RC_CLOSING = 205
    NNTP_RC_GROUPSELECTED = 211
    NNTP_RC_LIST = 215
    NNTP_RC_ARTICLE = 220
    NNTP_RC_HEAD = 221
    NNTP_RC_BODY = 222
    NNTP_RC_STAT = 223
    NNTP_RC_DATAFOLLOWS = 224
    NNTP_RC_NEWARTICLES = 230
    NNTP_RC_NEWGROUPS = 231
    NNTP_RC_TRANSFERRED = 235
    NNTP_RC_POSTEDOK = 240
    NNTP_RC_AUTHENTICATIONACCEPTED = 281
    NNTP_RC_SENDARTICLETRANSFERRED = 335
    NNTP_RC_SENDARTICLE = 340
    NNTP_RC_MOREAUTHENTICATIONREQUIRED = 381
    NNTP_RC_SERVICEDISCONTINUED = 400
    NNTP_RC_NOSUCHNEWSGROUP = 411
    NNTP_RC_NOTINANEWSGROUP = 412
    NNTP_RC_NOARTICLESELECTED = 420
    NNTP_RC_NONEXTARTICLE = 421
    NNTP_RC_NOLASTARTICLE = 422
    NNTP_RC_BADARTICLENUMBER = 423
    NNTP_RC_ARTICLENOTFOUND = 430
    NNTP_RC_DONTSEND = 435
    NNTP_RC_TRANSFERFAILED = 436
    NNTP_RC_ARTICLEREJECTED = 437
    NNTP_RC_POSTINGNOTALLOWD = 440
    NNTP_RC_POSTINGFAILED = 441
    NNTP_RC_AUTHENTICATIONREQUIRED = 480
    NNTP_RC_AUTHENTICATIONREJECTED = 482
    NNTP_RC_NOSUCHCOMMAND = 500
    NNTP_RC_SYNTAXERROR = 501
    NNTP_RC_ACCESSRESTRICTED = 502
    NNTP_RC_PROGRAMNOTFAULT = 503
    NNTP_RC_CONNECTIONFAILURE1 = 504
    NNTP_RC_CONNECTIONFAILURE2 = 505
End Enum

'kind of content in data
Private Enum TNNTPData
    NNTP_DATA_EMPTY = 0 'empty
    NNTP_DATA_GROUP = 1 'newsgroups
    NNTP_DATA_XOVER = 2 'headers
    NNTP_DATA_TEXT = 3 'article, header, body
End Enum

'events
Public Event DoneConnect(Rc As TNNTPRc)
Public Event DoneDisconnect(Rc As TNNTPRc)
Public Event DoneSelectGroup(Rc As TNNTPRc)
Public Event DoneGetList(Rc As TNNTPRc)
Public Event DoneGetXover(Rc As TNNTPRc)
Public Event DoneGetStat(Rc As TNNTPRc)
Public Event DoneGetArticle(Rc As TNNTPRc)
Public Event DoneGetHeader(Rc As TNNTPRc)
Public Event DoneGetBody(Rc As TNNTPRc)
Public Event DoneGotoFirst(Rc As TNNTPRc)
Public Event DoneGotoNext(Rc As TNNTPRc)
Public Event DoneGotoLast(Rc As TNNTPRc)
Public Event DonePost(Rc As TNNTPRc)
Public Event DoneSend(Rc As TNNTPRc)
Public Event Receive(Count As Long)
Public Event Debugger(Text As String)



Private Sub Tmr_Timer()
    On Error Resume Next
    
    'IsConnected = State1 = NNTP_STAT_CONNECTED
    
    
    'connection refresh if connected but idle (prevents disconnect by server)
    If State1 = NNTP_STAT_CONNECTED Then
        'refresh time set up?
        If Refresh1 > 0 Then
            'refresh time reached?
            If Timer - LastReceive > Refresh1 Then
                'yes, sending HELP command (that tells the server that i am still there :-)
                KeepAlive = True
                Winsock(1).SendData "HELP" + vbCrLf
            End If
        End If
    
    'Online and Busy?
    ElseIf State1 <> NNTP_STAT_DISCONNECTED Then
        'timeout time set up?
        If Timeout1 > 0 Then
            'timeout time reached?
            If Timer - LastReceive > Timeout1 Then
                'yes, disconnect
                DisconnectNow NNTP_RC_TIMEOUT
            End If
        End If
        
    End If
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    'initial values when creating a new control instance
    Refresh1 = 60
    Timeout1 = 300
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    'reading control setup
    With PropBag
        Debugger1 = .ReadProperty("Debugger", False)
        Refresh1 = .ReadProperty("Refresh", 60)
        Timeout1 = .ReadProperty("Timeout", 300)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    'save control properties
    With PropBag
        .WriteProperty "Debugger", Debugger1
        .WriteProperty "Refresh", Refresh1, 60
        .WriteProperty "Timeout", Timeout1, 300
    End With
End Sub


Private Sub UserControl_Resize()
    On Error Resume Next
    'resize the control in the form to an appropriate size
    UserControl.Width = Screen.TwipsPerPixelX * 28
    UserControl.Height = Screen.TwipsPerPixelY * 28
End Sub

Public Sub Connect()
    On Error Resume Next
    
    'set start/end times for statistics
    CmdStatistic True, True
    
    'already connected
    If State1 <> NNTP_STAT_DISCONNECTED Then
        RaiseEvent DoneConnect(NNTP_RC_BUSY)
        
    'server IP or name missing
    ElseIf Host1 = "" Then
        RaiseEvent DoneConnect(NNTP_RC_HOSTNOTFOUND)
    
    'connect
    Else
        Received1 = 0
        LastReceive = Timer
        SendUsername = Username1 <> ""
        SendPassword = Password1 <> ""
        State1 = NNTP_STAT_CONNECTING
    
        Load Winsock(1)
        Winsock(1).Close
        Winsock(1).RemoteHost = Host1
        Winsock(1).RemotePort = 119
        Winsock(1).Connect
        Tmr.Enabled = True
        
    End If
End Sub

Public Sub Disconnect()
    On Error Resume Next
    
    'set start/end times for statistics
    CmdStatistic True, True
    
    'not connected
    If State1 = NNTP_STAT_DISCONNECTED Then
        RaiseEvent DoneDisconnect(NNTP_RC_NOTCONNECTED)
        
    'busy -> cancel (force unload)
    ElseIf State1 <> NNTP_STAT_CONNECTED Then
        DisconnectNow NNTP_RC_CANCELED
        
    'disconnect
    Else
        State1 = NNTP_STAT_DISCONNECTING
        WaitKeepAlive
        Winsock(1).SendData "QUIT" + vbCrLf
        
    End If
End Sub

Public Sub GetList()
    On Error Resume Next
    
    'set start/end times for statistics
    CmdStatistic True, True
    
    'not connected
    If State1 = NNTP_STAT_DISCONNECTED Then
        RaiseEvent DoneGetList(NNTP_RC_NOTCONNECTED)
        
    'busy
    ElseIf State1 <> NNTP_STAT_CONNECTED Then
        RaiseEvent DoneGetList(NNTP_RC_BUSY)
    
    'get list of newsgroups
    Else
        State1 = NNTP_STAT_GETLIST
        DataClear
        WaitKeepAlive
        'all newsgroups (LIST)
        If DateTime1 = 0 Then
            Winsock(1).SendData "LIST" + vbCrLf
        'new newsgroups only (NEWGROUPS)
        Else
            Winsock(1).SendData "NEWGROUPS " + Format(DateTime1, "yymmdd hhmmss") + " GMT " + vbCrLf
        End If
        
    End If
End Sub

Public Sub GetArticle()
    On Error Resume Next
    
    'set start/end times for statistics
    CmdStatistic True, True
    
    'not connected
    If State1 = NNTP_STAT_DISCONNECTED Then
        RaiseEvent DoneGetArticle(NNTP_RC_NOTCONNECTED)
        
    'busy
    ElseIf State1 <> NNTP_STAT_CONNECTED Then
        RaiseEvent DoneGetArticle(NNTP_RC_BUSY)
    
    'get article
    Else
        State1 = NNTP_STAT_GETARTICLE
        DataClear
        WaitKeepAlive
        'by message number
        If MsgNo1 > 0 Then
            Winsock(1).SendData "ARTICLE " + CStr(MsgNo1) + vbCrLf
        'by message id
        Else
            Winsock(1).SendData "ARTICLE " + MsgId1 + vbCrLf
        End If
        
    End If
End Sub

Public Sub GetBody()
    On Error Resume Next
    
    'set start/end times for statistics
    CmdStatistic True, True
    
    'not connected
    If State1 = NNTP_STAT_DISCONNECTED Then
        RaiseEvent DoneGetBody(NNTP_RC_NOTCONNECTED)
        
    'busy
    ElseIf State1 <> NNTP_STAT_CONNECTED Then
        RaiseEvent DoneGetBody(NNTP_RC_BUSY)
    
    'get body
    Else
        State1 = NNTP_STAT_GETBODY
        DataClear
        WaitKeepAlive
        'by message number
        If MsgNo1 > 0 Then
            Winsock(1).SendData "BODY " + CStr(MsgNo1) + vbCrLf
        'by message id
        Else
            Winsock(1).SendData "BODY " + MsgId1 + vbCrLf
        End If
        
    End If
End Sub

Public Sub GetHeader()
    On Error Resume Next
    
    'set start/end times for statistics
    CmdStatistic True, True
    
    'not connected
    If State1 = NNTP_STAT_DISCONNECTED Then
        RaiseEvent DoneGetHeader(NNTP_RC_NOTCONNECTED)
        
    'busy
    ElseIf State1 <> NNTP_STAT_CONNECTED Then
        RaiseEvent DoneGetHeader(NNTP_RC_BUSY)
    
    'get header
    Else
        State1 = NNTP_STAT_GETHEADER
        DataClear
        WaitKeepAlive
        'by message number
'        If MsgNo1 > 0 Then
            Winsock(1).SendData "HEAD " + CStr(Easy) + vbCrLf
        'by message id
'        Else
'            Winsock(1).SendData "HEAD " + MsgId1 + vbCrLf
'        End If
        
    End If
End Sub

Public Sub GetXover()
    On Error Resume Next
    
    'set start/end times for statistics
    CmdStatistic True, True
    
    'not connected
    If State1 = NNTP_STAT_DISCONNECTED Then
        RaiseEvent DoneGetXover(NNTP_RC_NOTCONNECTED)
        
    'busy
    ElseIf State1 <> NNTP_STAT_CONNECTED Then
        RaiseEvent DoneGetXover(NNTP_RC_BUSY)
    
    'receive headers by XOVER command
    Else
        State1 = NNTP_STAT_GETXOVER
        DataClear
        WaitKeepAlive
        Winsock(1).SendData "XOVER " + Format(FirstArticle1, "00000000") + "-" + Format(LastArticle1, "00000000") + vbCrLf
        
    End If
End Sub

Public Sub GetStat()
    On Error Resume Next
    
    'set start/end times for statistics
    CmdStatistic True, True
    
    'not connected
    If State1 = NNTP_STAT_DISCONNECTED Then
        RaiseEvent DoneGetStat(NNTP_RC_NOTCONNECTED)
        
    'busy
    ElseIf State1 <> NNTP_STAT_CONNECTED Then
        RaiseEvent DoneGetStat(NNTP_RC_BUSY)
    
    'get state
    Else
        State1 = NNTP_STAT_GETSTAT
        DataClear
        WaitKeepAlive
        'by message number
        If MsgNo1 > 0 Then
            Winsock(1).SendData "STAT " + CStr(MsgNo1) + vbCrLf
        'by message id
        Else
            Winsock(1).SendData "STAT " + MsgId1 + vbCrLf
        End If
        
    End If
End Sub


Public Sub GotoLast()
    On Error Resume Next
    
    'set start/end times for statistics
    CmdStatistic True, True
    
    'not connected
    If State1 = NNTP_STAT_DISCONNECTED Then
        RaiseEvent DoneGotoLast(NNTP_RC_NOTCONNECTED)
        
    'busy
    ElseIf State1 <> NNTP_STAT_CONNECTED Then
        RaiseEvent DoneGotoLast(NNTP_RC_BUSY)
    
    'last article
    Else
        State1 = NNTP_STAT_GOTOLAST
        DataClear
        WaitKeepAlive
        Winsock(1).SendData "LAST" + vbCrLf
        
    End If
End Sub

Public Sub GotoNext()
    On Error Resume Next
    
    'set start/end times for statistics
    CmdStatistic True, True
    
    'not connected
    If State1 = NNTP_STAT_DISCONNECTED Then
        RaiseEvent DoneGotoNext(NNTP_RC_NOTCONNECTED)
        
    'busy
    ElseIf State1 <> NNTP_STAT_CONNECTED Then
        RaiseEvent DoneGotoNext(NNTP_RC_BUSY)
    
    'next article
    Else
        State1 = NNTP_STAT_GOTONEXT
        DataClear
        WaitKeepAlive
        Winsock(1).SendData "NEXT" + vbCrLf
        
    End If
End Sub



Public Sub SelectGroup()
    On Error Resume Next
    
    'set start/end times for statistics
    CmdStatistic True, True
    
    'not connected
    If State1 = NNTP_STAT_DISCONNECTED Then
        RaiseEvent DoneSelectGroup(NNTP_RC_NOTCONNECTED)
        
    'busy
    ElseIf State1 <> NNTP_STAT_CONNECTED Then
        RaiseEvent DoneSelectGroup(NNTP_RC_BUSY)
    
    'select newsgroup
    Else
        State1 = NNTP_STAT_SELECTING
        WaitKeepAlive
        Winsock(1).SendData "GROUP " + Newsgroup1 + vbCrLf
        
    End If
End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    Dim Rcv As String, Lines() As String, i As Long
  
    If bytesTotal > 0 Then
    
        'keep time of last data receival
        LastReceive = Timer
        
        'count reveived bytes (total & since last command)
        Received1 = Received1 + bytesTotal
        Bytes1 = Bytes1 + bytesTotal
        
        'get received data
        Winsock(1).GetData Rcv, vbString
        
        'process data line by line
        Lines() = Split(Holdover + Rcv, vbCrLf, , vbBinaryCompare)
        For i = LBound(Lines) To UBound(Lines)
            
            'line complete
            If i < UBound(Lines) Then
                ProcessLine Lines(i)
                
                'raise debug event
                If Debugger1 Then
                    RaiseEvent Debugger(Lines(i))
                End If

            'line incomplete (spool)
            Else
                Holdover = Lines(i)
            End If
        Next i
        
        'raise receive-event
        RaiseEvent Receive(Data.Count)
    
    End If
    
End Sub

Private Sub Winsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    'socket error (also raised when server not found)
    DisconnectNow NNTP_RC_SOCKERR
End Sub

Private Sub DataClear()
    On Error Resume Next
    'reset data variables
    Data.Clear
    DataSize1 = 0
    Content = NNTP_DATA_EMPTY
    Holdover = ""
    'reset some article variables (not the article header vars)
    Text1 = ""
    Message1 = ""
End Sub

Public Property Get Newsgroup() As String
Attribute Newsgroup.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the name of current newsgroup
    Newsgroup = Newsgroup1
End Property

Public Property Let Newsgroup(ByVal vNewValue As String)
    On Error Resume Next
    'sets the name of newsgroup
    Newsgroup1 = Trim$(vNewValue)
End Property

Public Property Get Received() As Long
Attribute Received.VB_MemberFlags = "400"
    On Error Resume Next
    'returns total amount of received data
    Received = Received1
End Property

Public Property Let Received(vNewValue As Long)
    On Error Resume Next
    '(re)sets total amount of received data
    Received1 = vNewValue
End Property

Public Property Get FirstArticle() As Long
Attribute FirstArticle.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the number of the first article
    FirstArticle = FirstArticle1
End Property

Public Property Let FirstArticle(ByVal vNewValue As Long)
    On Error Resume Next
    'sets the number of the first article
    FirstArticle1 = vNewValue
End Property

Public Property Get LastArticle() As Long
Attribute LastArticle.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the number of the last article
    LastArticle = LastArticle1
End Property

Public Property Let LastArticle(ByVal vNewValue As Long)
    On Error Resume Next
    'sets the number of the last article
    LastArticle1 = vNewValue
End Property

Public Property Get ArticleCount() As Long
Attribute ArticleCount.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the number of received articles
    ArticleCount = ArticleCount1
End Property

Public Sub PostArticle()
    On Error Resume Next
    Dim Dat As String
    
    'set start/end times for statistics
    CmdStatistic True, True
    
    'not connected
    If State1 = NNTP_STAT_DISCONNECTED Then
        RaiseEvent DonePost(NNTP_RC_NOTCONNECTED)
        
    'busy
    ElseIf State1 <> NNTP_STAT_CONNECTED Then
        RaiseEvent DonePost(NNTP_RC_BUSY)
    
    'name of newsgroup missing
    ElseIf Newsgroup1 = "" Then
        RaiseEvent DonePost(NNTP_RC_NOTINANEWSGROUP)
    
    'anything else missing?
    ElseIf Author1 = "" Or Subject1 = "" Or Text1 = "" Then
        RaiseEvent DonePost(NNTP_RC_POSTINGFAILED)
    
    'ok so far
    Else
        'increase counter of message id to avoid duplicate
        MsgIdCnt = MsgIdCnt + 1
        
        'format date
        Dat = Mid$("SunMonTueWedThuFriSat", Weekday(Now) * 3 - 2, 3) + ", " + _
              Format(Now, "dd") + " " + _
              Mid$("JanFebMarAprMayJunJulAugSepOctNovDec", Month(Now) * 3 - 2, 3) + " " + _
              Format(Now, "yy hh:mm:ss") + " GMT"
        
        'build header data
        Send1 = "From: " + Author1 + vbCrLf + _
                "Newsgroups: " + Newsgroup1 + vbCrLf + _
                "Subject: " + Subject1 + vbCrLf + _
                "Message-ID: <" + Format(Now(), "yymmddhhmmss" + CStr(MsgIdCnt) + "@") + Host1 + ">" + vbCrLf + _
                "Date: " + Dat + vbCrLf + _
                IIf(Newsreader1 = "", "", "X-Newsreader: " + Newsreader1 + vbCrLf) + _
                IIf(References1 = "", "", "References: " + References1 + vbCrLf) + _
                vbCrLf + _
                Text1 + vbCrLf + _
                "." + vbCrLf

        'post now
        State1 = NNTP_STAT_POSTING
        WaitKeepAlive
        Winsock(1).SendData "POST" + vbCrLf
    End If
End Sub

Public Property Get Timeout() As Integer
Attribute Timeout.VB_Description = "After this time of inactivity, the control sends a timeout message to the application."
    On Error Resume Next
    'returns the time in seconds until a timeout is detected
    Timeout = Timeout1
End Property

Public Property Let Timeout(ByVal vNewValue As Integer)
    On Error Resume Next
    'invalid value
    If vNewValue < 0 Then
        Err.Raise 9
    'set the timeout time in secondes
    Else
        Timeout1 = vNewValue
    End If
End Property


Public Property Get Refresh() As Integer
Attribute Refresh.VB_Description = "In case of inactivity for the specified time, the control sends a command to the host to keep the connection alive."
    On Error Resume Next
    'returns the refresh time in seconds
    Refresh = Refresh1
End Property

Public Property Let Refresh(vNewValue As Integer)
    On Error Resume Next
    'invalid value
    If vNewValue < 0 Then
        Err.Raise 9
    'set the refresh time in seconds
    Else
        Refresh1 = vNewValue
    End If
End Property

Public Property Get Host() As String
Attribute Host.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the name or ip-address of the newsserver
    Host = Host1
End Property

Public Property Let Host(ByVal vNewValue As String)
    On Error Resume Next
    'sets the name or ip-address of the newsserver
    Host1 = Trim$(vNewValue)
End Property

Public Property Get Username() As String
Attribute Username.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the login username
    Username = Username1
End Property

Public Property Let Username(ByVal vNewValue As String)
    On Error Resume Next
    'sets the login username
    Username1 = Trim$(vNewValue)
End Property

Public Property Get Password() As String
Attribute Password.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the login password
    Password = Password1
End Property

Public Property Let Password(ByVal vNewValue As String)
    On Error Resume Next
    'sets the login password
    Password1 = Trim$(vNewValue)
End Property

Public Property Get MsgNo() As Long
Attribute MsgNo.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the message number of the current article
    MsgNo = MsgNo1
End Property

Public Property Let MsgNo(ByVal vNewValue As Long)
    On Error Resume Next
    'set the message number of the current article
    If MsgNo1 <> vNewValue Then
        MsgNo1 = vNewValue
    End If
End Property

Public Property Get MsgId() As String
Attribute MsgId.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the message-id of the current article
    MsgId = MsgId1
End Property

Public Property Let MsgId(ByVal vNewValue As String)
    On Error Resume Next
    'sets the message-id
    vNewValue = Trim$(vNewValue)
    If MsgId1 <> vNewValue Then
        MsgId1 = vNewValue
    End If
End Property


Public Property Get References() As String
Attribute References.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the reference field of the article header
    References = References1
End Property

Public Property Let References(ByVal vNewValue As String)
    On Error Resume Next
    'sets the reference field of the article header
    References1 = Trim$(vNewValue)
End Property


Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the complete article as received (header/body)
    Dim i As Long, j As Long
    
    'only applicable if contains text
    If Content = NNTP_DATA_TEXT Then
        
        'join the string only once
        If Text1 = "" Then
            
            'data received?
            If Data.Count > 0 Then
                
                'create new string
                Text1 = String$(DataSize1 + (Data.Count - 1) * 2, Chr$(0))
                
                'set line data into string
                j = 1
                For i = 1 To Data.Count
                    
                    'insert text
                    If Len(Data.Item(i)) > 0 Then
                        Mid$(Text1, j) = Data.Item(i)
                        j = j + Len(Data.Item(i))
                    End If
                    
                    'inset carriage return / line feed
                    If i < Data.Count Then
                        Mid$(Text1, j) = vbCrLf
                        j = j + 2
                    End If
                    
                Next i
            
            End If
        
        End If
        
        'return value
        Text = Text1
    
    End If
End Property

Public Property Let Text(ByVal vNewValue As String)
    On Error Resume Next
    'sets the articles text
    Text1 = vNewValue
    If vNewValue = "" Then
        Content = NNTP_DATA_EMPTY
    Else
        Content = NNTP_DATA_TEXT
    End If
End Property

Public Property Get Debugger() As Boolean
    On Error Resume Next
    'returns the debug mode
    Debugger = Debugger1
End Property

Public Property Let Debugger(ByVal vNewValue As Boolean)
    On Error Resume Next
    'sets the debug mode (when set to true, a debug event will be raised for each received line)
    Debugger1 = vNewValue
End Property

Public Property Get Hostinfo() As String
Attribute Hostinfo.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the hostinfo string (it will be received during login)
    Hostinfo = Hostinfo1
End Property

Private Sub CmdStatistic(SetStart As Boolean, SetEnd As Boolean)
    On Error Resume Next
    'reset the command start time
    If SetStart Then
        CmdStartTime = Timer
        CmdStartBytes = Received1
    End If
    'reset the command end time
    If SetEnd Then
        CmdEndTime = Timer
        CmdEndBytes = Received1
    End If
End Sub

Public Property Get Bps() As Single
Attribute Bps.VB_MemberFlags = "400"
    On Error Resume Next
    'calculate bytes per second
    If CmdEndTime <= CmdStartTime Then
        Bps = 0
    Else
        Bps = (CmdEndBytes - CmdStartBytes) / (CmdEndTime - CmdStartTime)
    End If
End Property

Public Property Get State() As TNNTPState
Attribute State.VB_MemberFlags = "400"
    On Error Resume Next
    'returns current control state
    State = State1
End Property

Public Function SaveData(Filename As String) As Boolean
    On Error GoTo Fehler
    'write the data content quickly to a file
    Dim i As Long, Fnr As Integer
    
    'get free file number
    Fnr = FreeFile
    
    'save data to filename
    Open Filename For Output As #Fnr
    For i = 1 To Data.Count
        Print #Fnr, Data.Item(i)
        DoEvents
    Next i
    
    'ready
    Close #Fnr
    SaveData = True
Exit Function

Fehler:
    Close #Fnr
    SaveData = False
End Function

Public Property Get RcString(Rc As TNNTPRc) As String
Attribute RcString.VB_MemberFlags = "400"
    On Error Resume Next
    'returns a description to a given return code
    Select Case Rc
    Case NNTP_RC_NONE: RcString = ""
    Case NNTP_RC_NOTCONNECTED: RcString = "Not connected"
    Case NNTP_RC_HOSTNOTFOUND: RcString = "Host not found"
    Case NNTP_RC_BUSY: RcString = "Busy"
    Case NNTP_RC_SOCKERR: RcString = "Socket error"
    Case NNTP_RC_CANCELED: RcString = "Canceled"
    Case NNTP_RC_TIMEOUT: RcString = "Timeout"
    Case NNTP_RC_POSTING: RcString = "Posting allowed"
    Case NNTP_RC_NOPOSTING: RcString = "No posting allowed"
    Case NNTP_RC_CLOSING: RcString = "Connection closed"
    Case NNTP_RC_GROUPSELECTED: RcString = "Newsgroup selected"
    Case NNTP_RC_LIST: RcString = "List follows"
    Case NNTP_RC_ARTICLE: RcString = "Article follows"
    Case NNTP_RC_HEAD: RcString = "Head follows"
    Case NNTP_RC_BODY: RcString = "Body follows"
    Case NNTP_RC_STAT: RcString = "Statistic follows"
    Case NNTP_RC_DATAFOLLOWS: RcString = "Data follows"
    Case NNTP_RC_NEWARTICLES: RcString = "New articles follow"
    Case NNTP_RC_NEWGROUPS: RcString = "New newsgroups follow"
    Case NNTP_RC_TRANSFERRED: RcString = "Transferred"
    Case NNTP_RC_POSTEDOK: RcString = "Posted"
    Case NNTP_RC_AUTHENTICATIONACCEPTED: RcString = "Authentication accepted"
    Case NNTP_RC_SENDARTICLETRANSFERRED: RcString = "Send article to be transferred"
    Case NNTP_RC_SENDARTICLE: RcString = "Send article to be posted"
    Case NNTP_RC_MOREAUTHENTICATIONREQUIRED: RcString = "More authentication required"
    Case NNTP_RC_SERVICEDISCONTINUED: RcString = "Service discontinued"
    Case NNTP_RC_NOSUCHNEWSGROUP: RcString = "No such newsgroup"
    Case NNTP_RC_NOTINANEWSGROUP: RcString = "Not in a newsgroup"
    Case NNTP_RC_NOARTICLESELECTED: RcString = "No article selected"
    Case NNTP_RC_NONEXTARTICLE: RcString = "No next article"
    Case NNTP_RC_NOLASTARTICLE: RcString = "No last article"
    Case NNTP_RC_BADARTICLENUMBER: RcString = "Bad article number"
    Case NNTP_RC_ARTICLENOTFOUND: RcString = "Article not found"
    Case NNTP_RC_DONTSEND: RcString = "Do not send"
    Case NNTP_RC_TRANSFERFAILED: RcString = "Transfer failed"
    Case NNTP_RC_ARTICLEREJECTED: RcString = "Article rejected"
    Case NNTP_RC_POSTINGNOTALLOWD: RcString = "Posting permitted"
    Case NNTP_RC_POSTINGFAILED: RcString = "Posting failed"
    Case NNTP_RC_AUTHENTICATIONREQUIRED: RcString = "Authentication required"
    Case NNTP_RC_AUTHENTICATIONREJECTED: RcString = "Authentication rejected"
    Case NNTP_RC_NOSUCHCOMMAND: RcString = "No such command"
    Case NNTP_RC_SYNTAXERROR: RcString = "Syntax error"
    Case NNTP_RC_ACCESSRESTRICTED: RcString = "Access restricted"
    Case NNTP_RC_PROGRAMNOTFAULT: RcString = "Program not fault"
    Case NNTP_RC_CONNECTIONFAILURE1: RcString = "Connection failture"
    Case NNTP_RC_CONNECTIONFAILURE2: RcString = "Connection failture"
    Case Else: RcString = "Error number " + CStr(Rc)
    End Select
End Property

Private Sub DisconnectNow(Rc As TNNTPRc)
    On Error Resume Next
    'forcely disconnects from a newsserver
    Tmr.Enabled = False
    Unload Winsock(1)
    'set start/end times for statistics
    CmdStatistic False, True
    ClearVars 'Newsgruppen und Artikelrelevante Variabeln löschen
    State1 = NNTP_STAT_DISCONNECTED
    RaiseEvent DoneDisconnect(Rc)
End Sub

Public Property Get Author() As String
Attribute Author.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the name of the author
    Author = Author1
End Property

Public Property Let Author(ByVal vNewValue As String)
    On Error Resume Next
    'sets the name of the author
    Author1 = Trim$(vNewValue)
End Property

Public Property Get Subject() As String
Attribute Subject.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the subject of the current article
    Subject = Subject1
End Property

Public Property Let Subject(ByVal vNewValue As String)
    On Error Resume Next
    'sets the subject of the current article (for posting)
    Subject1 = Trim$(vNewValue)
End Property

Private Sub WaitKeepAlive()
    On Error Resume Next
    'wait until the keep-alive command is completed
    While KeepAlive
        DoEvents
    Wend
End Sub

Private Sub ClearVars()
    On Error Resume Next
    
    'initialize newsgroup variables
    Newsgroup1 = ""
    FirstArticle1 = 0
    LastArticle1 = 0
    ArticleCount1 = 0
    
    'initialize article variables
    MsgNo1 = 0
    Subject1 = ""
    Author1 = ""
    MsgId1 = ""
    References1 = ""
    Bytes1 = 0
    Lines1 = 0
    XRef1 = ""
    Newsreader1 = ""
    Text1 = ""
    Message1 = ""
    
    'initialize internal variables
    Received1 = 0
End Sub

Public Property Get Copyright() As String
Attribute Copyright.VB_Description = "This is copyright information"
Attribute Copyright.VB_ProcData.VB_Invoke_Property = ";Verschiedenes"
    On Error Resume Next
    'returns the developers info (you dont remove this, do you?)
    Copyright = "©2001 U.Keller, http://www.uwekeller.de"
End Property

Public Property Let Copyright(vNewValue As String)
    'copyright is write protected
End Property

Public Property Get Newsreader() As String
Attribute Newsreader.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the name of the used newsreader
    Newsreader = Newsreader1
End Property

Public Property Let Newsreader(ByVal vNewValue As String)
    On Error Resume Next
    'sets the name of the used newsreader (set before post)
    Newsreader1 = Trim$(vNewValue)
End Property

Public Property Get Message() As String
Attribute Message.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the clear message without any header or binary attachment
    Dim i As Long, TextStart As Boolean
    
    'is data available?
    If Content = NNTP_DATA_TEXT Then
        
        'extract message part if not done yet
        If Message1 = "" Then
        
            'detect & extract
            For i = 1 To Data.Count
                
                'first empty line is end of header
                If TextStart = False And Data(i) = "" Then
                    TextStart = True
                
                'first text line found
                ElseIf TextStart = False And InStr(Data(i), ": ") = 0 Then
                    TextStart = True
                    Message1 = Data(i) + vbCrLf
                    
                'MIME text block found
                ElseIf Left$(Data(i), 24) = "Content-Type: text/plain" Then
                    TextStart = True
                    
                'UUE binary found
                ElseIf Left$(Data(i), 7) = "begin 6" Then
                    Exit For
                
                'MIME binary found
                ElseIf Left$(Data(i), 14) = "Content-Type: " Then
                    Exit For
                
                'get message
                ElseIf TextStart Then
                    Message1 = Message1 + Data(i) + vbCrLf
                    
                End If
            Next i
            
        End If
    
        Message = Message1
    End If
End Property

Public Property Let Message(ByVal vNewValue As String)
    On Error Resume Next
    'use the text property instead
    Text = vNewValue
End Property

Public Property Get DataCount() As Long
Attribute DataCount.VB_MemberFlags = "400"
    On Error Resume Next
    'empty, no data available
    If Content = NNTP_DATA_EMPTY Then
        DataCount = 0
    'data is text (header, article, body)
    ElseIf Content = NNTP_DATA_TEXT Then
        DataCount = 1
    'data are newsgroups or headers
    Else
        DataCount = Data.Count
    End If
End Property

Public Property Get DataIndex() As Long
Attribute DataIndex.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the pointer to the current data line
    DataIndex = DataIndex1
End Property

Public Property Let DataIndex(ByVal vNewValue As Long)
    On Error Resume Next
    'sets the pointer to the current data line
    Dim f() As String
    
    'out of range
    If (vNewValue < 1 Or vNewValue > Data.Count) Or (Content = NNTP_DATA_TEXT And vNewValue <> 1) Then
        Err.Raise 9
    
    'newsgroups
    ElseIf Content = NNTP_DATA_GROUP Then
        'initialize variables
        Newsgroup1 = ""
        LastArticle1 = 0
        FirstArticle1 = 0
        Posting1 = False
        'set new data
        DataIndex1 = vNewValue
        f() = Split(Data(DataIndex1), " ")
        Newsgroup1 = f(0)
        LastArticle1 = f(1)
        FirstArticle1 = f(2)
        Posting1 = f(3) = "y"
        
    'headers (XOVER)
    ElseIf Content = NNTP_DATA_XOVER Then
        'initialize variables
        MsgNo1 = 0
        Subject1 = ""
        Author1 = ""
        DateTime1 = 0
        MsgId1 = 0
        References1 = ""
        Bytes1 = 0
        Lines1 = 0
        XRef1 = ""
        'set new data
        DataIndex1 = vNewValue
        f() = Split(Data(DataIndex1), vbTab)
        MsgNo1 = f(0)
        Subject1 = f(1)
        Author1 = f(2)
        DateTime1 = CorrectDate(f(3))
        MsgId1 = f(4)
        References1 = f(5)
        Bytes1 = f(6)
        Lines1 = f(7)
        XRef1 = f(8)
    
    End If
End Property

Public Property Get Posting() As Boolean
Attribute Posting.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the posting flag of the current newsserver or newsgroup
    Posting = Posting1
End Property

Public Property Get DateTime() As Date
Attribute DateTime.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the date time (article or newsgroups)
    DateTime = DateTime1
End Property

Public Property Let DateTime(ByVal vNewValue As Date)
    On Error Resume Next
    'sets the date/time of the article or newsgroups
    DateTime1 = vNewValue
End Property

Public Property Get Bytes() As Long
Attribute Bytes.VB_MemberFlags = "400"
    'returns the size of the article or other received data in bytes
    On Error Resume Next
    'return size
    Bytes = Bytes1
End Property

Public Property Get Lines() As Long
Attribute Lines.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the size of article in lines
    Lines = Lines1
End Property

Public Property Get XRef() As String
Attribute XRef.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the cross-reference data
    XRef = XRef1
End Property

Public Property Get DataText() As String
Attribute DataText.VB_MemberFlags = "400"
    On Error Resume Next
    'returns the complete text of the current data line
    If Content <> NNTP_DATA_EMPTY Then
        If DataIndex1 >= 1 Or DataIndex1 <= Data.Count Then
            DataText = Data(DataIndex1)
        End If
    End If
End Property


Private Sub ProcessLine(Line As String)
    On Error Resume Next
    Dim w() As String
    'process a received line
    
    'according current control state
    Static Rc As TNNTPRc
    
    'keep-alive
    If KeepAlive Then
        If Line = "." Then
            KeepAlive = False
        End If
        
    'connecting
    ElseIf State1 = NNTP_STAT_CONNECTING Then
        Rc = Val(Left$(Line, 3))
        'store hostinfo
        If Rc = NNTP_RC_POSTING Or Rc = NNTP_RC_NOPOSTING Then
            Hostinfo1 = Mid$(Line, 5)
            Posting1 = (Rc = NNTP_RC_POSTING)
        'error during login
        ElseIf Rc >= 500 Then
            DisconnectNow Rc
            Exit Sub
        End If
        'authorization required?
        If SendUsername = False And SendPassword = False Then
            Select Case Rc
            'login ok
            Case NNTP_RC_POSTING, NNTP_RC_NOPOSTING, NNTP_RC_AUTHENTICATIONACCEPTED
                ClearVars
                State1 = NNTP_STAT_CONNECTED
                'set start/end times for statistics
                CmdStatistic False, True
                RaiseEvent DoneConnect(Rc)
            'login not possible
            Case Else
                DisconnectNow Rc
            End Select
        'send username
        ElseIf SendUsername Then
            SendUsername = False
            Winsock(1).SendData "authinfo user " + Username1 + vbCrLf
        'send password
        ElseIf SendPassword Then
            SendPassword = False
            Winsock(1).SendData "authinfo pass " + Password1 + vbCrLf
        End If
        
    'disconnecting
    ElseIf State1 = NNTP_STAT_DISCONNECTING Then
        Rc = Val(Left$(Line, 3))
        DisconnectNow Rc
    
    'getting list of newsgroups
    ElseIf State1 = NNTP_STAT_GETLIST Then
        'first line
        If Content = NNTP_DATA_EMPTY Then
            Content = NNTP_DATA_GROUP
            Rc = Val(Left$(Line, 3))
            'error
            If Rc <> NNTP_RC_LIST And Rc <> NNTP_RC_NEWGROUPS Then
                State1 = NNTP_STAT_CONNECTED
                'set start/end times for statistics
                CmdStatistic False, True
                RaiseEvent DoneGetList(Rc)
                Exit Sub
            End If
        'done
        ElseIf Line = "." Then
            State1 = NNTP_STAT_CONNECTED
            'set start/end times for statistics
            CmdStatistic False, True
            RaiseEvent DoneGetList(Rc)
        'store line in data
        Else
            Data.Add Line
            DataSize1 = DataSize1 + Len(Line)
        End If
                
    'selecting newsgroup
    ElseIf State1 = NNTP_STAT_SELECTING Then
        w() = Split(Line)
        Rc = w(0)
        'ok
        If Rc = NNTP_RC_GROUPSELECTED Then
            State1 = NNTP_STAT_CONNECTED
            ArticleCount1 = w(1)
            FirstArticle1 = w(2)
            LastArticle1 = w(3)
            MsgNo1 = FirstArticle1
        'error
        Else
            State1 = NNTP_STAT_CONNECTED
            ArticleCount1 = 0
            FirstArticle1 = 0
            LastArticle1 = 0
        End If
        'set start/end times for statistics
        CmdStatistic False, True
        RaiseEvent DoneSelectGroup(Rc)
    
    'Artikelliste
    ElseIf State1 = NNTP_STAT_GETXOVER Then
        'first line
        If Content = NNTP_DATA_EMPTY Then
            Content = NNTP_DATA_XOVER
            Rc = Val(Left$(Line, 3))
            'error
            If Rc <> NNTP_RC_DATAFOLLOWS Then
                State1 = NNTP_STAT_CONNECTED
                'set start/end times for statistics
                CmdStatistic False, True
                RaiseEvent DoneGetXover(Rc)
                Exit Sub
            End If
        'done
        ElseIf Line = "." Then
            State1 = NNTP_STAT_CONNECTED
            'set start/end times for statistics
            CmdStatistic False, True
            RaiseEvent DoneGetXover(Rc)
        'store line in data
        Else
            Data.Add Line
            DataSize1 = DataSize1 + Len(Line)
        End If
        
    'getting state
    ElseIf State1 = NNTP_STAT_GETSTAT Then
        w() = Split(Line)
        Rc = w(0)
        If Rc = NNTP_RC_STAT Then
            MsgNo1 = w(1)
            MsgId1 = w(2)
        End If
        State1 = NNTP_STAT_CONNECTED
        'set start/end times for statistics
        CmdStatistic False, True
        RaiseEvent DoneGetStat(Rc)

    'getting article
    ElseIf State1 = NNTP_STAT_GETARTICLE Then
        'first line
        If Content = NNTP_DATA_EMPTY Then
            Content = NNTP_DATA_TEXT
            w() = Split(Line)
            Rc = w(0)
            'ok
            If Rc = NNTP_RC_ARTICLE Then
                MsgNo1 = w(1)
                MsgId1 = w(2)
            'error
            Else
                State1 = NNTP_STAT_CONNECTED
                'set start/end times for statistics
                CmdStatistic False, True
                RaiseEvent DoneGetArticle(Rc)
                Exit Sub
            End If
        'done
        ElseIf Line = "." Then
            State1 = NNTP_STAT_CONNECTED
            'set start/end times for statistics
            CmdStatistic False, True
            RaiseEvent DoneGetArticle(Rc)
        'store line in data
        Else
            Data.Add Line
            DataSize1 = DataSize1 + Len(Line)
        End If

    'getting article header
    ElseIf State1 = NNTP_STAT_GETHEADER Then
        'first line
        If Content = NNTP_DATA_EMPTY Then
            Content = NNTP_DATA_TEXT
            w() = Split(Line)
            Rc = w(0)
            'ok
            If Rc = NNTP_RC_HEAD Then
                MsgNo1 = w(1)
                MsgId1 = w(2)
            'error
            Else
                State1 = NNTP_STAT_CONNECTED
                'set start/end times for statistics
                CmdStatistic False, True
                RaiseEvent DoneGetHeader(Rc)
                Exit Sub
            End If
        'done
        ElseIf Line = "." Then
            State1 = NNTP_STAT_CONNECTED
            'set start/end times for statistics
            CmdStatistic False, True
            RaiseEvent DoneGetHeader(Rc)
        'store line in data
        Else
            Data.Add Line
            DataSize1 = DataSize1 + Len(Line)
        End If

    'getting body
    ElseIf State1 = NNTP_STAT_GETBODY Then
        'first line
        If Content = NNTP_DATA_EMPTY Then
            Content = NNTP_DATA_TEXT
            w() = Split(Line)
            Rc = w(0)
            'ok
            If Rc = NNTP_RC_BODY Then
                MsgNo1 = w(1)
                MsgId1 = w(2)
            'error
            Else
                State1 = NNTP_STAT_CONNECTED
                'set start/end times for statistics
                CmdStatistic False, True
                RaiseEvent DoneGetBody(Rc)
                Exit Sub
            End If
        'done
        ElseIf Line = "." Then
            State1 = NNTP_STAT_CONNECTED
            'set start/end times for statistics
            CmdStatistic False, True
            RaiseEvent DoneGetBody(Rc)
        'store line in data
        Else
            Data.Add Line
            DataSize1 = DataSize1 + Len(Line)
        End If

    
    'seeking last article
    ElseIf State1 = NNTP_STAT_GOTOLAST Then
        w() = Split(Line)
        Rc = w(0)
        'ok
        If Rc = NNTP_RC_STAT Then
            MsgNo1 = w(1)
            MsgId1 = w(2)
        End If
        State1 = NNTP_STAT_CONNECTED
        'set start/end times for statistics
        CmdStatistic False, True
        RaiseEvent DoneGotoLast(Rc)
    
    'seeking next article
    ElseIf State1 = NNTP_STAT_GOTONEXT Then
        w() = Split(Line)
        Rc = w(0)
        'ok
        If Rc = NNTP_RC_STAT Then
            MsgNo1 = w(1)
            MsgId1 = w(2)
        End If
        State1 = NNTP_STAT_CONNECTED
        'set start/end times for statistics
        CmdStatistic False, True
        RaiseEvent DoneGotoNext(Rc)
    
    'posting article
    ElseIf State1 = NNTP_STAT_POSTING Then
        w() = Split(Line)
        Rc = w(0)
        'ok, post now
        If Rc = NNTP_RC_SENDARTICLE Then
            WaitKeepAlive
            Winsock(1).SendData Send1 'Send1 ist schon beim PostArticle-Befehl formatiert worden
        'ready (240=NNTP_RC_POSTEDOK, 440=NNTP_RC_POSTINGNOTALLOWED, 441=NNTP_RC_POSTINGFAILED)
        Else
            State1 = NNTP_STAT_CONNECTED
            'set start/end times for statistics
            CmdStatistic False, True
            RaiseEvent DonePost(Rc)
        End If
    
    'unexpected
    Else
        Rc = Val(Left$(Line, 3))
        'disconnect required
        If Rc = NNTP_RC_PROGRAMNOTFAULT Or Rc = NNTP_RC_CLOSING Then
            DisconnectNow Rc
        'any other
        Else
            'ignore
        End If
    End If
End Sub

Private Function CorrectDate(ByVal String1 As String, Optional Timeshift As Boolean = False) As String
    On Error Resume Next
    'corrects a given date
    Dim i As Integer, j As Integer
    Dim Day As Integer, Month As Integer, Year As Integer, Hour As Integer, Minute As Integer, Second As Integer
    Dim Monthdays
    
    'use original by default
    CorrectDate = String1
    
    'find first sign of day
    For i = 1 To Len(String1)
        Select Case Mid$(String1, i, 1)
        Case "0" To "9"
            Exit For
        End Select
    Next i
    
    'day not found
    If i >= Len(String1) Then
        Exit Function
        
    'find last sign of day
    Else
        For j = i + 1 To Len(String1)
            Select Case Mid$(String1, j, 1)
            Case "0" To "9"
                'nop
            Case Else
                Exit For
            End Select
        Next j
        Day = Val(Mid$(String1, i, j - i))
        
        'find first sign of month
        For i = j To Len(String1)
            Select Case Mid$(String1, i, 1)
            Case "0" To "9", "a" To "z"
                Exit For
            End Select
        Next i
        
        'month not found
        If i >= Len(String1) Then
            Exit Function
        
        'find last sign of month
        Else
            Select Case Mid$(String1, i, 1)
            Case "0" To "9"
                For j = i + 1 To Len(String1)
                    Select Case Mid$(String1, j, 1)
                    Case "0" To "9"
                        'nop
                    Case Else
                        Exit For
                    End Select
                Next j
                Month = Val(Mid$(String1, i, j - 1 - 1))
            Case "a" To "z", "ä", "ö", "ü", "ß"
                j = i + 3
                Month = 1 + InStr("janfebmaraprmayjunjulaugsepoctnovdec", Mid$(String1, i, 3)) / 3
            'error
            Case Else
                Exit Function
            End Select
            
            'find first sign of year
            For i = j To Len(String1)
                Select Case Mid$(String1, i, 1)
                Case "0" To "9"
                    Exit For
                End Select
            Next i
            
            'year not found
            If i >= Len(String1) Then
                Exit Function
            
            'find last signe of year
            Else
                For j = i To Len(String1)
                    Select Case Mid$(String1, j, 1)
                    Case "0" To "9"
                        'nop
                    Case Else
                        Exit For
                    End Select
                Next j
                Year = Val(Mid$(String1, i, j - i))
                If Year < 80 Then
                    Year = Year + 2000
                ElseIf Year < 100 Then
                    Year = Year + 1900
                End If
                
                'find time
                For i = j To Len(String1)
                    Select Case Mid$(String1, i, 1)
                    Case "0" To "9"
                        Exit For
                    End Select
                Next i
                If i < Len(String1) Then
                    Hour = Val(Mid$(String1, i, 2))
                    Minute = Val(Mid$(String1, i + 3, 2))
                    Second = Val(Mid$(String1, i + 6, 2))
                    
                    
                    'take time shift into account
                    If Timeshift Then
                        Monthdays = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
                        j = InStr(i, String1, "-")
                        If j > 0 Then
                            Hour = Hour + Val(Mid$(String1, j + 1, 2))
                            If Hour > 23 Then
                                Hour = Hour - 24
                                Day = Day + 1
                                If Day > Monthdays(Month) Then
                                    Day = 1
                                    Month = Month + 1
                                    If Month > 12 Then
                                        Month = 1
                                        Year = Year + 1
                                    End If
                                End If
                            End If
                        End If
                        
                        'time shift +
                        j = InStr(i, String1, "+")
                        If j > 0 Then
                            Hour = Hour - Val(Mid$(String1, j + 1, 2))
                            If Hour < 0 Then
                                Hour = Hour + 24
                                Day = Day - 1
                                If Day < 1 Then
                                    Month = Month - 1
                                    If Month < 1 Then
                                        Month = 12
                                        Year = Year - 1
                                    End If
                                    Day = Monthdays(Month)
                                End If
                            End If
                        End If
                        
                    End If
                End If
            End If
        End If
    End If
    
    'date found - return formatted date
    CorrectDate = Format(Year, "0000") + "/" + Format(Month, "00") + "/" + Format(Day, "00") + " " + Format(Hour, "00") + ":" + Format(Minute, "00") + ":" + Format(Second, "00")
End Function


Public Property Get DataSize() As Long
Attribute DataSize.VB_MemberFlags = "400"
    On Error Resume Next
    'return the size of the current data
    DataSize = DataSize1
End Property

