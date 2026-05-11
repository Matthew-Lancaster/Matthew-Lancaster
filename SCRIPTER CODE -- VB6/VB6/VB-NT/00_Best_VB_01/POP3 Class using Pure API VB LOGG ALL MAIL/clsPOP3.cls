VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPOP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Event DataArrival(sData As String) ' Used to Inform client APP Data arrived
Public Event DataSent(sData As String)    ' Used to Inform client APP Data was sent
Public Event Progress(dCurrectBytes As Double, dTotalSize As Double) ' Fires Progress
Public Event POP3Error(ByVal sData As String) ' Used to inform client APP an Error Occured

Private mvarServer            As String   ' Stores the Server Name/IP
Private mvarUserName          As String   ' Stores UserName used
Private mvarPassword          As String   ' Stores the Password used
Private mvarPopPort           As Integer  ' Stores Port used
Private mvarTimeOut           As Long     ' The Timeout for Communications in Seconds
Private mvarLastErrorCode     As Integer  ' Holds the Last error code
Private mvarisConnected       As Boolean  ' Is set to True when Connected to Server
Private mvarAPOPSupported     As Boolean  ' Is set to True if the Server supports APOP

Private udtWinsockData        As WSAData  ' API Socket Data
Private lngSocketHandle       As Long     ' The Connected Socket in Use

Private dblCurrentMessageSize As Double   ' Holds the Current Downloading Message size in bytes
Private lTotalBytes           As Double   ' Used to count the Download progress
Private bCancel               As Boolean  ' Set to True when cancelling out of connection
Private bBusy                 As Boolean  ' Set to True when socket is busy

' List of Error codes this class could produce
Public Enum POPErrors
   ERR_NoError = 0
   ERR_InvalidServer = -10001
   ERR_InvalidUsername = -10002
   ERR_InvalidPassword = -10003
   ERR_InvalidLogon = -10004
   ERR_APOPLoginFailed = -10005
   ERR_FailedConnect = -10006
   ERR_UnKnownError = -10050
   ERR_ErrorSTAT = -10101
   ERR_ErrorRETR = -10102
   ERR_ErrorDELE = -10103
   ERR_ErrorHEADER = -10104
   ERR_ErrorTOP = -10105
   ERR_ErrorUIDL = -10106
   ERR_ErrorNOOP = -10107
   ERR_ErrorRSET = -10108
   ERR_ErrorLIST = -10109
   ERR_InvalidMessageID = -10110
   ERR_BadParameter = -10111
   ERR_MailBoxLocked = -10199
   ERR_SocketCreateError = -10500
   ERR_SocketConnectFail = -10501
   
   ERR_Timeout = -10999
End Enum


' *************Start of Properties ****************
' -------------------------------------------------------------
Public Property Get APOPSupported() As Boolean
   APOPSupported = mvarAPOPSupported
End Property
Public Property Get IsConnected() As Boolean
   IsConnected = mvarisConnected
End Property
Public Property Let Password(ByVal vData As String)
Attribute Password.VB_Description = "The POP server password"
   mvarPassword = vData
End Property
Public Property Get Password() As String
   Password = mvarPassword
End Property
Public Property Let TimeOut(ByVal vData As Long)
Attribute TimeOut.VB_Description = "The Timeout value to use"
   mvarTimeOut = vData
End Property
Public Property Get TimeOut() As Long
   TimeOut = mvarTimeOut
End Property
Public Property Let PopPort(ByVal vData As Integer)
Attribute PopPort.VB_Description = "Port to connect to the server"
   mvarPopPort = vData
End Property
Public Property Get PopPort() As Integer
   PopPort = mvarPopPort
End Property
Public Property Let UserName(ByVal vData As String)
Attribute UserName.VB_Description = "The Username of the mailbox"
   mvarUserName = vData
End Property
Public Property Get UserName() As String
   UserName = mvarUserName
End Property
Public Property Let Server(ByVal vData As String)
Attribute Server.VB_Description = "The Server address to connect to"
   mvarServer = vData
End Property
Public Property Get Server() As String
   Server = mvarServer
End Property
' *****************End Properties*******************
' -------------------------------------------------------------


' ==============================================================
' Procedure:    RETR
'
' Created on:   29 April 2005    By  Darren Lawrence
'
' Function :-
' '         Retreive a selected Message from the POP server
' ==============================================================
Public Function RETR(ByVal MessageID As Long) As String
Dim strResult As String
Dim strOutPut As String
Dim dblSize    As Double
      
100   If CheckMessageID(MessageID, dblSize) = True Then
110      Debug.Print ("[clsPOP3] (Sub) RETR")
120      On Error GoTo Err_RETR:

130      dblCurrentMessageSize = dblSize ' Record the Message Size for the progress event

140      strOutPut = "RETR " & MessageID
        
150      strResult = SendReceiveSocket(strOutPut, True, False, True)
160      strOutPut = Split(strResult, vbCrLf)(0)   ' We only want to send out the Status line to the DataArrival Event
170      RaiseEvent DataArrival(strOutPut)         ' Output the Message Size
180      RaiseEvent Progress(dblCurrentMessageSize, dblCurrentMessageSize) ' Set to Complete Size
190      If TestForError(strResult) <> ERR_NoError Then
            ' Some Error happend
200         RaiseEvent POP3Error(GetLastError)
210         RETR = GetLastError ' Return the Error string
            If mvarLastErrorCode = POPErrors.ERR_Timeout Then
               ' If Timout happens on RETR it may be that we just didnt wait long enough
               ' and if thats the case then we better close the session
               ' so any data that was waiting is discarded
               QUIT
            End If
         End If
220      strResult = Split(TidyReply(strResult), vbCrLf, 2)(1) ' Remove the top part of the message (+OK etc)
230      strResult = TidyReply(strResult)
240      RETR = strResult ' Return the Message

      Else
250      RaiseEvent POP3Error(GetLastError)
260      RETR = GetLastError
      End If
      Exit Function
      
280 Err_RETR:
      Debug.Print Err, Err.Description, Erl

End Function

' ==============================================================
' Procedure:    DELE
'
' Created on:   29 April 2005    By  Darren Lawrence
'
' Function :-
' '         Delete a message from the server
' '         The actual delete will only happen if you issue the QUIT command
' '         To start the Update stage.
' ==============================================================
Public Function DELE(ByVal MessageID As Long) As Boolean
Exit Function

Dim strResult As String
            
      strResult = SendReceiveSocket("DELE " & MessageID, False, True)
      
      DELE = True
      
      If TestForError(strResult) <> ERR_NoError Then
         ' Error occcured
         RaiseEvent POP3Error(GetLastError)
         DELE = False
      End If

End Function

' ==============================================================
' Procedure:    Top
'
' Created on:   23 April 2005    By  Darren Lawrence
'
' Function :-
' '         The Header command will return the header of a message
' '         plus the specified number of lines from the body
' '         Returns a Header Object
' ==============================================================
Public Function Header(ByVal MessageID As Long) As Email_Header
Dim strResult           As String
Dim strTemp             As String
Dim clsTemp_EmailHeader As Email_Header

      strResult = "TOP " & MessageID & " 0"
      strResult = SendReceiveSocket(strResult, True, False)

      If TestForError(strResult) <> ERR_NoError Then
         ' Error occcured - Return a blank header
         Set Header = New Email_Header
         RaiseEvent POP3Error(GetLastError)
         Exit Function
      End If
      On Error Resume Next
      strTemp = Split(strResult, vbCrLf)(0)
      If Err.Number > 0 Then
    Set clsTemp_EmailHeader = Nothing

      Exit Function
      End If
      RaiseEvent DataArrival(strTemp)
      
      strResult = Split(TidyReply(strResult), vbCrLf, 2)(1)
      
      Set clsTemp_EmailHeader = New Email_Header
      clsTemp_EmailHeader.PopulateHeader (strResult)
      Set Header = clsTemp_EmailHeader

      Set clsTemp_EmailHeader = Nothing

End Function

' ==============================================================
' Procedure:    Top
'
' Created on:   23 April 2005    By  Darren Lawrence
'
' Function :-
' '         The Top command will return the header of a message
' '         plus the specified number of lines from the message body
'
' ==============================================================
Public Function TOP(ByVal MessageID As Long, Optional ByVal Lines As Integer = 0) As String
Dim strResult As String
Dim strOutPut As String

      strOutPut = "TOP " & MessageID & " " & Lines

      strResult = SendReceiveSocket(strOutPut, True, False)

      If TestForError(strResult) <> ERR_NoError Then
         TOP = strResult
         RaiseEvent POP3Error(GetLastError)
         Exit Function
      End If
            
      strOutPut = Split(strResult, vbCrLf)(0)
      RaiseEvent DataArrival(strOutPut)
      
      strResult = Split(TidyReply(strResult), vbCrLf, 2)(1)
      
      strResult = Mid$(strResult, InStr(1, strResult, vbCrLf) + 2)
      strResult = Mid$(strResult, 1, Len(strResult) - 3)

      TOP = Split(TidyReply(strResult), vbCrLf, 2)(1)

End Function

' ==============================================================
' Procedure:    STAT
'
' Created on:   29 April 2005    By  Darren Lawrence
'
' Function :-
' '         Retrieve a STAT of the POP server
'
' '         This will return a line that indicates how many messages and how many octets
' '         in in the selected maildrop
'
' '         eg
' '         +OK 8 7362
' ==============================================================
Public Function STAT() As String
Dim strResult As String

      strResult = SendReceiveSocket("STAT", False, True)

      If TestForError(strResult) <> ERR_NoError Then
         ' Error occcured
         mvarLastErrorCode = ERR_ErrorSTAT
         STAT = strResult
         RaiseEvent POP3Error(GetLastError)
         Exit Function
      End If
      STAT = TidyReply(strResult)
      
End Function

' ==============================================================
' Procedure:    UIDL
'
' Created on:   29 April 2005    By  Darren Lawrence
'
' Function :-
' '      This will return a Unique ID List - this returns a Unique ID for each message
' '      or individual Message , this ID will be the same at all times even
' '      across sessions.
'
' ==============================================================
Public Function UIDL(Optional ByVal MessageID As Long = -1) As String
Dim strResult     As String
Dim strOutPut     As String
Dim bWaitForTerm  As Boolean

      strOutPut = "UIDL"
      bWaitForTerm = True
      If MessageID > -1 Then
         strOutPut = strOutPut & " " & MessageID
         bWaitForTerm = False
      End If
      
      strResult = SendReceiveSocket(strOutPut, bWaitForTerm, False)
         
      If TestForError(strResult) <> ERR_NoError Then
         ' Error occcured
         UIDL = strResult
         RaiseEvent POP3Error(GetLastError)
         Exit Function
      End If
      
      strOutPut = Split(strResult, vbCrLf)(0) ' We only want to return in the Event the top line
      RaiseEvent DataArrival(strOutPut)
      
      UIDL = strResult
      If InStr(UIDL, vbCrLf) Then
         UIDL = Split(UIDL, vbCrLf, 2)(1) ' Return result less top line
      End If
      UIDL = TidyReply(UIDL)
End Function

' ==============================================================
' Procedure:    NOOP
'
' Created on:   29 April 2005    By  Darren Lawrence
'
' Function :-
' '      Perform A NOOP,. this is used to test if the connection is still alive
' '      to a POP server
' ==============================================================
Public Function NOOP() As Boolean
Dim strResult As String

      strResult = SendReceiveSocket("NOOP", True, True)

      If TestForError(strResult) <> ERR_NoError Then
         ' Error occcured
         NOOP = False
         mvarLastErrorCode = ERR_ErrorNOOP
         RaiseEvent POP3Error(GetLastError)
         Exit Function
      End If
      NOOP = True

End Function

' ==============================================================
' Procedure:    RESET
'
' Created on:   29 April 2005    By  Darren Lawrence
'
' Function :-
' '      Perform RSET on POP server
' '      Resets previously emails set to delete to not delete
' ==============================================================
Public Function RESET() As Boolean
Dim strResult As String

      bBusy = False
      bCancel = True
      DoEvents
      bCancel = False
      strResult = SendReceiveSocket("RSET", True, True)

      If TestForError(strResult) <> ERR_NoError Then
         ' Error occcured
         mvarLastErrorCode = ERR_ErrorRSET
         RESET = False
         RaiseEvent POP3Error(GetLastError)
         Exit Function
      End If
      RESET = True

End Function

' ==============================================================
' Procedure:    LIST
'
' Created on:   29 April 2005    By  Darren Lawrence
'
' Function :-
' '      Return a list of Messages or the size of an individual Message
' ==============================================================
Public Function List(Optional ByVal MessageID As Long = -1) As String
Dim strResult     As String
Dim strOutPut     As String
Dim bWaitForTerm  As Boolean

      strOutPut = "LIST"
      bWaitForTerm = True
      If MessageID > -1 Then
         strOutPut = strOutPut & " " & MessageID
         bWaitForTerm = False
      End If
      
      strResult = SendReceiveSocket(strOutPut, bWaitForTerm, True)

      If TestForError(strResult) <> ERR_NoError Then
         ' Error occcured
         List = strResult
         RaiseEvent POP3Error(GetLastError)
         Exit Function
      End If
      strOutPut = Split(strResult, vbCrLf)(0)
      RaiseEvent DataArrival(strOutPut)
      ' LIST = strResult
      If InStr(strResult, vbCrLf) Then
         strResult = Split(strResult, vbCrLf, 2)(1) ' Return result less top line
      End If
      List = TidyReply(strResult)
End Function

' ==============================================================
' Procedure:    QUIT
'
' Created on:   29 April 2005    By  Darren Lawrence
'
' Function :-
' '      Terminate a POP session and perform any updates required
' ==============================================================
Public Sub QUIT()
Dim strResult As String

      bBusy = False
      bCancel = True
      DoEvents
      bCancel = False
      If mvarisConnected = True Then
         If lngSocketHandle = 0 Then Exit Sub
         bCancel = True ' Attempt to stop anything happening
         DoEvents
         strResult = SendReceiveSocket("QUIT")
         RaiseEvent DataArrival(strResult)
         mvarisConnected = False
         Call closesocket(lngSocketHandle)
         lngSocketHandle = 0
      End If
End Sub

' ==============================================================
' Procedure:    Connect
'
' Created on:   29 April 2005    By  Darren Lawrence
'
' Function :-
' '      Connect to a POP3 server and login
' ==============================================================
Public Function Connect(Optional APOPConnect As Boolean = True, Optional ByVal Server As String, Optional ByVal UserName As String, Optional ByVal Password As String, Optional ByVal PopPort As Integer = 110) As Boolean
Dim lngType          As Long
Dim lngProtocol      As Long
Dim lngRetValue      As Long
Dim lngBytesReceived As Long
Dim strResult        As String
Dim strOutPut        As String
Dim strAPOPDigest    As String

      ' start up winsock service
      lngRetValue = WSAStartup(&H201, udtWinsockData)
      If lngRetValue <> 0 Then
         Select Case lngRetValue
            Case WSASYSNOTREADY
               strResult = "The underlying network subsystem is not " & _
               "ready for network communication."
            Case WSAVERNOTSUPPORTED
               strResult = "The version of Windows Sockets API support " & _
               "requested is not provided by this particular " & _
               "Windows Sockets implementation."
            Case WSAEINVAL
               strResult = "The Windows Sockets version specified by the " & _
               "application is not supported by this DLL."
         End Select
100      Err.Raise WSASYSNOTREADY, , strResult
110      Exit Function
      End If
      
120   On Error GoTo Err_Connect:

130   If Server <> "" Then
140      mvarServer = Server
      End If
150   If UserName <> "" Then
160      mvarUserName = UserName
      End If
170   If Password <> "" Then
180      mvarPassword = Password
      End If
190   mvarPopPort = PopPort
200   mvarisConnected = False
      
210   If mvarServer = "" Then
220      mvarLastErrorCode = ERR_InvalidServer
230      Connect = False
240      RaiseEvent POP3Error(GetLastError)
250      Exit Function
      End If
260   If mvarUserName = "" Then
270      mvarLastErrorCode = ERR_InvalidUsername
280      Connect = False
290      RaiseEvent POP3Error(GetLastError)
300      Exit Function
      End If
310   If mvarPassword = "" Then
320      mvarLastErrorCode = ERR_InvalidPassword
330      Connect = False
340      RaiseEvent POP3Error(GetLastError)
350      Exit Function
      End If
      
      ' Create the Socket
360   lngSocketHandle = modWinsockAPI.vbSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
370   If lngSocketHandle = SOCKET_ERROR Then
380      Connect = False
390      mvarLastErrorCode = lngRetValue
410      RaiseEvent POP3Error(GetErrorDescription(Err.LastDllError))
420      Exit Function
      End If
      
430   lngRetValue = modWinsockAPI.vbConnect(lngSocketHandle, mvarServer, mvarPopPort)
440   If lngRetValue = SOCKET_ERROR Then
450      Connect = False
460      mvarLastErrorCode = lngRetValue
480      RaiseEvent POP3Error(GetErrorDescription(Err.LastDllError))
490      Exit Function
      End If
      
      mvarisConnected = True
      
      strResult = SendReceiveSocket() ' Wait for the Greeting from Server
                  
560   If TestForError(strResult) <> ERR_NoError Then
         ' Error occcured
570      Connect = False
580      RaiseEvent POP3Error(GetLastError)
590      Exit Function
      End If
      
      ' APOP Authentication ------------------------------------------------------
      
600   If InStr(1, strResult, "<") > 0 And InStr(1, strResult, ">") > 0 Then
610      mvarAPOPSupported = True
      End If

630   If APOPConnect = True Then

         ' If the connection was requested as APOP then lets
         ' Check if the server supports APOP first and then work out the
         ' response to the digest string issued if there was one.
         ' A server identifies if it supports APOP by issuing a greeting
         ' in the banner with a string enclosed in <>
         ' eg
         ' +OK POP3 Server Ready <16cb.1328.7c17102>
         ' You must add the password to this string and MD5 the result and then
         ' issue the APOP command followed by mailbox name then the MD5'd string
         
640      If mvarAPOPSupported Then

650         strAPOPDigest = Mid$(strResult, InStr(1, strResult, "<"), 9999)
            
            ' Do APOP connection
            
            ' Create the correct response string to the servers digest string
660         strOutPut = "APOP " & mvarUserName & " " & strOutPut & MDFIVE(strAPOPDigest & Password)

            strResult = SendReceiveSocket(strOutPut) ' Send APOP login String
            
780         If TestForError(strResult) <> ERR_NoError Then
               ' Error occcured
790            If mvarLastErrorCode <> POPErrors.ERR_MailBoxLocked Then
800               mvarLastErrorCode = POPErrors.ERR_APOPLoginFailed
               End If
810            Connect = False
820            RaiseEvent POP3Error(GetLastError)
830            Exit Function
            End If
840         mvarisConnected = True
850         Connect = True
860         Exit Function
         Else
            ' APOP is not supported so lets login as normal
870         mvarAPOPSupported = False
         End If
      End If
      
      ' END APOP Authentication ------------------------------------------------------

      ' Connected - Send User
880   strOutPut = "USER " & mvarUserName

      strResult = SendReceiveSocket(strOutPut)

1000  If TestForError(strResult) <> ERR_NoError Then
         ' Error occcured
         ' Test The Error
1010     If mvarLastErrorCode <> POPErrors.ERR_MailBoxLocked Then
1020        mvarLastErrorCode = POPErrors.ERR_InvalidLogon
         End If
1030     Connect = False
1040     RaiseEvent POP3Error(GetLastError)
1050     Exit Function
      End If
      
      ' User ok  - Send Password
1060  strOutPut = "PASS " & mvarPassword

      strResult = SendReceiveSocket(strOutPut)
      
1180  If TestForError(strResult) <> ERR_NoError Then
         ' Error occcured
1190     mvarLastErrorCode = POPErrors.ERR_InvalidLogon
1200     Connect = False
1210     RaiseEvent POP3Error(GetLastError)
1220     Exit Function
      End If
1230  mvarisConnected = True
1240  Connect = True
            
1250  Exit Function

1260 Err_Connect:
1270  Debug.Print Err.Description, Erl
      
1280 Exit_Connect:
End Function

' ==============================================================
' Procedure:    GetLastError
'
' Created on:   29 April 2005    By  Darren Lawrence
'
' Function :-
' '      Used to return a String error of the last reported error state
' ==============================================================
Public Function GetLastError() As String
      Select Case mvarLastErrorCode
         Case POPErrors.ERR_InvalidServer
            GetLastError = "Invalid Server Specified"
         Case POPErrors.ERR_InvalidLogon
            GetLastError = "Invalid Logon - Wrong Username or Password"
         Case POPErrors.ERR_Timeout
            GetLastError = "Timeout waiting for response"
         Case POPErrors.ERR_InvalidUsername
            GetLastError = "Invalid Username supplied"
         Case POPErrors.ERR_InvalidPassword
            GetLastError = "Invalid Password supplied"
         Case POPErrors.ERR_APOPLoginFailed
            GetLastError = "Failed to Login Using APOP"
         Case POPErrors.ERR_FailedConnect
            GetLastError = "Failed to Connect to Server"
         Case POPErrors.ERR_ErrorSTAT
            GetLastError = "Error with STAT Command"
         Case POPErrors.ERR_ErrorRETR
            GetLastError = "Error Retrieving Message"
         Case POPErrors.ERR_ErrorDELE
            GetLastError = "Error Deleting Message"
         Case POPErrors.ERR_ErrorHEADER
            GetLastError = "Error Retrieving Message Header"
         Case POPErrors.ERR_ErrorTOP
            GetLastError = "Error with TOP Command"
         Case POPErrors.ERR_ErrorUIDL
            GetLastError = "Error Fetching UIDL of Messages"
         Case POPErrors.ERR_ErrorNOOP
            GetLastError = "Error with NOOP Command"
         Case POPErrors.ERR_ErrorRSET
            GetLastError = "Error With Reset"
         Case POPErrors.ERR_ErrorLIST
            GetLastError = "Error Fetching LIST of Messages"
         Case POPErrors.ERR_MailBoxLocked
            GetLastError = "MailBox Locked"
         Case POPErrors.ERR_BadParameter
            GetLastError = "Bad Parameter specified"
         Case POPErrors.ERR_InvalidMessageID
            GetLastError = "Message does not Exist"
         Case POPErrors.ERR_SocketCreateError
            GetLastError = "Failed to Create Socket API"
         Case POPErrors.ERR_SocketConnectFail
            GetLastError = "Failed to Connect to Socket API"
      End Select
End Function

' This does a basic tidyup of the recevied data
Private Function TidyReply(ByVal strResult As String) As String
      
      If Left$(strResult, 5) = "+OK" & vbCrLf Then
         strResult = Mid$(strResult, 6, Len(strResult))
      End If
      If Left$(strResult, 4) = "+OK " Then
         strResult = Mid$(strResult, 5, Len(strResult))
      End If
      If Right$(strResult, 5) = vbCrLf & vbCrLf & "." Then
         strResult = Mid$(strResult, 1, Len(strResult) - 5)
      End If
      If Right$(strResult, 3) = vbCrLf & "." Then
         strResult = Mid$(strResult, 1, Len(strResult) - 3)
      End If

      TidyReply = strResult
End Function

' Checks the Message Exists on the Server
' Used for the RETR Method and returns True if OK and the Message Size
' This is required for the progress monitor to work
Private Function CheckMessageID(MessageID As Long, dMessageSize As Double) As Boolean
Dim strResult As String

      CheckMessageID = True
      
      strResult = SendReceiveSocket("LIST " & MessageID)
      
      If TestForError(strResult) <> ERR_NoError Then
         CheckMessageID = False
      End If
      
      dMessageSize = Split(strResult, " ")(2)
      
End Function

Private Function TestForError(ByVal strResult As String) As POPErrors
      ' to check for an Error state and deal with error responses
      
      mvarLastErrorCode = POPErrors.ERR_NoError
      
      If Left$(strResult, 4) = "-ERR" Then
         ' Error occcured
         strResult = Mid$(strResult, 6, Len(strResult)) ' Take off -ERR
         Select Case UCase$(strResult)
            Case "INVALID MESSAGE NUMBER"
               mvarLastErrorCode = POPErrors.ERR_InvalidMessageID
            Case "BAD PARAMETER"
               mvarLastErrorCode = POPErrors.ERR_BadParameter
            Case "MAILBOX LOCKED"
               mvarLastErrorCode = POPErrors.ERR_MailBoxLocked
            Case "TIMEOUT"
               mvarLastErrorCode = POPErrors.ERR_Timeout
            Case Else
               mvarLastErrorCode = POPErrors.ERR_UnKnownError
         End Select
      End If
      TestForError = mvarLastErrorCode
End Function

Private Sub Class_Initialize()
      lngSocketHandle = 0
End Sub

' ==============================================================
' Procedure:    SendReceiveSocket
'
' Created on:   03 July 2005    By  Darren Lawrence
'
' Function :-
' '      This function is used to wait for a response from the POP server
' '      it will wait for a given timeout period before giving up Using the API
' ==============================================================
Private Function SendReceiveSocket(Optional sSendString As String, Optional bWaitForTerm As Boolean = False, Optional bRaiseRecieved As Boolean = True, Optional bTrackProgress As Boolean = False) As String
Dim lngType          As Long
Dim lngProtocol      As Long    ' Protocol being used
Dim lngRetValue      As Long    'Dummy Variable
Dim lngBytesReceived As Long    ' Number of Bytes of data received
Dim strResult        As String
Dim strTempReceived  As String  ' Used to store Received Data Temporarily - Helps speed up
Dim strReceivedData  As String  ' Actual Received Data
Dim bExit            As Boolean ' Exit the receiving
Dim dteTimeDiff      As Date    ' Used for Timeout Counter
Dim lngTotalBytes    As Double  ' Total Bytes recevied Count
Dim dteTimeDiff2

100   On Error GoTo Err_WaitResponse:

110   Do
120      DoEvents
      Loop Until Not bBusy

130   bBusy = True
140   bExit = False
150   strReceivedData = ""

160   If sSendString <> "" Then

         ' Send the data to the Server first
170      RaiseEvent DataSent(sSendString)
180      If modWinsockAPI.vbSendString(lngSocketHandle, sSendString & vbCrLf) = SOCKET_ERROR Then
190         mvarLastErrorCode = lngRetValue
200         RaiseEvent POP3Error(GetErrorDescription(Err.LastDllError))
210         GoTo Exit_WaitResponse
         End If
      End If
      'time
      '  Dim dteTimeDiff2
'        dteTimeDiff2 = DateDiff("s", Now - TimeSerial(0, 0, mvarTimeOut), Now())
220   dteTimeDiff = VBA.DateAdd("s", CDbl(mvarTimeOut), Now())  ' Setup the Time the routine will stop
230   DoEvents
240   Do

         ' Retreive a Chunk of Data
         
250      lngBytesReceived = modWinsockAPI.vbRecv(lngSocketHandle, strResult, mvarTimeOut)
260      If lngBytesReceived = SOCKET_ERROR Then
270         mvarLastErrorCode = POPErrors.ERR_SocketConnectFail
280         RaiseEvent POP3Error(GetErrorDescription(Err.LastDllError))
290         GoTo Exit_WaitResponse
         End If
300      If lngBytesReceived = SOCKET_TIMEOUT Then
310         SendReceiveSocket = "-ERR TimeOut"
320         GoTo Exit_WaitResponse
         End If
         
330      If lngBytesReceived > 0 Then
            ' Data received so lets deal with it and reset the Timout Counter
            
340         If bWaitForTerm Then
               ' We will Exit the Loop only if the Termination sequence of characters are at the end
               ' of the returned data if specified to do so
350            If Right$(strResult, 5) = vbCrLf & "." & vbCrLf Then
360               bExit = True
               End If
            Else
               ' Exit the wait if the received data starts with +OK
370            If Left$(strResult, 3) = "+OK" Then
380               bExit = True
               End If
            End If

            ' This will help speed up processing
390         strTempReceived = strTempReceived & strResult
400         If Len(strTempReceived) > 5000 Then
410            strReceivedData = strReceivedData & strTempReceived
420            strTempReceived = ""
            End If
            
            ' 320         strReceivedData = strReceivedData & strResult

430         dteTimeDiff = VBA.DateAdd("s", mvarTimeOut, Now())
440         If bTrackProgress Then
450            lngTotalBytes = lngTotalBytes + lngBytesReceived
460            RaiseEvent Progress(lngTotalBytes, dblCurrentMessageSize)
470            DoEvents
            End If
         End If

480      If Left$(strResult, 4) = "-ERR" Then
            ' Oh dear an error was sent
490         bExit = True
         End If
500      strResult = ""
         ' DoEvents
      If dteTimeDiff > 0 Then
        dteTimeDiff2 = DateDiff("s", Now - TimeSerial(0, 0, mvarTimeOut), Now())
        'TimeOutVar = mvarTimeOut - dteTimeDiff2
        frmMain.Label1 = Trim(Str(dteTimeDiff2)) + " Secs -- Timeout"
      End If
      Loop Until dteTimeDiff < Now() Or bExit Or bCancel
      
      ' Before we continue make sure the Temp buffer is added to the actual received
      ' data
510   If strTempReceived <> "" Then
520      strReceivedData = strReceivedData & strTempReceived
530      strTempReceived = ""
      End If
            
540   If bCancel Then
550      GoTo Exit_WaitResponse
      End If
      
560   If Right$(strReceivedData, 2) = vbCrLf Then
         ' Remove any excess CRLF from end of recevied data
570      strReceivedData = Left$(strReceivedData, Len(strReceivedData) - 2)
      End If
      
580   SendReceiveSocket = strReceivedData

590   If bRaiseRecieved Then
         ' Let the client know data was received if required
600      RaiseEvent DataArrival(strReceivedData)
      End If
      
610   GoTo Exit_WaitResponse

Err_WaitResponse:
      Debug.Print Err.Description, Erl

Exit_WaitResponse:
      bBusy = False
End Function

Private Sub Class_Terminate()
      Call closesocket(lngSocketHandle)
      Call modWinsockAPI.WSACleanup
End Sub
