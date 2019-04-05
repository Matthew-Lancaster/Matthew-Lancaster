VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hyper Scanner"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTimeout 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   30
      Text            =   "64"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtMaxConn 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   28
      Text            =   "32"
      Top             =   3345
      Width           =   495
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   1200
   End
   Begin VB.CommandButton cmdSaveLog 
      Caption         =   "SaveLog"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   26
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdClearLog 
      Caption         =   "Clear Log"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   2880
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox rtfLog 
      Height          =   3495
      Left            =   3600
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1800
      TabIndex        =   23
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtEndPort 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2400
      MaxLength       =   5
      TabIndex        =   18
      Text            =   "2048"
      Top             =   825
      Width           =   975
   End
   Begin VB.TextBox txtStartPort 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   16
      Text            =   "1"
      Top             =   825
      Width           =   975
   End
   Begin VB.TextBox txtEndIP 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   2880
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "254"
      Top             =   465
      Width           =   495
   End
   Begin VB.TextBox txtEndIP 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   12
      Text            =   "123"
      Top             =   465
      Width           =   495
   End
   Begin VB.TextBox txtEndIP 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "168"
      Top             =   465
      Width           =   495
   End
   Begin VB.TextBox txtEndIP 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "192"
      Top             =   465
      Width           =   495
   End
   Begin VB.TextBox txtStartIP 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   2880
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "1"
      Top             =   105
      Width           =   495
   End
   Begin VB.TextBox txtStartIP 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "123"
      Top             =   105
      Width           =   495
   End
   Begin VB.TextBox txtStartIP 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "168"
      Top             =   105
      Width           =   495
   End
   Begin VB.TextBox txtStartIP 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "192"
      Top             =   105
      Width           =   495
   End
   Begin MSWinsockLib.Winsock wskScan 
      Index           =   0
      Left            =   3000
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblTimeout 
      Caption         =   "Timeout (ms):"
      Height          =   255
      Left            =   1800
      TabIndex        =   29
      Top             =   3375
      Width           =   1095
   End
   Begin VB.Label lblMaxConn 
      Caption         =   "Max Pipelines:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Line Line3 
      X1              =   3480
      X2              =   3480
      Y1              =   120
      Y2              =   3600
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3360
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255.255.255.255:65535"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label lblCurrentIPLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Current IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblDash 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   840
      Width           =   135
   End
   Begin VB.Label lblPortRange 
      Alignment       =   1  'Right Justify
      Caption         =   "Port Range:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblEndingIP 
      Alignment       =   1  'Right Justify
      Caption         =   "Ending IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblDot 
      Alignment       =   2  'Center
      Caption         =   "."
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   13
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblDot 
      Alignment       =   2  'Center
      Caption         =   "."
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   11
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblDot 
      Alignment       =   2  'Center
      Caption         =   "."
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   9
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblDot 
      Alignment       =   2  'Center
      Caption         =   "."
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblDot 
      Alignment       =   2  'Center
      Caption         =   "."
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblDot 
      Alignment       =   2  'Center
      Caption         =   "."
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblIPStart 
      Alignment       =   1  'Right Justify
      Caption         =   "Starting IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************
'* -={Hyper Scanner}=- port scanner by Warren Genik        *
'* Created July 31, 2004 - British Colombia, Canada        *
'*                                                         *
'* Description:                                            *
'*  Hyper Scanner is a port scanner designed to be fast    *
'*  and easily configurable.  I've used several other      *
'*  scanners that don't allow configuration of pipelines   *
'*  or timeouts, so I created this.                        *
'* Warning!!!!                                             *
'*  Setting the pipelines too high may lock up your PC and *
'*  possibly annoy your ISP.  I usually use settings of    *
'*  about 64-128 pipelines and 200ms timeout.  That runs   *
'*  on my P4 3.0ghz 1.5gb DDR400 system :->                *
'*                                                         *
'* Anyways, enjoy!                                         *
'*                                            Warren Genik *
'***********************************************************



'***********************************************************
'* Declarations                                            *
'***********************************************************

Option Explicit 'Always use this!
Dim intCurrIP(3) As Integer 'Variables to keep track of current IP
Dim lngCurrPort As Long 'Current port
Dim strCurrIP As String 'String to keep actual IP (x.x.x.x) instead of partial
Dim intMaxConnections As Integer 'Pipeline variable
Dim intTimeout As Integer 'Timeout variable



'***********************************************************
'* Form startup / shutdown                                 *
'***********************************************************

Private Sub Form_Load()
   'Read registry settings
   txtStartIP(0) = GetSetting("Hyper Scanner", "Settings", "StartIP0", 0)
   txtStartIP(1) = GetSetting("Hyper Scanner", "Settings", "StartIP1", 0)
   txtStartIP(2) = GetSetting("Hyper Scanner", "Settings", "StartIP2", 0)
   txtStartIP(3) = GetSetting("Hyper Scanner", "Settings", "StartIP3", 0)
   txtEndIP(0) = GetSetting("Hyper Scanner", "Settings", "EndIP0", 255)
   txtEndIP(1) = GetSetting("Hyper Scanner", "Settings", "EndIP1", 255)
   txtEndIP(2) = GetSetting("Hyper Scanner", "Settings", "EndIP2", 255)
   txtEndIP(3) = GetSetting("Hyper Scanner", "Settings", "EndIP3", 255)
   txtStartPort = GetSetting("Hyper Scanner", "Settings", "StartPort", 1024)
   txtEndPort = GetSetting("Hyper Scanner", "Settings", "EndPort", 10001)
   txtMaxConn = GetSetting("Hyper Scanner", "Settings", "MaxConn", 16)
   txtTimeout = GetSetting("Hyper Scanner", "Settings", "Timeout", 200)
   lblCurrent = "Inactive"
   'Prepare the log
   Log "-={ Hyper Scanner }=- by Warren Genik", 2
   Log "-------------------------------------", 2
   Log Time & " - Activated.", 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Save all the settings to registry
   SaveSetting "Hyper Scanner", "Settings", "StartIP0", txtStartIP(0)
   SaveSetting "Hyper Scanner", "Settings", "StartIP1", txtStartIP(1)
   SaveSetting "Hyper Scanner", "Settings", "StartIP2", txtStartIP(2)
   SaveSetting "Hyper Scanner", "Settings", "StartIP3", txtStartIP(3)
   SaveSetting "Hyper Scanner", "Settings", "EndIP0", txtEndIP(0)
   SaveSetting "Hyper Scanner", "Settings", "EndIP1", txtEndIP(1)
   SaveSetting "Hyper Scanner", "Settings", "EndIP2", txtEndIP(2)
   SaveSetting "Hyper Scanner", "Settings", "EndIP3", txtEndIP(3)
   SaveSetting "Hyper Scanner", "Settings", "StartPort", txtStartPort
   SaveSetting "Hyper Scanner", "Settings", "EndPort", txtEndPort
   SaveSetting "Hyper Scanner", "Settings", "MaxConn", txtMaxConn
   SaveSetting "Hyper Scanner", "Settings", "Timeout", txtTimeout
End Sub



'***********************************************************
'* Scan engine startup / shutdown                          *
'***********************************************************

Private Sub cmdStart_Click()
   Dim i As Integer
   'Read variables
   intMaxConnections = txtMaxConn
   intTimeout = txtTimeout
   'Display some info to the user
   Log Time & " - Inititializing engine...", 1
   Log Time & " - (" & intMaxConnections & " pipelines, " & intTimeout & "ms timeout)", 1
   'Figure out the starting IP (4 parts)
   For i = 0 To 3
      intCurrIP(i) = txtStartIP(i)
      txtStartIP(i).Enabled = False 'Disallow editing while running
      txtEndIP(i).Enabled = False 'Ditto
   Next
   lngCurrPort = txtStartPort 'More variables
   'The actual engine startup...
   For i = 1 To intMaxConnections 'Create one Winsock for every pipeline
      Load wskScan(i) 'Makes the new Winsock (in an array)
      wskScan(i).Tag = 0 'Reset timeout counter
      DoEvents 'Give some idle time to the rest of the system
   Next
   strCurrIP = intCurrIP(0) & "." & intCurrIP(1) & "." & intCurrIP(2) & "." & intCurrIP(3) 'Make a string version of the IP
   'Disable buttons/textboxes while running
   cmdStart.Enabled = False
   cmdStop.Enabled = True
   txtStartPort.Enabled = False
   txtEndPort.Enabled = False
   txtMaxConn.Enabled = False
   txtTimeout.Enabled = False
   Log Time & " - Engine initialized.", 1 'Done initializing
   tmrMain.Enabled = True 'Start the main loop
End Sub

Private Sub cmdStop_Click()
   Dim i As Integer
   'Stop and shutdown scanning engine
   Log Time & " - Stopping engine...", 1
   tmrMain.Enabled = False
   For i = 0 To 3
      'Reenable textboxes
      txtStartIP(i).Enabled = True
      txtEndIP(i).Enabled = True
   Next
   txtEndIP(0).Enabled = False
   txtEndIP(1).Enabled = False
   For i = 1 To intMaxConnections
      wskScan(i).Close 'Close all open ports
      DoEvents
      Unload wskScan(i) 'Destroy pipelines
   Next
   'Reenable settings/buttons
   lblCurrent = "Inactive"
   cmdStart.Enabled = True
   cmdStop.Enabled = False
   txtStartPort.Enabled = True
   txtEndPort.Enabled = True
   txtMaxConn.Enabled = True
   txtTimeout.Enabled = True
   'Engine stopped
   Log Time & " - Engine stopped.", 1
End Sub



'***********************************************************
'* Core scanning code                                      *
'***********************************************************

Private Sub tmrMain_Timer()
   Dim i As Integer
   'Loop through each pipeline
   For i = 0 To intMaxConnections
      If wskScan(i).State = 0 Then 'Check to see if the pipeline is closed (not in use)
         lblCurrent = strCurrIP & ":" & lngCurrPort 'Show what is currently being tested
         wskScan(i).Connect strCurrIP, lngCurrPort 'Start the actual connection
         DoEvents 'Idle time
         lngCurrPort = lngCurrPort + 1 'Select a new port to scan (increase sequentially)
         If lngCurrPort > txtEndPort Then 'Check port range
            lngCurrPort = txtStartPort 'Restart the port #
            intCurrIP(3) = intCurrIP(3) + 1 'Increase the IP Class D (last #)
            If intCurrIP(3) > txtEndIP(3) Then 'Check IP Class D range
               intCurrIP(3) = txtStartIP(3) 'Restart Class D range
               intCurrIP(2) = intCurrIP(2) + 1 'Increase IP Class C (2nd last #)
               If intCurrIP(2) > txtEndIP(2) Then 'Check IP Class C range
                  'Scan is now completed (all ranges started)
                  cmdStop_Click 'Stop the engine
                  Log Time & " - Scan completed.", 2
                  Exit Sub
               End If
            End If
         End If
         strCurrIP = intCurrIP(0) & "." & intCurrIP(1) & "." & intCurrIP(2) & "." & intCurrIP(3) 'Make the new IP string
         wskScan(i).Tag = 0 'Reset the timeout counter
      Else 'This pipeline is in use (trying to connect)
         wskScan(i).Tag = wskScan(i).Tag + 1 'Increase timeout
         If wskScan(i).Tag > intTimeout Then
            'Pipeline timed out (as specified by intTimeout)
            wskScan(i).Tag = 0 'Reset the timeout
            wskScan(i).Close 'Kill the connection
         End If
      End If
   Next
End Sub

Private Sub wskScan_Connect(Index As Integer)
   'Connected successfully to the specified IP/port
   Log Time & " - " & wskScan(Index).RemoteHostIP & ":" & wskScan(Index).RemotePort, 4 'Inform the user
   wskScan(Index).Tag = 0 'Reset timeout
   wskScan(Index).Close 'Close (and free) the pipeline
End Sub

Private Sub wskScan_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   'Some sort of error (possibly timeout)
   wskScan(Index).Tag = 0 'Reset timeout
   wskScan(Index).Close 'Close (and free) the pipeline
End Sub



'***********************************************************
'* Miscellaneous subs                                      *
'***********************************************************

Sub Log(strMsg As String, Optional intColor As Integer)
   'Log messages to the user
   rtfLog.SelStart = Len(rtfLog)
   rtfLog.SelColor = QBColor(intColor) 'In color!
   rtfLog.SelText = strMsg & vbCrLf
End Sub

Private Sub cmdClearLog_Click()
   rtfLog.Text = "" 'Erase the log
End Sub



'***********************************************************
'* Key filtering / verification                            *
'***********************************************************

Private Sub txtEndIP_KeyPress(Index As Integer, KeyAscii As Integer)
   'Filter/verify the number
   Select Case KeyAscii
   Case vbKey0 To vbKey9
   Case vbKeyDelete
   Case vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub txtEndPort_KeyPress(KeyAscii As Integer)
   'Filter/verify the number
   Select Case KeyAscii
   Case vbKey0 To vbKey9
   Case vbKeyDelete
   Case vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub txtMaxConn_KeyPress(KeyAscii As Integer)
   'Filter/verify the number
   Select Case KeyAscii
   Case vbKey0 To vbKey9
   Case vbKeyDelete
   Case vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub txtStartIP_KeyPress(Index As Integer, KeyAscii As Integer)
   'Filter/verify the number
   Select Case KeyAscii
   Case vbKey0 To vbKey9
   Case vbKeyDelete
   Case vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub txtStartPort_KeyPress(KeyAscii As Integer)
   'Filter/verify the number
   Select Case KeyAscii
   Case vbKey0 To vbKey9
   Case vbKeyDelete
   Case vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub txtTimeout_KeyPress(KeyAscii As Integer)
   'Filter/verify the number
   Select Case KeyAscii
   Case vbKey0 To vbKey9
   Case vbKeyDelete
   Case vbKeyBack
   Case Else
      KeyAscii = 0
   End Select
End Sub

Private Sub txtEndIP_Change(Index As Integer)
   'Filter/verify the number
   If txtEndIP(Index) = "" Then Exit Sub
   If txtEndIP(Index) > 255 Then txtEndIP(Index) = 255
End Sub

Private Sub txtStartIP_Change(Index As Integer)
   'Filter/verify the number
   If txtStartIP(Index) = "" Then Exit Sub
   If txtStartIP(Index) > 255 Then txtStartIP(Index) = 255
   txtEndIP(Index) = txtStartIP(Index)
End Sub
