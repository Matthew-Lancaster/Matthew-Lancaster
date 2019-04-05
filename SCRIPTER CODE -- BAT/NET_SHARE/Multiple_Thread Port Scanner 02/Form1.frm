VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi Port Scanner"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   5490
   Begin VB.Frame Frame5 
      Caption         =   "IP Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   120
      TabIndex        =   33
      Top             =   3960
      Width           =   5295
      Begin MSComctlLib.ListView ListView1 
         Height          =   2196
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   5052
         _ExtentX        =   8916
         _ExtentY        =   3863
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Timer tmrscan 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5040
      Top             =   120
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "Start Scan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   8916
      Width           =   5295
   End
   Begin VB.Frame Frame4 
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   21
      Top             =   6636
      Width           =   5295
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   255
         Left            =   1080
         TabIndex        =   31
         Top             =   1800
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label17 
         Caption         =   "Progress:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lbltime 
         Caption         =   "00:00:00"
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label15 
         Caption         =   "Rem. Time:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblspeed 
         Caption         =   "0 ip/sec"
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label Label13 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lbldone 
         Caption         =   "0 of 0"
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label11 
         Caption         =   "IP's Done:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblstatus 
         Caption         =   "idle..."
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label9 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   1800
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Speed Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   5295
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Min             =   1
         Max             =   500
         SelStart        =   255
         TickStyle       =   3
         Value           =   255
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   100
         Min             =   1
         Max             =   10000
         SelStart        =   3000
         TickStyle       =   3
         Value           =   3000
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "3000"
         Height          =   255
         Left            =   4680
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Timeout:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "255"
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Threads:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Port List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5295
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   4440
         TabIndex        =   13
         Text            =   "30"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3480
         TabIndex        =   11
         Text            =   "1"
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Range to list"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4200
         TabIndex        =   8
         Text            =   "80"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add to list:"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox lstports 
         Height          =   840
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "to:"
         Height          =   255
         Left            =   4200
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "from:"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "IP Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.Timer tmrtime 
         Interval        =   1000
         Left            =   3960
         Top             =   0
      End
      Begin VB.Timer tmrstat 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4440
         Top             =   0
      End
      Begin VB.TextBox txtstop 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Text            =   "84.144.255.255"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtstart 
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Text            =   "84.144.1.1"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Stop:"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Start:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KILLING_THREADS_UNLOAD
Dim ON_THE_WAY_OUT
Public lstcounter As Long
Public maxxx As Long
Public tick As Long
Public ipdone As Long
'sry for the bad style and the bugs

Private Sub Form_Activate()

reset
Call Set_Start

With ListView1
    .ColumnHeaders.Add , , "IP Address", 1200
    .ColumnHeaders.Add , , "OpenPort", 900
    .ColumnHeaders.Add , , "IP_ResolveName", 3000
    .View = lvwReport
End With

For i = 1 To 255
    Load ws(i) 'load the threads
    DoEvents
    lblstatus.Caption = "Loading Threads... [" & i & "]"
Next i

For i = 1 To 255
    ws(i).Close
    DoEvents 'spare the cpu
    ipdone = Trim(Str(i)): maxxx = Trim(Str(255))
    lblstatus.Caption = "Scanning " & txtstart.Text & " on Port " & lstports.List(lstcounter)
    lbldone.Caption = ipdone & " of " & maxxx & " __ IP Range Found is" & Str(ListView1.ListItems.Count + 1)
    
    ws(i).Connect txtstart.Text, lstports.List(lstcounter)
    DoEvents
    trynext
Next

KILLING_THREADS_UNLOAD = True
For z = 1 To 255
    ws(z).Close
    DoEvents
    Unload ws(z)
    DoEvents
    'lblstatus.Caption = "Killing Threads... [" & z & "]"
Next z

tmrtime.Enabled = False
cmdstart.Enabled = False

End Sub

'----
'Multiple Thread/IP/Port Scanner by GaveDigga (from psc cd)
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=64292&lngWId=1
'----
'Fri 22 September 2017 13:15:40----------



Private Sub Form_Load()



End Sub


Public Function AddOne(IP As String) As String 'add one to ip address
On Error Resume Next
    If ipdone < maxxx Then 'stop the scan
        ipdone = ipdone + 1
        Bar1.Value = Bar1.Value + 1
    End If
    lblstatus.Caption = "Scanning " & txtstart.Text & " on Port " & lstports.List(lstcounter)
    lbldone.Caption = ipdone & " of " & maxxx + " __ IP Found in Ranger is" + Str(ListView1.ListItems.Count)
    
    '__ IP's Done: __ ipdone = "0"

    If Val(ipdone) > maxxx Or txtstart.Text = txtstop.Text Then 'stop the scan
        'Call tmrstat_Timer
        Call Cmdstop_Click
        'Call Script_Ender
        
        AddOne = IP
        Exit Function
    End If


    Dim A As Long, B As Long, C As Long, D As Long
    Dim A1 As Long, B1 As Long, C1 As Long
    A1 = InStr(1, IP, ".")
    A = Mid(IP, 1, A1)
    IP = Mid(IP, A1 + 1, Len(IP) - A1)
    B1 = InStr(1, IP, ".")
    B = Mid(IP, 1, B1)
    IP = Mid(IP, B1 + 1, Len(IP) - B1)
    C1 = InStr(1, IP, ".")
    C = Mid(IP, 1, C1)
    IP = Mid(IP, C1 + 1, Len(IP) - C1)
    D = IP


    If D >= 255 Then

        If C >= 255 Then

            If B >= 255 Then

                If A >= 255 Then
                
                Else
                    A = A + 1
                End If

                B = 0
            Else
                B = B + 1
            End If

            C = 0
        Else
            C = C + 1
        End If

        D = 0
    Else
        D = D + 1
    End If

    AddOne = A & "." & B & "." & C & "." & D
End Function


Private Sub Command1_Click() 'add a new port to the port List
    If Text3.Text = "" Or IsNumeric(Text3.Text) = False Or Val(Text3.Text) > "65535" Or Text3.Text = "0" Then Exit Sub
    lstports.AddItem Text3.Text
    Text3.Text = ""
End Sub

Private Sub Command2_Click() 'range ports to the port list
On Error Resume Next
    If Val(Text4.Text) > "65535" Or Val(Text5.Text) > "65535" Then Exit Sub
    For i = Text4.Text To Text5.Text
    lstports.AddItem i
    Next i
End Sub

Private Sub Cmdstart_Click()
On Error Resume Next
    'Call reset
    Call Set_Start
    
    If lstports.ListCount = "0" Then MsgBox "add some ports to list!": Exit Sub

    cmdstart.Enabled = False
    'reset
    disall

    'calculate the number of ip's to scan
    Dim bIP1() As String
    Dim bIP2() As String
    Dim lResult As Double
    bIP1 = Split(txtstart.Text, ".")
    bIP2 = Split(txtstop.Text, ".")
    lResult = (bIP2(3) - bIP1(3)) + (bIP2(2) - bIP1(2)) * 255 + (bIP2(1) - bIP1(1)) * 255 * 255 + (bIP2(0) - bIP1(0)) * 255 * 255 * 255
    maxxx = lResult
    Bar1.Max = maxxx

Select Case cmdstart.Caption
Case "Start Scan"
    cmdstart.Caption = "Stop Scan"
    For i = 1 To Slider1.Value
        Load ws(i) 'load the threads
        DoEvents
        lblstatus.Caption = "Loading Threads... [" & i & "]"
    Next i
    tmrscan.Enabled = True
    tmrstat.Enabled = True
    tmrtime.Enabled = True
    cmdstart.Enabled = True

Case "Stop Scan"
    On Error Resume Next
    cmdstart.Caption = "Start Scan"
    tmrscan.Enabled = False
    tmrstat.Enabled = False
    tmrtime.Enabled = False
    For z = 1 To Slider1.Value
        ws(z).Close
        DoEvents
        Unload ws(z)
        DoEvents
        lblstatus.Caption = "Killing Threads... [" & z & "]"
    Next z
    enall
    'reset
    cmdstart.Enabled = True
    Call Set_Start
    
End Select
End Sub


Private Sub Cmdstop_Click()

'Case "Stop Scan"
    On Error Resume Next
    cmdstart.Caption = "Start Scan"
    '-----------------------------------------------------------------------------
    '-----------------------------------------------------------------------------
    'CHECK HERE LIKE A FLIPFLOP TO RUN ONCE OR STUCK IN A TIMER LOOP TO EXIT AFTER
    'GET A BUFFER OF CALL FROM THE THREAD CLOSER
    '-----------------------------------------------------------------------------
    'If tmrscan.Enabled = True Then Stop
    If tmrscan.Enabled = False Then
        Exit Sub
    End If
    '-----------------------------------------------------------------------------
    tmrscan.Enabled = False
    tmrstat.Enabled = False
    tmrtime.Enabled = False
    For z = 1 To Slider1.Value
        ws(z).Close
        DoEvents
        Unload ws(z)
        DoEvents
        lblstatus.Caption = "Killing Threads... [" & z & "]"
    Next z
    enall
    'reset
    cmdstart.Enabled = True
    Call Set_Start
'End Select

End Sub


Sub Set_Start()

'ipdone = "0"
txtstart.Text = "192.168.1.0"
txtstop.Text = "192.168.1.255"
Dim Port_Array(10)
Port_Array(1) = "139"
For r = 0 To lstports.ListCount - 1
    If lstports.List(r) = Port_Array(1) Then Port_Array(1) = ""
Next
For r = 1 To UBound(Port_Array)
    If Port_Array(r) <> "" Then lstports.AddItem Port_Array(r)
Next

End Sub


Private Sub Slider1_Change() 'shows the value of the slider1
Label6.Caption = Slider1.Value
End Sub

Private Sub Slider2_Change() 'shows the value of the slider2
Label8.Caption = Slider2.Value
tmrscan.Interval = Slider2.Value
End Sub

Private Sub reset() 'reset the stats
lblstatus.Caption = "idle..."
lbldone.Caption = "0 of 0"
lblspeed.Caption = "0 ip/sec"
lbltime.Caption = "00:00:00"
Bar1.Value = "0"

maxxx = "0"
tick = "0"
lstcounter = "0"
ipdone = "0"
End Sub

Private Sub disall() 'disable buttons, sliders and texts
txtstart.Enabled = False
txtstop.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Slider1.Enabled = False
Slider2.Enabled = False
End Sub

Private Sub enall() 'enable buttons, sliders and texts
txtstart.Enabled = True
txtstop.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Slider1.Enabled = True
Slider2.Enabled = True
End Sub

Private Sub tmrscan_Timer()
On Error Resume Next
For i = 1 To Slider1.Value
    ws(i).Close
    DoEvents 'spare the cpu
    ws(i).Connect txtstart.Text, lstports.List(lstcounter)
    DoEvents
    trynext
    
    'If Val(ipdone) >= maxxx Or txtstart.Text = txtstop.Text Then
    '    Exit For 'stop the scan
    'End If
    
Next i
    
'tmrscan.Enabled

'------------------------------
'AFTER THE LOOP FOR NEXT THREAD
'ENDER IS AROUND HERE
'------------------------------
ON_THE_WAY_OUT = ON_THE_WAY_OUT + 1
'tmrtime.Enabled = False
Debug.Print ON_THE_WAY_OUT

Exit Sub
'
'    If ipdone >= maxxx Or txtstart.Text = txtstop.Text Then 'stop the scan
'        tmrscan.Enabled = False
'        tmrstat.Enabled = False
'        Call Set_Start
'
'        'Call tmrstat_Timer
'        'Call Script_Ender
''        AddOne = IP
''        Exit Function
''        Exit For
'    End If
'
'Next i


End Sub

Sub Script_Ender()

'Called From __ Function AddOne

lblstatus.Caption = "Scanning " & txtstart.Text & " on Port " & lstports.List(lstcounter)
lbldone.Caption = ipdone & " of " & maxxx
lbldone.Caption = ipdone & " of " & maxxx + " __ IP Found in Ranger is" + Str(ListView1.ListItems.Count)

'ipdone = "0"
'__ IP's Done: __ ipdone = "0"



Exit Sub


'If cmdstart.Enabled = True Then
'Exit Sub
'Stop
'End If
'
''Error here editoed On
'If ipdone >= maxxx Or txtstart.Text = txtstop.Text Then 'stop the scan
'
'    On Error Resume Next
'    cmdstart.Caption = "Start Scan"
'    tmrscan.Enabled = False
'    tmrstat.Enabled = False
'    tmrtime.Enabled = False
'    For z = 1 To Slider1.Value
'        ws(z).Close
'        'DoEvents
'        Unload ws(z)
'        DoEvents
'        lblstatus.Caption = "Killing Threads... [" & z & "]"
'        tmrscan.Enabled = False
'        tmrstat.Enabled = False
'        tmrtime.Enabled = False
'        'If ipdone >= maxxx Then Exit For 'stop the scan
'
'    Next z
'    enall
'    'reset
''    Call Set_Start
'    cmdstart.Enabled = True
'End If
End Sub

'Private Sub tmrstat_Timer() 'shows some stats
''lblstatus.Caption = "Scanning " & txtstart.Text & " on Port " & lstports.List(lstcounter)
''lbldone.Caption = ipdone & " of " & maxxx
''__ IP's Done: __ ipdone = "0"
'
'Exit Sub




'Error here edited On
'If Val(ipdone) >= maxxx Or txtstart.Text = txtstop.Text Then 'stop the scan
'
'On Error Resume Next
'    cmdstart.Caption = "Start Scan"
'    tmrscan.Enabled = False
'    tmrstat.Enabled = False
'    tmrtime.Enabled = False
'    For z = 1 To Slider1.Value
'        ws(z).Close
'        'DoEvents
'        Unload ws(z)
'        DoEvents
'        lblstatus.Caption = "Killing Threads... [" & z & "]"
'        If ipdone >= maxxx Then Exit For 'stop the scan
'    Next z
'    enall
'    'reset
'    cmdstart.Enabled = True
'    ListView1.Refresh
'    End If
'End Sub

Private Sub tmrtime_Timer() 'calculate the speed and the rem. time
On Error Resume Next
Dim bps As String
Dim rt As String
tick = tick + 1
bps = ipdone / tick
rt = (maxxx - ipdone) / bps

lblspeed.Caption = Round(bps) & " ip/sec"
lbltime.Caption = FormatSeconds(rt)
End Sub

Private Sub ws_Close(Index As Integer)
If KILLING_THREADS_UNLOAD = False Then
    trynext
End If
End Sub

Private Sub ws_Connect(Index As Integer) 'add the host and the port to the listview

ListView1.ListItems.Add , , ws(Index).RemoteHost
ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ws(Index).RemotePort
'ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ws(Index).LocalHostName
'ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , ws(Index).Name

Dim obj As clsIPResolve
Set obj = New clsIPResolve

ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , obj.AddressToName(ws(Index).RemoteHost)
Set obj = Nothing

trynext

Exit Sub

'Error here edited On
'If ipdone >= maxxx Or txtstart.Text = txtstop.Text Then 'stop the scan
'        tmrscan.Enabled = False
'        tmrstat.Enabled = False
''        Call Set_Start
'
'        'Call tmrstat_Timer
'        'Call Script_Ender
''        AddOne = IP
''        Exit Function
'    End If
End Sub

Private Sub ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
trynext
'Error here edited On

'Exit Sub
'If ipdone >= maxxx Or txtstart.Text = txtstop.Text Then 'stop the scan
'        tmrscan.Enabled = False
'        tmrstat.Enabled = False
''        Call Set_Start
'
'        'Call tmrstat_Timer
'        'Call Script_Ender
''        AddOne = IP
''        Exit Function
'    End If
End Sub

Private Sub trynext() 'goto the next port or add one to the ip address
If lstcounter = lstports.ListCount - 1 Then
    lstcounter = "0"
    txtstart.Text = AddOne(txtstart.Text)
Else
    lstcounter = lstcounter + 1
End If
End Sub

Public Function FormatSeconds(ByVal nSeconds As Long, _
  Optional ByVal sFormat As String = "hh:nn:ss") As String 'format the seconds in the hh:mm:ss format

  sFormat = Replace(sFormat, "mm", "nn")

  FormatSeconds = Format$(DateAdd("s", nSeconds, CDate("00:00:00")), sFormat)
End Function

Public Function Round( _
                ByVal Number As Double, _
                Optional ByVal NumDigitsAfterDecimal As Integer = 0 _
                ) As Double 'round the speed number
    Round = Int(Number * 10 ^ NumDigitsAfterDecimal + 0.5) _
                / 10 ^ NumDigitsAfterDecimal
End Function

