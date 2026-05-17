VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "XNTP Control Test"
   ClientHeight    =   5295
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NNTP-Test.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7620
   StartUpPosition =   2  'Bildschirmmitte
   Begin NNTPTEST.NNTP NNTP1 
      Left            =   5580
      Top             =   3900
      _ExtentX        =   741
      _ExtentY        =   741
      Debugger        =   0   'False
   End
   Begin VB.CheckBox chkDebug 
      Appearance      =   0  '2D
      Caption         =   "Debug &mode"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6300
      TabIndex        =   22
      Top             =   4020
      Width           =   1215
   End
   Begin VB.TextBox txtNewsgroup 
      Height          =   345
      Left            =   1080
      TabIndex        =   8
      Text            =   "alt.test"
      Top             =   3480
      Width           =   6435
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   6360
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3060
      Width           =   1155
   End
   Begin VB.TextBox txtUsername 
      Height          =   345
      Left            =   4260
      TabIndex        =   4
      Top             =   3060
      Width           =   1155
   End
   Begin VB.TextBox txtServer 
      Height          =   345
      Left            =   1080
      TabIndex        =   2
      Top             =   3060
      Width           =   2175
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last"
      Height          =   375
      Left            =   5460
      TabIndex        =   19
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   4380
      TabIndex        =   18
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdStat 
      Caption         =   "S&tat"
      Height          =   375
      Left            =   3300
      TabIndex        =   17
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdArticle 
      Caption         =   "&Article"
      Height          =   375
      Left            =   60
      TabIndex        =   14
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdHeader 
      Caption         =   "&Header"
      Height          =   375
      Left            =   1140
      TabIndex        =   15
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6540
      TabIndex        =   20
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdBody 
      Caption         =   "&Body"
      Height          =   375
      Left            =   2220
      TabIndex        =   16
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   375
      Left            =   3300
      TabIndex        =   12
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdXover 
      Caption         =   "&Xover"
      Height          =   375
      Left            =   4380
      TabIndex        =   13
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&List"
      Height          =   375
      Left            =   2220
      TabIndex        =   11
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      ForeColor       =   &H80000017&
      Height          =   2835
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   7455
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Disconnect"
      Height          =   375
      Left            =   1140
      TabIndex        =   10
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   60
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Ne&wsgroup"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3540
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Password"
      Height          =   195
      Left            =   5520
      TabIndex        =   5
      Top             =   3120
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Username"
      Height          =   195
      Left            =   3420
      TabIndex        =   3
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Newsser&ver"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label labInfo 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   60
      TabIndex        =   21
      Top             =   4920
      Width           =   7455
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

Private Sub chkDebug_Click()
    'Toggle the debugger mode
    NNTP1.Debugger = chkDebug.Value = 1
End Sub

Private Sub cmdArticle_Click()
    'Get the current article (header and body)
    Msg "Getting article..."
    NNTP1.GetArticle
End Sub

Private Sub cmdBody_Click()
    'get the current article (header only)
    Msg "Getting body..."
    NNTP1.GetBody
End Sub

Private Sub cmdConnect_Click()
    'Connecting to a newsserver
    Msg "Connecting..."
    NNTP1.Host = txtServer
    NNTP1.Username = txtUsername
    NNTP1.Password = txtPassword
    NNTP1.Connect
End Sub


Private Sub cmdDisconnect_Click()
    'Disconnecting from a newsserver
    Msg "Disconnecting..."
    NNTP1.Disconnect
End Sub

Private Sub cmdEnd_Click()
    'Quit the test program
    Unload Me
    End
End Sub

Private Sub cmdHeader_Click()
    'get the current article header
    Msg "Getting header..."
    NNTP1.GetHeader
End Sub

Private Sub cmdLast_Click()
    'goto the previous article in the selected group
    Msg "Seeking last article..."
    NNTP1.GotoLast
End Sub

Private Sub cmdList_Click()
    'get a list of newsgroups
    Msg "Receiving list of newsgroups..."
    NNTP1.DateTime = Now() - 92
    NNTP1.GetList
End Sub

Private Sub cmdNext_Click()
    'goto the next article in the selected group
    Msg "Seeking next article..."
    NNTP1.GotoNext
End Sub

Private Sub cmdPost_Click()

End Sub

Private Sub cmdSelect_Click()
    'Selecting a newsgroup
    Msg "Selecting a newsgroup..."
    NNTP1.Newsgroup = txtNewsgroup
    NNTP1.SelectGroup
End Sub

Private Sub cmdStat_Click()
    'get the current article state
    Msg "Getting statistic..."
    NNTP1.GetStat
End Sub

Private Sub cmdXover_Click()
    'Retrieving all available article headers
    Msg "Getting article headers..."
    NNTP1.GetXover
End Sub


Private Sub Form_Load()
    'loading user-entered data
    Dim file As String, f As Integer, v As Variant
    file = App.Path + "\setup.ini"
    If Dir(file, vbHidden And Not vbDirectory) <> "" Then
        f = FreeFile()
        Open file For Input As #f
        Line Input #f, v: txtServer = v
        Line Input #f, v: txtUsername = v
        Line Input #f, v: txtPassword = v
        Line Input #f, v: txtNewsgroup = v
        Input #f, v: chkDebug = v
        Close #f
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'save user-entered data
    Dim file As String, f As Integer
    file = App.Path + "\setup.ini"
    f = FreeFile()
    Open file For Output As #f
    Print #f, txtServer
    Print #f, txtUsername
    Print #f, txtPassword
    Print #f, txtNewsgroup
    Print #f, chkDebug
    Close #f
End Sub

Private Sub NNTP1_Debugger(Text As String)
    'Print received data when in debugging mode
    Debug.Print Text
End Sub

Private Sub NNTP1_DoneConnect(Rc As TNNTPRc)
    Msg "Connected with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Msg "Server Msg=" + NNTP1.Hostinfo
End Sub

Private Sub NNTP1_DoneDisconnect(Rc As TNNTPRc)
    Msg "Disconnected with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Close #1
End Sub

Public Sub Msg(String1 As String)
    With Text1
        If Len(.Text) > 16384 Then
            .Text = Mid(.Text, 2000)
        End If
        If .Text = "" Then
            .Text = String1 + vbCrLf
        Else
            .Text = .Text + Right(String1, 16384) + vbCrLf
        End If
        .SelStart = Len(.Text)
    End With
End Sub

Private Sub NNTP1_DoneGetArticle(Rc As TNNTPRc)
    Msg "Article got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
End Sub

Private Sub NNTP1_DoneGetBody(Rc As TNNTPRc)
    On Error Resume Next
    Msg "Body got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Msg "Speed is " & CLng(NNTP1.Bps)
End Sub

Private Sub NNTP1_DoneGetHeader(Rc As TNNTPRc)
    Msg "Header got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Msg "Posting date was " & NNTP1.DateTime
End Sub

Private Sub NNTP1_DoneGetList(Rc As TNNTPRc)
    Msg "List of newsgroups received with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    labInfo.Caption = CStr(NNTP1.DataCount) + " newsgroups received"
End Sub

Private Sub NNTP1_DoneGetStat(Rc As TNNTPRc)
    Msg "Stat got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
End Sub

Private Sub NNTP1_DoneGetXover(Rc As TNNTPRc)
    Msg "Article headers got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    labInfo = CStr(NNTP1.DataCount) + " headers received"
End Sub

Private Sub NNTP1_DoneGotoLast(Rc As TNNTPRc)
    Msg "Goto last article with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
End Sub

Private Sub NNTP1_DoneGotoNext(Rc As TNNTPRc)
    Msg "Goto next article with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
End Sub

Private Sub NNTP1_DoneList(Rc As TNNTPRc)
    Msg "List got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Msg NNTP1.DataCount & " newsgroups received"
End Sub

Private Sub NNTP1_DoneSelectGroup(Rc As TNNTPRc)
    Msg "Newsgroup selected with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Msg "FirstArticle=" & NNTP1.FirstArticle
    Msg "LastArticle=" & NNTP1.LastArticle
    Msg "ArticleCount=" & NNTP1.ArticleCount
End Sub

Private Sub NNTP1_DoneXover(Rc As TNNTPRc)
    Msg "Xover done with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    NNTP1.DataIndex = 1
    Msg "FirstHeader=" + NNTP1.Subject
    NNTP1.DataIndex = NNTP1.DataCount
    Msg "LastHeader=" + NNTP1.Subject
End Sub

Private Sub NNTP1_DoneSend(Rc As TNNTPRc)
    Msg NNTP1.Text
End Sub

Private Sub NNTP1_Receive(Count As Long)
    labInfo.Caption = CStr(Count) & " lines"
End Sub

