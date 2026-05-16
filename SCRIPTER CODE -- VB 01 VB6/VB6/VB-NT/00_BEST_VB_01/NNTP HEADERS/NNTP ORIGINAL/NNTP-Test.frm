VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XNTP Control Test"
   ClientHeight    =   7845
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9945
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
   ScaleHeight     =   7845
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   330
      Left            =   45
      TabIndex        =   41
      Top             =   5100
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      ForeColor       =   &H80000017&
      Height          =   2115
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2940
      Width           =   9840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&First"
      Height          =   375
      Left            =   5730
      TabIndex        =   23
      Top             =   7080
      Width           =   630
   End
   Begin NNTPTEST.NNTP NNTP1 
      Left            =   5580
      Top             =   6540
      _extentx        =   741
      _extenty        =   741
      debugger        =   0   'False
   End
   Begin VB.CheckBox chkDebug 
      Appearance      =   0  'Flat
      Caption         =   "Debug &mode"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6300
      TabIndex        =   22
      Top             =   6660
      Width           =   1215
   End
   Begin VB.TextBox txtNewsgroup 
      Height          =   345
      Left            =   1080
      TabIndex        =   8
      Text            =   "alt.test"
      Top             =   6120
      Width           =   6435
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   6360
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   5700
      Width           =   1155
   End
   Begin VB.TextBox txtUsername 
      Height          =   345
      Left            =   4260
      TabIndex        =   4
      Top             =   5700
      Width           =   1155
   End
   Begin VB.TextBox txtServer 
      Height          =   345
      Left            =   1080
      TabIndex        =   2
      Top             =   5700
      Width           =   2175
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last"
      Height          =   375
      Left            =   5055
      TabIndex        =   19
      Top             =   7080
      Width           =   630
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   4380
      TabIndex        =   18
      Top             =   7080
      Width           =   630
   End
   Begin VB.CommandButton cmdStat 
      Caption         =   "S&tat"
      Height          =   375
      Left            =   3300
      TabIndex        =   17
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdArticle 
      Caption         =   "&Article"
      Height          =   375
      Left            =   60
      TabIndex        =   14
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdHeader 
      Caption         =   "&Header"
      Height          =   375
      Left            =   1140
      TabIndex        =   15
      Top             =   7080
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
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdBody 
      Caption         =   "&Body"
      Height          =   375
      Left            =   2220
      TabIndex        =   16
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   375
      Left            =   3300
      TabIndex        =   12
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdXover 
      Caption         =   "&Xover"
      Height          =   375
      Left            =   4380
      TabIndex        =   13
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&List"
      Height          =   375
      Left            =   2220
      TabIndex        =   11
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      ForeColor       =   &H80000017&
      Height          =   2835
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   9840
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Disconnect"
      Height          =   375
      Left            =   1140
      TabIndex        =   10
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   60
      TabIndex        =   9
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Rev"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7530
      TabIndex        =   40
      Top             =   7545
      Width           =   1245
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "SZ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7530
      TabIndex        =   39
      Top             =   7260
      Width           =   1245
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "KMan Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7530
      TabIndex        =   38
      Top             =   6690
      Width           =   1245
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7530
      TabIndex        =   37
      Top             =   6405
      Width           =   1245
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "KMan Kook"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7530
      TabIndex        =   36
      Top             =   6975
      Width           =   1245
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   8805
      TabIndex        =   35
      Top             =   7545
      Width           =   1155
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   8805
      TabIndex        =   34
      Top             =   7260
      Width           =   1155
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   8805
      TabIndex        =   33
      Top             =   6690
      Width           =   1155
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   8805
      TabIndex        =   32
      Top             =   6405
      Width           =   1155
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   8805
      TabIndex        =   31
      Top             =   6975
      Width           =   1155
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "To Go To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7530
      TabIndex        =   30
      Top             =   6105
      Width           =   1245
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Now"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7530
      TabIndex        =   29
      Top             =   5820
      Width           =   1245
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tott"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7530
      TabIndex        =   28
      Top             =   5535
      Width           =   1245
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8805
      TabIndex        =   27
      Top             =   6105
      Width           =   1155
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8805
      TabIndex        =   26
      Top             =   5535
      Width           =   1155
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8805
      TabIndex        =   25
      Top             =   5820
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Ne&wsgroup"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   6180
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Password"
      Height          =   195
      Left            =   5520
      TabIndex        =   5
      Top             =   5760
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Username"
      Height          =   195
      Left            =   3420
      TabIndex        =   3
      Top             =   5760
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Newsser&ver"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label labInfo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   60
      TabIndex        =   21
      Top             =   7500
      Width           =   6300
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim Buffer1 As String
Dim Buffer2 As String

Option Compare Text

Public TextDebug, TextDebug2, XXString, Text3, XRef, LinesCount, GODO, LinesCount2



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

Private Sub cmdHeader_Item_Click()
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
    'NNTP1.GotoNext
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
    
    Main.Show
    DoEvents
    'loading user-entered data
    Dim file As String, f As Integer, v As Variant
    
    f = FreeFile()
    Open App.Path + "\TxTPost.txt" For Output As #f
    Print #f, "----------------------------------------------------------------------------------#"
    Print #f, "### Process Batch ---" + Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p")
    Print #f, "----------------------------------------------------------------------------------#"
    Close #f
    
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

txtServer = "news.btinternet.com"
Call cmdConnect_Click
Do
    Call Sleep(100)
    DoEvents
Loop Until InStr(Text1, "Server Msg=news.bt.com") > 0

Dim NG(5)

NG(1) = "alt.philosophy"
NG(2) = "24hoursupport.helpdesk"
NG(3) = "alt.usenet.kooks"
NG(4) = "alt.support.schizophrenia"
NG(5) = "alt.slack"

For tt5 = 1 To 5

txtNewsgroup = NG(tt5)
XXString = ""

Call cmdSelect_Click
Do
    Call Sleep(100)
    DoEvents
Loop Until InStr(XXString, "ArticleCount") > 0



'NNTP1.LastArticle
    
Dim LastStart
Dim EasyRem
    
On Error Resume Next
file = App.Path + "\LastArtical -- " + NG(tt5) + ".txt"
f = FreeFile()
Open file For Input As #f
Line Input #f, v: LastStart = v
Close #f
On Error GoTo 0

EasyRem = NNTP1.LastArticle
Easy = NNTP1.LastArticle
If LastStart = "" Then LastStart = Easy
If LastStart = 0 Then LastStart = Easy

Label6 = Trim(Str(NNTP1.LastArticle))

tt = 0

Do
    TextDebug = ""
    TextDebug2 = ""
    XXString = ""
    DoEvents
    Text2 = ""
    Text3 = ""
    Label5 = Trim(Str(Easy))
    Call cmdHeader_Item_Click
    Easy = Easy - 1
    'If InStr(Text2, "borg@gone.com") > 0 Then Stop

    Do
        Call Sleep(100)
        DoEvents
        If InStr(XXString, "Disconnected") > 0 Or InStr(XXString, "Not connected") > 0 Then
            outhere = True
            XXString = ""
            txtServer = "news.btinternet.com"
            Call cmdConnect_Click
            Do
                Call Sleep(100)
                DoEvents
            If InStr(XXString, "Connected") > 0 Then Exit Do
            Loop Until InStr(XXString, "Server Msg=news.bt.com") > 0 Or InStr(XXString, "Disconnected") > 0

            Call Sleep(500)

            Call cmdSelect_Click
            Do
                Call Sleep(100)
                DoEvents
            Loop Until InStr(XXString, "ArticleCount") > 0
            Call Sleep(500)
            
            EasyRem = NNTP1.LastArticle
            Easy = NNTP1.LastArticle
            
            If LastStart = "" Then LastStart = Easy
            If LastStart = 0 Then LastStart = Easy
            Label6 = Trim(Str(NNTP1.LastArticle))
            'LastStart = EasyRem
        
        End If
        
    If Text2 <> "" And InStrRev(Text2, ".") > Len(Text2) - 4 And InStr(Text2, "Xref: ") = 0 Then Easy = Easy + 1: falseheader = True: Exit Do
    If InStr(TextDebug, NG(tt5)) > 0 Then Exit Do
    Loop Until InStr(TextDebug, ".") > 0 And InStr(Text2, "Xref: ") > 0 Or InStr(Text2, "423 no such article") > 0 Or outhere = True
    outhere = False
    Call Sleep(100)
    
    GODO = 0
    If InStr(Text2, "borg@gone.com") > 0 Then GODO = 1
    If InStr(Text2, "matt.lan@btinternet.com") > 0 Then GODO = 1
    If InStr(Text2, "Kadaitcha Man") > 0 Then GODO = 1
    If InStr(Text2, "gellie618@webtv.net") > 0 Then GODO = 1
    If InStr(Text2, "Rev. 11D Meow!") > 0 Then GODO = 1
    If InStr(Text2, "chessucat") > 0 Then GODO = 1
    If InStr(Text2, "Daniel Urtiz") > 0 Then GODO = 1
    
    If falseheader = True Then falseheader = False: GODO = 0
    
    If GODO = 1 Then
        TextDebug = ""
        Text3 = ""
        Call cmdArticle_Click

        Do
            DoEvents
            Call Sleep(100)
            DoEvents
            
            If InStr(Text3, "412 no group selected") > 0 Then
                Call cmdSelect_Click
                Do
                    Call Sleep(100)
                    DoEvents
                Loop Until InStr(Text1, "ArticleCount") > 0
                Call Sleep(500)
            End If
            
        Loop Until (InStr(TextDebug, ".") > 0 And Text3 <> "" And LinesCount2 >= LinesCount + 3) Or InStr(TextDebug, "time out") > 0
        Call Sleep(100)
    
        tad = 0
        If InStr(XXString, "RC=3 (Busy)") > 0 Then
            Easy = Easy + 1
            tad = 1
        End If
    
        If tad = 0 Then
            Text3 = "## -- " + NG(tt5) + vbCrLf + Text3
            Text3 = Mid(Text3, 1, Len(Text3) - 4)
            ff2 = "## ---------------------------------------------" + vbCrLf + vbCrLf + vbCrLf
            ww = InStr(Text3, ff2)
            If ww > 0 Then
            Text3 = Mid(Text3, 1, (ww + Len(ff2)) - 3) + Mid(Text3, ww + Len(ff2))
            End If
        
            L2(tt5 - 1) = Val(L2(tt5 - 2)) + 1
        
            f = FreeFile()
            Open App.Path + "\TxTPost.txt" For Append As #f
            Print #f, Text3
            Print #f, "-----------------------------------------------------------------------------#" + vbCrLf
            Print #f, "-----------------------------------------------------------------------------#"
            Close #f
        
    
    
    End If
    End If 'tad
    
'If tt5 = 5 Then offg = 200
offg = 250
Label7 = Trim(Str(LastStart - offg))
    
On Error Resume Next
PBar1.Max = LastStart + 1
PBar1.Min = LastStart - offg
PBar1.Max = LastStart + 1
PBar1.Value = (PBar1.Max - Easy) + PBar1.Min '- (PBar1.Max - PBar1.Min)
On Error GoTo 0
    
Loop Until Easy < LastStart - offg





file = App.Path + "\LastArtical -- " + NG(tt5) + ".txt"
f = FreeFile()
Open file For Output As #f
Print #f, EasyRem
Close #f

Dim A1 As String, A2 As String
A1 = App.Path + "\TxTPost.txt"
A2 = App.Path + "\USENET POSTS.txt"

Call JoinFiles(A1, A2)

Next

labInfo = "EndE..."
labInfo.BackColor = QBColor(10)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
'    Call cmdDisconnect_Click
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
End
End Sub

Private Sub NNTP1_Debugger(Text As String)
    'Print received data when in debugging mode
    'Debug.Print Text
    
    TextDebug = Text + vbCrLf
    If Text <> "." Then
        TextDebug2 = Text + vbCrLf
        
        If InStr(Text, "From: ") > 0 Then
            Text3 = Text3 + vbCrLf
        End If
        
        If InStr(Text, "Lines:") > 0 Then
            LinesCount = Val(Mid(Text, 7))
        End If
        
        Text3 = Text3 + Text + vbCrLf
        
        If InStr(Text, "Date: ") > 0 Then
            Text3 = Text3 + vbCrLf
        End If
        
        If InStr(Text, "Xref: ") > 0 Then
            LinesCount2 = 0
            Text3 = Text3 + vbCrLf + "## ---------------------------------------------"
            Text3 = Text3 + vbCrLf + "## USENET - Message------------------"
            Text3 = Text3 + vbCrLf + "## ---------------------------------------------" + vbCrLf
        End If
    End If
    LinesCount2 = LinesCount2 + 1
    
    Text2 = Text2 + Text + vbCrLf
    Text2.SelStart = Len(Text2)

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
    XXString = XXString + String1
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

Private Sub NNTP1_DoneGotoFirst(Rc As TNNTPRc)
    Msg "Goto last article with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
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

Private Sub Text1_Change()
A = A
End Sub

Public Function JoinFiles(Source1 As String, Source2 As String) As Boolean
    
    'Join 1 in front of 2
    
    'On Error GoTo errorhandler
    
    'FileInUseDelay (Source1)
    Open Source1 For Binary Access Read As #1
    'FileInUseDelay (Source2)
   Open Source2 For Binary Access Read As #2
   
   
   Buffer1 = Space(LOF(1))
   Get #1, , Buffer1
   
   Buffer2 = Space(LOF(2))
   Get #2, , Buffer2
   
   Close #1, #2
   
    'FileInUseDelay (Source2)
    Kill Source2
    'FileInUseDelay (Source2)
    Open Source2 For Binary Access Write As #2
    Put #2, , Buffer1
    Put #2, , Buffer2
    Close #2
   
   JoinFiles = True

'Text1 = ""

Exit Function
ErrorHandler:
    MsgBox "Error Join File" + vbCrLf + Source1 + vbCrLf + Source2
    Close #1
    Close #2
    Close #3

End Function

