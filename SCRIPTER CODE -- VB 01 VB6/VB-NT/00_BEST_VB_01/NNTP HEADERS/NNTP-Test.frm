VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   Caption         =   "XNTP Control Test"
   ClientHeight    =   8385
   ClientLeft      =   165
   ClientTop       =   450
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
   ScaleHeight     =   8385
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4080
      Top             =   1440
   End
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
      ScrollBars      =   2  'Vertical
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1080
      TabIndex        =   8
      Text            =   "alt.test"
      Top             =   5910
      Width           =   6435
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   6360
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   5535
      Width           =   1155
   End
   Begin VB.TextBox txtUsername 
      Height          =   345
      Left            =   4260
      TabIndex        =   4
      Top             =   5535
      Width           =   1155
   End
   Begin VB.TextBox txtServer 
      Height          =   345
      Left            =   1080
      TabIndex        =   2
      Top             =   5535
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
      ScrollBars      =   2  'Vertical
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
   Begin NNTPTEST.NNTP NNTP1 
      Left            =   6795
      Top             =   7635
      _ExtentX        =   741
      _ExtentY        =   741
      Debugger        =   0   'False
   End
   Begin VB.Label DATELab 
      BackColor       =   &H00E0E0E0&
      Caption         =   "DATE"
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
      TabIndex        =   42
      Top             =   7905
      Width           =   6300
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
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   5970
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Password"
      Height          =   195
      Left            =   5520
      TabIndex        =   5
      Top             =   5595
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Username"
      Height          =   195
      Left            =   3420
      TabIndex        =   3
      Top             =   5595
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Newsser&ver"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   5595
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
Public ACE, TT5, NOTYET
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Dim HH()
Dim NG()
Dim LastStart
Dim EasyRem
Dim A1 As String, A2 As String, Text4, Text5, Text6
Dim Buffer1 As String
Dim Buffer2 As String, DONEARTICALORHEADER

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
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Main.Show
    DoEvents
    'loading user-entered data
    Dim file As String, f As Integer, v As Variant
    On Error Resume Next
    Kill App.Path + "\USENET ONE BATCH.txt"
    On Error GoTo 0
    
ReDim HH(30)
ReDim NG(5)
    
    YT = 0
    'YT = YT + 1: HH(YT) = "borg"
    YT = YT + 1: HH(YT) = "borg@gone.com"
    YT = YT + 1: HH(YT) = "matt.lan@btinternet.com"
    YT = YT + 1: HH(YT) = "anon@no.email" 'Kadaitcha Man
    YT = YT + 1: HH(YT) = "gellie618@webtv.net"
    YT = YT + 1: HH(YT) = "Rev. 11D Meow!"
    YT = YT + 1: HH(YT) = "chessucat"
    YT = YT + 1: HH(YT) = "Daniel Urtiz"
    YT = YT + 1: HH(YT) = "Anonymous"
    YT = YT + 1: HH(YT) = "Nomen Nescio"
    YT = YT + 1: HH(YT) = """%"""
    YT = YT + 1: HH(YT) = "jalees@easynet.co.uk"
    YT = YT + 1: HH(YT) = "judy"
    
    
    ReDim Preserve HH(YT)
    
    For R = 1 To UBound(HH)
        HH(R) = LCase(HH(R))
    Next
    
    On Error Resume Next
    Kill "TxTPost --*.TXT"
    On Error GoTo 0
    
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
DONEARTICALORHEADER = False
Do
    Call Sleep(100)
    DoEvents
Loop Until DONEARTICALORHEADER = True
'Loop Until InStr(Text1, "Server Msg=news.bt.com") > 0


NG(1) = "alt.usenet.kooks"
NG(2) = "alt.philosophy"
NG(3) = "24hoursupport.helpdesk"
NG(4) = "alt.support.schizophrenia"
NG(5) = "alt.slack"

'For tt5 = 1 To 5

TT5 = 1

Call GETNEWSGROUPSTART

    
    
    
End Sub

Sub GETNEWSGROUPSTART()


NOTYET = False

On Error Resume Next
Kill App.Path + "\USENET HEADERS -- " + txtNewsgroup + ".txt"
On Error GoTo 0


txtNewsgroup = NG(TT5)
XXString = ""
Label7 = ""

DONEARTICALORHEADER = False
Call cmdSelect_Click
Do
    Call Sleep(100)
    DoEvents
Loop Until DONEARTICALORHEADER = True
'Loop Until InStr(XXString, "ArticleCount") > 0



'NNTP1.LastArticle
    
    
On Error Resume Next
file = App.Path + "\LastArtical -- " + NG(TT5) + ".txt"
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

    TextDebug = ""
    TextDebug2 = ""
    XXString = ""
    DoEvents
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Label5 = Trim(Str(Easy))
    
    
    
    
    DONEARTICALORHEADER = False
    '##
    Call cmdHeader_Item_Click
    Do
        Call Sleep(100)
        DoEvents
    Loop Until DONEARTICALORHEADER = True

    NOTYET = True
    
    Call cmdXover_Click
    
    
    
End Sub


Sub TIMER1_TIMER()
    
If ACE = False Then Exit Sub
    

If TT5 < 5 Then
    TT5 = TT5 + 1: ACE = FASLE: Exit Sub
End If
    
    
    
    
    
    
    
    
    
        
        
    
    
'If tt5 = 5 Then offg = 200
'offg = 250

Label7 = Trim(Str(LastStart - offg))
    






file = App.Path + "\LastArtical -- " + NG(TT5) + ".txt"
FileInUseDelay (file)
f = FreeFile()
Open file For Output As #f
Print #f, PBar1.Value
Close #f






'Simple Copy File
'Set fs = CreateObject("Scripting.FileSystemObject")
'fs.CopyFile App.Path + "\USENET POSTS -- " + txtNewsgroup + " ORIGINAL.txt", App.Path + "\LAST_BATCH\USENET POSTS -- " + txtNewsgroup + " ORIGINAL.txt"


labInfo = "EndE..."
labInfo.BackColor = QBColor(10)
labInfo = "EndE..."
DoEvents

Me.WindowState = 0




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
    
    If NOTYET = FASLE Then Exit Sub
    
    A1 = App.Path + "\USENET HEADERS -- " + txtNewsgroup + ".txt"
    FileInUseDelay (A1)
    f = FreeFile()
    Open A1 For Append As #f
    Print #f, Text + vbCrLf
    Close #f

    'Debug.Print Text
    'TIMEVAL1 = Mid(Text, InStr(Text, "> "))
    'TIMEVAL2 = Mid(TIMEVAL1, 1, InStr(Text, " <"))

    DoEvents

    On Error Resume Next
    PBar1.Max = NNTP1.LastArticle
    PBar1.Min = NNTP1.FirstArticle
    PBar1.Value = Val(Mid(Text, 1, 8))
    Label5 = PBar1.Value
    
    'DATELab = NNTP1.DateTime

    On Error GoTo 0

    
    
    Exit Sub
    
    TextDebug = Text + vbCrLf
    If Text <> "." Then
        TextDebug2 = Text + vbCrLf
        
        If InStr(Text, "Lines:") > 0 Then
            LinesCount = Val(Mid(Text, 7))
        End If
        
        DIY = True
        
        
        ii = 0
        If InStr(Text, "From:") > 0 Then ii = 1
        If InStr(Text, "Date:") > 0 Then ii = 1
        If InStr(Text, "Subject:") > 0 Then ii = 1
        If InStr(Text, "Newsgroups:") > 0 Then ii = 1
        
        If InStr(Text, "NNTP -Posting - Date:") > 0 Then ii = 2: IX = 2
        If InStr(Text, "Message-ID:") > 0 Then ii = 2: IX = 1
        If InStr(Text, "In-Reply-To:") > 0 Then ii = 2: IX = 2
        If InStr(Text, "References:") > 0 Then ii = 2: IX = 2
        
        If InStr(Text, "Newsreader:") > 0 Then ii = 3
        If InStr(Text, "Lines:") > 0 Then ii = 3
        If InStr(Text, "Bytes:") > 0 Then ii = 3
        
        If ii = 1 Then
            Text4 = Text4 + Text + vbCrLf
            DIY = False
        End If
        
        If ii = 2 Then
            If IX = 1 Then Text5 = Text + Text5 + vbCrLf
            If IX = 2 Then Text5 = Text5 + Text + vbCrLf
            DIY = False
        End If
        
        If ii = 3 Then
            Text6 = Text6 + Text + vbCrLf
            DIY = False
        End If
        
        If DIY = True Then Text3 = Text3 + Text + vbCrLf
        
        If InStr(Text, "Xref: ") > 0 Then
            LinesCount2 = 0
            TXR = "--------------------------" + vbCrLf
            
            Text3 = Text3 + TXR + Text5 + TXR + Text6 + TXR + Text4
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
    DONEARTICALORHEADER = True
End Sub

Private Sub NNTP1_DoneDisconnect(Rc As TNNTPRc)
    Msg "Disconnected with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Close #1
    DONEARTICALORHEADER = True
End Sub
Public Sub Msg(String1 As String)
    'XXString = XXString + String1
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
    DONEARTICALORHEADER = True
End Sub

Private Sub NNTP1_DoneGetBody(Rc As TNNTPRc)
    On Error Resume Next
    Msg "Body got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Msg "Speed is " & CLng(NNTP1.Bps)
End Sub

Private Sub NNTP1_DoneGetHeader(Rc As TNNTPRc)
    Msg "Header got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Msg "Posting date was " & NNTP1.DateTime
    DONEARTICALORHEADER = True
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
    
    'ACE = True

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
    DONEARTICALORHEADER = True
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
    
    FileInUseDelay (Source1)
    Open Source1 For Binary Access Read As #1
    FileInUseDelay (Source2)
    Open Source2 For Binary Access Read As #2
   
   
   Buffer1 = Space(LOF(1))
   Get #1, , Buffer1
   
   Buffer2 = Space(LOF(2))
   Get #2, , Buffer2
   
   Close #1, #2
   
    FileInUseDelay (Source2)
    Kill Source2
    FileInUseDelay (Source2)
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








Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function

Sub FileInUseDelay(Tx$)
        
    qq = Now + TimeSerial(0, 1, 30)
    Do
'        DoEvents
        ii = FileInUse(Tx$)
        If ii = True Then Sleep (200)
    Loop Until ii = False Or qq < Now
    
    If ii = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Tx$ + vbCrLf + "Longer than 1 min 30 secs"
        Stop
    End If
End Sub

