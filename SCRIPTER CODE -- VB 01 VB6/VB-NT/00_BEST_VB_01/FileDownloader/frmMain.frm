VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Downloader Test Application"
   ClientHeight    =   7815
   ClientLeft      =   150
   ClientTop       =   465
   ClientWidth     =   6855
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   " Resume "
      Height          =   1260
      Left            =   240
      TabIndex        =   32
      Top             =   3600
      Width           =   2895
      Begin VB.TextBox txtMaxBytes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   36
         Text            =   "-1"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtResumeFrom 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   34
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "(negative number = all bytes)"
         Height          =   195
         Left            =   720
         TabIndex        =   38
         Top             =   960
         Width           =   2025
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bytes"
         Height          =   195
         Left            =   2400
         TabIndex        =   37
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Read maximum"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resume from byte"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   270
         Width           =   1275
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " HTTP Authentication "
      Height          =   1020
      Left            =   3360
      TabIndex        =   27
      Top             =   3600
      Width           =   3255
      Begin VB.TextBox txtHttpPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   31
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtHttpUser 
         Height          =   285
         Left            =   600
         TabIndex        =   30
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pass:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   270
         Width           =   375
      End
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   240
      TabIndex        =   25
      Text            =   "http://download.microsoft.com/download/.netframesdk/Redist/1.0/W98NT42KMeXP/EN-US/dotnetredist.exe"
      Top             =   180
      Width           =   6345
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   24
      Top             =   5280
      Width           =   6375
   End
   Begin VB.Frame Frame1 
      Caption         =   " Proxy Server "
      Height          =   2865
      Left            =   3720
      TabIndex        =   12
      Top             =   600
      Width           =   2895
      Begin VB.ComboBox cmbProxyServer 
         Height          =   315
         ItemData        =   "frmMain.frx":1CFA
         Left            =   600
         List            =   "frmMain.frx":1D0A
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Text            =   "Administrator"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   16
         Text            =   "MyPassword"
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtProxy 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   15
         Text            =   "10.0.0.5"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtProxyPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Text            =   "1080"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkRemoteDNS 
         Caption         =   "Use remote DNS"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   270
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proxy:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   750
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   195
         Left            =   1920
         TabIndex        =   19
         Top             =   750
         Width           =   330
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Download Type "
      Height          =   1815
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   3255
      Begin VB.OptionButton optDlType 
         Caption         =   "Download to File"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   480
         ScaleHeight     =   45
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   7
         Top             =   480
         Width           =   2655
         Begin VB.OptionButton optDlToFile 
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   10
            Top             =   15
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.TextBox txtFilename 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   9
            Text            =   "M:\01 Installations\Installation programs\#00 New Install Progs\01 Downloads 2009"
            Top             =   0
            Width           =   2415
         End
         Begin VB.OptionButton optDlToFile 
            Caption         =   "Use temporary file"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   8
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.OptionButton optDlType 
         Caption         =   "Use a download buffer"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
      End
      Begin VB.OptionButton optDlType 
         Caption         =   "Use a download stream"
         Height          =   255
         Index           =   2
         Left            =   225
         TabIndex        =   5
         Top             =   1440
         Value           =   -1  'True
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Progress "
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   3255
      Begin VB.PictureBox picProgress 
         AutoRedraw      =   -1  'True
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   100
         TabIndex        =   2
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblDownloadStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Downloaded x bytes from y bytes"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdFetch 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fetch URL"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the URL of the file to download:"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   -15
      Width           =   2670
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "File"
      Begin VB.Menu Mnu_DownLoad_VB 
         Caption         =   "DownLoad To VB"
      End
      Begin VB.Menu Mnu_Installs 
         Caption         =   "Download to Installs"
      End
      Begin VB.Menu Mnu_HomeMove 
         Caption         =   "HomeMove"
      End
   End
   Begin VB.Menu Mnu_Reset 
      Caption         =   "Reset"
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_OPEN_DOWNLOAD_FILE 
      Caption         =   "OPEN DOWNLOAD FILE"
   End
   Begin VB.Menu MNU_FOLDER 
      Caption         =   "OPEN FOLDER"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents RemoteFile As DataConnection, BasePath
Attribute RemoteFile.VB_VarHelpID = -1
Dim DlType As DOWNLOAD_TYPE
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim STACKFETCH

Private Sub Form_Load()
    'Initialize WinSock... (this*must* be done)
    StartWinsock vbNullString
    'Create a new DataConnection class
    Set RemoteFile = New DataConnection
    'Set default proxy to 'No Proxy'
    cmbProxyServer.ListIndex = 0
    'Set default download type to Stream Buffer
    optDlType_Click 0
    optDlType(0) = vbChecked
    
    txtURL = "http://download.microsoft.com/download/.netframesdk/Redist/1.0/W98NT42KMeXP/EN-US/dotnetredist.exe"
    If Clipboard.GetText <> "" Then txtURL = Clipboard.GetText
    
        
    BasePath = "D:\VB6\VB-NT\00_Best_VB_01\"
    BasePath = "M:\01 Installations\Installation programs\#00 New Install Progs\01 Downloads 2009\"
    Call txtURL_Change
    Me.Show
    DoEvents


    Call Mnu_HomeMove_BRHhomemove_Click
'    Call Mnu_HomeMove_EBhomemove_Click
'    Call Mnu_HomeMove_CHIhomemove_Click
'    Call Mnu_HomeMove_Fivehomemove_Click
'    Call Mnu_HomeMove_MShomemove_Click



'
Exit Sub
    Comm$ = Command$
    Comm$ = "http://www.rightmove.co.uk/property-to-rent/property-32019662.html"
    If Comm$ <> "" Then
        BasePath = "E:\01 Favs\#00 Agents\WORK 2011-04-11\AGENT URL DOWNLOAD\"
        txtURL = Comm$
        Call txtURL_Change
        Call cmdFetch_Click
        End
    End If

End Sub



Private Sub Mnu_HomeMove_BRHhomemove_Click()
'SEACH FOR THE VAR
'STACKFETCH


txtURL = "http://www.homemove.org.uk/uploads/BRHhomemove.pdf"
BasePath = "D:\# MY DOCS\# 01 My Documents\03 PDF Files\HomeMove\Brighton & Hove City Council\"

uu1 = DateValue("03-09-10") - 28
uu1 = DateValue("10-09-10") - 28
uu2 = (Now) - 1
Do
uu1 = uu1 + 14
Loop Until uu1 - 1 > uu2
uu1 = uu1 - 14


txtFilename = BasePath + "BRHhomemove " + Format(uu1, "YYYY-MM-DD") + ".pdf"

On Error Resume Next
Kill txtFilename

Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.Getfile(txtFilename)
F.Delete True
'txtResumeFrom.Text = Trim(Str(F.Size))
txtResumeFrom.Text = Trim(Str(0))



If Dir(txtFilename) <> "" Then
'    Call Mnu_HomeMove_EBhomemove_Click
End If

If Dir(txtFilename) = "" Then
'    Call cmdFetch_Click
'
'
'    Do
'        Sleep 500
'        DoEvents
'    Loop Until RemoteFile.DownloadLength <> -1
'
'    Do
'        DoEvents
'    Loop Until RemoteFile.DownloadLength = lByteCount
    
End If

End Sub



Private Sub Mnu_HomeMove_EBhomemove_Click()
txtURL = "http://www.homemove.org.uk/uploads/EBhomemove.pdf"
BasePath = "D:\# MY DOCS\# 01 My Documents\03 PDF Files\HomeMove\Eastbourne Borough Council\"

uu1 = DateValue("10-09-10") - 28
uu2 = (Now) - 1
Do
uu1 = uu1 + 14
Loop Until uu1 - 1 > uu2
uu1 = uu1 - 14


txtFilename = BasePath + "EBhomemove " + Format(uu1, "YYYY-MM-DD") + ".pdf"
On Error Resume Next
Kill txtFilename
Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.Getfile(txtFilename)
F.Delete True
'txtResumeFrom.Text = Trim(Str(F.Size))
txtResumeFrom.Text = Trim(Str(0))

If Dir(txtFilename) <> "" Then
'    Call Mnu_HomeMove_CHIhomemove_Click
End If

If Dir(txtFilename) = "" Then
'    Call cmdFetch_Click
'
'
'    Do
'        DoEvents
'    Loop Until RemoteFile.DownloadLength <> -1
'
'    Do
'        DoEvents
'    Loop Until RemoteFile.DownloadLength = lByteCount
End If
End Sub


Private Sub Mnu_HomeMove_CHIhomemove_Click()
txtURL = "http://www.homemove.org.uk/uploads/CHIhomemove.pdf"
BasePath = "D:\# MY DOCS\# 01 My Documents\03 PDF Files\HomeMove\Chichester, Lewes, Mid Sussex, Rother District Council\"

uu1 = DateValue("10-09-10") - 28
uu2 = (Now) - 1
Do
uu1 = uu1 + 14
Loop Until uu1 - 1 > uu2
uu1 = uu1 - 14


txtFilename = BasePath + "CHIhomemove " + Format(uu1, "YYYY-MM-DD") + ".pdf"
On Error Resume Next
Kill txtFilename
Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.Getfile(txtFilename)
F.Delete True
'txtResumeFrom.Text = Trim(Str(F.Size))
txtResumeFrom.Text = Trim(Str(0))

If Dir(txtFilename) <> "" Then
'    Call Mnu_HomeMove_Fivehomemove_Click
End If

If Dir(txtFilename) = "" Then
'    Call cmdFetch_Click
'
'    Do
'        DoEvents
'    Loop Until RemoteFile.DownloadLength <> -1
'
'    Do
'        DoEvents
'    Loop Until RemoteFile.DownloadLength = lByteCount
End If

End Sub

Private Sub Mnu_HomeMove_Fivehomemove_Click()
txtURL = "http://www.homemove.org.uk/uploads/Fivehomemove.pdf"
BasePath = "D:\# MY DOCS\# 01 My Documents\03 PDF Files\HomeMove\Adur, Arun, Chichester, Lewes, Mid Sussex, Rother, Wealden District, Hastings, Worthing Borough Council\"

uu1 = DateValue("10-09-10") - 28
uu2 = (Now) - 1
Do
uu1 = uu1 + 14
Loop Until uu1 - 1 > uu2
uu1 = uu1 - 14


txtFilename = BasePath + "Fivehomemove " + Format(uu1, "YYYY-MM-DD") + ".pdf"
On Error Resume Next
Kill txtFilename
Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.Getfile(txtFilename)
F.Delete True
'txtResumeFrom.Text = Trim(Str(F.Size))
txtResumeFrom.Text = Trim(Str(0))

If Dir(txtFilename) <> "" Then
'    Call Mnu_HomeMove_MShomemove_Click
End If

If Dir(txtFilename) = "" Then
'    Call cmdFetch_Click
'
'    Do
'        DoEvents
'    Loop Until RemoteFile.DownloadLength <> -1
'
'    Do
'        DoEvents
'    Loop Until RemoteFile.DownloadLength = lByteCount
End If

End Sub


Private Sub Mnu_HomeMove_MShomemove_Click()
txtURL = "http://www.homemove.org.uk/uploads/MShomemove.pdf"
BasePath = "D:\# MY DOCS\# 01 My Documents\03 PDF Files\HomeMove\Mid Sussex\"

uu1 = DateValue("10-09-10") - 28
uu2 = (Now) - 1
Do
uu1 = uu1 + 14
Loop Until uu1 - 1 > uu2
uu1 = uu1 - 14


txtFilename = BasePath + "MShomemove " + Format(uu1, "YYYY-MM-DD") + ".pdf"
On Error Resume Next
Kill txtFilename
Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.Getfile(txtFilename)
F.Delete True
'txtResumeFrom.Text = Trim(Str(F.Size))
txtResumeFrom.Text = Trim(Str(0))

If Dir(txtFilename) <> "" Then
'
End If

If Dir(txtFilename) = "" Then
'    Call cmdFetch_Click
'
'
'    Do
'        DoEvents
'    Loop Until RemoteFile.DownloadLength <> -1
'
'    Do
'        DoEvents
'    Loop Until RemoteFile.DownloadLength = lByteCount
End If
'cmdFetch.BackColor = QBColor(14)
'cmdFetch.Caption = cmdFetch.Caption + " -- DONE"
End Sub




Private Sub Form_Unload(Cancel As Integer)
    'Clean up...
    Set RemoteFile = Nothing
    EndWinsock
End Sub

Private Sub Mnu_DownLoad_VB_Click()

BasePath = "D:\VB6\VB-NT\00_Best_VB_01\"
Call txtURL_Change

End Sub


Private Sub MNU_FOLDER_Click()

Shell "EXPLORER.EXE /select," + """D:\# MY DOCS\# 01 My Documents\03 PDF Files\HomeMove""", vbNormalFocus

End Sub

Private Sub Mnu_Installs_Click()

BasePath = "M:\01 Installations\Installation programs\#00 New Install Progs\01 Downloads 2009\"
Call txtURL_Change
End Sub
Private Sub MNU_OPEN_DOWNLOAD_FILE_Click()
'Shell "EXPLORER.EXE " + txtFilename, vbNormalFocus
Shell "EXPLORER.EXE /select," + """D:\# MY DOCS\# 01 My Documents\03 PDF Files\HomeMove""", vbNormalFocus
End Sub

Private Sub Mnu_Reset_Click()
Unload frmMain
frmMain.Show
txtURL = ""
BasePath = ""
txtFilename = ""

End Sub

Private Sub MNU_VB_Click()
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
End
End Sub

Private Sub optDlType_Click(Index As Integer)
    optDlToFile(0).Enabled = (Index = 0)
    optDlToFile(1).Enabled = (Index = 0)
    txtFilename.Enabled = (Index = 0)
    DlType = Index
End Sub
Private Sub cmbProxyServer_Click()
    chkRemoteDNS.Enabled = (cmbProxyServer.ListIndex = 2 Or cmbProxyServer.ListIndex = 3)
    txtPassword.Enabled = (cmbProxyServer.ListIndex = 3)
    txtUsername.Enabled = chkRemoteDNS.Enabled
    txtProxy.Enabled = (cmbProxyServer.ListIndex <> 0)
    txtProxyPort.Enabled = (cmbProxyServer.ListIndex <> 0)
End Sub
Private Sub cmdFetch_Click()
    
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set F = fs.Getfile(txtFilename)
    txtResumeFrom.Text = Trim(Str(F.Size))
    On Error GoTo 0
    
    
    txtOutput.Text = ""
    With RemoteFile
        .DownloadType = DlType
        .Filename = txtFilename.Text
        .UseTempFile = optDlToFile(1).Value
        .ProxyType = cmbProxyServer.ListIndex
        .ProxyHostname = txtProxy.Text
        .ProxyPort = Val(txtProxyPort.Text)
        .ProxyUsername = txtUsername.Text
        .ProxyPassword = txtPassword.Text
        .ProxyUseRemoteDNS = (chkRemoteDNS.Value = vbChecked)
        .AutoRedirect = True
        .AllowCache = False
        .UseResume = False
        picProgress.Cls
        .Disconnect
        .HttpUser = txtHttpUser.Text
        .HttpPass = txtHttpPass.Text
        .UseHttpAuthorization = True
        .ResumeFrom = Val(txtResumeFrom.Text)
        .UseResume = True
        .PacketSize = 10000
        .FetchURLString txtURL.Text
        .MaxDownload = Val(txtMaxBytes)
    End With



End Sub
Private Sub RemoteFile_BytesReceived(lByteCount As Long, ID As Long)
    'If the script knows how many bytes that it has to receive then
    If RemoteFile.DownloadLength > 0 Then
        '... draw a progress bar
        lblDownloadStatus.Caption = "Downloaded " + CStr(lByteCount) + " bytes from " + CStr(RemoteFile.DownloadLength) + "."
        picProgress.ScaleWidth = RemoteFile.DownloadLength
        picProgress.Line (0, 0)-(lByteCount, 1), , BF
    'Or else, simply show how many bytes we have received
    Else
        lblDownloadStatus.Caption = "Downloaded " + CStr(lByteCount) + " bytes."
    End If
End Sub
Private Sub RemoteFile_Connected(ID As Long)
    'Successfully connected to the remote host
    Debug.Print "Connected"
End Sub
Private Sub RemoteFile_Disconnected(ID As Long)
    Debug.Print "Disconnected"
    'If we have downloaded everything to a buffer...
    If RemoteFile.DownloadType = dtToBuffer Then
        '... then show it
        txtOutput.Text = RemoteFile.GetBufferAsString
        RemoteFile.ClearBuffer
    End If
    'Tell the user the download is complete
    If RemoteFile.DownloadType = dtToFile And RemoteFile.UseTempFile Then
        MsgBox "Download completed!" + vbCrLf + "Result saved to " + RemoteFile.Filename, vbInformation
    Else
        MsgBox "Download completed!", vbInformation
        
        
        
'        Call Mnu_HomeMove_BRHhomemove_Click
        If STACKFETCH = 0 Then Call Mnu_HomeMove_EBhomemove_Click
        If STACKFETCH = 1 Then Call Mnu_HomeMove_CHIhomemove_Click
        If STACKFETCH = 2 Then Call Mnu_HomeMove_Fivehomemove_Click
        'If STACKFETCH = 3 Then Call Mnu_HomeMove_MShomemove_Click
        If STACKFETCH = 3 Then
            cmdFetch.BackColor = QBColor(14)
            cmdFetch.Caption = cmdFetch.Caption + " -- DONE"
            MsgBox "ALL COMPLETE - END:"
        End If
        STACKFETCH = STACKFETCH + 1
        
    End If

    Set fs = CreateObject("Scripting.FileSystemObject")
    If InStr(txtFilename, "BRHhomemove") > 0 Then
        On Error Resume Next
        fs.copyfile txtFilename, BasePath + "BRHhomemove Always.pdf"
        
    End If



End Sub
Private Sub RemoteFile_DownloadFailed(Why As DOWNLOAD_FAILURE, ID As Long)
    'Uhoh... the download failed :(
    Debug.Print "Download Failed"
    MsgBox "The download failed... Error code " + CStr(Why), vbExclamation
    MsgBox RemoteFile.HTTPReply
End Sub
Private Sub RemoteFile_ObjectMoved(sNewUrl As String, ID As Long)
    Debug.Print "Object moved to " & sNewUrl
End Sub
Private Sub RemoteFile_StreamBytes(lByteCount As Long, bBytes() As Byte, ID As Long)
    'Show the received bytes
    txtOutput.Text = txtOutput.Text + StrConv(bBytes, vbUnicode)
End Sub

Private Sub txtURL_Change()

If textfilename = "" Then
tq$ = Mid$(txtURL, InStrRev(txtURL, "/") + 1)
Mid$(tq$, 1, 1) = UCase(Mid$(tq$, 1, 1))
txtFilename = BasePath + tq$
On Error Resume Next
Set fs = CreateObject("Scripting.FileSystemObject")
Set F = fs.Getfile(txtFilename)
txtResumeFrom.Text = Trim(Str(F.Size))
End If

End Sub
