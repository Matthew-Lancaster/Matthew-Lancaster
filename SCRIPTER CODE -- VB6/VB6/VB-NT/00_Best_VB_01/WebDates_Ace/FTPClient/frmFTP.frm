VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Not Connected"
   ClientHeight    =   7065
   ClientLeft      =   2400
   ClientTop       =   480
   ClientWidth     =   8280
   Icon            =   "frmFTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   8280
   Visible         =   0   'False
   Begin VB.CommandButton cmdRename 
      Caption         =   "Rename File"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton cmdCreateDir 
      Caption         =   "Create Directory"
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   6600
      Width           =   2175
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":0442
            Key             =   "dir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":059C
            Key             =   "file"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":09EE
            Key             =   "dll"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":0E40
            Key             =   "web"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":35F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection Properties"
      Height          =   2535
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   8055
      Begin VB.Frame Frame3 
         Caption         =   "Connection Mode"
         Height          =   855
         Left            =   4440
         TabIndex        =   23
         Top             =   1440
         Width           =   3255
         Begin VB.OptionButton optPassive 
            Caption         =   "Passive Connection"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   480
            Width           =   2295
         End
         Begin VB.OptionButton optActive 
            Caption         =   "Active Connection"
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "File Transfer Mode"
         Height          =   855
         Left            =   600
         TabIndex        =   22
         Top             =   1440
         Width           =   3375
         Begin VB.OptionButton optBin 
            Caption         =   "Binary File Transfer"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton optAscii 
            Caption         =   "ASCII File Transfer"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "&Disconnect"
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5760
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   5760
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         Height          =   255
         Left            =   4440
         TabIndex        =   21
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "User:"
         Height          =   255
         Left            =   4440
         TabIndex        =   20
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Server:"
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   390
         Width           =   975
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5760
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Last Accessed"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Attributes"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete item"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download file from Server"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh Directory Listing"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload file to server"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6600
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3360
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDir 
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   8055
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mFTP As cFTP
Attribute mFTP.VB_VarHelpID = -1


' ------------------------------------------------------------------------------------------------------------------------------
'
' If you have any questions or possible improvements to this code or the FTP class, email me: mike@ediy.co.nz
' For help on any of the class API functions, an excellent reference is available here:
' http://msdn.microsoft.com/library/default.asp?url=/workshop/networking/wininet/reference/win32_ref_entry.asp
' ------------------------------------------------------------------------------------------------------------------------------


Private Sub cmdConnect_Click()
   mWait
   If mFTP.OpenConnection(txtServer.Text, txtUser.Text, txtPassword.Text) Then
      'mFTP.SetFTPDirectory "/"
      Me.Caption = "Connected to " & txtServer.Text & " as " & txtUser.Text & " at " & Now
      ''RefreshDirectoryListing
      SaveSetting "eDIY FTP Client", "Settings", "Server", txtServer.Text
      SaveSetting "eDIY FTP Client", "Settings", "User", txtUser.Text
      SaveSetting "eDIY FTP Client", "Settings", "Password", txtPassword.Text
   End If
   mOk
End Sub

Private Sub cmdCreateDir_Click()
   Dim sName As String
   sName = Trim(InputBox("Please enter a name for this directory:"))
   If sName <> "" Then
      If mFTP.CreateFTPDirectory(sName) Then
         MsgBox "Directory " & sName & " was successfully created.", vbInformation
         'RefreshDirectoryListing
      Else
         MsgBox mFTP.GetLastErrorMessage
      End If
   End If
         
End Sub

Private Sub cmdDelete_Click()
   If Not (lv.SelectedItem Is Nothing) Then
      If MsgBox("Are you sure you want to delete " & lv.SelectedItem.Text & "?", vbQuestion + vbYesNo, "Delete?") = vbYes Then
         If mFTP.Directory(lv.SelectedItem.Text).Directory Then
            If mFTP.RemoveFTPDirectory(lv.SelectedItem.Text) Then
               MsgBox "Directory " & lv.SelectedItem.Text & " was successfully removed.", vbInformation
               'RefreshDirectoryListing
            Else
               MsgBox mFTP.GetLastErrorMessage
            End If
         Else
            If mFTP.DeleteFTPFile(lv.SelectedItem.Text) Then
               MsgBox "The file " & lv.SelectedItem.Text & " was successfully deleted.", vbInformation
               'RefreshDirectoryListing
            Else
               MsgBox mFTP.GetLastErrorMessage
            End If
         End If
      End If
   End If
End Sub

Private Sub cmdDisconnect_Click()
    mFTP.CloseConnection
    lv.ListItems.Clear
    Me.Caption = "Not Connected."
End Sub

Private Sub cmdDownload_Click()
   Dim lTimer As Long
   Dim strRemote As String
   Dim strLocal As String
   
   If Not (lv.SelectedItem Is Nothing) Then
     If lv.SelectedItem.Text <> ".." Then
        If Not mFTP.Directory(lv.SelectedItem.Text).Directory Then
           strRemote = lv.SelectedItem.Text
           cd.Filename = lv.SelectedItem.Text
           cd.ShowSave
           strLocal = Trim(cd.Filename)
           pb.Visible = True
           lTimer = Timer
           
           If strLocal <> "" Then
              mWait
              If Not mFTP.FTPDownloadFile(strLocal, strRemote) Then
                   mOk
                   MsgBox mFTP.GetLastErrorMessage
                Else
                   mOk
                   MsgBox "Download complete in " & Format(Timer - lTimer, "###,##0.00") & " seconds."
              End If
           End If
           pb.Visible = False
           'RefreshDirectoryListing
        End If
     End If
   End If
End Sub

Private Sub cmdRefresh_Click()
    RefreshDirectoryListing
End Sub

Private Sub cmdRename_Click()
    If Not (lv.SelectedItem Is Nothing) Then
       If lv.SelectedItem.Text <> ".." Then
            If Not mFTP.Directory(lv.SelectedItem.Text).Directory Then
               Dim sNewName As String
               sNewName = Trim(InputBox("Please enter a new name for this file:"))
               If sNewName <> "" Then
                  If mFTP.RenameFTPFile(lv.SelectedItem.Text, sNewName) Then
                     MsgBox "The file " & lv.SelectedItem.Text & "was successfuly renamed to " & sNewName
                     'RefreshDirectoryListing
                  Else
                     MsgBox mFTP.GetLastErrorMessage
                  End If
               End If
            End If
         End If
      End If
End Sub

Private Sub cmdUpload_Click()
   Dim lTimer As Long
   Dim strRemote As String
   
   mWait
   'cd.ShowOpen
   'strRemote = Trim(cd.FileTitle)

   
mFTP.SetFTPDirectory "/httpdocs/LoveFolder/Media/"
'RefreshDirectoryListing

'Exit Sub
   
Dim HALO As String
HALO = "D:\VB6\VB-NT\00_Best_VB_01\WebDates_Ace\FTPClient\Fast Show 05 Politics.mp3"
strRemote = "Fast Show 05 Politics.mp3"
   pb.Visible = True
   lTimer = Timer
   If strRemote <> "" Then
'       If Not mFTP.FTPUploadFile(cd.Filename, strRemote) Then
       If Not mFTP.FTPUploadFile(HALO, strRemote) Then
           MsgBox mFTP.GetLastErrorMessage
         Else
           MsgBox "Upload complete in " & Format(Timer - lTimer, "###,##0.00") & " seconds."
       End If
   End If
   pb.Visible = False
   mOk
   
RefreshDirectoryListing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mFTP.CloseConnection
End Sub


Private Sub lv_DblClick()
   mWait
   If Not (lv.SelectedItem Is Nothing) Then
      If lv.SelectedItem.Text = ".." Then
         mFTP.SetFTPDirectory lv.SelectedItem.Text
      Else
         If mFTP.Directory(lv.SelectedItem.Text).Directory Then
            mFTP.SetFTPDirectory lv.SelectedItem.Text
         End If
         
         If Not mFTP.Directory(lv.SelectedItem.Text).Directory Then
            cmdDownload_Click
         End If
      End If
   End If
   mOk
   
   'RefreshDirectoryListing
End Sub

Public Sub mFTP_FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)
    pb.Max = lTotalBytes
    pb.Min = 0
    If lTotalBytes >= lCurrentBytes Then
        pb.Value = lCurrentBytes
    End If
    DoEvents
End Sub

Private Sub RefreshDirectoryListing()
   Dim Item As cDirItem
   Dim lstX As ListItem
   Dim sAttr As String
   
   mWait
   lblDir.Caption = "Current Directory: " & Trim(mFTP.GetFTPDirectory)
   mFTP.GetDirectoryListing "*.*"
   
   lv.ListItems.Clear
   Set lstX = lv.ListItems.Add(, , "..")
   lstX.SmallIcon = 1
   For Each Item In mFTP.Directory
      sAttr = ""
      With Item
         If .Archive Then sAttr = sAttr & " A " Else sAttr = sAttr & " - "
         If .Compressed Then sAttr = sAttr & " C " Else sAttr = sAttr & " - "
         If .Directory Then sAttr = sAttr & " D " Else sAttr = sAttr & " - "
         If .Hidden Then sAttr = sAttr & " H " Else sAttr = sAttr & " - "
         If .Normal Then sAttr = sAttr & " N " Else sAttr = sAttr & " - "
         If .Offline Then sAttr = sAttr & " O " Else sAttr = sAttr & " - "
         If .ReadOnly Then sAttr = sAttr & " R " Else sAttr = sAttr & " - "
         If .System Then sAttr = sAttr & " S " Else sAttr = sAttr & " - "
         If .Temporary Then sAttr = sAttr & " T " Else sAttr = sAttr & " - "
      End With
      
      Set lstX = lv.ListItems.Add(, , Item.Filename)
      With lstX
      If Item.Directory Then
         .SmallIcon = 1
         .SubItems(1) = "< Directory >"
      Else
         .SmallIcon = 2
         .SubItems(1) = Item.FileSize
         .SubItems(2) = Item.LastWriteTime
         .SubItems(3) = sAttr
      End If
      End With
   Next
   mOk
End Sub

Private Sub Form_Load()
   Exit Sub
   Set mFTP = New cFTP
   
   
   txtServer.Text = "matthewlan.com" 'GetSetting("eDIY FTP Client", "Settings", "Server", "")
   txtUser.Text = "matthewlan" 'GetSetting("eDIY FTP Client", "Settings", "User", "")
   txtPassword.Text = "interned243" 'GetSetting("eDIY FTP Client", "Settings", "Password", "")
   mFTP.SetModeActive
   mFTP.SetTransferBinary
   Me.Caption = "Not Connected."
    cmdConnect_Click
End Sub

Private Sub optActive_Click()
   mFTP.SetModeActive
End Sub

Private Sub optPassive_Click()
   mFTP.SetModePassive
End Sub

Private Sub optAscii_Click()
   mFTP.SetTransferASCII
End Sub

Private Sub optBin_Click()
   mFTP.SetTransferBinary
End Sub

Private Sub txtPassword_GotFocus()
   txtPassword.SelStart = 0
   txtPassword.SelLength = 255
End Sub

Private Sub txtServer_GotFocus()
   txtServer.SelStart = 0
   txtServer.SelLength = 255
End Sub

Private Sub txtUser_GotFocus()
   txtUser.SelStart = 0
   txtUser.SelLength = 255
End Sub

Private Sub mWait()
   Screen.MousePointer = vbHourglass
End Sub

Private Sub mOk()
   Screen.MousePointer = vbDefault
End Sub

