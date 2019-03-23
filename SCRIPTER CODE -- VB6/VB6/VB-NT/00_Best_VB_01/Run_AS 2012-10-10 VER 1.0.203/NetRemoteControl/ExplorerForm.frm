VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ExplorerForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "Remote Explorer"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   Icon            =   "ExplorerForm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   8715
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2460
      Top             =   2250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   2400
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":1036
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":1488
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":3192
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":32EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2400
      Top             =   3450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":373E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckExplorer 
      Left            =   3990
      Top             =   3540
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   2820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":3A58
            Key             =   "Computer3"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":629A
            Key             =   "REMOVABLE"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":66EC
            Key             =   "FIXED"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":6B3E
            Key             =   "CDROM"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":6F90
            Key             =   "REMOTE"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":73E2
            Key             =   "DisconnectedNetDrive"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":7834
            Key             =   "closefolder"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExplorerForm.frx":7C86
            Key             =   "openfolder"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   3570
      TabIndex        =   2
      Top             =   690
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   11033
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList2"
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6255
      Left            =   60
      TabIndex        =   1
      Top             =   690
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   11033
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "download"
            Object.ToolTipText     =   "Download File from Remote System"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "send"
            Object.ToolTipText     =   "Send File to Remote System"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "createfolder"
            Object.ToolTipText     =   "Create New Folder"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "run"
            Object.ToolTipText     =   "Run Selected File"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "properties"
            Object.ToolTipText     =   "Show File Properties"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete File or Folder"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ExplorerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DownloadFileName As String
Dim DownloadFileSize As Long, FileNumber As Integer

Private Sub Form_Load()
    sckExplorer.RemoteHost = RemoteForm.sckIdentifier.RemoteHost
    sckExplorer.RemotePort = 8001
    sckExplorer.Connect
    TreeView1.Nodes.Add , , UCase(sckExplorer.RemoteHost), sckExplorer.RemoteHost, 1
    If sckExplorer.State = sckClosed Then _
        Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    sckExplorer.Close
    Unload ProgressForm01
End Sub

Private Sub Form_Resize()
On Error Resume Next
    TreeView1.Height = (Me.ScaleHeight - (Screen.TwipsPerPixelY * 48))
    ListView1.Width = (Me.ScaleWidth - (Screen.TwipsPerPixelX * 238))
    ListView1.Height = (Me.ScaleHeight - (Screen.TwipsPerPixelY * 48))
End Sub

Private Sub sckExplorer_Close()
    'MsgBox "Remote Connection Lost!", vbInformation, "Remote Exlorer"
    Unload Me
End Sub

Private Sub sckExplorer_Connect()
    sckExplorer.SendData "GetDrives#"
End Sub

Private Sub sckExplorer_DataArrival(ByVal bytesTotal As Long)
Dim Command As String, txtData As String
Dim Position As Integer

On Error GoTo ConnectionError
    sckExplorer.getData txtData, vbString
    Position = InStr(1, txtData, "#")
    Command = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    Select Case Command
        Case "GetDrives"
            Me.MousePointer = vbHourglass
            TreeView1.Enabled = False
            GetDrives (txtData)
            TreeView1.Enabled = True
            Me.MousePointer = vbNormal
        Case "GetFolders"
            Me.MousePointer = vbHourglass
            TreeView1.Enabled = False
            TreeView1.MousePointer = vbHourglass
            GetFolder (txtData)
            sckExplorer.SendData "GetFiles#" & TreeView1.SelectedItem.Key
            Me.MousePointer = vbNormal
            TreeView1.Enabled = True
            TreeView1.MousePointer = vbNormal
        Case "GetFiles"
            Me.MousePointer = vbHourglass
            ListView1.Enabled = False
            ListView1.MousePointer = ccHourglass
            GetFiles (txtData)
            ListView1.Enabled = True
            Me.MousePointer = vbNormal
            ListView1.MousePointer = ccDefault
        Case "FileSize"
            DoEvents
            CancelDownload = False
            Me.Enabled = False
            DownloadFileSize = CLng(txtData)
            ProgressForm01.MousePointer = vbHourglass
            ProgressForm01.ProgressBar1.Max = DownloadFileSize + 512
            ProgressForm01.FilePath.Caption = DownloadFileName
            ProgressForm01.Show
        Case "DownloadFile"
            Put #FileNumber, , txtData
            ProgressForm01.ProgressBar1.Value = ProgressForm01.ProgressBar1.Value + Len(txtData)
            ProgressForm01.ProgressLabel.Caption = Format(ProgressForm01.ProgressBar1.Value / ProgressForm01.ProgressBar1.Max, "###% Completed.")
            DoEvents
        Case "DownloadComplete"
            DoEvents
            Close #FileNumber
            ProgressForm01.ProgressBar1.Value = ProgressForm01.ProgressBar1.Max
            Unload ProgressForm01
            If CancelDownload = False Then _
                MsgBox "File recieved successfuly from remote system.", vbInformation, "Remote Download"
            Me.Enabled = True
        Case "CreateDirectory"
            TreeView1.Nodes.Add TreeView1.SelectedItem.Key, tvwChild, UCase(TreeView1.SelectedItem.Key & txtData & "\"), txtData, "closefolder", "openfolder"
            MsgBox "Directory successfuly created.", vbInformation
    End Select
    Exit Sub
ConnectionError:
    'MsgBox "Remote Connection Lost!", vbExclamation, "Remote Control"
End Sub

Sub GetDrives(drvs As String)
Dim Drive As String, Position As String
Dim DriveType As String

    Position = InStr(1, drvs, "|")
    While Position > 0
        Drive = Left(drvs, Position - 1)
        drvs = Mid(drvs, Position + 1)
        Position = InStr(1, drvs, vbCrLf)
        DriveType = Left(drvs, Position - 1)
        TreeView1.Nodes.Add UCase(sckExplorer.RemoteHost), tvwChild, UCase(Drive), Drive, DriveType
        drvs = Mid(drvs, Position + 2)
        Position = InStr(1, drvs, "|")
    Wend
End Sub

Sub GetFolder(Folders As String)
Dim Position As Integer, Folder As String
On Error Resume Next
    Position = InStr(1, Folders, "#")
    While Position > 0
        Folder = Left(Folders, Position - 1)
        TreeView1.Nodes.Add UCase(TreeView1.SelectedItem.Key), tvwChild, UCase(TreeView1.SelectedItem.Key & Folder & "\"), Folder, "closefolder", "openfolder"
        Folders = Mid(Folders, Position + 1)
        Position = InStr(1, Folders, "#")
    Wend
End Sub

Sub GetFiles(Files As String)
Dim Position As Integer, File As String
Dim Item As ListItem
On Error Resume Next
    ListView1.ListItems.Clear
    Position = InStr(1, Files, "#")
    While Position > 0
        File = Left(Files, Position - 1)
        ListView1.ListItems.Add , UCase(File), File, 1
        Files = Mid(Files, Position + 1)
        Position = InStr(1, Files, "#")
    Wend
    ListView1.Refresh
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim FileName As String, Respond As Integer
Dim DirName As String

    If TreeView1.SelectedItem.Key = UCase(sckExplorer.RemoteHost) Then _
        Exit Sub
On Error GoTo CancelError
    Select Case Button.Key
        Case "download"
            FileName = (TreeView1.SelectedItem.Key & ListView1.SelectedItem.Text)
            With CommonDialog1
                .CancelError = True
                .Flags = cdlOFNCreatePrompt
                .DialogTitle = "Save remote file to:"
                .DefaultExt = Right(ListView1.SelectedItem.Text, 3)
                .FilterIndex = 1
                .FileName = FileName
                .ShowSave
                If Trim(.FileName) = "" Then _
                    Exit Sub
                If Len(Dir(.FileName)) <> 0 Then
                    Respond = MsgBox(.FileName & " is exist!. Do you wish to overwrite this file?", vbQuestion + vbYesNo, "REMOTE CONTROL")
                    If Respond = vbNo Then Exit Sub
                End If
            End With
            sckExplorer.SendData "DownloadFile#" & FileName
            DownloadFileName = CommonDialog1.FileName
            FileNumber = FreeFile()
            Open DownloadFileName For Binary As #FileNumber
        Case "send"
            CancelSend = False
            With CommonDialog1
                .CancelError = True
                .Flags = cdlOFNFileMustExist
                .DialogTitle = "Select File to Send:"
                .Filter = "All Files|*.*"
                .FilterIndex = 1
                .ShowOpen
                If .FileTitle = "" Then _
                    Exit Sub
                FileName = TreeView1.SelectedItem.Key & .FileTitle
                DoEvents
                sckExplorer.SendData "SendFileName#" & FileName
                Sleep (10)
                SendFile .FileName, sckExplorer
            End With
        Case "createfolder"
            DirName = InputBox("Enter directory name", "Remote Explorer")
            If Trim(DirName) = "" Then _
                DirName = "Folder01"
            sckExplorer.SendData "CreateDirectory#" & TreeView1.SelectedItem.Key & DirName
        Case "run"
            FileName = TreeView1.SelectedItem.Key & ListView1.SelectedItem.Text
            sckExplorer.SendData "RunFile#" & FileName & Chr(0)
        Case "delete"
            Respond = MsgBox("Are you sour you want to delete remote system's file.", vbExclamation + vbYesNo, "Remote Explorer")
            If Respond = vbNo Then Exit Sub
            FileName = TreeView1.SelectedItem.Key & ListView1.SelectedItem.Text
            sckExplorer.SendData "DeleteFile#" & FileName & Chr(0)
    End Select
    Exit Sub
CancelError:
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = UCase(sckExplorer.RemoteHost) Then Exit Sub
    sckExplorer.SendData "GetFolders#" & Node.Key
End Sub
