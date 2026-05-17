VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Hallsoft Process Commander 1.0"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   8325
   Icon            =   "frmCentral.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   555
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Left            =   1575
      Top             =   2430
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   1665
      Top             =   4185
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentral.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentral.frx":28FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentral.frx":50B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentral.frx":5506
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentral.frx":5662
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentral.frx":7E16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2295
      Top             =   4185
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentral.frx":7F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentral.frx":83C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentral.frx":8CBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentral.frx":B46E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picInfo 
      Height          =   1950
      Left            =   3045
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   450
      Width           =   3885
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   6
         Left            =   1575
         TabIndex        =   18
         Top             =   1665
         Width           =   120
      End
      Begin VB.Shape shpValue 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   6
         Left            =   1485
         Top             =   1620
         Width           =   7500
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   5
         Left            =   1575
         TabIndex        =   17
         Top             =   1395
         Width           =   120
      End
      Begin VB.Shape shpValue 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   5
         Left            =   1485
         Top             =   1350
         Width           =   7500
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   4
         Left            =   1575
         TabIndex        =   16
         Top             =   1125
         Width           =   120
      End
      Begin VB.Shape shpValue 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   4
         Left            =   1485
         Top             =   1080
         Width           =   7500
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   3
         Left            =   1575
         TabIndex        =   15
         Top             =   855
         Width           =   120
      End
      Begin VB.Shape shpValue 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   3
         Left            =   1485
         Top             =   810
         Width           =   7500
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   2
         Left            =   1575
         TabIndex        =   14
         Top             =   585
         Width           =   120
      End
      Begin VB.Shape shpValue 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   2
         Left            =   1485
         Top             =   540
         Width           =   7500
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   1
         Left            =   1575
         TabIndex        =   13
         Top             =   315
         Width           =   120
      End
      Begin VB.Shape shpValue 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   1
         Left            =   1485
         Top             =   270
         Width           =   7500
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   0
         Left            =   1575
         TabIndex        =   12
         Top             =   45
         Width           =   120
      End
      Begin VB.Shape shpValue 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   0
         Left            =   1485
         Top             =   0
         Width           =   7500
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   11
         Top             =   1665
         Width           =   1020
      End
      Begin VB.Shape shpInfo 
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   6
         Left            =   0
         Top             =   1620
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Version"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   10
         Top             =   1395
         Width           =   1125
      End
      Begin VB.Shape shpInfo 
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   5
         Left            =   0
         Top             =   1350
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Description"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   9
         Top             =   1125
         Width           =   1080
      End
      Begin VB.Shape shpInfo 
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   4
         Left            =   0
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Write Time"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   8
         Top             =   855
         Width           =   1110
      End
      Begin VB.Shape shpInfo 
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   3
         Left            =   0
         Top             =   810
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   585
         Width           =   1125
      End
      Begin VB.Shape shpInfo 
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   2
         Left            =   0
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Width           =   300
      End
      Begin VB.Shape shpInfo 
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   1
         Left            =   0
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Path"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   45
         Width           =   615
      End
      Begin VB.Shape shpInfo 
         BorderColor     =   &H80000010&
         Height          =   285
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   1500
      End
   End
   Begin MSComctlLib.StatusBar statMain 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   4980
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Terminate"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Suspend"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Resume"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tProcess 
      Height          =   4425
      Left            =   45
      TabIndex        =   1
      Top             =   450
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   7805
      _Version        =   393217
      Indentation     =   476
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgList"
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lwProcess 
      Height          =   2400
      Left            =   3015
      TabIndex        =   3
      Top             =   2475
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Owner Process"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Thread"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Usage"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Process"
      Begin VB.Menu mnuFileTerminate 
         Caption         =   "&Terminate"
      End
      Begin VB.Menu mnuFileRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFileS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSuspend 
         Caption         =   "Suspend &thread"
      End
      Begin VB.Menu mnuFileResume 
         Caption         =   "R&esume thread"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Visit Hallsoft"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public wS_FileName          As New mclsFileProperties
Public wSuspendMode         As Boolean
Const sLocation As String = "frmMain"

Public Sub Init()
    mdlProcess.List_ActiveProcess tProcess
    mdlProcess.List_ActiveModules tProcess
    Dim i As Long
    For i = 0 To lblValue.Count - 1
        lblValue(i).Caption = ""
    Next i
    For i = 1 To tlbMain.Buttons.Count
        tlbMain.Buttons(i).ToolTipText = tlbMain.Buttons(i).Key
    Next i
End Sub

Private Sub Form_Load()

'        wS_FileName.FindFileInfo Node.Tag, False
        wS_FileName.FindFileInfo "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\URL Logger.exe", False
        'lblValue(0).Caption = Node.Tag
        lblValue(0).Caption = "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\URL Logger.exe"
        lblValue(1).Caption = CStr(CInt(wS_FileName.ByteSize / 1024)) & " Kb"
        lblValue(2).Caption = wS_FileName.CompanyName
        lblValue(3).Caption = wS_FileName.LastWriteTime
        lblValue(4).Caption = wS_FileName.FileDescription
        lblValue(5).Caption = wS_FileName.ProductVersion
        lblValue(6).Caption = wS_FileName.ProductName
        
        '// Get threads
        'If Node.Image = 1 Then
            mdlProcess.List_ActiveThreads lwProcess, GetParentIndex(tProcess, 1)
        'End If


Exit Sub
    
    Init
End Sub

Private Sub Form_Resize()
Exit Sub
On Error Resume Next
    With tProcess
        .Height = Me.ScaleHeight - tlbMain.Height - 4 - statMain.Height
        '.Width = Me.ScaleWidth - picInfo.Width - 4
    End With
    With picInfo
        .Left = tProcess.Left + tProcess.Width + 2
        .Width = Me.ScaleWidth - tProcess.Width - 8
    End With
    With lwProcess
        .Left = picInfo.Left
        .Width = picInfo.Width
        .Height = Me.ScaleHeight - picInfo.Height - tlbMain.Height - 8 - statMain.Height
    End With
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileRefresh_Click()
    Action_Call "Refresh"
End Sub

Private Sub mnuFileResume_Click()
    Action_Call "Resume"
End Sub

Private Sub mnuFileSuspend_Click()
    Action_Call "Suspend"
End Sub


Private Sub mnuFileTerminate_Click()
    Action_Call "Terminate"
End Sub

Private Sub mnuHelpAbout_Click()
    Action_Call "About"
End Sub

Private Sub mnuHelpWeb_Click()
    Action_Call "WWW"
End Sub


Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Action_Call Button.Key
End Sub

Private Sub tlbMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    With statMain.Panels(1)
        Select Case tlbMain.Buttons(1).Key
        Case "Terminate"
            .Text = "Terminates the selected process, the system may become unstable."
        Case "Suspend"
            .Text = "Selected thread will be suspended until resumed."
        Case "Resume"
            .Text = "Resumes the suspended thread, they are marked with red color."
        Case "Refresh"
            .Text = "Refresh the tree view."
        Case "Help"
            .Text = "Displays help window."
        Case "About"
            .Text = "Displays about window."
        End Select
    End With
End Sub


Private Sub tmrRefresh_Timer()
    List_ActiveProcess tProcess
    List_ActiveModules tProcess
    
    tmrRefresh.Enabled = False
End Sub

Private Sub tProcess_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo VB_Error
    Dim msg
    
    '// Check if you have suspended some app
    If wSuspendMode = True Then
        msg = MsgBox("You have suspended a thread, you should resume it before going to other threads. You shuld cancel ?", vbExclamation Or vbOKCancel, "Warning !")
    Else
    '// It's ok
        msg = vbOK
    End If
    If msg = vbOK Then
        wS_FileName.FindFileInfo Node.Tag, False
        lblValue(0).Caption = Node.Tag
        lblValue(1).Caption = CStr(CInt(wS_FileName.ByteSize / 1024)) & " Kb"
        lblValue(2).Caption = wS_FileName.CompanyName
        lblValue(3).Caption = wS_FileName.LastWriteTime
        lblValue(4).Caption = wS_FileName.FileDescription
        lblValue(5).Caption = wS_FileName.ProductVersion
        lblValue(6).Caption = wS_FileName.ProductName
        
        '// Get threads
        If Node.Image = 1 Then
            mdlProcess.List_ActiveThreads lwProcess, GetParentIndex(tProcess, 1)
        End If
    End If
Exit Sub
VB_Error:
    Err_Vb Err.Number, Err.Description, sLocation, "Node_Click"
    Resume Next
End Sub

Public Sub Action_Call(Key As String)
On Error GoTo era
    Dim msg
    Select Case Key
    Case "Terminate"
        msg = MsgBox("Are you sure you want to terminate the selected process ? This action will kill process, NOT module(s) !", vbExclamation Or vbYesNo, "Warning !")
        If msg = vbYes Then mdlProcess.Process_Kill ProcessId(mdlMisc.GetParentIndex(tProcess, 1) - 1): tmrRefresh.Enabled = True
    Case "Suspend"
        Thread_Suspend CLng(lwProcess.SelectedItem.SubItems(1))
        lwProcess.SelectedItem.ForeColor = vbRed
        lwProcess.SelectedItem.Bold = True
        wSuspendMode = True
    Case "Resume"
        Thread_Resume CLng(lwProcess.SelectedItem.SubItems(1))
        lwProcess.SelectedItem.ForeColor = vbWindowText
        lwProcess.SelectedItem.Bold = False
        wSuspendMode = False
    Case "Refresh"
        tmrRefresh.Enabled = True
    Case "About"
        frmAbout.Show vbModal, Me
    Case "WWW"
        File_Start "www.geocities.com/hallsoftware", "Open"
    End Select
Exit Sub
era:
    If Err = 91 Then
        MsgBox "There are no selected threads !", vbExclamation, "Suspend"
    Else
        Call Err_Vb(Err, Err.Description, sLocation, "tlbMain_ButtonClick")
    End If
End Sub
