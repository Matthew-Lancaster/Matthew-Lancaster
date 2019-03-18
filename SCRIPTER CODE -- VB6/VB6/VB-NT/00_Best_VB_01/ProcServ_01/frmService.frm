VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServices 
   Caption         =   "Services"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmService.frx":0000
   ScaleHeight     =   7905
   ScaleWidth      =   11175
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Resume 
      Caption         =   "Resume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   7080
      Top             =   360
   End
   Begin VB.CommandButton Start 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Pause 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Stop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwService 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Click on column header to sort"
      Top             =   1320
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11245
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   23
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Service Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Service Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Service State"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "AcceptStop"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "AcceptPause"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Caption"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "CheckPoint"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "CreationClassName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "DesktopInteract"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "DisplayName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ErrorControl"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ExitCode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "InstallDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "PathName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "ProcessID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "StartMode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "StartName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "SystemCreationClassName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "SystemName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "TagID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "WaitHint"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   6120
      X2              =   6120
      Y1              =   120
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   120
      X2              =   6120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "ACTIONS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   3855
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   120
      X2              =   6120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   1200
   End
End
Attribute VB_Name = "frmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Item As ListItem
Private TimerCount As Integer
Private ServiceObject As SWbemObject
Private ServiceName As String

Public Sub ShowForm()
' Services

    Dim Enumerator As SWbemObjectSet
    Dim Object As SWbemObject
    Dim Item As ListItem
    
    On Error Resume Next
    
    MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_Service")
    
    If Err.Number = 91 Then
        MsgBox "Please CONNECT to Local or Remote computer first", vbOKOnly, "Connect please!"
        Exit Sub
    End If
    
    lvwService.ListItems.Clear
    
    For Each Object In Enumerator
        Set Item = lvwService.ListItems.Add(, Object.Name, Object.Name)
        With Item
            .SubItems(1) = Object.Description
            .SubItems(2) = Object.State
            .SubItems(3) = Object.AcceptPause
            .SubItems(4) = Object.AcceptStop
            .SubItems(5) = Object.Caption
            .SubItems(6) = Object.CheckPoint
            .SubItems(7) = Object.CreationClassName
            .SubItems(8) = Object.DesktopInteract
            .SubItems(9) = Object.DisplayName
            .SubItems(10) = Object.ErrorControl
            .SubItems(11) = Object.ExitCode
            .SubItems(12) = Object.InstallDate
            .SubItems(13) = Object.PathName
            .SubItems(14) = Object.ProcessID
            .SubItems(15) = Object.StartMode
            .SubItems(16) = Object.StartName
            .SubItems(17) = Object.State
            .SubItems(18) = Object.Status
            .SubItems(19) = Object.SystemCreationClassName
            .SubItems(20) = Object.SystemName
            .SubItems(21) = Object.TagID
            .SubItems(22) = Object.WaitHint
        End With
    Next
    
    Me.Show

End Sub

Private Sub Form_Load()

    frmServices.Height = MDIProcServ.Height - 200
    frmServices.Width = MDIProcServ.Width - 275

End Sub

Private Sub lvwService_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lvwService.SortOrder = lvwAscending Then
        lvwService.SortOrder = lvwDescending
    Else
        lvwService.SortOrder = lvwAscending
    End If
    
    lvwService.Sorted = True
    lvwService.SortKey = ColumnHeader.Index - 1

End Sub

Private Sub Resume_Click()
    
    On Error Resume Next
    ServiceName = lvwService.SelectedItem.Text
    If Err.Number = 0 Then
    
        ' Note how the CIM method "StopService" of Win32_Service
        ' is executed as if it were an automation method of SWbemObject
        Set ServiceObject = MDIProcServ.WBEMServices.Get("Win32_Service='" & ServiceName & "'")
        ServiceObject.ResumeService
        
        Set Item = frmServices.lvwService.FindItem(ServiceName)
        Item.SubItems(2) = ServiceObject.TargetInstance.State
        Timer1.Enabled = True
    End If

End Sub

Private Sub Timer1_Timer()

    TimerCount = TimerCount + 1
    If TimerCount > 10 Then
        Item.Bold = False
        Timer1.Enabled = False
        TimerCount = 0
        Me.ShowForm
    Else
        TimerCount = TimerCount + 1
        Item.Bold = Not Item.Bold
    End If

End Sub

Private Sub Pause_Click()
    
    On Error Resume Next
    ServiceName = lvwService.SelectedItem.Text
    If Err.Number = 0 Then
        Set ServiceObject = MDIProcServ.WBEMServices.Get("Win32_Service='" & ServiceName & "'")
        ' Note how the CIM method "PauseService" of Win32_Service
        ' is executed as if it were an automation method of SWbemObject
        ServiceObject.PauseService
        
        Set Item = frmServices.lvwService.FindItem(ServiceName)
        Item.SubItems(2) = ServiceObject.TargetInstance.State
        Timer1.Enabled = True
    End If

End Sub

Private Sub Start_Click()

    On Error Resume Next
    ServiceName = lvwService.SelectedItem.Text
    If Err.Number = 0 Then
        ' Note how the CIM method "StartService" of Win32_Service
        ' is executed as if it were an automation method of SWbemObject
        Set ServiceObject = MDIProcServ.WBEMServices.Get("Win32_Service='" & ServiceName & "'")
        ServiceObject.StartService
        
        Set Item = frmServices.lvwService.FindItem(ServiceName)
        Item.SubItems(2) = ServiceObject.TargetInstance.State
        Timer1.Enabled = True
    End If

End Sub

Private Sub Stop_Click()
    
    On Error Resume Next
    ServiceName = lvwService.SelectedItem.Text
    If Err.Number = 0 Then
    
        ' Note how the CIM method "StopService" of Win32_Service
        ' is executed as if it were an automation method of SWbemObject
        Set ServiceObject = MDIProcServ.WBEMServices.Get("Win32_Service='" & ServiceName & "'")
        ServiceObject.StopService
        
        Set Item = frmServices.lvwService.FindItem(ServiceName)
        Item.SubItems(2) = ServiceObject.TargetInstance.State
        Timer1.Enabled = True
    End If

End Sub


