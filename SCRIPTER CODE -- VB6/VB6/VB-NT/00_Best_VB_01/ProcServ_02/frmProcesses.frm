VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmProcesses 
   Caption         =   "Processes"
   ClientHeight    =   7908
   ClientLeft      =   60
   ClientTop       =   528
   ClientWidth     =   11172
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7908
   ScaleWidth      =   11172
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvwProcess 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19706
      _ExtentY        =   13780
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
      NumItems        =   42
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ProcessName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Caption"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CreationClassName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CreationDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "CSName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ExecutablePath"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Handle"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "HandleCount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "InstallDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "KernelMode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "MaximumWorkingSetSize"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "MinimumWorkingSetSize"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "OSCreationClassName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "OSName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "OtherOperationCount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "OtherTransferCount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "PageFaults"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "PageFileUsage"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "ParentProcessID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "PeakPageFileUsage"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "PeakVirtualSize"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "PeakWorkingSetSize"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Priority"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "PrivatePageCount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "ProcessID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "QuotaNonPagedPoolUsage"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "QuotaPagedPoolUsage"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "QuotaPeakNonPagedPoolUsage"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "QuotaPeakPagedPoolUsage"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "ReadOperationCount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "ReadTransferCount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "SessionID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "TerminationDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "ThreadCount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "UserModeTime"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "VirtualSize"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "WindowsVersion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Text            =   "WorkingSetSize"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Text            =   "WriteOperationCount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Text            =   "WriteTransferCount"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ShowForm()
' Processes

    Dim Enumerator As SWbemObjectSet
    Dim Object As SWbemObject
    Dim Item As ListItem

    On Error Resume Next
    
    MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_Process")
    
    If Err.Number = 91 Then
        MsgBox "Please CONNECT to Local or Remote computer first", vbOKOnly, "Connect please!"
        Exit Sub
    End If
    
    lvwProcess.ListItems.Clear
    
    For Each Object In Enumerator
        Set Item = lvwProcess.ListItems.Add(, Object.Name, Object.Name)
        With Item
            .SubItems(1) = Object.Caption
            .SubItems(2) = Object.CreationClassName
            .SubItems(3) = Object.CreationDate
            .SubItems(4) = Object.CSName
            .SubItems(5) = Object.Description
            .SubItems(6) = Object.ExecutablePath
            .SubItems(7) = Object.Handle
            .SubItems(8) = Object.HandleCount
            .SubItems(9) = Object.InstallDate
            .SubItems(10) = Object.KernelModeTime
            .SubItems(11) = Object.MaximumWorkingSetSize
            .SubItems(12) = Object.MinimumWorkingSetSize
            .SubItems(13) = Object.OSCreationClassName
            .SubItems(14) = Object.OSName
            .SubItems(15) = Object.OtherOperationCount
            .SubItems(16) = Object.OtherTransferCount
            .SubItems(17) = Object.PageFaults
            .SubItems(18) = Object.PageFileUsage
            .SubItems(19) = Object.ParentProcessID
            .SubItems(20) = Object.PeakPageFileUsage
            .SubItems(21) = Object.PeakVirtualSize
            .SubItems(22) = Object.PeakWorkingSetSize
            .SubItems(23) = Object.Priority
            .SubItems(24) = Object.PrivatePageCount
            .SubItems(25) = Object.ProcessID
            .SubItems(26) = Object.QuotaNonPagedPoolUsage
            .SubItems(27) = Object.QuotaPagedPoolUsage
            .SubItems(28) = Object.QuotaPeakNonPagedPoolUsage
            .SubItems(29) = Object.QuotaPeakPagedPoolUsage
            .SubItems(30) = Object.ReadOperationCount
            .SubItems(31) = Object.ReadTransferCount
            .SubItems(32) = Object.SessionID
            .SubItems(33) = Object.Status
            .SubItems(34) = Object.TerminationDate
            .SubItems(35) = Object.ThreadCount
            .SubItems(36) = Object.UserModeTime
            .SubItems(37) = Object.VirtualSize
            .SubItems(38) = Object.WindowsVersion
            .SubItems(39) = Object.WorkingSetSize
            .SubItems(40) = Object.WriteOperationCount
            .SubItems(41) = Object.WriteTransferCount
        End With
    Next
'    For R = 1 To 41
'    Debug.Print Item.SubItems(R)
'    Next
    Me.Show

End Sub

Private Sub Form_Load()

    frmProcesses.Height = MDIProcServ.Height - 200
    frmProcesses.Width = MDIProcServ.Width - 275

End Sub

Private Sub lvwProcess_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lvwProcess.SortOrder = lvwAscending Then
        lvwProcess.SortOrder = lvwDescending
    Else
        lvwProcess.SortOrder = lvwAscending
    End If
    
    lvwProcess.Sorted = True
    lvwProcess.SortKey = ColumnHeader.Index - 1

End Sub
