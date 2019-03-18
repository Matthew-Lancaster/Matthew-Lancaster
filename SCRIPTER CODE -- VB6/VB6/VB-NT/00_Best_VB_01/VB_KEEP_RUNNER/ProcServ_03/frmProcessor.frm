VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmProcessor 
   Caption         =   "Processor"
   ClientHeight    =   7908
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11172
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7908
   ScaleWidth      =   11172
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvwProcessor 
      Height          =   7935
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   13991
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   44
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "DeviceID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "AddressWidth"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Architecture"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Availability"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Caption"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ConfigManagerErrorCode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ConfigManagerUserConfig"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "CpuStatus"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "CreationClassName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "CurrentClockSpeed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "CurrentVoltage"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "DataWidth"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "ErrorCleared"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "ErrorDescription"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "ExtClock"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Family"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "InstallDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "L2CacheSize"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "L2CacheSpeed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "LastErrorCode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Level"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "LoadPercentage"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Manufacturer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "MaxClockSpeed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "OtherFamilyDescription"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "PNPDeviceID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "PowerManagementCapabilities"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "PowerManagementSupported"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "ProcessorID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "ProcessorType"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "Revision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Role"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "SocketDesignation"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "StatusInfo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Text            =   "Stepping"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Text            =   "SystemCreationClassName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Text            =   "SystemName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Text            =   "UniqueID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Text            =   "UpgradeMethod"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Text            =   "Version"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Text            =   "VoltageCaps"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ShowForm()
' Processor

    Dim Enumerator As SWbemObjectSet
    Dim Object As SWbemObject
    Dim Item As ListItem

    On Error Resume Next
    
    MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_Processor")
    
    If Err.Number = 91 Then
        MsgBox "Please CONNECT to Local or Remote computer first", vbOKOnly, "Connect please!"
        Exit Sub
    End If
    
    On Error Resume Next
    If Form1.VAR_FORM1_EXIST = True Then
        If Err.Number = 0 Then VAR_FORM1_EXIST = Form1.VAR_FORM1_EXIST
    End If
    On Error GoTo 0
    If VAR_FORM1_EXIST = True Then
        Me.Hide
        MDIProcServ.Hide
    End If

    
    lvwProcessor.ListItems.Clear
    
    For Each Object In Enumerator
        Set Item = lvwProcessor.ListItems.Add(, Object.DeviceID, Object.DeviceID)

        With Item
            .SubItems(1) = Object.AddressWidth
            .SubItems(2) = Object.Architecture
            .SubItems(3) = Object.Availability
            .SubItems(4) = Object.Caption
'            .SubItems(5) = Object.ConfigManagerErrorCode
'            .SubItems(6) = Object.ConfigManagerUserConfig
            .SubItems(7) = Object.CpuStatus
            .SubItems(8) = Object.CreationClassName
            .SubItems(9) = Object.CurrentClockSpeed
            .SubItems(10) = Object.CurrentVoltage
            .SubItems(11) = Object.DataWidth
            .SubItems(12) = Object.Description
'            .SubItems(13) = Object.ErrorCleared
'            .SubItems(14) = Object.ErrorDescription
            .SubItems(15) = Object.ExtClock
            .SubItems(16) = Object.Family
'            .SubItems(17) = Object.InstallDate
            .SubItems(18) = Object.L2CacheSize
'            .SubItems(19) = Object.L2CacheSpeed
'            .SubItems(20) = Object.LastErrorCode
            .SubItems(21) = Object.Level
            .SubItems(22) = Object.LoadPercentage
            .SubItems(23) = Object.Manufacturer
            .SubItems(24) = Object.MaxClockSpeed
            .SubItems(25) = Object.Name
'            .SubItems(26) = Object.OtherFamilyDescription
'            .SubItems(27) = Object.PNPDeviceID
'            .SubItems(28) = Object.PowerManagementCapabilities
'            .SubItems(29) = Object.PowerManagementSuported
            .SubItems(30) = Object.ProcessorID
            .SubItems(31) = Object.ProcessorType
'            .SubItems(32) = Object.Revision
            .SubItems(33) = Object.Role
            .SubItems(34) = Object.SocketDesignation
            .SubItems(35) = Object.Status
            .SubItems(36) = Object.StatusInfo
'            .SubItems(37) = Object.Stepping
            .SubItems(38) = Object.SystemCreationClassName
            .SubItems(39) = Object.SystemName
'            .SubItems(40) = Object.UniqueID
            .SubItems(41) = Object.UpgradeMethod
            .SubItems(42) = Object.Version
'            .SubItems(43) = Object.VoltageCaps
        End With
    
        For R = 1 To 43
            Debug.Print Str(R) + " -- " + Item.SubItems(R)
        Next
        
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "Description")
        Item_2.SubItems(1) = Object.Description
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "Name")
        Item_2.SubItems(1) = Object.Name
        
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "AddressWidth")
        Item_2.SubItems(1) = Trim(Str(Object.AddressWidth)) + " Bit"
        
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "Manufacturer")
        Item_2.SubItems(1) = Object.Manufacturer
        
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "L2CacheSize")
        Item_2.SubItems(1) = Trim(Str(Object.L2CacheSize))
        'Item_2.SubItems(1) = Trim(Str(Object.L2CacheSize)) + "    AddressWidth_" + Trim(Str(Object.AddressWidth)) + "   SocketDesignation_" + Object.SocketDesignation

        'Item_2.SubItems(1) = Trim(Str(Object.AddressWidth)) + "     L2CacheSize" + Trim(Str(Object.L2CacheSize)) + "         Level " + Trim(Str(Object.Level)) + "        SocketDesignation " + Object.SocketDesignation
'        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "Level")
'        Item_2.SubItems(1) = Object.Level
        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "SocketDesignation")
        Item_2.SubItems(1) = Object.SocketDesignation
        
        
        'Item_2.SubItems(1) = Object.Manufacturer + "   Level_" + Trim(Str(Object.Level)) + "   ProcessorID_" + Object.ProcessorID
'        Set Item_2 = Form1.ListView_CPU_INFO.ListItems.Add(, , "ProcessorID")
'        Item_2.SubItems(1) = Object.ProcessorID
        
        TT_1 = ""
        TT_1 = TT_1 + "Description ___ " + Item.SubItems(12) + vbCrLf
        TT_1 = TT_1 + "Name _______ " + Item.SubItems(25) + vbCrLf
        TT_1 = TT_1 + "L2CacheSize _ " + Item.SubItems(18) + " __ "
        TT_1 = TT_1 + "AddressWidth _ " + Item.SubItems(1) + " __ "
        TT_1 = TT_1 + "Level _ " + Item.SubItems(21) + vbCrLf
        TT_1 = TT_1 + "Manufacturer __ " + Item.SubItems(23) + " __ "
        TT_1 = TT_1 + "SocketDesignation _ " + Item.SubItems(34) + vbCrLf
        TT_1 = TT_1 + "ProcessorID __ " + Item.SubItems(30)
        
    
    Next
    
'    Form1.Text_CPU_INFO.Text = TT_1

    'Me.Show

    Call Form1.LV_AutoSizeColumn(Form1.ListView_CPU_INFO, Form1.ListView_CPU_INFO.ColumnHeaders.Item(1))
    Call Form1.LV_AutoSizeColumn(Form1.ListView_CPU_INFO, Form1.ListView_CPU_INFO.ColumnHeaders.Item(2))


End Sub

Private Sub Form_Load()

    frmProcessor.height = MDIProcServ.height - 200
    frmProcessor.width = MDIProcServ.width - 275

    On Error Resume Next
        If Form1.VAR_FORM1_EXIST = True Then
            If Err.Number = 0 Then VAR_FORM1_EXIST = Form1.VAR_FORM1_EXIST
        End If
    On Error GoTo 0
    
    If VAR_FORM1_EXIST = True Then
        MDIProcServ.height = 0
        MDIProcServ.width = 0
        MDIProcServ.Top = -1000
    End If

End Sub

Private Sub lvwProcessor_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lvwProcessor.SortOrder = lvwAscending Then
        lvwProcessor.SortOrder = lvwDescending
    Else
        lvwProcessor.SortOrder = lvwAscending
    End If
    
    lvwProcessor.Sorted = True
    lvwProcessor.SortKey = ColumnHeader.Index - 1

End Sub

