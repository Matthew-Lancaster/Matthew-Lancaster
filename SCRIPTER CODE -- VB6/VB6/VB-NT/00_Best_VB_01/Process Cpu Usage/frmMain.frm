VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "CPU Usage Monitor"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   9210
   Visible         =   0   'False
   Begin VB.Timer tmrRefresh 
      Interval        =   500
      Left            =   5850
      Top             =   2220
   End
   Begin ComctlLib.ListView lstProcess 
      Height          =   7695
      Left            =   -15
      TabIndex        =   1
      Top             =   390
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   13573
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Processes"
         Object.Width           =   7938
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Utilisation CPU%"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Process ID"
         Object.Width           =   2030
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "List of the processes running and their CPU usage :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   15
      TabIndex        =   0
      Top             =   135
      Width           =   5910
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                                        Matthieu Napoli
'      /\   ^   .^. +---._  /\   ^   .^. .-----..-----.
'     /  \ / \ /   \|     \/  \ / \ /   \|__  __\__  __|
'    /    M   |  A  |  D  /    M   |  A  \ T  T   |  |
'   /  /\  /\ | |-\ |    /  /\  /\ | |-\  \|  |   |  |
'  /__/  ><  \|_|  \|__./__/  ><  \|_|  \__|__|   |__|
'                                   madmatt_12@msn.com

' Program made by MadMatt
' clsPHDQuery class made by ShareVB (i lightly modified it)
' Thanks to ShareVB and the MSDN

' I'm french, so sorry if my english is not perfect ;-)
    

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim PDHQuery As clsPDHQuery, AA

Private Sub Form_Load()
    If Loaded = False And IsIDE = False Then Sleep 1000: Loaded = True
    
    ' Create the query
    Set PDHQuery = New clsPDHQuery
    ' Make the list of the processes

    Dim Item As ListItem
    Dim str() As String
    Dim Index As Long
    
    Dim tabID() As Long
    Dim i As Long
    Dim Counter As String
    Dim indexCounter As String
    Dim pName As String
    Dim Key As String
    Dim Ret As Long, XXi
    ' Get the processes list
    mpListProcess tabID()
    
    AA = LCase("Jpeg Slides PJpeg ART.exe winamp.exe Cid-RunMe.exe Explorer.EXE WMICPU2 MINI.exe VolumeBar ClipBoard Logger.exe Drive_Detach.exe URL Logger.exe Drives_Gig.exe WMICPU1 BIG.exe Active Idle.exe taskmgr.exe Tidal.exe BBCWeather.exe Cid-RunMe.exe vb6.exe OUTLOOK.EXE FeedDemon.exe CPUProcess.exe IEXPLORE.EXE WebDates.exe ViceVersa.exe TrueCrypt.exeThumbs.exe GoogleEarth.exe filezilla.exe ezcddaxfc.exe msimn.exe EXCEL.EXE mmc.exe WinRAR.exe WinHTTrack.exe Picasa3.exe cool2000.exe WaveEdit.exe")
        
    'lstProcess.ListItems.Clear
    For i = 0 To UBound(tabID)
        ' Get the process name
        pName = mpGetProcessName(tabID(i))
        ' Create the counter
        XXi = 0
        If InStr(AA, LCase(pName)) > 0 Then XXi = 1
        If InStr(LCase(pName), "winamp") > 0 Then XXi = 1
        
        If XXi = 1 Then
            Counter = PDHQuery.GetFormatedCounter(COUNTERPERF_PROCESS, GetFileNameWithoutExtension(pName), COUNTERPERF_PERCENTPROCESSORTIME)
                        
            ' Add the counter
            Ret = PDHQuery.AddCounter(Counter)
            ' Get his index
            indexCounter = PDHQuery.Count - 1
            If indexCounter < 0 Then indexCounter = 0
            If Ret = -1 Then indexCounter = -1   ' in case of error
            ' We save the process ID and his counter index in the item key
            Key = str(tabID(i)) + "//" + str(indexCounter)
            ' Add the item to the listview
            lstProcess.ListItems.Add , Key, pName
            Set Item = lstProcess.ListItems.Item(Key)
            Item.SubItems(2) = tabID(i)
        End If
    Next i
    
    ' Ask a first collect
    PDHQuery.Collect

    For i = 1 To lstProcess.ListItems.Count
        Set Item = lstProcess.ListItems.Item(i)
        
        ' Get the counter index
        str() = Split(Item.Key, "//")
        Index = Val(str(1))
        ' Fill the subitem
        If Index > -1 Then
            Item.SubItems(1) = PDHQuery.GetCounterData(Index)

        End If
        If XXi = 1 Then
            'If Val(Item.SubItems(1)) > 85 Then Stop 'killJPeg: Unload Me
        End If
    Next i


    Cmd = Command$
    'Cmd = "Quite"
    If IsIDE = False Or Cmd <> "" Then frmMain.Hide


End Sub

Private Sub Form_Resize()
x = 1
y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next
On Error GoTo 0

Me.Width = (x + 80)
If mnu = 1 Then
    pluso = 720: pluso = 1100 'Sometimes Different
Else
    pluso = 480
End If

Me.Height = (y + pluso)
Me.Refresh
DoEvents

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PDHQuery.Class_Terminate
    Set PDHQuery = Nothing
    lstProcess.ListItems.Clear
    
    Reset
    If MUSTUNLOAD = True Then Unload frmMain2
End Sub

Sub Reload()
    Exit Sub
    PDHQuery.Class_Terminate
    Set PDHQuery = Nothing
    'lstProcess.ListItems.Clear
    'Call Form_Load
    Set PDHQuery = New clsPDHQuery
    ' Get the processes list
    Dim tabID() As Long
    mpListProcess tabID()
    ' Ask a first collect
    PDHQuery.Collect

End Sub

Sub killJPeg()

End Sub

' Refresh the list
Public Sub tmrRefresh_Timer()
    
    
    
    Dim Item As ListItem
    Dim i As Long
    Dim str() As String
    Dim pName As String
    Dim Index As Long
    Dim Counter As String
    ' Get the CPU usage
    PDHQuery.Collect
    For i = 1 To lstProcess.ListItems.Count
        Set Item = lstProcess.ListItems.Item(i)
        XXi = 0
        
        If InStr(AA, LCase(pName)) > 0 Then XXi = 1

        ' Get the counter index
        str() = Split(Item.Key, "//")
        Index = Val(str(1))
        ' Fill the subitem
        If Index > -1 Then
            Item.SubItems(1) = PDHQuery.GetCounterData
