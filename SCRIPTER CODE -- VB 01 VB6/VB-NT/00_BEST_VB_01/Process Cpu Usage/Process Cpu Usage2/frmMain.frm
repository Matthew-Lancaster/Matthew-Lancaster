VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "CPU Usage - MadMatt"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   11520
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView lstProcess 
      Height          =   8295
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   14631
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Processus"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Utilisation CPU"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   500
      Left            =   4710
      Top             =   45
   End
   Begin ComctlLib.ListView lstProcess2 
      Height          =   8295
      Left            =   630
      TabIndex        =   2
      Top             =   390
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   14631
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Processus"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Utilisation CPU"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Old Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6735
      TabIndex        =   6
      Top             =   1395
      Width           =   1425
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6735
      TabIndex        =   5
      Top             =   1845
      Width           =   225
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Peek CPU USage On WinAmp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6705
      TabIndex        =   4
      Top             =   450
      Width           =   4455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6720
      TabIndex        =   3
      Top             =   855
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "List of the processes running and their CPU usage :"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3630
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


Dim PDHQuery As clsPDHQuery

Private Sub Form_Load()
    
    lstProcess2.Top = lstProcess.Top
    lstProcess2.Left = lstProcess.Left
    lstProcess2.Width = lstProcess.Width
    lstProcess2.Height = lstProcess.Height
    ' Create the query
    Set PDHQuery = New clsPDHQuery
    ' Make the list of the processes
    Dim tabID() As Long
    Dim i As Long
    Dim Counter As String
    Dim indexCounter As String
    Dim pName As String
    Dim Key As String
    Dim Ret As Long
    ' Get the processes list
    mpListProcess tabID()
    lstProcess.ListItems.Clear
    For i = 0 To UBound(tabID)
        ' Get the process name
        pName = mpGetProcessName(tabID(i))
        ' Create the counter
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
    Next i
    ' Ask a first collect
    PDHQuery.Collect
End Sub

Sub PRoRemove()
    
    ' Create the query
    Set PDHQuery = New clsPDHQuery
    
    ' Make the list of the processes
    Dim tabID() As Long
    Dim i As Long
    Dim Counter As String
    Dim indexCounter As String
    Dim pName As String
    Dim Key As String
    Dim Ret As Long
    ' Get the processes list
    mpListProcess tabID()
    lstProcess.ListItems.Clear
    For i = 0 To UBound(tabID)
        ' Get the process name
        pName = mpGetProcessName(tabID(i))
        ' Create the counter
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
    Next i
    ' Ask a first collect
    PDHQuery.Collect

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set PDHQuery = Nothing
End Sub

' Refresh the list
Private Sub tmrRefresh_Timer()
    ' Listitem
    Dim Item As ListItem
    Dim i As Long
    Dim str() As String
    Dim pName As String
    Dim Index As Long
    Dim Counter As String
    ' Get the CPU usage
    PDHQuery.Collect
    'lstProcess2.ListItems.Clear
    flag = 0
    For i = 1 To lstProcess.ListItems.Count
        
        Set Item = lstProcess.ListItems.Item(i)
        'If UCase(Item) = "WINAMP.EXE" Then
        If UCase(Item) = UCase("Jpeg Slides PJpeg ART.exe") Then
            flag = 1
        End If
        If UCase(Item) = "TIDAL.EXE" Then
            a = a 'flag = 1
        End If
        'if item="Winamp"
        ' Get the counter index
        str() = Split(Item.Key, "//")
        Index = Val(str(1))
        ' Fill the subitem
        'If InStr(UCase(Item), "WINAMP.EXE") > 0 Then Stop
        'If Index > -1 Then
            Err.Clear
            On Error Resume Next
            ace = PDHQuery.GetCounterData(Index)
            'If Val(ace) = 0 And IsNumeric(ace) = False Then Stop
            'If IsNumeric(ace) Then
                value1 = Val(ace)
                Item.SubItems(1) = value1
            'End If
            
                
        'End If
    Next i
    
    If flag = 0 Then
    Label4 = Label2: Label2 = "0"
    End If
                
    
    For i = 1 To lstProcess.ListItems.Count
        a1$ = lstProcess.ListItems.Item(i).SubItems(1)
        b1$ = lstProcess.ListItems.Item(i)
        c1$ = lstProcess.ListItems.Item(i).Key
        'If Val(a1$) > 0 Then
            If UCase(b1$) = UCase("Jpeg Slides PJpeg ART.exe") Then
                If Val(a1$) > Val(Label2) Then Label2.Caption = a1$
            End If
            
            index2 = index2 + 1
            lstProcess2.ListItems.Add , c1$, b1$
            lstProcess2.ListItems.Item(index2).SubItems(1) = a1$
        'End If
    Next
    
    'If 1 = 2 Then
    For i = lstProcess.ListItems.Count To 1 Step -1
    If PDHQuery.GetCounterData(i) = "End" Then
            b1$ = lstProcess.ListItems.Item(i)
            
            'lstProcess2.ListItems.Remove (i)
            For ii = lstProcess2.ListItems.Count To 1 Step -1
                b2$ = lstProcess2.ListItems.Item(ii)
                If b1$ = b2$ Then
                    lstProcess2.ListItems.Remove (ii)
                End If
            Next
            
            lstProcess.ListItems.Remove (i)
    
    End If
    Next
    'End If
    
    For i = lstProcess2.ListItems.Count To 1 Step -1
        a1$ = lstProcess2.ListItems.Item(i).SubItems(1)
        If a1$ = "" Or Val(a1$) = 0 Then
            lstProcess2.ListItems.Remove (i)
        End If
    Next
    
    
End Sub
