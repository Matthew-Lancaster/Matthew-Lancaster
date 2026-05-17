VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "enPOP3 Logger"
   ClientHeight    =   8595
   ClientLeft      =   2010
   ClientTop       =   3765
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "END"
      Height          =   525
      Left            =   11790
      TabIndex        =   24
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add Contact"
      Default         =   -1  'True
      Height          =   510
      Left            =   7980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   975
      UseMaskColor    =   -1  'True
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   1755
      Top             =   1920
   End
   Begin MSComctlLib.ListView lvHeader 
      Height          =   3705
      Left            =   6210
      TabIndex        =   16
      Top             =   5010
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   6535
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Element"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LIST"
      Height          =   345
      Left            =   5040
      TabIndex        =   14
      Top             =   780
      Width           =   825
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UIDL"
      Height          =   345
      Left            =   5910
      TabIndex        =   13
      Top             =   1140
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HEADER"
      Height          =   345
      Left            =   6870
      TabIndex        =   12
      Top             =   780
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TOP"
      Height          =   345
      Left            =   5040
      TabIndex        =   11
      Top             =   1140
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RETR"
      Height          =   345
      Left            =   5910
      TabIndex        =   10
      Top             =   780
      Width           =   915
   End
   Begin VB.CheckBox chkAPOP 
      Caption         =   "Use APOP"
      Height          =   225
      Left            =   2520
      TabIndex        =   9
      Top             =   270
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CommandButton cmdUIDL 
      Caption         =   "UIDL"
      Height          =   435
      Left            =   2640
      TabIndex        =   8
      Top             =   1050
      Width           =   2355
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   330
      Left            =   270
      TabIndex        =   7
      Top             =   8400
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   6870
      TabIndex        =   6
      Text            =   "1"
      Top             =   1170
      Width           =   975
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List"
      Height          =   435
      Left            =   240
      TabIndex        =   5
      Top             =   1050
      Width           =   2355
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H00FFC0C0&
      Height          =   3315
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   5040
      Width           =   5745
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   2490
      TabIndex        =   3
      Text            =   "mypasswod"
      Top             =   525
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Text            =   "mrenigma"
      Top             =   510
      Width           =   2325
   End
   Begin VB.TextBox txtServer 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Text            =   "myemail"
      Top             =   150
      Width           =   2295
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Connect"
      Height          =   375
      Left            =   5010
      TabIndex        =   0
      Top             =   120
      Width           =   2205
   End
   Begin MSComctlLib.ListView lvMessages 
      Height          =   3465
      Left            =   165
      TabIndex        =   15
      Top             =   1515
      Width           =   17850
      _ExtentX        =   31485
      _ExtentY        =   6112
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Subject"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvHeader2 
      Height          =   3705
      Left            =   12120
      TabIndex        =   17
      Top             =   5010
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   6535
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Element"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label LabEmail 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   9180
      TabIndex        =   23
      Top             =   1245
      Width           =   7500
   End
   Begin VB.Label LabUser 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   9180
      TabIndex        =   22
      Top             =   960
      Width           =   7500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   9570
      TabIndex        =   20
      Top             =   420
      Width           =   2070
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   9570
      TabIndex        =   19
      Top             =   120
      Width           =   2070
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   7260
      TabIndex        =   18
      Top             =   105
      Width           =   2070
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit


Dim LL$, LL2$, LV As ListItem, FINISHED, WRITEMADE
Dim FROMT, TT, SUBJ, MFOLD
Dim UU, UU1, UU2, dDATE, StartDate, EndDate, GoodEnd
Public LoggonGood
Public WithEvents cPOPTest As clsPOP3
Attribute cPOPTest.VB_VarHelpID = -1
Dim bCollecting As Boolean
Dim sMessage As String

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)







Private Sub Command7_Click()

FINISHED = True
Unload Me

End Sub

'---------------------
'Count = 166 -- Mon 27-Jul-2009 13:25:53
'---------------------
'Microsoft Visual Basic
'---------------------
'Current Public/Instancing/DataSourceBehavior/Persistable property values are not valid in this type of project ('enPOP3'). The property values have been changed.
'---------------------------
'Got this error msg when convert from group project to a whole VBP one
'No Visiual Error Reported Seen


Private Sub Form_Load()
      
    Set FS = CreateObject("Scripting.FileSystemObject")
      
    Set cPOPTest = New clsPOP3
      
    Me.lvMessages.ColumnHeaders(1).Width = "1000" 'count
    Me.lvMessages.ColumnHeaders(2).Width = "3000" 'mail addy #1
    Me.lvMessages.ColumnHeaders(2).Width = "4000" 'mail addy #2
      Me.lvMessages.ColumnHeaders(3).Width = "4800" 'subject
      Me.lvMessages.ColumnHeaders(4).Width = "2600" 'date
      Me.lvMessages.ColumnHeaders(5).Width = "1600" 'size
      
      Me.lvHeader.ColumnHeaders(1).Width = "1500"
      Me.lvHeader.ColumnHeaders(2).Width = "8000"
      Me.lvHeader2.ColumnHeaders(1).Width = "1500"
      Me.lvHeader2.ColumnHeaders(2).Width = "8000"
      'Me.txtServer = GetSetting("enPOP3", "Settings", "Server", "myemail")
      'Me.txtUsername = GetSetting("enPOP3", "Settings", "Username", "mrenigma")
      'Me.txtPassword = GetSetting("enPOP3", "Settings", "Password", "mypassword")
    
        
    If Command$ <> "" Then Me.WindowState = 1
        
        On Error Resume Next
        
        frmMain.Show
        
        
        'Code to Auto Size Form based on controls used Including Menu Bar Sketchy Style
'Make sure form set to scale Twips
'put in your form load

x = 1
y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.TOP > y Then y = Control.Height + Control.TOP
        If InStr(Control.name, "Mnu_") > 0 Then mnu = 1
    End If
Next
'On Error GoTo 0

Me.Width = (x + 80)
If mnu = 1 Then
    pluso = 720 ': pluso = 1100 'Sometimes Different
Else
    pluso = 450
End If

Me.Height = (y + pluso)
Me.Refresh
DoEvents

Me.Left = (Screen.Width - Me.Width) / 2
Me.TOP = (Screen.Height - Me.Height) / 2

        
        
    'Me.WindowState = 1
        
    On Error Resume Next
        
        
        
        DoEvents
        txtServer = "mail.btinternet.com"
        txtUsername = "matt.lan@btinternet.com"
        txtPassword = ""
        DoEvents
        
        'Call KWab.cmdGo2_Click
        
        Call cmdGo_Click
        
        If LoggonGood = "" Then Unload Form1: Unload Me: Exit Sub
        
      
        'D:\VB6\VB-NT\00_Best_VB_01\VPN_AutoDial IP\#1-IP Data.txt
      
      
        
      
      
        On Error Resume Next
        Open "M:\00 Loggers Text\#00 EMAIL DATA POP3\#00 DATE INDEX.txt" For Input As #1
            Line Input #1, LL$
            Line Input #1, LL2$
        Close #1
      
        'IF LL2 ="" THEN LL=1 -- LATER
        LL2$ = ""
        
        
        StartDate = DateValue(LL$) + TimeValue(LL$)
      
        Call cmdList_Click
      
End Sub



Private Sub cmdGo_Click()
Dim sReturn As String
Dim Records As Long
Dim TotalSize As Double

      cPOPTest.TimeOut = 20
      cPOPTest.QUIT
      If cPOPTest.Connect(Me.chkAPOP.Value, Me.txtServer, Me.txtUsername, Me.txtPassword, 110) = True Then
         ' Connected Get Stat
         sReturn = cPOPTest.STAT
         If sReturn <> "" Then
            Records = Split(sReturn, " ")(0)
            TotalSize = Split(sReturn, " ")(1)
         
         End If
         PrintStatus vbCrLf & "Total Emails:" & Records
         PrintStatus "Total Mailbox Size:" & Format(TotalSize / 1024, 0) & " KB"

      End If

LoggonGood = sReturn

End Sub

Private Sub cmdList_Click()

'Exit Sub

'ALWAYS
GRAB_STORE_MESSAGES = True


'On Local Error GoTo 0
Timer1.Enabled = False
Dim sResult As String
Dim asMessages() As String
Dim liD As Long
Dim lSize As Long
Dim sSize As String
Dim lMessageCount As Long
Dim i As Long, ii
Dim HeaderItem As Email_Header
Dim xNode As Node
Dim R
ReDim HeaderArry(1)

      If cPOPTest.IsConnected = True Then
         bCollecting = True
         sResult = cPOPTest.STAT
         
         If lMessageCount > 0 Then
            'lMessageCount = lMessageCount + 10
            'If lMessageCount > Split(sResult, " ")(0) Then lMessageCount = Split(sResult, " ")(0)
         End If
         
         If lMessageCount = 0 Then lMessageCount = Split(sResult, " ")(0)

         Label2 = Trim(Str(lMessageCount)) + " Msgs"
         sResult = cPOPTest.List

         If lMessageCount > 0 Then
            ' Messages
            'Me.lvMessages.ListItems.Clear
            'For R = 1 To 30
            '    Me.lvMessages.ListItems.Add , , " "
            'Next
            
            Me.PB.Max = lMessageCount
         
            asMessages = Split(sResult, vbCrLf)
            
            'For i = lMessageCount - 140 To lMessageCount
            'If LL2$ = "" Then LL2$ = "1"
            
            ll3 = Val(LL2$) - 5
            If ll3 < 1 Then ll3 = 1
            
            
            
            For i = lMessageCount To 1 Step -1
            
            
            DoEvents
            DoEvents
               
               
               txtID = Trim(Str(i))
               
                'If i = 3255 Then Stop
               
               'Me.PB.Value = i
            
               Set HeaderItem = cPOPTest.Header(i)
               
                Dim aq(3)
               
                aq(0) = UCase("matt.lan@btinternet.com")
                aq(1) = UCase("matthew.lancaster@btopenworld.com")
                aq(2) = UCase("matthew.lancaster@btinternet.com")
                aq(3) = UCase("matthew.lancaster2@btinternet.com")
                
                Dim RS, WD, WD2, WD3, Mag
                
                Mag = 0
                
                WD = UCase(HeaderItem.ElementValue("From"))
                'If InStr(WD, UCase("codeproject.com")) > 0 Then Stop
                'Just the Domain name
                WD3 = Trim(Mid$(WD, InStrRev(WD, "@") + 1))
                If InStr(WD3, ">") > 0 Then WD3 = Mid$(WD3, 1, InStr(WD3, ">") - 1)
                
                
                'If InStr(KWab.txtRes, WD3) > 0 Then Mag = 1
                DoEvents
                
            '    For RS = 1 To UBound(aq)
            '        If InStr(WD, aq(RS)) Then
            '            Mag = 0
            '        End If
            '    Next
               
            '   Mag = 0
               
            '   If Mag = 0 Then
               ii = ii + 1
                ReDim Preserve HeaderArry(ii)
                HeaderArry(ii) = i
               
                
                dDATE = Trim(HeaderItem.ElementValue("date"))
                If IsNumeric(Mid$(dDATE, 1, 1)) Then dDATE = Format(DateValue(Mid(dDATE, 1, 20)), "DDD  ") + dDATE
                dDATE = Replace(dDATE, ",", " ")
               
               
               lSize = Split(asMessages(ii - 1), " ")(1)
               
               If lSize > 1024 Then
                  sSize = Format(lSize / 1024, ".##") & " kb"
               Else
                  sSize = lSize & " b"
               End If
               
                If GRAB_STORE_MESSAGES = True Then sMessage = cPOPTest.RETR(i)
                If GRAB_STORE_MESSAGES = False Then sMessage = cPOPTest.RETR(i)
            
            
            'Clipboard.Clear
            'Clipboard.SetText (sMessage)
            
            KK = InStr(sMessage, "X-Apparently")
            dDATE = Mid(sMessage, KK, InStr(KK, sMessage, vbCr) - KK)
            dDATE = Mid(dDATE, InStr(dDATE, ";") + 2)
                
            dDATE = Replace(dDATE, "GMT", "")
            dDATE = Replace(dDATE, "", "")
            dDATE = Replace(dDATE, "(BST)", "")
            UU = dDATE
            If InStr(dDATE, "-") > 0 Then
                UU = Mid(dDATE, 1, InStr(dDATE, "-") - 2)
                UU1 = Mid(dDATE, InStr(dDATE, "-") + 1, 2)
                UU = Mid(UU, 6)
                UU = DateValue(UU) + TimeValue(UU)
                UU2 = Val(UU1)
                UU = UU - TimeSerial(UU2, 0, 0)
            End If
            If InStr(dDATE, "+") > 0 Then
                UU = Mid(dDATE, 1, InStr(dDATE, "+") - 2)
                UU1 = Mid(dDATE, InStr(dDATE, "+") + 1, 2)
                UU = Mid(UU, 6)
                UU = DateValue(UU) + TimeValue(UU)
                UU2 = Val(UU1)
                UU = UU + TimeSerial(UU2, 0, 0)
            End If
                         
                         
            FROMT = Trim(HeaderItem.ElementValue("From"))
            DoEvents
            
            TT = InStrRev(FROMT, "@")
            If InStrRev(FROMT, "<", TT) > 0 Then
                FROMT = Mid(FROMT, InStrRev(FROMT, "<", TT) + 1)
                FROMT = Mid(FROMT, 1, InStr(FROMT, ">") - 1)
            End If
            
            FROMT = Replace(FROMT, "    ", "")
            FROMT = Replace(FROMT, vbCr, "")
            FROMT = Replace(FROMT, vbLf, "")
            FROMT = Replace(FROMT, Chr(9), "")
            FROMT = Replace(FROMT, """", "")
            FROMT = Replace(FROMT, "\", "")
            FROMT = Replace(FROMT, "?", "")
            FROMT = Replace(FROMT, ">", "")
            FROMT = Replace(FROMT, "<", "")
            FROMT = Replace(FROMT, """", "")
            FROMT = Trim(FROMT)
            
            SUBJ = Trim(HeaderItem.ElementValue("SUBJECT"))
            
            SUBJ = Replace(SUBJ, "    ", "")
            SUBJ = Replace(SUBJ, vbCr, "")
            SUBJ = Replace(SUBJ, vbLf, "")
            SUBJ = Replace(SUBJ, Chr(9), "")
            SUBJ = Replace(SUBJ, """", "")
            SUBJ = Replace(SUBJ, "\", "")
            SUBJ = Replace(SUBJ, "?", "")
            SUBJ = Replace(SUBJ, ":", "-")
            SUBJ = Replace(SUBJ, "/", "-")
            SUBJ = Replace(SUBJ, "=", "-")
            SUBJ = Replace(SUBJ, "*", "-")
            SUBJ = Replace(SUBJ, "<", "")
            SUBJ = Replace(SUBJ, ">", "")
            SUBJ = Replace(SUBJ, "|", "")
            'SUBJ = Replace(SUBJ, "'", "")
            'SUBJ = Replace(SUBJ, ",", "")
            
            If Len(SUBJ) > 80 Then SUBJ = Mid(SUBJ, 1, 75) + "##"
            
            'If UU > StartDate Then
            MFOLD = Format(UU, "YYYY-MM-MMM")
            On Error Resume Next
            MkDir "M:\00 Loggers Text\#00 EMAIL DATA POP3\" + MFOLD
            'On Error GoTo 0
            Dim FR1
            Err.Clear
                
            FileName = "M:\00 Loggers Text\#00 EMAIL DATA POP3\" + MFOLD + "\Email--" + Format(UU, "YYYY-MM-DD HH-MM-SS") + "-EMAIL-" + FROMT + "-#-SUBJECT-" + SUBJ + ".eml"
                
            'Shell "EXPLORER /SELECT,""" + FileName + """", vbNormalFocus
                
            'If Dir(FileName) <> "" Then A = A
            If FS.fileexists(FileName) = True Then Exit For
                    
            WRITEMADE = True
                    
            For RI = 1 To 100
            Err.Clear
            FR1 = FreeFile
            Open FileName For Output As #FR1
                Print #FR1, sMessage;
            Close #FR1
            Sleep 200
            If Err.Number = 0 Then Exit For
            Next
                
            If Err.Number > 0 Then Me.WindowState = 0: MsgBox "ERROR WRITE FILE " + vbCrLf + FileName, vbMsgBoxSetForeground: Stop
            
            'err.description
            'SET TO 1 IS IGNORED WHEN SORT IS USED
            'ii = Me.lvMessages.ListItems.Count
            'SORT ON 1ST COLUMN
            'lvwDESCENDING
                
                
                
               
           With Me.lvMessages
            Set LV = .ListItems.Add(, , Format(i, "00000"))
                LV.SubItems(1) = Trim(HeaderItem.ElementValue("From"))
                LV.SubItems(2) = Trim(HeaderItem.ElementValue("SUBJECT"))
                LV.SubItems(3) = Format(UU, "ddd DD-MMM-YYYY HH-MM-SS")
                LV.SubItems(4) = sSize
            End With
            DoEvents
            
            'Me.lvMessages.SortOrder = lvwDescending
            'Me.lvMessages.SortKey = 0
            'Me.lvMessages.Sorted = False
            Me.lvMessages.Sorted = True
            DoEvents

            Do
                'II2 = Me.lvMessages.ListItems.Count
                'If II2 >= 100 Then Me.lvMessages.ListItems.Remove (II2)
            Loop Until II2 < 100


            Me.lvMessages.ListItems(Me.lvMessages.ListItems.Count).Selected = True
            Me.lvMessages.SelectedItem.EnsureVisible
            
            'End If udate bigger start date
            
            
            'PUT THIS IF WANT
            'Me.txtResult = sMessage
               
            'If UU > EndDate Then EndDate = UU
            
            'Open "M:\00 Loggers Text\#00 EMAIL DATA POP3\#00 DATE INDEX.txt" For Output As #1
            '    Print #1, Format(EndDate, "DD-MM-YYYY HH:MM:SS")
            '    Print #1, Str(i - 1)
            'Close #1
               
               
            Next
         
         
         End If
         ' Me.vsMessages.AutoSize 0, Me.vsMessages.Cols - 1
      
        'ubound(asMessages())
        ReDim Preserve asMessages(ii)
        'ubound(asMessages())
        'lSize = Split(asMessages(ii - 1), " ")(1)
        
        For R = Me.lvMessages.ListItems.Count To 1 Step -1
            If Trim(Me.lvMessages.ListItems.item(R)) = "" Then
                Me.lvMessages.ListItems.Remove (R)
            End If
        Next
         
        Label3 = Trim(Str(Me.lvMessages.ListItems.Count)) + " Msgs When Filtered"
      
         bCollecting = False
      End If
      
      
      
      'If cPOPTest.IsConnected = True Then
      '   cPOPTest.QUIT
      'End If
      'Set cPOPTest = Nothing
Timer1.Enabled = True


'startdate

'If GoodEnd = True Then
'    Open App.Path + "\# EMAIL DATA\#00 DATE INDEX.txt" For Output As #1
'    Print #1, Format(EndDate, "DD-MM-YYYY HH:MM:SS")
'    Close #1
'End If

FINISHED = True

'Exit Sub

Unload Me

End Sub

Public Sub ReDoAfterAddy(HALO1EMailTest)
        Dim R

        For R = Me.lvMessages.ListItems.Count To 1 Step -1
        
'            goaty = Me.lvMessages.ListItems(ii).SubItems(1)
        
            
            
            Call StripEmailOutTheFromLine(R)
            If HALO1EMailTest = HALO1EMail Then
            
                Me.lvMessages.ListItems.Remove (R)
            
            End If
        
        
        Next
         
         Label3 = Trim(Str(Me.lvMessages.ListItems.Count)) + " Msgs When Filtered"


End Sub

Sub StripEmailOutTheFromLine(R)

Dim WD, WD3, WD4

HALO1EMail = lvMessages.ListItems(R).SubItems(1)

'No Name - No Brain - Faggots Brains - No Game
WD = HALO1EMail
HALO1User = "zz NO NAME"
If InStr(WD, "<") > 0 Then
                                
    
    'Just the Email
    WD3 = Mid$(WD, InStrRev(WD, "<") + 1)
    If InStr(WD3, ">") > 0 Then WD3 = Mid$(WD3, 1, InStr(WD3, ">") - 1)
    HALO1User = Trim(Mid$(WD, 1, InStr(WD, "<") - 1))
    
    Do
        WD4 = InStr(HALO1User, """")
        If WD4 > 0 Then Mid$(HALO1User, WD4, 1) = " "
    Loop Until WD4 = 0
    HALO1User = Trim(HALO1User)

    Do
        WD4 = InStr(HALO1User, " ")
        If WD4 > 0 Then Mid$(HALO1User, WD4, 1) = "_"
    Loop Until WD4 = 0
    HALO1User = Trim(HALO1User)

Else
    
    'The Full Email Addy is okay just a addy no padding
    WD3 = HALO1EMail

End If

HALO1EMail = WD3
If InStr(HALO1User, "@") > 0 And InStr(HALO1User, "<") = 0 And InStr(HALO1User, ">") = 0 And InStr(HALO1User, """") = 0 Then
    WD3 = Trim(Mid$(HALO1User, InStr(WD, "@") - 1))
    Mid(WD3, 2, 1) = UCase(Mid(WD3, 2, 1))
    HALO1User = WD3
End If

If HALO1User = "zz NO NAME" Then
    ''Just the Email Domian Later
    WD3 = Trim(Mid$(WD, InStr(WD, "@") + 1))
    If InStr(WD, "<") > 0 Then WD3 = Mid$(WD, InStr(WD, "<") + 1)
    Mid(WD3, 1, 1) = UCase(Mid(WD3, 1, 1))
    HALO1User = WD3
End If

LabUser = HALO1User
LabEmail = HALO1EMail
                

End Sub



Private Sub cmdUIDL_Click()
      Debug.Print cPOPTest.UIDL
End Sub

Private Sub Command1_Click()
Dim sMessage As String

        Exit Sub

      sMessage = cPOPTest.RETR(Me.txtID.Text)
      'DATER = cPOPTest.RETR(Me.)
      Open App.Path + "\Email--" + Format(Now, "YYYY-MM-DD HH-MM-SS") + ".eml" For Output As #1
      Print #1, sMessage;
      Close #1
      
      ' Me.PB.Value = Me.PB.Max
'      Me.txtResult = Left$(sMessage, 2000) ' Output first 2000 bytes
      
      Me.txtResult = sMessage
      ' Stop
End Sub

Private Sub Command2_Click()
      Debug.Print cPOPTest.TOP(Me.txtID.Text, 99)
End Sub

Private Sub Command3_Click()
Dim HeaderItem As Email_Header
Dim i As Long
Dim DumVar, A1$, B1$, WE

    If Trim(Me.txtID.Text) = "" Then Me.txtID.Text = "1"
      Set HeaderItem = cPOPTest.Header(HeaderArry(Me.txtID.Text))
      Me.lvHeader2.ListItems.Clear
        If lvHeader.ListItems.Count > 0 Then
        For WE = 1 To lvHeader.ListItems.Count
            A1$ = lvHeader.ListItems.item(WE).SubItems(1)
            B1$ = lvHeader.ListItems.item(WE)
            Me.lvHeader2.ListItems.Add (WE), , B1$
            Me.lvHeader2.ListItems.item(WE).SubItems(1) = A1$
        Next
        End If
          
      Me.lvHeader.ListItems.Clear
        
        On Error Resume Next
        DumVar = HeaderItem.ElementCount
        If Err.Number > 0 Then Exit Sub
        On Error GoTo 0
      
      If cPOPTest.GetLastError = "" Then
         For i = 1 To HeaderItem.ElementCount
            Me.lvHeader.ListItems.Add i, , HeaderItem.ElementNameFromIndex(i)
            Me.lvHeader.ListItems.item(i).SubItems(1) = HeaderItem.ElementValueFromIndex(i)
         Next
         'Me.lblFrom = "From:" & HeaderItem.ElementValue("From")
      Else
         Stop
      End If
    
'From
End Sub

Private Sub Command4_Click()
      Debug.Print cPOPTest.UIDL(Me.txtID.Text)

End Sub

Private Sub Command5_Click()
      Debug.Print cPOPTest.List(Me.txtID.Text)
End Sub

Private Sub Command6_Click()
'KWab.Show
'KWab2.Show
'Exit Sub
Command6.BackColor = QBColor(12)
Command6.Refresh
Dim HALO1EMailMark

HALO1EMailMark = HALO1EMail




'Call KWab.cmdAddContact_Click(HALO1User)

HALO1EMailMark = HALO1EMail
Call ReDoAfterAddy(HALO1EMailMark)


End Sub

Sub ISGoodEmail()

Exit Sub

Dim K, B1$, SameContact, K1, HALO4, XX, XXD

SameContact = False

If Trim(HALO1User) = "" And Trim(HALO1EMail) = "" Then Exit Sub
    
'TEMP OUT
'For K = 1 To KWab.lstContactsUser.ListCount
'    B1$ = KWab.lstContactsUser.List(K)
'    If InStr(B1$, HALO1User) > 0 Then
'        SameContact = True
'        If KWab.lstContactsEmail.List(K) = HALO1EMail Then MsgBox "You Got That Contact": Exit Sub
'            HALO4 = B1$ ' 'MsgBox "You Got That Contact But It Had/Got No Brain er Email" ': Exit For ': Exit Sub
'    End If
'Next


If SameContact = True Then
    HALO1User = HALO1User + "@" + Mid(HALO1EMail, 1, InStr(HALO1EMail, "@") - 1)
    Do
         XXD = HALO1User
         XX = InStr(HALO1User, ".")
         If XX > 0 Then Mid(HALO1User, XX + 1, 1) = UCase(Mid(HALO1User, XX + 1, 1))
        XX = InStr(HALO1User, "@")
        If XX > 0 Then Mid(HALO1User, XX + 1, 1) = UCase(Mid(HALO1User, XX + 1, 1))
    Loop Until XXD = HALO1User
    HALO4 = HALO1User
    SameContact = False
'TEMP OUT
'    For K = 1 To KWab.lstContactsUser.ListCount
'        B1$ = KWab.lstContactsUser.List(K)
'        If InStr(B1$, HALO1User) > 0 Then
'        SameContact = True
'        If KWab.lstContactsEmail.List(K) = HALO1EMail Then MsgBox "You Got That Contact": Exit Sub
'            HALO4 = B1$ ' 'MsgBox "You Got That Contact But It Had/Got No Brain er Email" ': Exit For ': Exit Sub
'        End If
'
'
'    Next
End If

If HALO4 <> "" Then HALO1User = HALO4
        
'Exit Sub
If SameContact = True Then
    K1 = 1
    K = InStrRev(HALO1User, " ")
    If K > 0 Then K1 = Mid$(HALO1User, K)
    If Mid$(K1, 2, 1) = "(" Then
        K = InStrRev(HALO1User, " "): K1 = Mid$(HALO1User, K + 2): K1 = Mid$(K1, 1, Len(K1) - 1)
        K = IsNumeric(K1)
        If K = True Then
            K = InStrRev(HALO1User, " ")
            HALO1User = Mid$(HALO1User, 1, K) + "(" + Trim(Str(Val(K1) + 1)) + ")"
        Else
            HALO1User = HALO1User + " (" + Trim(Str(Val(K1) + 1)) + ")"
        End If
        Else
'        K = InStrRev(HALO1Email, " "): K1 = Mid$(HALO1Email, K + 2): K1 = Mid$(K1, 1, Len(K1) - 1)
'        K = IsNumeric(K1)
            'If K = False Then MsgBox "problems isNumeric": Stop
            K1 = 1
            HALO1User = HALO1User + " (" + Trim(Str(Val(K1) + 1)) + ")"
    End If
End If

'HALO1User =  HALO1User
LabUser = HALO1User
LabEmail = HALO1EMail


End Sub


Private Sub cPOPTest_DataArrival(sData As String)
      ' If Not (bCollecting) Then
      PrintStatus "<" & sData
      ' End If
End Sub

Private Sub cPOPTest_DataSent(sData As String)
      ' If Not (bCollecting) Then
      PrintStatus ">" & sData
      ' End If
End Sub

Private Sub cPOPTest_Progress(dCurrectBytes As Double, dTotalSize As Double)
      On Error Resume Next
      Me.PB.Max = dTotalSize
      Me.PB.Value = dCurrectBytes
End Sub

Private Sub cPOPTest_POP3Error(ByVal sData As String)
    'If cPOPTest.GetLastError = "Timeout waiting for response" Then
    '    Call cmdGo_Click
    '    If LoggonGood = "" Then Exit Sub
    'End If
      Me.WindowState = 0
      'MsgBox ("Error occured - " & cPOPTest.GetLastError), vbMsgBoxSetForeground
        'Stop
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
        
       If LoggonGood <> "" Then
        
        
        'Unload KWab
        'Unload KWab2
        On Error Resume Next
      If cPOPTest.IsConnected = True Then
         cPOPTest.QUIT
      End If
      Set cPOPTest = Nothing
      
        End If
      
'      Call SaveSetting("enPOP3", "Settings", "Server", Me.txtServer)
'      Call SaveSetting("enPOP3", "Settings", "Username", Me.txtUsername)
'      Call SaveSetting("enPOP3", "Settings", "Password", Me.txtPassword)
    
    If WRITEMADE = True Then
        'SYNC MAIL DOES THIS
        Shell "D:\VB6\VB-NT\00_Best_VB_01\Sync Files\Sync Files.exe", vbMinimizedNoFocus
    
    
    End If
    
    
    If FINISHED = True Then End
      'End
End Sub

Private Sub PrintStatus(sText As String)
      Me.txtResult.Text = txtResult.Text & sText & vbCrLf
      Me.txtResult.SelStart = Len(Me.txtResult.Text)
End Sub

Private Sub Label1_Click()
'LABEL1
End Sub

Private Sub lvMessages_Click()
Exit Sub

Dim R, WD, WD3, WD4

For R = 1 To lvMessages.ListItems.Count
    If lvMessages.ListItems(R).Selected = True Then Exit For
Next

Me.txtID.Text = Str(R)

Call StripEmailOutTheFromLine(R)

'Make good Email should be
Call ISGoodEmail

Call Command3_Click
End Sub

Private Sub Timer1_Timer()
txtResult.Text = txtResult.Text + "** Saty Alive" + vbCrLf
Call cmdUIDL_Click
'Label1 = TimeOutVar
End Sub

Private Sub txtResult_Change()
Dim TT, TTH, TX
TX = ""
TT = InStr(txtResult, TX)
If TT > 0 Then
TTH = txtResult.Text
Mid(TTH, TT, Len(TX)) = String(Len(TX) - 5, "*") + String(5, " ")
txtResult.Text = TTH
End If

End Sub
