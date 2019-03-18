VERSION 5.00
Begin VB.Form dlgConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect Dialog"
   ClientHeight    =   3420
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   6252
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6252
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optRemote 
      Caption         =   "Remote Machine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.OptionButton optLocal 
      Caption         =   "Local Machine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.TextBox txtConnect 
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   0
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Do you wish to connect to a local machine or remote?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Name of remote computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   3135
   End
End
Attribute VB_Name = "dlgConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()

    txtConnect.Enabled = False
    
End Sub

Private Sub OKButton_Click()

    StrConnect = txtConnect.Text
    MDIProcServ.ConnectWBEM (txtConnect.Text)
    
    If Err.Number <> 0 Then
        Exit Sub
    End If
    
'    frmComputerSystem.Visible = False
'    frmEventLogFile.Visible = False
'    frmLogicalDisk.Visible = False
    frmOperatingSystem.Visible = False
'    frmProcesses.Visible = False
'    frmProcessor.Visible = False
'    frmServices.Visible = False
    
    MDIProcServ.Caption = "ProcServ - building - this may take a minute or two ..."
    
'    frmComputerSystem.ShowForm
'    frmEventLogFile.ShowForm
'    frmLogicalDisk.ShowForm
    frmOperatingSystem.ShowForm
'    frmProcesses.ShowForm
'    frmProcessor.ShowForm
'    frmServices.ShowForm
    
    With MDIProcServ
        .Caption = "Done building information"
'        .Services.Enabled = True
'        .Processes.Enabled = True
'        .ComputerSystem.Enabled = True
'        .LogicalDisk.Enabled = True
        .OperatingSystem.Enabled = True
'        .Processor.Enabled = True
'        .EventLog.Enabled = True
    End With
    
    Unload Me

End Sub

Private Sub optLocal_Click()

    txtConnect.Text = "."
    txtConnect.Enabled = False

End Sub

Private Sub optRemote_Click()

    txtConnect.Enabled = True
    txtConnect.Text = ""
    txtConnect.SetFocus

End Sub

