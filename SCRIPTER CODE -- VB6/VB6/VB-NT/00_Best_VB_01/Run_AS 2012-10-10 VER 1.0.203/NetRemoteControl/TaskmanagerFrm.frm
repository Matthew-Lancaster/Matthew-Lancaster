VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form TaskmanagerFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Taskmanager"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TaskmanagerFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7395
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   60
      Picture         =   "TaskmanagerFrm.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   30
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "        Remote Taskmanager"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   105
      TabIndex        =   4
      Top             =   120
      Width           =   7125
      Begin VB.Label Label1 
         Caption         =   "Remote TaskManager provides information about programs and processes running on remote computer!!!"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   6525
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Running Tasks"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3855
      Left            =   105
      TabIndex        =   2
      Top             =   1200
      Width           =   7185
      Begin MSWinsockLib.Winsock sckTaskManager 
         Left            =   1500
         Top             =   1530
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList SmallIconList 
         Left            =   2760
         Top             =   2220
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "TaskmanagerFrm.frx":0884
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   1980
         Top             =   1500
      End
      Begin MSComctlLib.ListView TaskListView 
         Height          =   3405
         Left            =   75
         TabIndex        =   3
         Top             =   300
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   6006
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "SmallIconList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1170
         Picture         =   "TaskmanagerFrm.frx":0CD6
         Top             =   1770
         Width           =   480
      End
   End
   Begin VB.CommandButton KittBttn 
      Caption         =   "Kill Task"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3015
      Picture         =   "TaskmanagerFrm.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1365
   End
End
Attribute VB_Name = "TaskmanagerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LItem As ListItem
Private Sub Form_Load()
Dim LItem As ListItem
    
    sckTaskManager.RemoteHost = RemoteForm.sckIdentifier.RemoteHost
    sckTaskManager.RemotePort = 8004
    sckTaskManager.Connect
    
With TaskListView
    .ColumnHeaders.Add , , "Task", .Width / 2 * 1.5
    .ColumnHeaders.Add , , "Status", .Width / 3
    .View = lvwReport
End With
End Sub

Private Sub KittBttn_Click()
On Error Resume Next
    Set LItem = TaskListView.SelectedItem
    If Len(LItem.Text) > 0 Then
        sckTaskManager.SendData "KillTask#" & LItem.Text
    End If
End Sub

Private Sub sckTaskManager_Close()
    Unload Me
End Sub

Private Sub sckTaskManager_Connect()
    sckTaskManager.SendData "GetTaskList#"
End Sub

Private Sub sckTaskManager_DataArrival(ByVal bytesTotal As Long)
Dim Command As String, txtData As String
Dim Position As Integer
    
    sckTaskManager.getData txtData
    
    Position = InStr(1, txtData, "#")
    Command = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Select Case Command
        Case "GetTaskList"
            AddinList (Mid(txtData, InStr(1, txtData, "#") + 1))
        Case "KillTask"
            TaskListView.ListItems.Remove LItem.Index
            sckTaskManager.SendData "GetTaskList#"
    End Select
End Sub

Private Sub TaskListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    TaskListView.SortKey = ColumnHeader.Index - 1
    If TaskListView.SortOrder = lvwAscending Then
        TaskListView.SortOrder = lvwDescending
    Else
        TaskListView.SortOrder = lvwAscending
    End If
    TaskListView.Sorted = True
End Sub
