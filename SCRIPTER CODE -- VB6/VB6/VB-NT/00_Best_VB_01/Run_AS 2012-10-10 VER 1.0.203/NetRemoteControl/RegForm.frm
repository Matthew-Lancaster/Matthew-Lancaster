VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form RegForm 
   Caption         =   "Remote Registry Editor"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   Icon            =   "RegForm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   7710
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   6645
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   7305
      Begin MSWinsockLib.Winsock sckReg 
         Left            =   660
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Frame Frame7 
         Caption         =   "Commands:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Left            =   1035
         TabIndex        =   11
         Top             =   4560
         Width           =   5175
         Begin VB.Frame Frame9 
            Height          =   1635
            Left            =   2640
            TabIndex        =   16
            Top             =   210
            Width           =   2385
            Begin VB.CommandButton DeleteKeyBttn 
               Caption         =   "D&elete Sub Key"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   120
               TabIndex        =   18
               Top             =   660
               Width           =   2115
            End
            Begin VB.CommandButton AddKeyBttn 
               Caption         =   "Add &Sub Key"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   120
               TabIndex        =   17
               Top             =   180
               Width           =   2115
            End
         End
         Begin VB.Frame Frame8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1635
            Left            =   150
            TabIndex        =   12
            Top             =   210
            Width           =   2385
            Begin VB.CommandButton ShowValueBttn 
               Caption         =   "Sh&ow Value"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   120
               TabIndex        =   15
               Top             =   1140
               Width           =   2115
            End
            Begin VB.CommandButton DeleteValueBttn 
               Caption         =   "&Delete Value"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   120
               TabIndex        =   14
               Top             =   660
               Width           =   2115
            End
            Begin VB.CommandButton NewValueBttn 
               Caption         =   "&Add Value"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   120
               TabIndex        =   13
               Top             =   180
               Width           =   2115
            End
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Value Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   705
         Left            =   525
         TabIndex        =   9
         Top             =   3810
         Width           =   6195
         Begin VB.TextBox ValueDataTxt 
            Height          =   315
            Left            =   120
            MaxLength       =   255
            TabIndex        =   10
            Top             =   240
            Width           =   5955
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Select Value Data Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   705
         Left            =   2160
         TabIndex        =   7
         Top             =   3090
         Width           =   2925
         Begin VB.ComboBox ValueTypeCombo 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2685
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Value Name or Sub Key Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   705
         Left            =   525
         TabIndex        =   5
         Top             =   2340
         Width           =   6195
         Begin VB.TextBox KeyNameTxt 
            Height          =   315
            Left            =   120
            MaxLength       =   255
            TabIndex        =   6
            Top             =   240
            Width           =   5955
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sub Key:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   705
         Left            =   525
         TabIndex        =   3
         Top             =   1620
         Width           =   6195
         Begin VB.TextBox SubKeyTxt 
            Height          =   315
            Left            =   120
            MaxLength       =   255
            TabIndex        =   4
            Top             =   240
            Width           =   5955
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Root Keys:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   705
         Left            =   1725
         TabIndex        =   1
         Top             =   900
         Width           =   3795
         Begin VB.ComboBox RootKeyCombo 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   3555
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   555
         Left            =   210
         TabIndex        =   19
         Top             =   240
         Width           =   6885
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "RegForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_SZ = 1                         ' Unicode nul terminated string

Private Sub AddKeyBttn_Click()
Dim txtData As String, idx As Integer

    If Trim(KeyNameTxt.Text) = "" Then
        MsgBox "Pleae Enter Key Name", vbExclamation
        KeyNameTxt.SetFocus
        Exit Sub
    End If
    idx = RootKeyCombo.ListIndex
    txtData = "AddKey#" & RootKeyCombo.Text & "#" & SubKeyTxt.Text & "#" & KeyNameTxt.Text & "#"
    sckReg.SendData txtData
End Sub

Private Sub DeleteKeyBttn_Click()
Dim txtData As String, idx As Integer

    If Trim(KeyNameTxt.Text) = "" Then
        MsgBox "Pleae Enter Key Name", vbExclamation
        KeyNameTxt.SetFocus
        Exit Sub
    End If
    idx = RootKeyCombo.ListIndex
    txtData = "DeleteKey#" & RootKeyCombo.Text & "#" & SubKeyTxt.Text & "#" & KeyNameTxt.Text & "#"
    sckReg.SendData txtData
End Sub

Private Sub DeleteValueBttn_Click()
Dim txtData As String, idx As Integer
    If Trim(KeyNameTxt.Text) = "" Then
        MsgBox "Pleae Enter Value Name", vbExclamation
        KeyNameTxt.SetFocus
        Exit Sub
    End If
    idx = RootKeyCombo.ListIndex
    txtData = "DeleteValue#" & RootKeyCombo.Text & "#" & SubKeyTxt.Text & "#" & KeyNameTxt.Text & "#"
    sckReg.SendData txtData
End Sub

Private Sub Form_Load()
    
    sckReg.RemoteHost = RemoteForm.sckIdentifier.RemoteHost
    sckReg.RemotePort = 8002
    sckReg.Connect
    'If sckReg.State = sckNotConnected Then _
        Unload Me
    
    RootKeyCombo.AddItem "HKEY_CLASSES_ROOT"
    RootKeyCombo.ItemData(RootKeyCombo.NewIndex) = HKEY_CLASSES_ROOT
    RootKeyCombo.AddItem "HKEY_CURRENT_USER"
    RootKeyCombo.ItemData(RootKeyCombo.NewIndex) = HKEY_CURRENT_USER
    RootKeyCombo.AddItem "HKEY_LOCAL_MACHINE"
    RootKeyCombo.ItemData(RootKeyCombo.NewIndex) = HKEY_LOCAL_MACHINE
    RootKeyCombo.AddItem "HKEY_USERS"
    RootKeyCombo.ItemData(RootKeyCombo.NewIndex) = HKEY_USERS
    RootKeyCombo.AddItem "HKEY_CURRENT_CONFIG"
    RootKeyCombo.ItemData(RootKeyCombo.NewIndex) = HKEY_CURRENT_CONFIG
    RootKeyCombo.ListIndex = 1
    
    ValueTypeCombo.AddItem "String Value"
    ValueTypeCombo.ItemData(ValueTypeCombo.NewIndex) = REG_SZ
    ValueTypeCombo.AddItem "Binary Value"
    ValueTypeCombo.ItemData(ValueTypeCombo.NewIndex) = REG_BINARY
    ValueTypeCombo.AddItem "DWORD Value"
    ValueTypeCombo.ItemData(ValueTypeCombo.NewIndex) = REG_DWORD
    ValueTypeCombo.ListIndex = 0
    Label1.Caption = "Remote Registry Editor is an advanced tool for changing settings in remote system registry, which contains information about how remote computer runs."
End Sub

Private Sub Form_Resize()
Dim X As Single, Y As Single
    
    X = (Me.ScaleWidth - Frame1.Width) / 2
    Y = (Me.ScaleHeight - Frame1.Height) / 2
    Frame1.Move X, Y
    
End Sub

Private Sub NewValueBttn_Click()
Dim txtData As String, idx1 As Integer, idx2 As Integer

    If Trim(KeyNameTxt.Text) = "" Then
        MsgBox "Please Enter Value Name", vbExclamation
        KeyNameTxt.SetFocus
        Exit Sub
    End If
    idx1 = RootKeyCombo.ListIndex
    idx2 = ValueTypeCombo.ListIndex
    txtData = "AddValue#" & RootKeyCombo.Text & "#" & SubKeyTxt.Text & "#" & KeyNameTxt.Text & "#" & ValueTypeCombo.Text & "#" & ValueDataTxt.Text & "#"
    sckReg.SendData txtData
End Sub

Private Sub sckReg_Close()
    'MsgBox "Remote Connection Lost", vbInformation, "Remote Registry Editor"
    Unload Me
End Sub

Private Sub sckReg_DataArrival(ByVal bytesTotal As Long)
Dim txtData As String, Command As String
Dim Position As Integer
    
    sckReg.getData txtData
    Position = InStr(1, txtData, "#")
    Command = Left(txtData, Position - 1)
    txtData = Mid(txtData, Position + 1)
    
    Select Case Command
        Case "ERROR"
            MsgBox txtData, vbCritical
        Case "SUCCESS"
            MsgBox txtData, vbInformation
        Case "ShowValue"
            ValueDataTxt.Text = Left(txtData, InStr(1, txtData, "#") - 1)
            txtData = Mid(txtData, InStr(1, txtData, "#") + 1)
            ValueTypeCombo.Text = txtData
            MsgBox "Value Successfuly read."
    End Select
End Sub

Private Sub ShowValueBttn_Click()
Dim txtData As String, idx As Integer
    If Trim(KeyNameTxt.Text) = "" Then
        MsgBox "Pleae Enter Value Name", vbExclamation
        KeyNameTxt.SetFocus
        Exit Sub
    End If
    idx = RootKeyCombo.ListIndex
    txtData = "ShowValue#" & RootKeyCombo.Text & "#" & SubKeyTxt.Text & "#" & KeyNameTxt.Text & "#"
    sckReg.SendData txtData
End Sub
