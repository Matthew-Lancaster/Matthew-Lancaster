VERSION 5.00
Begin VB.Form frmFilterParams 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filter Parameters"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnDisable 
      BackColor       =   &H80000014&
      Caption         =   "Disable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.Frame fraFilterParams 
      Caption         =   "Parameters"
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3015
      Begin VB.CheckBox chkEnabled 
         Alignment       =   1  'Right Justify
         Caption         =   "Enabled"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   390
         Value           =   1  'Checked
         Width           =   1200
      End
      Begin VB.TextBox txtParam 
         Height          =   285
         Index           =   4
         Left            =   1170
         TabIndex        =   7
         Text            =   "0"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtParam 
         Height          =   285
         Index           =   3
         Left            =   1170
         TabIndex        =   6
         Text            =   "0"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtParam 
         Height          =   285
         Index           =   2
         Left            =   1170
         TabIndex        =   5
         Text            =   "0"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtParam 
         Height          =   285
         Index           =   1
         Left            =   1170
         TabIndex        =   4
         Text            =   "0"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtParam 
         Height          =   285
         Index           =   0
         Left            =   1170
         TabIndex        =   3
         Text            =   "0"
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblParam 
         Caption         =   "Param 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   200
         TabIndex        =   14
         Top             =   2190
         Width           =   735
      End
      Begin VB.Label lblParam 
         Caption         =   "Param 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   200
         TabIndex        =   13
         Top             =   1830
         Width           =   735
      End
      Begin VB.Label lblParam 
         Caption         =   "Param 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   200
         TabIndex        =   12
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label lblParam 
         Caption         =   "Param 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   200
         TabIndex        =   11
         Top             =   1110
         Width           =   735
      End
      Begin VB.Label lblParam 
         Caption         =   "Param 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   0
         Left            =   195
         TabIndex        =   2
         Top             =   750
         Width           =   735
      End
   End
   Begin VB.Label lblFilterName 
      Caption         =   "< Filter Name >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmFilterParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()

frmFilterParams.Tag = False
frmFilterParams.Hide

End Sub

Private Sub btnOK_Click()

frmFilterParams.Tag = True
frmFilterParams.Hide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Cancel = 1
Me.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)

Cancel = 1
Me.Hide

End Sub
