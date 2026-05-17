VERSION 5.00
Begin VB.Form FRM_zz_DEBUG_VAR_TIMER_SET 
   Caption         =   "FORM_DEBUG_VAR_TIMER_SET"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   480
   ClientWidth     =   12690
   LinkTopic       =   "Form2"
   ScaleHeight     =   7755
   ScaleWidth      =   12690
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "DEBUG_VAR_TIMER_SET.frx":0000
      Top             =   0
      Width           =   12060
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6300
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "DEBUG_VAR_TIMER_SET.frx":0008
      Top             =   1020
      Width           =   12150
   End
End
Attribute VB_Name = "FRM_zz_DEBUG_VAR_TIMER_SET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    'Exit Sub
    If IsIDE = True Then End
End Sub

