VERSION 5.00
Begin VB.Form Form2_ANY_SPECIAL_FOLDER 
   Caption         =   "Jump Any Special Folder"
   ClientHeight    =   5664
   ClientLeft      =   228
   ClientTop       =   552
   ClientWidth     =   9672
   LinkTopic       =   "Form2"
   ScaleHeight     =   5664
   ScaleWidth      =   9672
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   4200
      Top             =   2880
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4368
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   9615
   End
   Begin VB.Label Label1 
      Caption         =   "Press Jump Any Special Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "Form2_ANY_SPECIAL_FOLDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
