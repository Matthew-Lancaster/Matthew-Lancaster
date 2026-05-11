VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Email Domain Collector"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4245
   LinkTopic       =   "Form2"
   ScaleHeight     =   1785
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   420
      Left            =   2205
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   420
      Left            =   180
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1230
      Left            =   2205
      TabIndex        =   1
      Top             =   450
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1230
      Left            =   180
      TabIndex        =   0
      Top             =   450
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
