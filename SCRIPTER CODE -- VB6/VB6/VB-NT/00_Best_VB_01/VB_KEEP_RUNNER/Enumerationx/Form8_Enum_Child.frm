VERSION 5.00
Begin VB.Form Form8_Enum_Child 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patch"
   ClientHeight    =   3375
   ClientLeft      =   3465
   ClientTop       =   3780
   ClientWidth     =   2475
   LinkTopic       =   "Form8_Enum_Child"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   2475
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Write the text to be replaced below:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Write the handle of the window whose text you want to change:"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form8_Enum_Child"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------
'           Author: Muhammad Abubakar
'           http://go.to/abubakar
'           <joehacker@yahoo.com>
'------------------------------------
Option Explicit

Private Sub Command1_Click()
    Unload Me
    
End Sub

Private Sub Command2_Click()
    SendMessage Val(Text1.Text), WM_SETTEXT, 0, ByVal Text2.Text
    Unload Me
    
End Sub

