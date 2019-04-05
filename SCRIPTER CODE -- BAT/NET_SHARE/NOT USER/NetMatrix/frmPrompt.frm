VERSION 5.00
Begin VB.Form frmPrompt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Now scanning..."
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmPrompt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPrompt.frx":0442
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmPrompt.frx":04CB
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

