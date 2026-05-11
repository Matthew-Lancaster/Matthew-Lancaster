VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSck 
   Caption         =   "frmSck"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   2250
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock WinSck 
      Left            =   840
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Form included as a container for the Winsock control.
'
' The Class module instantiates the Winsock Control
' WithEvents and handles all events locally.
'
' Dynamically loading the Winsock Control at run time
' without utilizing a container Form is supported by
' the Class module, however there are currently unresolved
' deployment issues when using the 'formless' method.
'
' See the notation in the Sub Class_Initialize() in
' the class module for additional information.
'
