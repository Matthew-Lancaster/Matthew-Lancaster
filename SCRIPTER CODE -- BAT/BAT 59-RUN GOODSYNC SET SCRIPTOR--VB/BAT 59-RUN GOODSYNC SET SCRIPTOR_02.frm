VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   Icon            =   "BAT 59-RUN GOODSYNC SET SCRIPTOR_02.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : Form1
'    Project    : BAT 45-SCRIPT RUN GITHUB
'
'    Description: BAT 45-SCRIPT RUN GITHUB
'
'    Author     : Matthew Lanacster _ Matt.Lan@Btinternet.com
'    Modified   : Sunday 04:05:20 Pm_07 October 2018
'--------------------------------------------------------------------------------

' WANT ICON AND THEN HAVE A FORM

Private Sub Form_Load()

Call Module1.Main

Unload Me

End Sub
