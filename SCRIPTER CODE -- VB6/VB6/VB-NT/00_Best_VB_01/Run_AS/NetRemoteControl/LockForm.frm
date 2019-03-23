VERSION 5.00
Begin VB.Form LockForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "LockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Private Sub Form_Load()
    Me.Show
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Screen.Width, _
    Screen.Height, SWP_SHOWWINDOW Or SWP_NOOWNERZORDER
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SystemParametersInfo SPI_SCREENSAVERRUNNING, 0&, Null, _
    SPIF_UPDATEINIFILE
End Sub
