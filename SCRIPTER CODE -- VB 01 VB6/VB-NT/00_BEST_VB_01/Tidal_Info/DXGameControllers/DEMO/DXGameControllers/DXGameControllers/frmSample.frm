VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DXGameControllers"
   ClientHeight    =   3336
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8784
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3336
   ScaleWidth      =   8784
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFoc 
      Caption         =   "Ignore Background Events"
      Height          =   375
      Left            =   4980
      TabIndex        =   3
      Top             =   2880
      Width           =   2355
   End
   Begin VB.ListBox lboEvt 
      Height          =   2736
      ItemData        =   "frmSample.frx":0000
      Left            =   4980
      List            =   "frmSample.frx":0002
      TabIndex        =   2
      Top             =   60
      Width           =   3735
   End
   Begin VB.CommandButton cmdScn 
      Caption         =   "Rescan For Devices"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   2355
   End
   Begin VB.ListBox lboLst 
      Height          =   2736
      ItemData        =   "frmSample.frx":0004
      Left            =   120
      List            =   "frmSample.frx":0006
      TabIndex        =   0
      Top             =   60
      Width           =   4755
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Name:    frmSample
' Author:  Michael A Seel
' Date:    Feb 2004

Option Explicit

Implements DirectXEvent8  'Allow form to receive device event notification

Private WithEvents mLst As cControllerList  'Our controller list
Attribute mLst.VB_VarHelpID = -1
Private mFoc As Boolean  'True = Allow background events (needed only for this sample)


Private Sub cmdFoc_Click()
   Dim i As Integer
   mFoc = Not mFoc
   If mFoc Then
      cmdFoc.Caption = "Allow Background Events"
   Else
      cmdFoc.Caption = "Ignore Background Events"
   End If
   For i = 1 To mLst.Count
      mLst.Item(i).FocusEventsOnly = mFoc
   Next
End Sub

Private Sub cmdScn_Click()
   lboLst.Clear
   lboEvt.Clear
   Me.Refresh
   mLst.RescanForDevices  'Rescan
   RefreshList
End Sub

Private Sub DirectXEvent8_DXCallback(ByVal EventID As Long)
'EventID is specific to each device if DirectXEvent8 is implemented on the "EventForm"
'If it is not implemented, all devices will have an EventID of '0', and this procedure
'is never called.
   mLst.CheckForEvents EventID  'Check for events on device with EventID
End Sub

Private Sub Form_Load()
   ChDir (App.Path)
   ChDrive (App.Path)
   Me.Show
   DoEvents
   Set mLst = New cControllerList  'Create an instance of the controller list.
   mLst.Initialize Me, False  'Initialize the list using this form as the owner.
   RefreshList  'Update our controller display
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mLst.Terminate  'Clean up the controller list
   Set mLst = Nothing  'Discard the controller list
End Sub

Private Sub RefreshList()
   Dim i As Integer
   lboLst.Clear
   For i = 1 To mLst.Count  'Loop through each controller...
      With mLst.Item(i)
         lboLst.AddItem CStr(i) & ": " & .Name & " " & .GUID  '...and list some info
      End With
   Next
End Sub

Private Sub mLst_ElementChange(ByVal Index As Integer, ByVal Element As ControllerElement, ByVal Pressed As Boolean)
' Index: The number of the controller
' Element: Which element changed
' Pressed: The state of the element
   If Pressed Then
      lboEvt.AddItem "Controller " & CStr(Index) & ": Pressed " & ElementName(Element)
   Else
      lboEvt.AddItem "Controller " & CStr(Index) & ": Released " & ElementName(Element)
   End If
   lboEvt.ListIndex = lboEvt.NewIndex
   lboEvt.TopIndex = lboEvt.NewIndex
End Sub
