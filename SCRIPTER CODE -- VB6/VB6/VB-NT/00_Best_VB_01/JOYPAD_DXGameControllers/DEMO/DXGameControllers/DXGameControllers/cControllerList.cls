VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cControllerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Name:    cControllerList
' Author:  Michael A Seel
' Date:    Feb 2004

Option Explicit

Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Private mCol As Collection  'Connected controllers
Private mFrm As Form  'The owner form and (optional) object to receive event notification
Private mFoc As Boolean  'Flag set to ignore background events

Public Event ElementChange(ByVal Index As Integer, ByVal Element As ControllerElement, ByVal Pressed As Boolean)
Public Event ChangeBegin(ByVal Index As Integer)
Public Event ChangeEnd(ByVal Index As Integer)


Public Sub Initialize(ByVal EventForm As Form, Optional ByVal FocusEventsOnly As Boolean)
' EventForm: The owner form / object to receive event notification (if implemented)
   If Not (EventForm Is Nothing) Then  'If an owner is specified...
      Set mFrm = EventForm  'Remember it
      mFoc = FocusEventsOnly  'Set the background events flag
      RescanForDevices  'Scan for devices
   End If
End Sub

Public Sub RescanForDevices()
   Dim oDX As DirectX8  'DirectX8 instance
   Dim oEnm As DirectInputEnumDevices8  'DirectX8 input device list instance
   Dim i As Integer
   Dim oTmp As cController  'Temporary controller object
   
   For Each oTmp In mCol  'Loop through our list of controllers...
      oTmp.Terminate  'And terminate each one
   Next
   Set mCol = Nothing  'Discard the current list
   Set mCol = New Collection  'Reinitialize our list
   If Not (mFrm Is Nothing) Then  'If we still have an owner form...
      Set oDX = New DirectX8  'Create an instance of DirectX8
      Set oEnm = oDX.DirectInputCreate.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY)  'Get a list of attached game controllers
      For i = 1 To oEnm.GetCount  'Loop through each device entry
         Set oTmp = New cController  'Create an new instance of our controller class
         If oTmp.Initialize(oEnm.GetItem(i).GetGuidInstance, mFrm, mFoc, Me) Then  'If the GUID initializes the object...
            mCol.Add oTmp  'Copy the instance to our list
         End If
         Set oTmp = Nothing  'Discard the temporary instance
      Next
      Set oEnm = Nothing  'Discard the device list
      Set oDX = Nothing  'Discard the DirectX8 instance
   End If
End Sub

Public Sub CheckForEvents(Optional ByVal EventID As Long)
'EventID: Specifies the eventid of the controller that fired the event
'If events were not used (i.e., DirectXEvent8 wasn't implemented on the
'owner form) then every controller's EventID will be '0'.
   Dim oTmp As cController
   For Each oTmp In mCol  'Loop through our controllers
      If oTmp.EventID = EventID Then oTmp.CheckForChanges  'If the EventID's match, check for changes
   Next
End Sub

Public Sub Terminate()
   Dim oDX As DirectX8  'DirectX8 instance
   Dim oTmp As cController
   
   For Each oTmp In mCol  'Loop through each controller...
      oTmp.Terminate  'And terminate the connection
   Next
   Set mCol = Nothing  'Clear our list
   Set mCol = New Collection  'Reinitialize our list
   Set mFrm = Nothing  'Clear the owner form reference
   mFoc = False  'Clear the ignore background events flag
End Sub

Public Function Count() As Integer
   Count = mCol.Count  'Return the controller count
End Function

Public Function Item(ByVal Index As Integer) As cController
   Set Item = mCol(Index)  'Return the specified controller instance
End Function

Public Function Find(ByVal GUID As String) As Integer
'This is useful in determining if a specific game device is
'attached or not, and which controller number it is
   Dim i As Integer
   For i = 1 To mCol.Count  'Loop through our controllers
      If mCol(i).GUID = GUID Then  'If the GUID's match...
         Find = i  'Return our index
         Exit For  'Exit the loop, since the GUID should be unique
      End If
   Next
End Function


Public Sub zElementChange(Controller As cController, ByVal Element As ControllerElement, ByVal Pressed As Boolean)
'Locate and fire events for the child controller
   Dim iDx As Integer
   iDx = Locate(Controller)
   If iDx > 0 Then RaiseEvent ElementChange(iDx, Element, Pressed)
End Sub

Public Sub zChangeBegin(Controller As cController)
'Locate and fire events for the child controller
   Dim iDx As Integer
   iDx = Locate(Controller)
   If iDx > 0 Then RaiseEvent ChangeBegin(iDx)
End Sub

Public Sub zChangeEnd(Controller As cController)
'Locate and fire events for the child controller
   Dim iDx As Integer
   iDx = Locate(Controller)
   If iDx > 0 Then RaiseEvent ChangeEnd(iDx)
End Sub


Private Function Locate(Controller As cController) As Integer
'Internal. Used by the z* events called by the child controllers
   Dim i As Integer
   For i = 1 To mCol.Count  'Loop through each controller
      If mCol(i) Is Controller Then  'If it matches the Controller argument...
         Locate = i  'Return its index
         Exit For  'Exit the loop, since the same reference should be listed only once
      End If
   Next
End Function


Private Sub Class_Initialize()
   Set mCol = New Collection  'Initialize our list
End Sub

Private Sub Class_Terminate()
   Me.Terminate  'Clear everything up
   Set mCol = Nothing  'Discard our list
End Sub
