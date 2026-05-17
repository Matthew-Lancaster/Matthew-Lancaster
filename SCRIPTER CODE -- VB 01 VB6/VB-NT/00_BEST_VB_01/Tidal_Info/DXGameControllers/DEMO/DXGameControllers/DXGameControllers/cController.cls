VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Name:    cController
' Author:  Michael A Seel
' Date:    Feb 2004
'
' Note regarding analog controls:
' To keep this simple (like my gamepad;), I have reduced analog values to
' simple on/off switches. Therefore, sliders, POV's, analog sticks, etc.,
' won't work as you expect if your expecting full-on analog enjoyment.
' My code doesn't employ this because my gamepad is really, really simple
' and I would have no way to test it.
' However, this would be easy enough to implement by using the DIJOYSTATE
' data type instead of Boolean arrays, and by replacing the existing events
' with one that simply passes the current state along with a boolean array
' indicating where changes have occurred.


Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private mDX As DirectX8  'We need to retain the instance on a module-level for events to work
Private mPar As cControllerList  '(Optional) The parent list
Private mDvc As DirectInputDevice8  'A reference to the controller itself
Private mFrm As Form  'The owner form and (optional) object to receive event notification
Private mEvt As Long  '(Optional) The EventID for the controller
Private mFoc As Boolean  'Flag set to ignore events when form doesn't have focus
Private mSte() As Boolean  'The last retrieved state of the controller
Private mChg As Boolean  'Flag set when change events are firing

' The following events are for single WithEvents declarations of cController
Public Event ElementChange(ByVal Element As ControllerElement, ByVal Pressed As Boolean)
Public Event ChangeBegin()
Public Event ChangeEnd()


Public Function Initialize(ByVal GUIDInstance As String, ByVal EventForm As Form, Optional ByVal FocusEventsOnly As Boolean, Optional Parent As cControllerList) As Boolean
' GUIDInstance: The GUID of the controller
' EventForm: The owner form / object to receive event notification (if implemented)
   If Not (EventForm Is Nothing) Then  'This must point to an actual instance of a form
      Me.Terminate  'Clear everything first
      Set mDX = New DirectX8  'Create an instance of DirectX8 to keep until Terminate occurs
      On Error Resume Next
      ' Try to create an instance for the device
      Set mDvc = mDX.DirectInputCreate.CreateDevice(GUIDInstance)
      On Error GoTo 0
      If mDvc Is Nothing Then  'Did that work?
         Set mDX = Nothing  'DirectX8 reference won't be needed
      Else
         Set mFrm = EventForm  'Remember our owner/event form
         mFoc = FocusEventsOnly  'Remember whether to ignore background events or not
         mDvc.SetCommonDataFormat DIFORMAT_JOYSTICK  'Set the data format we will use later
         mDvc.SetCooperativeLevel mFrm.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE  'System-wide
         On Error Resume Next  'If it doesn't work, they must want to use a Timer or something ;)
         ' Try to use the form as an event
         mEvt = mDX.CreateEvent(mFrm)
         On Error GoTo 0
         If mEvt <> 0 Then mDvc.SetEventNotification mEvt  'Set the event, if created
         mDvc.Acquire  'Connect the reference to the device itself
         Initialize = True  'Return that the initialization worked
         mSte = GetState  'Get the current state of the device
         Set mPar = Parent  'Remember our parent, if there is one
      End If
   End If
End Function

Public Sub Terminate()
   If mEvt <> 0 Then  'If an event exists...
      On Error Resume Next
      mDX.DestroyEvent mEvt  '...destroy it...
      On Error GoTo 0
      mEvt = 0  '...and clear the value.
   End If
   Set mDX = Nothing  'Discard our DirectX8 reference, since we don't need it anymore.
   If Not (mDvc Is Nothing) Then  'If a device reference exists...
      mDvc.Unacquire  '...disconnect from the device...
      Set mDvc = Nothing  '...and discard the reference.
   End If
   mFoc = False  'Clear the flag
   mChg = False  'And this one, too
   ReDim mSte(1 To 56)  'Oh... and this too.
   Set mPar = Nothing  'Discard the parent reference
   Set mFrm = Nothing  'Discard the form reference
End Sub


Public Property Get IsInit() As Boolean
   IsInit = Not (mDvc Is Nothing)  'If there is no reference, there was no initialization
End Property

Public Property Get Name() As String
   If Not (mDvc Is Nothing) Then Name = mDvc.GetDeviceInfo.GetProductName  'Return the device's name
End Property

Public Property Get GUID() As String
   If Not (mDvc Is Nothing) Then GUID = mDvc.GetDeviceInfo.GetGuidInstance  'Return the device's GUID
End Property

Public Property Get EventForm() As Form
   Set EventForm = mFrm  'Return the owner/event form reference
End Property

Public Property Get EventID() As Long
   EventID = mEvt  'Return the EventID value
End Property

Public Property Let FocusEventsOnly(ByVal vData As Boolean)
   If IsInit Then mFoc = vData  'Set the flag for ignoring background events
End Property
Public Property Get FocusEventsOnly() As Boolean
   FocusEventsOnly = mFoc  'Return the flag for ignoring background events
End Property

Public Property Get Parent() As cControllerList
   Set Parent = mPar  'Return our parent list reference
End Property

Public Property Get Changing() As Boolean
   Changing = mChg  'Return the flag for if changes are taking place
End Property


Public Sub CheckForChanges()
   Dim bSte() As Boolean  'Buffer for current state
   Dim bChg() As Boolean  'Buffer for state changes
   Dim bEvt As Boolean  'Flag set if events should be fired
   Dim i As Integer
   
   If Not ((mDvc Is Nothing) Or mChg) Then  'If there is a device and we aren't already firing events...
      bSte = GetState  'Get the current state
      If StateChanged(mSte, bSte, bChg) Then  'If any changes occurred since the last state...
         If mFoc Then  'If we are ignoring background events...
            bEvt = (GetForegroundWindow = mFrm.hWnd)  'Then fire events if the form has focus
         Else  'If not...
            bEvt = True  'Fire events regardless
         End If
         If bEvt Then
            mChg = True  'Set our flag to indicate change events are firing
            RaiseEvent ChangeBegin  'Let single references know changes are coming
            If Not (mPar Is Nothing) Then mPar.zChangeBegin Me  'And let the parent know
            For i = 1 To 56  'Loop through all the controller elements
               If bChg(i) Then  'If this element has changed...
                  RaiseEvent ElementChange(i, bSte(i))  'Let references know what changed and how
                  If Not (mPar Is Nothing) Then mPar.zElementChange Me, i, bSte(i)  'And the parent, too
               End If
            Next
            RaiseEvent ChangeEnd  'Let references know the changes are all sent
            If Not (mPar Is Nothing) Then mPar.zChangeEnd Me  'And the parent, yeah
            mChg = False  'Clear our flag so we can do this again next time
         End If
         mSte = bSte  'Remember the current state
         Erase bChg  'Clear the changes buffer
      End If
      Erase bSte  'Clear the state buffer
   End If
End Sub

Private Function GetState() As Boolean()
   Dim vSte As DIJOYSTATE  'Data buffer of the format we specified during initialization
   Dim bSte() As Boolean  'Our proprietary state buffer
   Dim i As Integer
   
   ReDim bSte(1 To 56)  'Iniatilize the buffer for all elements
   If Not (mDvc Is Nothing) Then  'If we have the device...
      mDvc.GetDeviceStateJoystick vSte  'Get it's state
      With vSte
         For i = 1 To 32  'Loop through the buttons
            bSte(i) = CBool(vSte.Buttons(i - 1) And 128)  '0=Unpressed, 128=Pressed
         Next
         Select Case Round(.x / 32768)  '0 to 65535 = Left to Right
            Case 0 'Left
               bSte(ceXNeg) = True
            Case 2 'Right
               bSte(ceXPos) = True
         End Select
         Select Case Round(.y / 32768)  '0 to 65535 = Up to Down
            Case 0 'Up
               bSte(ceYNeg) = True
            Case 2 'Down
               bSte(ceYPos) = True
         End Select
         Select Case Round(.z / 32768)  '0 to 65535 = Out to In
            Case 0 'Out
               bSte(ceZNeg) = True
            Case 2 'In
               bSte(ceZPos) = True
         End Select
         Select Case Round(.rx / 32768)  '0 to 65535 = Left to Right
            Case 0
               bSte(ceRXNeg) = True
            Case 2
               bSte(ceRXPos) = True
         End Select
         Select Case Round(.ry / 32768)  '0 to 65535 = Up to Down
            Case 0
               bSte(ceRYNeg) = True
            Case 2
               bSte(ceRYPos) = True
         End Select
         Select Case Round(.rz / 32768)  '0 to 65535 = Out to In
            Case 0
               bSte(ceRZNeg) = True
            Case 2
               bSte(ceRZPos) = True
         End Select
         For i = 0 To 3
            Select Case Round(.POV(i) / 32768)  '0 to 65535 = ? to ?
               Case 0
                  bSte(cePOV0Neg + (i * 2)) = True
               Case 2
                  bSte(cePOV0Pos + (i * 2)) = True
            End Select
         Next
         For i = 0 To 1
            Select Case Round(.slider(i) / 32768)  '0 to 65535 = ? to ?
               Case 0 'Min
                  bSte(ceSlider0Neg + (i * 2)) = True
               Case 2 'Max
                  bSte(ceSlider0Pos + (i * 2)) = True
            End Select
         Next
      End With
   End If
   GetState = bSte
   Erase bSte
End Function

Private Function StateChanged(State1() As Boolean, State2() As Boolean, StateChange() As Boolean) As Boolean
   Dim i As Integer
   ReDim StateChange(1 To 56)  'Initialize the return buffer
   For i = 1 To 56  'Loop through each element
      If State1(i) <> State2(i) Then  'If the states are different...
         StateChange(i) = True  'Set the flag in the return buffer
         StateChanged = True  'Return that there is a change
      End If
   Next
End Function


Private Sub Class_Initialize()
   ReDim mSte(1 To 56)  'Initialize the state buffer
End Sub

Private Sub Class_Terminate()
   Me.Terminate  'Clear everything up
   Erase mSte  'Deallocate the state buffer
End Sub

