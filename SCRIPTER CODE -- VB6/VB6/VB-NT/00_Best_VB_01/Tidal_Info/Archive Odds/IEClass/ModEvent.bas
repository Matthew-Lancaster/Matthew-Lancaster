Attribute VB_Name = "Module1"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
Public cIEWPtr As Long, bCancel As Boolean
Public Enum IDEVENTS
   ID_BeforeNavigate = 1
   ID_NavigationComplete = 2
   ID_DownloadBegin = 3
   ID_DownloadComplete = 4
   ID_DocumentComplete = 5
   ID_MouseDown = 6
   ID_MouseUp = 7
   ID_ContextMenu = 8
   ID_CommandStateChange = 9
End Enum

Public Function CallEvent(nEvent As IDEVENTS, hwnd As Long, ParamArray EventInfo())
   Select Case nEvent
          Case ID_BeforeNavigate
               ResolvePointer(cIEWPtr).FireEvent nEvent, hwnd, EventInfo(0), EventInfo(1), EventInfo(2), EventInfo(3), EventInfo(4), EventInfo(5), CBool(EventInfo(6))
          Case ID_NavigationComplete, ID_DocumentComplete, ID_CommandStateChange
               ResolvePointer(cIEWPtr).FireEvent nEvent, hwnd, EventInfo(0), EventInfo(1)
          Case ID_MouseDown, ID_MouseUp
               ResolvePointer(cIEWPtr).FireEvent nEvent, hwnd, EventInfo(0), EventInfo(1), EventInfo(2), EventInfo(3)
          Case Else
               ResolvePointer(cIEWPtr).FireEvent nEvent, hwnd
   End Select
End Function

Private Function ResolvePointer(ByVal lpObj&) As cIEWindows
  Dim oIEW As cIEWindows
  CopyMemory oIEW, lpObj, 4&
  Set ResolvePointer = oIEW
  CopyMemory oIEW, 0&, 4&
End Function

