VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cIEWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"IE_Class"
Attribute VB_Ext_KEY = "Member0" ,"IE_Class"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private WithEvents winShell As SHDocVw.ShellWindows
Attribute winShell.VB_VarHelpID = -1
Private mCol As Collection
Dim bRefreshing As Boolean

Public Event IEWindowRegistered()
Public Event IEWindowRevoked()
Public Event IENavigationBegin(hwnd As Long, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Public Event IENavigationComplete(hwnd As Long, ByVal pDisp As Object, URL As Variant)
Public Event IEDocumentComplete(hwnd As Long, ByVal pDisp As Object, URL As Variant)
Public Event IEDownloadBegin(hwnd As Long)
Public Event IEDownloadComplete(hwnd As Long)
Public Event IEOnContextMenu(hwnd As Long)
Public Event IEMouseDown(hwnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event IEMouseUp(hwnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event IECommandStateChange(hwnd As Long, Button As CommandStateChangeConstants, Enable As Boolean)

Private Function Add(IEctl As SHDocVw.InternetExplorer) As IE_Class
    Dim objNewMember As IE_Class
    Set objNewMember = New IE_Class
    Set objNewMember.IEctl = IEctl
    mCol.Add objNewMember, CStr(objNewMember.IEHandle)
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get IE(vntIndexKey As Variant) As IE_Class
Attribute IE.VB_UserMemId = 0
  Do While bRefreshing
     DoEvents
  Loop
  Set IE = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
   Count = mCol.Count
End Property

Private Sub Remove(vntIndexKey As Variant)
   mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
  cIEWPtr = ObjPtr(Me)
  Refresh_Col
End Sub

Private Sub Class_Terminate()
   Set mCol = Nothing
   Set winShell = Nothing
End Sub

Private Sub winShell_WindowRegistered(ByVal lCookie As Long)
  Refresh_Col
  Do While bRefreshing
     DoEvents
  Loop
  RaiseEvent IEWindowRegistered
End Sub

Private Sub winShell_WindowRevoked(ByVal lCookie As Long)
  Refresh_Col
  Do While bRefreshing
     DoEvents
  Loop
  RaiseEvent IEWindowRevoked
End Sub

Private Sub Refresh_Col()
   bRefreshing = True
   Dim SWs As New SHDocVw.ShellWindows
   Dim var As SHDocVw.InternetExplorer
   Set mCol = Nothing
   Set mCol = New Collection
   For Each var In SWs
       Add var
   Next
   Set winShell = SWs
   Set SWs = Nothing
   Set var = Nothing
   bRefreshing = False
End Sub

Friend Function FireEvent(nEvent As IDEVENTS, hwnd As Long, ParamArray EventInfo())
   Select Case nEvent
          Case ID_BeforeNavigate
               RaiseEvent IENavigationBegin(hwnd, EventInfo(0), EventInfo(1), EventInfo(2), EventInfo(3), EventInfo(4), EventInfo(5), CBool(EventInfo(6)))
          Case ID_NavigationComplete
               RaiseEvent IENavigationComplete(hwnd, EventInfo(0), EventInfo(1))
          Case ID_DocumentComplete
               RaiseEvent IEDocumentComplete(hwnd, EventInfo(0), EventInfo(1))
          Case ID_DownloadBegin
               RaiseEvent IEDownloadBegin(hwnd)
          Case ID_DownloadComplete
               RaiseEvent IEDownloadComplete(hwnd)
          Case ID_ContextMenu
               RaiseEvent IEOnContextMenu(hwnd)
          Case ID_MouseDown
               RaiseEvent IEMouseDown(hwnd, CInt(EventInfo(0)), CInt(EventInfo(1)), CSng(EventInfo(2)), CSng(EventInfo(3)))
          Case ID_MouseUp
               RaiseEvent IEMouseUp(hwnd, CInt(EventInfo(0)), CInt(EventInfo(1)), CSng(EventInfo(2)), CSng(EventInfo(3)))
          Case ID_CommandStateChange
               RaiseEvent IECommandStateChange(hwnd, CLng(EventInfo(0)), CBool(EventInfo(1)))
   End Select
End Function

