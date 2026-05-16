VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IE_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Collection1"
Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Declare Function GetAsyncKeyState% Lib "user32" (ByVal vKey As Long)
Private Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Private Declare Function ScreenToClient& Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI)
Private Declare Function GetKeyState% Lib "user32" (ByVal nVirtKey As Long)

Private WithEvents IE As SHDocVw.InternetExplorer
Attribute IE.VB_VarHelpID = -1
Private WithEvents IEDoc As MSHTML.HTMLDocument
Attribute IEDoc.VB_VarHelpID = -1
Private mvarIEHandle As Long 'local copy
Private mvarEnableBack As Boolean 'local copy
Private mvarEnableForward As Boolean 'local copy

Private bLDown As Boolean, bRDown As Boolean

Public Property Get EnableForward() As Boolean
   EnableForward = mvarEnableForward
End Property

Public Property Get EnableBack() As Boolean
   EnableBack = mvarEnableBack
End Property

Public Property Set IEctl(IncomeIE As SHDocVw.InternetExplorer)
   Set IE = IncomeIE
   mvarIEHandle = IncomeIE.hwnd
   On Error Resume Next
   Set IEDoc = IncomeIE.Document
   bCancel = True
   Err.Clear
   IE.GoForward
   If Err Then
      mvarEnableForward = False
      bCancel = False
   Else
      IE.ExecWB OLECMDID_STOPDOWNLOAD, OLECMDEXECOPT_DONTPROMPTUSER
   End If
   On Error GoTo 0
End Property

Public Property Get IEctl() As SHDocVw.InternetExplorer
   Set IEctl = IE
End Property

Public Property Get IEHandle() As Long
    IEHandle = mvarIEHandle
End Property

Private Sub IE_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    CallEvent ID_BeforeNavigate, mvarIEHandle, pDisp, URL, Flags, TargetFrameName, PostData, Headers, Cancel
    If bCancel Then
       Cancel = True
       bCancel = False
    End If
End Sub

Private Sub IE_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
   Select Case Command
     Case CSC_NAVIGATEBACK
          mvarEnableBack = Enable
     Case CSC_NAVIGATEFORWARD
          mvarEnableForward = Enable
   End Select
   CallEvent ID_CommandStateChange, mvarIEHandle, Command, Enable
End Sub

Private Sub IE_DocumentComplete(ByVal pDisp As Object, URL As Variant)
  If pDisp Is IE Then
     CallEvent ID_DocumentComplete, mvarIEHandle, pDisp, URL
  End If
End Sub

Private Sub IE_DownloadBegin()
  CallEvent ID_DownloadBegin, mvarIEHandle
End Sub

Private Sub IE_DownloadComplete()
  CallEvent ID_DownloadComplete, mvarIEHandle
End Sub

Private Sub IE_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
  CallEvent ID_NavigationComplete, mvarIEHandle, pDisp, URL
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  Set IE = Nothing
End Sub

Private Function IEDoc_oncontextmenu() As Boolean
   CallEvent ID_ContextMenu, mvarIEHandle
   IEDoc_oncontextmenu = Not bCancel
   If bCancel = True Then bCancel = False
End Function

Private Sub IEDoc_onmousedown()
   Dim pt As POINTAPI
   Dim btn As Integer, i As Integer
   GetCursorPos pt
   ScreenToClient mvarIEHandle, pt
   If GetAsyncKeyState(vbKeyLButton) < 0 Then
      btn = 1: bLDown = True
   Else
      btn = 2: bRDown = True
   End If
   If GetAsyncKeyState(vbKeyShift) Then i = 1
   If GetAsyncKeyState(vbKeyControl) Then i = 2
   If GetAsyncKeyState(vbKeyMenu) Then i = 4
   CallEvent ID_MouseDown, mvarIEHandle, btn, i, CSng(pt.X), CSng(pt.Y)
End Sub

Private Sub IEDoc_onmouseup()
   Dim pt As POINTAPI
   Dim btn As Integer, i As Integer
   GetCursorPos pt
   ScreenToClient mvarIEHandle, pt
   If bLDown Then
      btn = 1: bLDown = False
   Else
      btn = 2: bRDown = False
   End If
   If GetAsyncKeyState(vbKeyShift) Then i = 1
   If GetAsyncKeyState(vbKeyControl) Then i = 2
   If GetAsyncKeyState(vbKeyMenu) Then i = 4
   CallEvent ID_MouseUp, mvarIEHandle, btn, i, CSng(pt.X), CSng(pt.Y)
End Sub
