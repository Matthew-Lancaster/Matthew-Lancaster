VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIECtrl 
   Caption         =   "IE Control"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   6
      Left            =   5520
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   6000
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   6615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   5
      Left            =   5520
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   4
      Left            =   5520
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   2
      Left            =   5520
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":015C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4895
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmIECtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetForegroundWindow& Lib "user32" (ByVal hwnd As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function PostMessage& Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any)

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOREDRAW = &H8

Const WM_CLOSE = &H10
Const LB_FINDSTRING = &H18F

Dim WithEvents IEWin As cIEWindows
Attribute IEWin.VB_VarHelpID = -1
Dim xClick As Single, yClick As Single
Dim SurfingTime As Date

Private Sub Command1_Click(Index As Integer)
  Select Case Index
      Case 0
         If Text1 <> "" Then IEWin.IE(Mid$(LV1.SelectedItem.Key, 2)).IEctl.Navigate Text1
      Case 1
         IEWin.IE(Mid$(LV1.SelectedItem.Key, 2)).IEctl.GoHome
      Case 2
         IEWin.IE(Mid$(LV1.SelectedItem.Key, 2)).IEctl.GoBack
      Case 3
         IEWin.IE(Mid$(LV1.SelectedItem.Key, 2)).IEctl.GoForward
      Case 4
         IEWin.IE(Mid$(LV1.SelectedItem.Key, 2)).IEctl.GoSearch
      Case 5
         IEWin.IE(Mid$(LV1.SelectedItem.Key, 2)).IEctl.Refresh
      Case 6
         IEWin.IE(Mid$(LV1.SelectedItem.Key, 2)).IEctl.Quit
  End Select
End Sub

Private Sub Command2_Click()
  If Command2.Caption = "Show forbidden URLs" Then
     Command2.Caption = "Hide forbidden URLs"
     Height = Height - ScaleHeight + Command3.Top + Command3.Height + 60
  Else
     Command2.Caption = "Show forbidden URLs"
     Height = Height - ScaleHeight + Command2.Top + Command2.Height + 60
  End If
End Sub

Private Sub Command3_Click()
   Dim sTemp As String
   sTemp = InputBox("Enter restricted URL", "Adding to forbidden list", "http://")
   If sTemp <> "" Then List1.AddItem sTemp
End Sub

Private Sub Command4_Click()
  If List1.ListIndex > -1 Then List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub Form_Load()
  Timer1.Enabled = False
  Caption = "Total surfing time = " & TimeSerial(0, 0, SurfingTime)
  Text1 = "http://www.freevbcode.com"
  Command2_Click
  Command1(0).Caption = "Navigate"
  Command1(1).Caption = "Go Home"
  Command1(2).Caption = "Go Back"
  Command1(3).Caption = "Go Forward"
  Command1(4).Caption = "Go Search"
  Command1(5).Caption = "Update"
  Command1(6).Caption = "Close"
  Command3.Caption = "Add to list"
  Command4.Caption = "Remove from list"
  Set IEWin = New cIEWindows
  Set LV1.SmallIcons = ImageList1
  With LV1
    .View = lvwReport
    .HideSelection = False
    .FullRowSelect = True
    .LabelEdit = lvwManual
    .ColumnHeaders.Add , , "Current IE windows"
    .ColumnHeaders.Add , , "Current URL"
    .ColumnHeaders.Add , , "Status"
    .ColumnHeaders(1).Width = 2500
    .ColumnHeaders(3).Width = 1000
  End With
  List1.AddItem "http://www.freevbcode.com/"
  List1.AddItem "http://www.altavista.com/"
  UpdateIEList LV1
  Window_SetAlwaysOnTop False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set IEWin = Nothing
  Window_SetAlwaysOnTop False
End Sub

Private Function Window_SetAlwaysOnTop(bAlwaysOnTop As Boolean) As Boolean
   Window_SetAlwaysOnTop = SetWindowPos(hwnd, -2 - bAlwaysOnTop, 0, 0, 0, 0, SWP_NOREDRAW Or SWP_NOSIZE Or SWP_NOMOVE)
End Function

Private Sub UpdateIEList(lv As ListView)
  lv.ListItems.Clear
  If IEWin.Count = 0 Then
     For i = 0 To 6
         Command1(i).Enabled = False
     Next i
     Exit Sub
  End If
  Dim tmpIE As IE_Class
  Dim itmx As ListItem
  For Each tmpIE In IEWin
             
      Set itmx = lv.ListItems.Add(, "K" & CStr(tmpIE.IEHandle), tmpIE.IEctl.LocationName)
      'itmx.SubItems(1) = tmpIE.IEctl.LocationURL
      
      'itmx.SubItems(2) = IIf(tmpIE.IEctl.Busy, "Busy", "Idle")
      'itmx.SmallIcon = IIf(tmpIE.IEctl.Busy, 1, 2)
  Next tmpIE
  For i = 0 To 6
      Command1(i).Enabled = True
  Next i
  
  
  'Command1(2).Enabled = IEWin.IE(Mid$(lv.SelectedItem.Key, 2)).EnableBack
  'Command1(3).Enabled = IEWin.IE(Mid$(lv.SelectedItem.Key, 2)).EnableForward
  Set itmx = Nothing
End Sub

Private Sub IEWin_IECommandStateChange(hwnd As Long, Button As SHDocVw.CommandStateChangeConstants, Enable As Boolean)
'   If Mid$(LV1.SelectedItem.Key, 2) = CStr(hwnd) And Button > 0 Then
'      Command1(4 - Button).Enabled = Enable
'   End If
      Command1(4 - Button).Enabled = Enable
End Sub

Private Sub IEWin_IEDocumentComplete(hwnd As Long, ByVal pDisp As Object, URL As Variant)
  Timer1.Enabled = False
  Dim itmx As ListItem
  Set itmx = LV1.ListItems("K" & CStr(hwnd))
  itmx.Text = IEWin.IE(CStr(hwnd)).IEctl.LocationName
  itmx.SmallIcon = 2
  itmx.SubItems(1) = URL
  itmx.SubItems(2) = "Idle"
  LV1.ListItems(itmx.Key).Selected = True
  Set itmx = Nothing
End Sub

Private Sub IEWin_IEDownloadBegin(hwnd As Long)
   If Timer1.Enabled = False Then
      Timer1.Enabled = True
      Dim itmx As ListItem
      Set itmx = LV1.ListItems("K" & CStr(hwnd))
      itmx.Text = IEWin.IE(CStr(hwnd)).IEctl.LocationName
      itmx.SmallIcon = 1
      itmx.SubItems(1) = IEWin.IE(CStr(hwnd)).IEctl.LocationURL
      itmx.SubItems(2) = "Downloading..."
      Set itmx = Nothing
   End If
End Sub

Private Sub IEWin_IEDownloadComplete(hwnd As Long)
  If Timer1.Enabled Then
     Timer1.Enabled = False
     Dim itmx As ListItem
     Set itmx = LV1.ListItems("K" & CStr(hwnd))
     itmx.Text = IEWin.IE(CStr(hwnd)).IEctl.LocationName
     itmx.SmallIcon = 2
     itmx.SubItems(1) = IEWin.IE(CStr(hwnd)).IEctl.LocationURL
     itmx.SubItems(2) = "Idle"
     Set itmx = Nothing
  End If
End Sub

Private Sub IEWin_IEMouseDown(hwnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim s As String
   If Button = 1 Then s = "Left "
   If Button = 2 Then s = "Right "
   s = s & "Mouse Down in " & IEWin.IE(CStr(hwnd)).IEctl.LocationName & vbCrLf
   s = s & "in position:" & vbCrLf
   s = s & "X = " & CStr(X) & vbCrLf
   s = s & "Y = " & CStr(Y)
   Debug.Print s
'   MsgBox s
End Sub

Private Sub IEWin_IEMouseUp(hwnd As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim s As String
   If Button = 1 Then s = "Left "
   If Button = 2 Then s = "Right "
   s = s & "Mouse Up in " & IEWin.IE(CStr(hwnd)).IEctl.LocationName & vbCrLf
   s = s & "in position:" & vbCrLf
   s = s & "X = " & CStr(X) & vbCrLf
   s = s & "Y = " & CStr(Y)
   Debug.Print s
'   MsgBox s
End Sub

Private Sub IEWin_IENavigationBegin(hwnd As Long, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
  Timer1.Enabled = True
  If IsForbidden(CStr(URL), List1) Then
     bCancel = True
     MsgBox "This URL is forbidden!!!", vbCritical, "Restricted URL"
     Exit Sub
  End If
  Dim itmx As ListItem
  Set itmx = LV1.ListItems("K" & CStr(hwnd))
  itmx.Text = IEWin.IE(CStr(hwnd)).IEctl.LocationName
  itmx.SmallIcon = 1
  itmx.SubItems(1) = URL
  itmx.SubItems(2) = "Navigating..."
  Set itmx = Nothing
End Sub

Private Sub IEWin_IENavigationComplete(hwnd As Long, ByVal pDisp As Object, URL As Variant)
  Dim itmx As ListItem
  Set itmx = LV1.ListItems("K" & CStr(hwnd))
  itmx.Text = IEWin.IE(CStr(hwnd)).IEctl.LocationName
  itmx.SmallIcon = 1
  itmx.SubItems(1) = IEWin.IE(CStr(hwnd)).IEctl.LocationURL
  itmx.SubItems(2) = "Downloading..."
  Set itmx = Nothing
End Sub

Private Sub IEWin_IEOnContextMenu(hwnd As Long)
   ret = MsgBox("Context Menu in " & IEWin.IE(CStr(hwnd)).IEctl.LocationName & " is about to appear." & vbCrLf & "Show context menu?", vbYesNo, "Context menu")
   If ret = vbNo Then bCancel = True
End Sub

Private Sub IEWin_IEWindowRegistered()
  Timer1.Enabled = True
  UpdateIEList LV1
End Sub

Private Sub IEWin_IEWindowRevoked()
  Timer1.Enabled = False
  UpdateIEList LV1
End Sub

Private Sub LV1_ItemClick(ByVal item As ListItem)
  Command1(2).Enabled = IEWin.IE(Mid$(item.Key, 2)).EnableBack
  Command1(3).Enabled = IEWin.IE(Mid$(item.Key, 2)).EnableForward
End Sub

Private Sub LV1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      xClick = X: yClick = Y
   End If
End Sub

Private Sub LV1_DblClick()
  Dim itmx As ListItem
  Set itmx = LV1.HitTest(xClick, yClick)
  If Not itmx Is Nothing Then
     SetForegroundWindow CLng(Mid$(itmx.Key, 2))
  End If
End Sub

Private Function IsForbidden(sUrl As String, lb As ListBox) As Boolean
  Dim i As Long
  i = SendMessage(lb.hwnd, LB_FINDSTRING, -1, ByVal sUrl)
  If i > -1 Then IsForbidden = True
End Function

Private Sub Timer1_Timer()
  If IsSurfing Then
     SurfingTime = SurfingTime + 1
     Caption = "Total surfing time = " & TimeSerial(0, 0, SurfingTime)
  End If
End Sub

Private Function IsSurfing() As Boolean
  Dim tmpIE As IE_Class
  For Each tmpIE In IEWin
      If Not tmpIE.IEctl Is Nothing Then
         If tmpIE.IEctl.Busy Then
            IsSurfing = True
            Exit For
         End If
      End If
  Next tmpIE
End Function
