VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form Messenger_Box 
   Caption         =   "MESSENGER BOX OF ELEITE SPY"
   ClientHeight    =   4344
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   7476
   LinkTopic       =   "Form1"
   ScaleHeight     =   4344
   ScaleWidth      =   7476
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_EXIT 
      Interval        =   40000
      Left            =   3144
      Top             =   1728
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1584
      Left            =   0
      TabIndex        =   1
      Top             =   2424
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   2794
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2280
      Top             =   1584
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C9FCFE&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7440
   End
   Begin VB.Menu MNU_12 
      Caption         =   "MENU BACK TO MAIN FORM"
   End
   Begin VB.Menu MNU_LIST1_LISTCOUNTING 
      Caption         =   "SCRIT ITEM COUNTER"
   End
End
Attribute VB_Name = "Messenger_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVS_EX_FULLROWSELECT = &H20
     
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Const BM_CLICK = &HF5&

Private Const WM_SETTEXT            As Long = &HC
Private Const WM_GETTEXT As Long = &HD
Private Const WM_GETTEXTLENGTH As Long = &HE

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, _
     ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszTitle As String) _
    As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByRef lParam As Any) _
    As Long
    
'Private Const WM_SETTEXT            As Long = &HC
'Private Const WM_GETTEXT = &HD

'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long  'MODULE 1141
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142
     
     
     
Private Function GetText(pHandle As Long) As String
       
    Dim Buffer As String
    Dim TextLength As Long
   
    TextLength = SendMessageAny(pHandle, WM_GETTEXTLENGTH, 0&, 0&)

    Buffer = String$(TextLength, 0&)
    SendMessageAny pHandle, WM_GETTEXT, TextLength + 1, Buffer
    GetText = Buffer
End Function
     
     
Private Sub SetFullRowSelection(ByVal hWndListView, ByVal bFullRow As Boolean)
   SendMessage hWndListView, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, ByVal CLng(bFullRow)
End Sub




'////////////////////////////////////////////////////////////////////
'//// FORM EVENTS
'////////////////////////////////////////////////////////////////////
       


Private Sub Form_Load()

    
Set FS = CreateObject("Scripting.FileSystemObject")

With ListView1
    .ColumnHeaders.Add , "INFO MESSENGER", "INFO MESSENGER", 4000, lvwColumnLeft
    .View = lvwReport
End With
   
'This is an example of how you can use this routine:
'enable full row selecting
SetFullRowSelection ListView1.hWnd, True
'disable full row selecting
'SetFullRowSelection ListView1.hwnd, False
   

'Label1.Caption = MsgBox_11

Call Timer1_Timer

For Each Control In Controls
    If InStr(UCase(Control.Name), "MNU_") > 0 Then
'        i_Menu_Count = i_Menu_Count + 1
        LABEL_44 = Trim(Control.Caption)
        LABEL_48 = Replace(LABEL_44, " ", "_")
        LABEL_48 = Replace(LABEL_48, "___", "__")
        LABEL_48 = "[__ " + LABEL_48 + " __]"
        LABEL_48 = Replace(LABEL_48, "[__ [__ ", "[__ ")
        LABEL_48 = Replace(LABEL_48, " __] __]", " __]")
        If LABEL_48 <> LABEL_44 Then
            Control.Caption = LABEL_48
        End If
    End If
Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

If frmMain.EXIT_TRUE = False Then
    Me.WindowState = vbMinimized
    Exit Sub
End If

End Sub

Private Sub MNU_12_Click()
    
    Me.WindowState = vbNormal
    Messenger_Box.Hide
    frmMain.WindowState = vbNormal

End Sub

Private Sub Timer_EXIT_Timer()
Unload Me
End Sub

Private Sub Timer1_Timer()

If Label1.Caption <> MsgBox_11 Then
    If MsgBox_11 = "" Then MsgBox_11 = "EMPTY BEGGINER"
        Label1.Caption = MsgBox_11
        STRING_VAR = Format(Now, "YYYY-MM-DD __ HH:MM:SS") + "__" + Replace(MsgBox_11, vbCrLf, "__")
        ListView1.ListItems.Add , , STRING_VAR
        Call LV_AutoSizeColumn(ListView1, ListView1.ColumnHeaders.Item(1))
    Exit Sub
End If

If MNU_LIST1_LISTCOUNTING.Caption <> "SCRIPT ITEM COUNT =" + Str(ListView1.ListItems.Count) Then
    MNU_LIST1_LISTCOUNTING.Caption = "[__ SCRIPT ITEM COUNT =" + Str(ListView1.ListItems.Count) + " __]"
End If

End Sub

Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column _
 As ColumnHeader = Nothing)

 Dim C As ColumnHeader
 If Column Is Nothing Then
  For Each C In LV.ColumnHeaders
   SendMessage LV.hWnd, LVM_FIRST + 30, C.Index - 1, -1
  Next
 Else
  SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
 End If
 LV.Refresh
End Sub


