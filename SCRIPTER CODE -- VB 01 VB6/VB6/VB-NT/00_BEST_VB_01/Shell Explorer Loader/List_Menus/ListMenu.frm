VERSION 5.00
Begin VB.Form frmListMenu 
   Caption         =   "ListMenu"
   ClientHeight    =   5616
   ClientLeft      =   2112
   ClientTop       =   2220
   ClientWidth     =   5100
   LinkTopic       =   "ListMenu"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5616
   ScaleWidth      =   5100
   Begin VB.TextBox Text1 
      Height          =   3495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuMore 
         Caption         =   "More Help"
         Begin VB.Menu Mnu_Double_Helper 
            Caption         =   "Double Helper"
         End
      End
   End
End
Attribute VB_Name = "frmListMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----
'[RESOLVED] List out all menus in vb6 application-VBForums
'http://www.vbforums.com/showthread.php?683620-RESOLVED-List-out-all-menus-in-vb6-application
'--------
'VB Helper: HowTo: List a form's menu items
'http://www.vb-helper.com/howto_list_menus.html
'----
'[ Tuesday 04:18:00 Pm_23 October 2018 ]

Option Explicit

Dim MenuName As New Collection
Dim MenuHandle As New Collection

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

Private Const MF_BYPOSITION = &H400&

Private Sub GetMenuInfo(hMenu As Long, Spaces As Integer, SubText, txt As String)
Dim num As Integer
Dim i As Integer
Dim length As Long
Dim sub_hmenu As Long
Dim sub_name As String
    
    num = GetMenuItemCount(hMenu)
    For i = 0 To num - 1
        ' Save this menu's info.
        sub_hmenu = GetSubMenu(hMenu, i)
        sub_name = Space$(256)
        'Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
        length = GetMenuString(hMenu, i, sub_name, Len(sub_name), MF_BYPOSITION)
        sub_name = Left$(sub_name, length)

        ' txt = txt & SubText & Space$(Spaces) & sub_name & vbCrLf
        txt = txt & SubText & Space$(Spaces) & sub_name & vbCrLf
        
        ' Get its child menu's names.
        GetMenuInfo sub_hmenu, Spaces + 4, "Sub Menu ----", txt
    Next i
End Sub

Public Sub GetMenuInfo_Not_Indented(hMenu As Long, Spaces As Integer, SubText, txt As String)
Dim num As Integer
Dim i As Integer
Dim length As Long
Dim sub_hmenu As Long
Dim sub_name As String
    
    num = GetMenuItemCount(hMenu)
    For i = 0 To num - 1
        ' Save this menu's info.
        sub_hmenu = GetSubMenu(hMenu, i)
        sub_name = Space$(256)
        length = GetMenuString(hMenu, i, sub_name, Len(sub_name), MF_BYPOSITION)
        
        sub_name = Left$(sub_name, length)

        ' txt = txt & SubText & Space$(Spaces) & sub_name & vbCrLf
        txt = txt & SubText & sub_name & vbCrLf
        
        ' Get its child menu's names.
        GetMenuInfo_Not_Indented sub_hmenu, Spaces + 4, "Sub Menu ----", txt
    Next i
End Sub



Private Sub Form_Load()
    
    ' Dim txt As String
    
    ' GetMenuInfo GetMenu(hwnd), 0, "", txt
    ' frmListMenu.GetMenuInfo_Not_Indented GetMenu(hWnd), 0, "", txt
    ' Text1.Text = txt

End Sub

Private Sub Form_Resize()
    Text1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuHelpAbout_Click()
    MsgBox "This example lists the program's menus."
End Sub


