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

Public EXIT_TRUE


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
    
    ' GetMenuInfo GetMenu(hWnd), 0, "", txt
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

Public Sub SET_MENU_PADD_WORK()

' Call frmListMenu.SET_MENU_PADD_WORK

Dim i_Menu_Count, i_Form_Counter
Dim i_Menu_Not_Visa_Count

Dim Control As Control, Label_44, LABEL_48

Dim R_NEXT

Dim Text_Checker_Form_Menu As String

Dim MENU_ITEM_VAR
Dim i

For i = 0 To Forms.Count - 1
    For Each Control In Forms(i).Controls
        If InStr(UCase(Control.Name), "MNU_") > 0 Then
            If Control.Visible = True Then
                i_Menu_Count = i_Menu_Count + 1
            End If
            If Control.Visible = False Then
                i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count + 1
            End If

        End If
    Next
Next

' i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count - i_Menu_Count

Dim Form As Form
For Each Form In Forms
    i_Menu_Not_Visa_Count = 0
    i_Menu_Count = 0
    For Each Control In Form.Controls
        Debug.Print Control.Name
        If InStr(UCase(Control.Name), "MNU_") > 0 Then
            If Control.Visible = True Then
                i_Menu_Count = i_Menu_Count + 1
            End If
            If Control.Visible = False Then
                i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count + 1
            End If
        End If
    Next
    
    ' ---------------------------------------------------------------------------------
    ' MNU_MENU_ITEM_COUNT IS A MENU ITEM THAT MUST BE IN EVERY FORM THAT HAVE MENU ITEM
    ' AND DO COUNT DISPLAY
    ' ---------------------------------------------------------------------------------
    If i_Menu_Not_Visa_Count > 0 Then
        Form.MNU_MENU_ITEM_COUNT.Caption = "MENU ITEM COUNT = " + Str(i_Menu_Count) + " &&" + Str(i_Menu_Not_Visa_Count) + " Not Visible"
    Else
        Debug.Print Form.Caption
        Debug.Print Form.Name
        ' --------------------------------------------------------------
        ' THE ERROR HANDLE SORT IT THAT NOT DROP OUT FOR ROUTINE
        ' AND CONTINUE WHEN FORM NOT GOT MNU ITEM TO DISPLAY COUNT THERE
        ' Tue 22-Sep-2020 11:12:00
        ' --------------------------------------------------------------
        On Error Resume Next
            Form.MNU_MENU_ITEM_COUNT.Caption = "MENU ITEM COUNT = " + Str(i_Menu_Count)
        On Error GoTo 0
    End If
Next


'Mnu_Form_Count.Caption = "Form Counter = " + Str(Forms.Count - 1) '  + " Really 7"
'Mnu_Form_Count.Visible = False

i_Menu_Count = 0

For i = 0 To Forms.Count - 1
    Text_Checker_Form_Menu = ""
    frmListMenu.GetMenuInfo_Not_Indented GetMenu(Forms(i).hWnd), 0, "", Text_Checker_Form_Menu
    Text_Checker_Form_Menu = UCase(Text_Checker_Form_Menu)
    For Each Control In Forms(i).Controls
        If InStr(UCase(Control.Name), "MNU_") > 0 Then
            MENU_ITEM_VAR = Replace(Control.Caption, "[__ ", "")
            MENU_ITEM_VAR = Replace(MENU_ITEM_VAR, " __]", "")
            MENU_ITEM_VAR = UCase(Trim(MENU_ITEM_VAR))
            If InStr(Text_Checker_Form_Menu, "SUB MENU ----" + MENU_ITEM_VAR) > 0 Then
                Control.Caption = UCase(Control.Caption)
            End If
            If InStr(Text_Checker_Form_Menu, "SUB MENU ----" + MENU_ITEM_VAR) = 0 Then
                
                'i_Menu_Count = i_Menu_Count + 1
                If InStr(Trim(Control.Caption), "[__ ") = 0 Then
                    Label_44 = Trim(Control.Caption)
                    'LABEL_48 = Replace(LABEL_44, " ", "_")
                    LABEL_48 = Label_44
                    LABEL_48 = Replace(LABEL_48, "___", "__")
                    LABEL_48 = "[__ " + LABEL_48 + " __]"
                    LABEL_48 = Replace(LABEL_48, "[__ [__ ", "[__ ")
                    LABEL_48 = Replace(LABEL_48, " __] __]", " __]")
                    If LABEL_48 <> Label_44 Then
                        Control.Caption = UCase(LABEL_48)
                    End If
                End If
            End If
        End If
    Next
Next

'Stop

''MNU_BRing_Front
''---------------
'i_Form_Counter = Forms.Count
'i_Form_Counter = 0
''for each f
''For i = 0 To Forms.Count - 1
''    Load Forms(i)
''    Show Forms(i)
''Next
'
'Dim Form As Form
'i_Form_Counter = 0
'For Each Form In Forms
'    i_Form_Counter = i_Form_Counter + 1
'    Load Form
'    Form.Show
'    Show Form
'    'Set Form = Nothing
'Next Form
'
'i_Form_Counter = 0
'For Each Form In Forms
'    i_Form_Counter = i_Form_Counter + 1
'    Load Form
'    Form.Show
'    'Set Form = Nothing
'Next Form

i_Form_Counter = Forms.Count - 1
Me.Refresh
End Sub
