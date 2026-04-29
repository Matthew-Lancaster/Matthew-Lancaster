VERSION 5.00
Begin VB.Form Form5_DISABLE_CLOSE_BUTTON_2 
   Caption         =   "Form5"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form5_DISABLE_CLOSE_BUTTON_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----
'Hide Max, Min and Close Buttons - Visual Basic (Microsoft) Win32API - Tek-Tips
'https://www.tek-tips.com/viewthread.cfm?qid=330760
'----

Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long

Private Const MIIM_ID = &H2
Private Const MIIM_TYPE = &H10
Private Const MFT_STRING = &H0

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Sub Form_Load()

    Dim hMenu As Long
    Dim lpmii As MENUITEMINFO
    
    hMenu = GetSystemMenu(hWnd, False)
    With lpmii
        .cbSize = Len(lpmii)
        .fMask = MIIM_TYPE Or MIIM_ID
        .fType = MFT_STRING
        .wID = 100                      'ID to check when subclassing
        .dwTypeData = "&Howdy"
        .cch = Len(.dwTypeData)
    End With
    Call InsertMenuItem(hMenu, GetMenuItemCount(hMenu), 1, lpmii)

End Sub
