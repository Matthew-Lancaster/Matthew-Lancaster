VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmParent 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Fractal Studio Alpha Version"
   ClientHeight    =   5130
   ClientLeft      =   165
   ClientTop       =   915
   ClientWidth     =   6510
   Icon            =   "frmParent.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tlbTools 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   794
      ButtonWidth     =   661
      ButtonHeight    =   635
      Appearance      =   1
      ImageList       =   "imlTools"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4875
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParent.frx":0A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParent.frx":0E44
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParent.frx":1286
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParent.frx":16C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParent.frx":1B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParent.frx":1F4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuTopFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTopDebug 
      Caption         =   "Debug"
   End
   Begin VB.Menu mnuTopWindow 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "Arrange"
      End
      Begin VB.Menu mnuWindowSep01 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindowShowparam 
         Caption         =   "Show Parameter Window"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTopHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About Fractal Studio"
      End
   End
End
Attribute VB_Name = "frmParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Terminate
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Cancel = 1
    Terminate
End Sub


Private Sub mnuFileExit_Click()
    Terminate
End Sub

Private Sub mnuFileNew_Click()
    AddFractal
End Sub

Private Sub mnuFileRender_Click()
    Calculate aData.nActive, True
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "Version " & App.Major & "." & App.Minor & App.Revision & " Alpha." & vbCrLf & "Copyright 2001 , Erling Andersen"
End Sub

Private Sub mnuTopDebug_Click()

If MsgBox("Do you want to register param file type (*.fs1) ?", vbYesNo, "File Associations") = vbYes Then
    RegisterFileExtension "fs1"
End If

End Sub

Private Sub mnuWindowArrange_Click()
    ArrangeWindows
End Sub

Private Sub mnuWindowShowparam_Click()
    frmParam.Show
    mnuWindowShowparam.Visible = False
    frmParent.mnuWindowSep01.Visible = False
End Sub

Private Sub tlbTools_ButtonClick(ByVal Button As MSComCtlLib.Button)
Dim sFile As String

Select Case Button.Index
Case 2
    'New
    AddFractal
Case 3
    'Open
    sFile = GetOpenFilePath("Open Fractal Param File", "Fstudio Param File|*.fs1")
    
    If sFile <> "" Then
        LoadParamFile sFile
    End If
Case 4
    'Save
    sFile = GetSaveFilePath("Save Fractal Param File", "Fstudio Param File|*.fs1")
    If FileExist(sFile) And sFile <> "" Then
        If MsgBox("The file already exists. Do you want to overwrite it?", _
            vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    If sFile <> "" Then
        SaveParamFile aData.nActive, sFile
    End If
Case 6
    'Arrow tool
    MsgBox "Not implemented yet.."
Case 7
    'Zoom tool
    MsgBox "Not implemented yet.."
Case 8
    'Move tool
    MsgBox "Not implemented yet.."
End Select

End Sub
