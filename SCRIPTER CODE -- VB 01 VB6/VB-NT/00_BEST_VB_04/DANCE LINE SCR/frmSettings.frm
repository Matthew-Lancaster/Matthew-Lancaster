VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dancing Line properties"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "frmSettings.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3690
      TabIndex        =   9
      Top             =   1740
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2610
      TabIndex        =   8
      Top             =   1740
      Width           =   975
   End
   Begin VB.TextBox txtLineSpeed 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtLineThickness 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.ComboBox cmbColorsList 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Line linSeparator 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   4905
      Y1              =   1605
      Y2              =   1605
   End
   Begin VB.Line linSeparator 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   -15
      X2              =   4905
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label lblSpeedRange 
      Caption         =   "(1 to 100)"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   1125
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Line speed:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblThicknessRange 
      Caption         =   "(1 to 6)"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   405
      Width           =   735
   End
   Begin VB.Label lblLineThickness 
      Caption         =   "Line thickness:"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblLineColor 
      Caption         =   "Line color:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    'Unload form
    Unload Me
    
End Sub

Public Sub cmdOK_Click()
    'Save the setting
    SaveSetting "Dancing Line Screen Saver", "Settings", "LINE_COLOR", cmbColorsList.ListIndex + 1
    SaveSetting "Dancing Line Screen Saver", "Settings", "LINE_THICKNESS", txtLineThickness.Text
    SaveSetting "Dancing Line Screen Saver", "Settings", "LINE_SPEED", txtLineSpeed.Text
    
    'Unload form
    Unload frmSettings
    
End Sub


Private Sub Form_Load()
    'Add colors names to the combo box
    With cmbColorsList
        .AddItem "Blue"
        .AddItem "Green"
        .AddItem "Cyan"
        .AddItem "Red"
        .AddItem "Violet"
        .AddItem "Orange"
        .AddItem "Gray"
        .AddItem "Dark Gray"
        .AddItem "Light Blue"
        .AddItem "Light Green"
        .AddItem "Light Cyan"
        .AddItem "Light Red"
        .AddItem "Light Violet"
        .AddItem "Yellow"
        .AddItem "White"
    End With
    
    'Select current color
    cmbColorsList.ListIndex = LineColor - 1
    
    'Type current thickness
    txtLineThickness.Text = LineThickness
    
    'Type current speed
    txtLineSpeed.Text = LineSpeed

End Sub


Private Sub imgVBPZ_Click()
    'Show the VBPZ ad
    MsgBox "This app was downloaded from VISUAL BASIC PROGRAMMING ZONE: visit us! URL: http://jump.to/vbprogzone", vbInformation, "VBPZ Message"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'When form is unloade, all Screen Saver must unload
    'End
    
End Sub


Private Sub txtLineSpeed_LostFocus()
    'Checks data
    If txtLineSpeed.Text = "" Then
        MsgBox "Enter a value for the line speed.", vbExclamation, "Error"
        Exit Sub
    End If
    If txtLineSpeed.Text > 100 Or _
    txtLineSpeed.Text < 1 Then
        MsgBox "Line speed range is 1..100", vbExclamation, "Error"
        Exit Sub
    End If

End Sub


Private Sub txtLineThickness_LostFocus()
    'Checks data
    If txtLineThickness.Text = "" Then
        MsgBox "Enter a value for the line thickness.", vbExclamation, "Error"
        Exit Sub
    End If
    If txtLineThickness.Text > 6 Or _
    txtLineThickness.Text < 1 Then
        MsgBox "Line thickness range is 1..6", vbExclamation, "Error"
        Exit Sub
    End If
    
End Sub


