VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl ColorChoose 
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   ScaleHeight     =   1530
   ScaleWidth      =   3135
   ToolboxBitmap   =   "ColorChoose.ctx":0000
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   1350
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frBorder 
      Caption         =   " Color Settings "
      Height          =   1455
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3075
      Begin VB.ComboBox cbNamedPalette 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   990
         Width           =   2715
      End
      Begin VB.TextBox txtColorValue 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   630
         Width           =   2265
      End
      Begin VB.CommandButton cmdChoose 
         Caption         =   "..."
         Height          =   285
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   630
         Width           =   375
      End
      Begin VB.Label lblValue 
         Caption         =   "Selected Color Value:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   1545
      End
   End
End
Attribute VB_Name = "ColorChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''' AUTHOR:       Nathan Moschkin                               '''''''
''''''' MODULE:       ColorChoose                                   '''''''
''''''' DESCRIPTION:  Color picker control with named-palette.      '''''''
'''''''                                                             '''''''
''''''' COPYRIGHT:    Datavations, LLC  2003                        '''''''
'''''''                                                             '''''''
'''''''                                                             '''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Option Base 0

Private mListed As Boolean

Private orgIndent As Long

Private orgHeight As Long

Private fNoChange As Boolean

Private m_Caption As String
Private m_Color As OLE_COLOR

Private m_Enabled As Boolean

Public Event ColorChanged()

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal varNew As Boolean)
    
    UserControl.Enabled = varNew
    m_Enabled = varNew
    
    frBorder.Enabled = varNew
    lblValue.Enabled = varNew
    cmdChoose.Enabled = varNew
    cbNamedPalette.Enabled = varNew
    txtColorValue.Enabled = varNew
    
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
    
End Property

Public Property Let Caption(ByVal varNew As String)
    m_Caption = varNew
    frBorder.Caption = " " & m_Caption & " "
    PropertyChanged
    
End Property

Public Property Get Color() As OLE_COLOR
Attribute Color.VB_UserMemId = 0
Attribute Color.VB_MemberFlags = "200"
    Color = m_Color
End Property

Public Property Let Color(ByVal varNew As OLE_COLOR)
    m_Color = varNew
    
    fNoChange = True
    SetupColor
    
    fNoChange = False
    PropertyChanged
    
    DoEvents
    
End Property

Private Sub SetupColor()

    Dim s As String, _
        e As String
        
    s = Hex(m_Color And &HFFFFFFFF)
    e = "&H" + String(8 - Len(s), "0") + s + "&"
    
    txtColorValue = e
    GetNamedColor
    
    If (m_Color = -1) Then
        cmdChoose.BackColor = vbButtonFace
    Else
        cmdChoose.BackColor = m_Color
    End If
    
    If (fNoChange = False) Then RaiseEvent ColorChanged
    
End Sub

Private Sub cbNamedPalette_Click()
''
    CheckNamedColor
    SetupColor

End Sub

Private Sub cbNamedPalette_LostFocus()

    Dim cColor As Long
    
    cColor = m_Color
    
    CheckNamedColor
    If (cColor = m_Color) Then
        fNoChange = True
    Else
        fNoChange = False
    End If
    
    SetupColor
    fNoChange = False
    
End Sub

Private Sub cbNamedPalette_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If (KeyCode = 13) And (Shift = 0) Then
        CheckNamedColor
        SetupColor
    End If
        
End Sub

Private Sub cmdChoose_Click()
''
    cDlg.Color = m_Color And &HFFFFFF
    cDlg.flags = &H3&
    
    cDlg.ShowColor
    
    If (cDlg.CancelError = False) Then
        m_Color = cDlg.Color
        SetupColor
    End If
    
    If (fNoChange = False) Then RaiseEvent ColorChanged
    
End Sub


Private Sub txtColorValue_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If (KeyCode = 13) And (Shift = 0) Then
        m_Color = Val(txtColorValue)
        SetupColor
    End If

End Sub

Private Sub txtColorValue_LostFocus()

    If (m_Color = Val(txtColorValue)) Then
        fNoChange = True
    End If
    
    m_Color = Val(txtColorValue)
    SetupColor
    
    fNoChange = False

End Sub

Private Sub UserControl_GotFocus()
    fNoChange = False
End Sub

Private Sub UserControl_EnterFocus()
    fNoChange = False
End Sub

Private Sub UserControl_ExitFocus()
    fNoChange = True
End Sub

Private Sub UserControl_LostFocus()
    fNoChange = True
End Sub

Private Sub UserControl_Initialize()

    ListNamedPalette
    
    orgIndent = txtColorValue.Left
    orgHeight = Height
    
    Color = vbGreen
    Caption = "Color Picker"
    
    m_Enabled = True

End Sub

Private Sub UserControl_Resize()
                                
    
    If (mListed = False) Then
        mListed = True
        
        
        If (g_ColorCount > 0) Then
            Dim i As Long, _
                j As Long
                
            Dim cmpVal As Long
            
            cmpVal = Val(txtColorValue)
            
            i = g_ColorCount - 1
            
            cbNamedPalette.Clear
            cbNamedPalette.AddItem "(auto)"
            
            For j = 0 To i
                cbNamedPalette.AddItem g_Colors(j).ColorName
            Next
        End If
    End If
                
    If (Width < 1965) Then
        Width = 1965
        Exit Sub
    End If
                
    If (Height <> orgHeight) Then
        Height = orgHeight
        Exit Sub
    End If
    
    frBorder.Move 0, 0, ScaleWidth, ScaleHeight
    
    With cmdChoose
        .Move ScaleWidth - (.Width + orgIndent), .Top, .Width, .Height
    End With
    
    With txtColorValue
        .Move .Left, .Top, ScaleWidth - (cmdChoose.Width + 45 + (2 * orgIndent)), .Height
    End With
    
    With cbNamedPalette
        .Move .Left, .Top, ScaleWidth - (2 * orgIndent)
    End With
    
End Sub

Private Function CheckNamedColor() As Boolean
''

    Dim i As Long, _
        j As Long
        
    Dim cmpVal As String
    
    cmpVal = cbNamedPalette.Text
    
    If (cmpVal = "(auto)") Then
        m_Color = -1&
        CheckNamedColor = True
        Exit Function
    End If
    
    i = g_ColorCount - 1
    
    For j = 0 To i
        If (LCase(g_Colors(j).ColorName) = LCase(cmpVal)) Then
            
            CheckNamedColor = True
            m_Color = g_Colors(j).Value
            
            Exit Function
        End If
    Next

    CheckNamedColor = False

End Function

Private Sub GetNamedColor()
''
    Dim i As Long, _
        j As Long
        
    Dim cmpVal As Long
        
    
    cmpVal = Val(txtColorValue)
    
    If (cmpVal = -1&) Then
        cbNamedPalette.Text = "(auto)"
        Exit Sub
    End If
    
    i = g_ColorCount - 1
    
    For j = 0 To i
        If (g_Colors(j).Value = cmpVal) Then
            cbNamedPalette.Text = g_Colors(j).ColorName
            Exit Sub
        End If
    Next
    
    cbNamedPalette.Text = "Undefined"
    
End Sub

Private Sub UserControl_Terminate()
''
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Caption = PropBag.ReadProperty("Caption", m_Caption)
    m_Color = PropBag.ReadProperty("Color", m_Color)
    m_Enabled = PropBag.ReadProperty("Enabled", m_Enabled)
    
    frBorder.Caption = " " & m_Caption & " "
    
    fNoChange = True
    SetupColor
    fNoChange = False
    
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "Caption", m_Caption
    PropBag.WriteProperty "Color", m_Color
    PropBag.WriteProperty "Enabled", m_Enabled

End Sub
