VERSION 5.00
Begin VB.UserControl Archimedes 
   BackColor       =   &H80000005&
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   PropertyPages   =   "PIControl.ctx":0000
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
End
Attribute VB_Name = "Archimedes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Coords() As POINTAPI, _
        m_CoordCount As Long

Private m_Incremental As Double, _
        m_MaxOrbits As Double, _
        m_StepInterval As Double
       
Private m_ForeColor As OLE_COLOR, _
        m_BackColor As OLE_COLOR
       
Private m_Base As Long, _
        m_Zoom As Double
       
Private m_Colors() As OLE_COLOR
       
Public Event CoordinateListChanged()

Public Property Get CoordinateListString() As String
    
    Dim a As Long, _
        t As String
    
    Dim o As POINTAPI
    
    o.x = ScaleWidth / 2
    o.y = ScaleHeight / 2
            
    t = t & "Origin: " & CLng(ScaleWidth / 2) & ", " & CLng(ScaleHeight / 2) & vbCrLf
    
    For a = 0 To m_CoordCount - 1
    
        If t <> "" Then t = t + vbCrLf
        
        If m_Coords(a).x < o.x Then
            t = t & "-" & (o.x - m_Coords(a).x) & ", "
        Else
            t = t & m_Coords(a).x & ", "
        End If
        
        If m_Coords(a).y < o.y Then
            t = t & "-" & (o.y - m_Coords(a).y)
        Else
            t = t & m_Coords(a).y
        End If
            
    Next
    
    CoordinateListString = t
    
End Property

Public Property Get Color(ByVal Index As Long) As OLE_COLOR
Attribute Color.VB_ProcData.VB_Invoke_Property = "NumberColors"
    If (Index < 0) Or (Index > 9) Then Exit Property
    
    Color = m_Colors(Index)
    
End Property

Public Property Let Color(ByVal Index As Long, ByVal varNew As OLE_COLOR)
    If (Index < 0) Or (Index > 9) Then Exit Property
    
    m_Colors(Index) = varNew
    PropertyChanged
    
End Property

Public Property Get CoordinateCount() As Long
    CoordinateCount = m_CoordCount
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal varNew As OLE_COLOR)
    m_ForeColor = varNew
    PropertyChanged
    
    Refresh
    
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = "NumberColors"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal varNew As OLE_COLOR)
    m_BackColor = varNew
    PropertyChanged
    
    Refresh
    
End Property

Public Property Get Base() As Long
    Base = m_Base
End Property

Public Property Let Base(ByVal varNew As Long)
    
    Dim a As Long
    
    a = m_Base - 1
    m_Base = varNew
    
    ReDim Preserve m_Colors(m_Base - 1)
        
    PropertyChanged
    
    Refresh
    
End Property

Public Property Get Zoom() As Double
    Zoom = m_Zoom
End Property

Public Property Let Zoom(ByVal varNew As Double)
    
    If varNew <= 0# Then varNew = 0.1
    m_Zoom = varNew
    
    CheckNotZero
    PropertyChanged
    
    Refresh
    
End Property

Public Property Get Incremental() As Double
    Incremental = m_Incremental
End Property

Public Property Let Incremental(ByVal varNew As Double)
    m_Incremental = varNew
    CheckNotZero
    PropertyChanged
    
    Refresh
    
End Property

Public Property Get MaxOrbits() As Double
    MaxOrbits = m_MaxOrbits
End Property

Public Property Let MaxOrbits(ByVal varNew As Double)
    m_MaxOrbits = varNew
    CheckNotZero
    PropertyChanged
    
    Refresh
    
End Property


Public Property Get StepInterval() As Double
    StepInterval = m_StepInterval
End Property

Public Property Let StepInterval(ByVal varNew As Double)
    m_StepInterval = varNew
    CheckNotZero
    PropertyChanged
    
    Refresh
    
End Property

Public Property Get DesignModePreview() As Boolean
    DesignModePreview = False
End Property


Public Property Let DesignModePreview(ByVal varNew As Boolean)
    
    If varNew = True Then
        Redraw True
    End If
    
End Property

Public Sub Refresh()
    Redraw
End Sub

Private Sub UserControl_Paint()
    
    Redraw
    
End Sub

Public Sub Redraw(Optional ByVal ForceDraw As Boolean)

    Dim p As CARTESEAN2D
        
    Dim lpBrush As LOGBRUSH, _
        hBrush As Long
            
    Dim x As Long, _
        y As Long
        
    Dim a As Double, _
        b As Double, _
        c As Long, _
        d As Long, _
        e
        
    Dim s As Double, _
        t As Double, _
        u As Double
    
    Dim lpRect As RECT
    
    Dim lCoord As POINTAPI, _
        pFile As Integer, _
        by As Byte, _
        sT As String
                
    Dim fDC As Long
    
    Dim z As Long, _
        r As Long
    
    On Error Resume Next
        
    fDC = hDC
        
    lpBrush.lbColor = GetActualColor(m_BackColor)
    lpBrush.lbStyle = BS_SOLID
    
    hBrush = CreateBrushIndirect(lpBrush)
    
    With lpRect
        .Right = ScaleWidth + 1
        .Bottom = ScaleHeight + 1
    End With
    
    FillRect fDC, lpRect, hBrush
    
    DeleteObject hBrush
        
    With lpRect
        .Right = .Right - 2
        .Bottom = .Bottom - 2
    End With
    
    Draw3DFrame_Sunken fDC, lpRect
    
    p.Origin.x = ScaleWidth / 2
    p.Origin.y = ScaleHeight / 2
    
    lCoord.x = p.Origin.x
    lCoord.y = p.Origin.y
    
    s = 0
    t = (360 * m_MaxOrbits)
    
    u = m_StepInterval
    e = 0@
    
'    pFile = FreeFile
    
'    Open "pif500.txt" For Binary Access Read Lock Write As pFile
    
    If (Ambient.UserMode = True) Or (ForceDraw = True) Then
        
        For a = s To t Step u
                        
            If (a > 90) Then p.Arc = 360 - (a Mod 360) Else p.Arc = a
            
            p.Distance = (a * b) * m_Zoom
'            p.Distance = a * b
            b = b + m_Incremental
            
            GetLinearPoints p
            
            If (c = 0) Then
                
    '            by = 0
    '            Do While IsNumeric(Chr(by)) = False
    '                Get #pFile, , by
    '            Loop
                
                d = Val(Chr(by))
                        
                ReDim m_Coords(0)
                
                CopyMemory m_Coords(0), p.Points, 8
                c = c + 1
                
                SetPixel fDC, p.Points.x, p.Points.y, GetActualColor(m_Colors(d))
                
            ElseIf (m_Coords(c - 1).x <> p.Points.x) Or (m_Coords(c - 1).y <> p.Points.y) Then
            
    '            by = 0
    '            Do While IsNumeric(Chr(by)) = False
    '                Get #pFile, , by
    '            Loop
                
    '            d = Val(Chr(by))
                        
                ReDim Preserve m_Coords(c)
                
                m_Coords(c).x = p.Points.x
                m_Coords(c).y = p.Points.y
                
                c = c + 1
                
                SetPixel fDC, p.Points.x, p.Points.y, GetActualColor(m_Colors(d))
                
            End If
            
            d = d + 1
            If d > (m_Base - 1) Then d = 0
                
            If (p.Points.x > ScaleWidth) Or (p.Points.x < 0) Or _
                (p.Points.y > ScaleHeight) Or (p.Points.y < 0) Then Exit For
            
        Next
        
    End If
    
'    Close #pFile
    
    m_CoordCount = c
    RaiseEvent CoordinateListChanged
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    m_Base = PropBag.ReadProperty("Base", m_Base)
    m_Zoom = PropBag.ReadProperty("Zoom", m_Zoom)

    m_Incremental = PropBag.ReadProperty("Incremental", m_Incremental)
    m_MaxOrbits = PropBag.ReadProperty("MaxOrbits", m_MaxOrbits)
    m_StepInterval = PropBag.ReadProperty("StepInterval", m_StepInterval)
    
    CheckNotZero
    
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_ForeColor)
    m_BackColor = PropBag.ReadProperty("BackColor", m_BackColor)
    
    Dim a As Long
    
    ReDim m_Colors(m_Base - 1)
    
    For a = 0 To m_Base - 1
        m_Colors(a) = PropBag.ReadProperty("Color" & a, m_Colors(a))
    Next
    
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "Incremental", m_Incremental
    PropBag.WriteProperty "MaxOrbits", m_MaxOrbits
    PropBag.WriteProperty "StepInterval", m_StepInterval
    
    PropBag.WriteProperty "ForeColor", m_ForeColor
    PropBag.WriteProperty "BackColor", m_BackColor
    
    PropBag.WriteProperty "Base", m_Base
    PropBag.WriteProperty "Zoom", m_Zoom
    
    Dim a As Long
    
    For a = 0 To m_Base - 1
        PropBag.WriteProperty "Color" & a, m_Colors(a)
    Next
    
End Sub

Private Sub CheckNotZero()

    If (m_StepInterval = 0) Then m_StepInterval = 1
    If (m_MaxOrbits = 0) Then m_MaxOrbits = 7
    If (m_Incremental = 0) Then m_Incremental = 0.00008
    
End Sub

Private Sub UserControl_Initialize()
    
    
    m_BackColor = vbWindowBackground
    m_ForeColor = vbWindowText
    
    CheckNotZero
    
    m_Base = 10
    
    m_Zoom = 1#
    
    ReDim m_Colors(m_Base - 1)
    
    m_Colors(0) = &H0&
    m_Colors(1) = &HCF&
    m_Colors(2) = &HCF7F&
    m_Colors(3) = &HCF00&
    m_Colors(4) = &H7FCF&
    m_Colors(5) = &HCF0000
    m_Colors(6) = &HCF7F00
    m_Colors(7) = &HCF7FCF
    m_Colors(8) = &HCFCF7F
    m_Colors(9) = &H7FCFCF
    
End Sub
