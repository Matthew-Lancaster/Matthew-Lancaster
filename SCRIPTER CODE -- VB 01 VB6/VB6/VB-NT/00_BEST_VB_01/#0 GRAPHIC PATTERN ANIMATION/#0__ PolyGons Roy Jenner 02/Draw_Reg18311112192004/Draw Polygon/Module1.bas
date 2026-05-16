Attribute VB_Name = "mdlDraw"
'****************************
'By     Jim Jose
'email  jimjosev33@yahoo.com
'****************************

'PLEASE READ THIS

'If you feel Satisfactory
'   Please 'Rate' this code
'Else
'   Give feedback to improve this code.
'End If
'Good luck
'****************************
Option Explicit

Public Type COORD   'The co-ordinates in Long
    X As Long
    Y As Long
End Type

Public Enum rgnFillStyle    'FillStyle Enum
    Transparent = 0
    HorizontalLine = 1
    VerticalLine = 2
    UpwardDiagonal = 3
    DownwardDiogonal = 4
    Cross = 5
    DiagonalCross = 6
    Solid = 7
    FillBitmap = 8
End Enum

Public Type LOGBRUSH    'Type for the API 'CreateBrushIndirect'
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Const ALTERNATE = 1 ' ALTERNATE and WINDING are
Const WINDING = 2 ' constants for FillMode.

'API'S used to draw
Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'======================================================================================================================
'< Ref >
    'This module can draw Regular Filled polygones
    'The polygons can be filled by 'LinePattern' and 'BitmapPattern'
    'The  API 'CreateBrushIndirect' is used to create 'Line Pattern' brush
    'The  API 'CreatePatternBrush' is used to create 'Bitmap Pattern' brush
    'If you have any questions please ask Me
'< Info >
    'You can draw unlimitted sides , in any size
    'You only need to specify 'Total Sides' and 'Side Length'
    'The 'BitmapPatterns' used are of the size 8 * 8 pixels
    'I didn't included an error message , since that will result in a continues loop ( in a timer )
    'If you know to fill with large bitmaps 'Please inform me'
'< Tips >
    'The border of the Polygon will be always the default Black
    'You can change the border color by redefining the ('VbBlack') in ('DrawLines:") in the last step of the function.
'======================================================================================================================

'This function can draw a regular polygon in  a PictureBox
Public Function DrawPolygon(Pic As PictureBox, _
                        ByVal X As Double, _
                        ByVal Y As Double, _
                        Optional IsCentre As Boolean = True, _
                        Optional ByVal Sides As Double = 4, _
                        Optional ByVal SideLength As Double = 100, _
                        Optional FillStyle As rgnFillStyle = 0, _
                        Optional ByVal Rotation As Double = 0, _
                        Optional imgData As StdPicture = 0) As Boolean
                        
On Error GoTo Handle
If Val(Sides) < 3 Then DrawPolygon = False: Exit Function   'Exiting none valid Sidecount

Dim Side  As Integer
Dim CtrAngle   As Double
Dim Pnt() As COORD

CtrAngle = 360 / Sides 'Getting the Centre angle
ReDim Pnt(1 To Sides) As COORD   'Getting the  vaiables for vertex of polygon

'-------------------------------------------------------------------------
'The bellow centering eqation I derived myself and may contain errors.
'Anyway it works Smoothly and perfectly under normal dimensions
'-------------------------------------------------------------------------
    If IsCentre = True Then 'if the given X,Y are the Orgine
        Dim VDist As Double
         'Determining the distance from the orgin to the centre of on side.
        VDist = SideLength / (2 * Tangent(CtrAngle / 2))
        'Setting the first vertex corresponding to the value of rotation
        Pnt(1).X = X + VDist * Cosine(270 + Rotation) + SideLength / 2 * Cosine(180 + Rotation)
        Pnt(1).Y = Y + VDist * Sine(270 + Rotation) + SideLength / 2 * Sine(180 + Rotation)
    Else 'If the given X,Y are one vertex
        Pnt(1).X = X: Pnt(1).Y = Y
    End If

'-------------------------------------------------------------------------
'Proceeding on determining other vertices ( In Math 'Vector Method' )  < x2 =x1 + (r * Theta) >
'-------------------------------------------------------------------------
    For Side = 2 To Sides
        Pnt(Side).X = Pnt(Side - 1).X + SideLength * Cosine((Side - 2) * CtrAngle + Rotation)
        Pnt(Side).Y = Pnt(Side - 1).Y + SideLength * Sine((Side - 2) * CtrAngle + Rotation)
    Next Side

'-------------------------------------------------------------------------
'Shading the sides ( We need to convert the values to Pixel mode.
'-------------------------------------------------------------------------
Dim Vtx() As COORD 'Defining  new verticies in Pixel Mode
ReDim Vtx(1 To Sides)

For Side = 1 To Sides
    Vtx(Side).X = Pnt(Side).X / Screen.TwipsPerPixelX
    Vtx(Side).Y = Pnt(Side).Y / Screen.TwipsPerPixelY
Next Side

Dim hBrush As Long, hRgn As Long    'Defining color filling Variables
    Select Case FillStyle
        Case 0  'Solid
            hBrush = 0
        Case 8  'Bitmap pattern
            hBrush = CreatePatternBrush(imgData)
        Case Else   'Other Patterns
            Dim Br As LOGBRUSH
            Br.lbColor = Pic.FillColor
            Br.lbHatch = FillStyle - 1
            Br.lbStyle = 2
            hBrush = CreateBrushIndirect(Br)
    End Select
    
    hRgn = CreatePolygonRgn(Vtx(1), Sides, ALTERNATE)   'Getting the polygon region from the Vertieces
    If hRgn Then FillRgn Pic.hdc, hRgn, hBrush  'Filling the determined area
    DeleteObject hRgn
    DeleteObject hBrush
    
'-------------------------------------------------------------------------
'Finally Drawing The each Side borders ( if we first try to draw the border
'then it  may covered by the shading
'-------------------------------------------------------------------------
DrawLines:
    For Side = 2 To Sides
        Pic.Line (Pnt(Side - 1).X, Pnt(Side - 1).Y)-(Pnt(Side).X, Pnt(Side).Y), vbBlack 'Drawing the sides
    Next Side
    Pic.Line (Pnt(Sides).X, Pnt(Sides).Y)-(Pnt(1).X, Pnt(1).Y), vbBlack 'Drawing the last  side and closing  the polygon
    DrawPolygon = True
Exit Function
Handle:
'MsgBox "Sorry ! .Unexpected error occured on drawing polygon." & vbCrLf & "Err." & Err.Number & " : " & Err.Description, vbCritical, "Error"
DrawPolygon = False
End Function
'-------------------------------------------------------------------------

'Getting  the Sine Value,  pie =  4 * Atn(1)
Public Function Sine(ByVal Angle As Double) As Double
    Sine = Sin(4 * Atn(1) / 180 * Angle)
End Function

'Getting  the CoSine Value,  pie =  4 * Atn(1)
Public Function Cosine(ByVal Angle As Double) As Double
    Cosine = Cos(4 * Atn(1) / 180 * Angle)
End Function

'Getting  the Tangent value. ( Tan= Sin/Cos )
Public Function Tangent(ByVal Angle As Double) As Double
On Error GoTo Handle
    Tangent = Sine(Angle) / Cosine(Angle)
Exit Function
Handle:
    Tangent = 90 'Because 'cose 90' is zero and will cause an error 'Divisible by Zero'
End Function
