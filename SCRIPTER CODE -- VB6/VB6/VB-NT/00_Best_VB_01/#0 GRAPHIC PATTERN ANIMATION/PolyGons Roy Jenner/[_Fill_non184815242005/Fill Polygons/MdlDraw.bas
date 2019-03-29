Attribute VB_Name = "mdlDraw"
'*********************************************************
'By     Jim Jose
'email  jimjosev33@yahoo.com
'*********************************************************

'It is made to get useful for anyone without 'much' changes.
'If you feel So
'   Please Rate this code
'Else
'   Give  feedback to improve  my code
'End If
'Good luck

'*********************************************************
'This code is made on request from PSC.
'Note:-
'   1)The FillRgn/CreatePolygonRgn APIs uses pixel values
'   2)You can select the 9 different styles to fill
'*********************************************************
Option Explicit

Public PicView As PictureBox


Public Pntcut() As Long
Public Pntcol() As Long
Public Pntcol2() As Long
Public Pnt() As POINTAPI
Public Pnt2() As POINTAPI
Public circx() As Double
Public circy() As Double


Public Type POINTAPI    'The co-ordinates
    X As Long
    Y As Long
End Type

'Public lop() As Pnt

Public Enum rgnFillStyle 'Fillstyle Enum
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

Public Type LOGBRUSH    'Api Type
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Const WINDING = 2 ' constants for FillMode.

'The APIs
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'This function could fill the polygon specified as the Pnts() which are the vertieces
Public Sub FillRegion(hDC As Long, Pnts() As POINTAPI, Optional ByVal FillColor As Long = 0, _
                        Optional ByVal FillStyle As rgnFillStyle = 7, Optional BitmapSource As StdPicture = -1)
Dim hBrush As Long, hRgn As Long
    Select Case FillStyle
        Case 0      'Transperant
            hBrush = 0
        Case 7     'Solid
            hBrush = CreateSolidBrush(FillColor)
        Case 8      'Bitmap
            If IsBitmap(BitmapSource) = False Then Err.Raise 1, , "Invalid Bitmap Object": Exit Sub
            hBrush = CreatePatternBrush(BitmapSource) 'Choose another source for bitmap.
        Case Else   'Hatching
            Dim Br As LOGBRUSH
            Br.lbColor = FillColor
            Br.lbHatch = FillStyle - 1
            Br.lbStyle = 2
            hBrush = CreateBrushIndirect(Br)
    End Select
        hRgn = CreatePolygonRgn(Pnts(LBound(Pnts)), UBound(Pnts) - LBound(Pnts), WINDING)
        If hRgn Then FillRgn hDC, hRgn, hBrush  'Filling the determined area
        DeleteObject hRgn
        DeleteObject hBrush
End Sub

'Returning True if the vMap is a standard picture
Private Function IsBitmap(vMap As StdPicture) As Boolean
On Error GoTo Handle
    If vMap.Height > 0 And vMap.Width > 0 Then IsBitmap = True
Exit Function
Handle:
    IsBitmap = False
End Function
