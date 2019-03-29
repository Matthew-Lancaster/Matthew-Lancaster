Attribute VB_Name = "Module1"
Global FileString As String

Type POINTAPI 'used for the polygon api call
X As Long
Y As Long
End Type

Type TDPoint 'a three dimensional point
X As Double
Y As Double
z As Double
End Type

Global Const pi = 3.14159265358979 'everyone's favourite constant
Global MapArray(0 To 255, 0 To 255) As Integer 'array of points
Global HeightArray(0 To 255, 0 To 255) As Integer 'height of polygon (with water level)
Global ViewX As Integer 'position in the viewing window
Global ViewY As Integer 'position in the viewing window
'faster than PSet
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal color As Long) As Long
'quick polygons
Declare Function Polygon Lib "gdi32" _
  (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

