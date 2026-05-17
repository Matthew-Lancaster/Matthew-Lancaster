VERSION 5.00
Begin VB.Form TerrainForm 
   Caption         =   "Terrain Builder"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   564
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Options"
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CheckBox chkshallow 
      Caption         =   "Complex Shallow Water"
      Height          =   615
      Left            =   6840
      TabIndex        =   40
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Redraw"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Redraw the picture"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CheckBox chkLoRes 
      Caption         =   "Always Low-Res on Modify"
      Height          =   255
      Left            =   8040
      TabIndex        =   38
      ToolTipText     =   "use low resolution when modifying at any resolution"
      Top             =   8040
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.HScrollBar scrbeachheight 
      Height          =   255
      LargeChange     =   5
      Left            =   1560
      Max             =   255
      TabIndex        =   37
      Top             =   6000
      Value           =   10
      Width           =   855
   End
   Begin VB.HScrollBar scrwaterdiff 
      Height          =   255
      LargeChange     =   5
      Left            =   1560
      Max             =   255
      TabIndex        =   34
      Top             =   5640
      Value           =   10
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Build 3D Map"
      Height          =   375
      Left            =   2520
      TabIndex        =   33
      Top             =   7200
      Width           =   1455
   End
   Begin VB.HScrollBar scrLum 
      Height          =   255
      Left            =   5160
      Max             =   15
      Min             =   1
      TabIndex        =   30
      Top             =   8040
      Value           =   4
      Width           =   1575
   End
   Begin VB.OptionButton OptMedRes 
      Caption         =   "Medium Resolution"
      Height          =   255
      Left            =   8040
      TabIndex        =   28
      ToolTipText     =   "medium resolution is much faster but at 25% detail"
      Top             =   7320
      Width           =   1935
   End
   Begin VB.OptionButton OptLoRes 
      Caption         =   "Low Resolution"
      Height          =   255
      Left            =   8040
      TabIndex        =   27
      ToolTipText     =   "Very Low detail but real time speed"
      Top             =   7680
      Width           =   1455
   End
   Begin VB.OptionButton OptHiRes 
      Caption         =   "High Resolution"
      Height          =   255
      Left            =   8040
      TabIndex        =   26
      ToolTipText     =   "High Resolution gives finest detail but is slower"
      Top             =   6960
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.VScrollBar ZoomBar 
      Height          =   6735
      LargeChange     =   10
      Left            =   10080
      Max             =   300
      Min             =   1
      TabIndex        =   25
      Top             =   120
      Value           =   20
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Roughen"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      ToolTipText     =   "Roughen the picture"
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Randomize"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox txtRandom 
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Text            =   "0"
      Top             =   7080
      Width           =   2055
   End
   Begin VB.TextBox thresh 
      Height          =   285
      Left            =   6840
      TabIndex        =   20
      Text            =   "15"
      ToolTipText     =   "Lower threshold means higher mountains, high threshold is low plateaus"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.HScrollBar scrollZ 
      Height          =   255
      LargeChange     =   10
      Left            =   5160
      Max             =   359
      TabIndex        =   16
      Top             =   7680
      Value           =   180
      Width           =   1575
   End
   Begin VB.HScrollBar scrollY 
      Height          =   255
      LargeChange     =   10
      Left            =   5160
      Max             =   359
      TabIndex        =   15
      Top             =   7440
      Value           =   180
      Width           =   1575
   End
   Begin VB.HScrollBar scrollX 
      Height          =   255
      LargeChange     =   10
      Left            =   5160
      Max             =   359
      TabIndex        =   14
      Top             =   7200
      Value           =   180
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Draw 3D"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   7680
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   6735
      Left            =   2520
      ScaleHeight     =   445
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   12
      Top             =   120
      Width           =   7455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   1560
      Max             =   255
      TabIndex        =   10
      Top             =   5280
      Value           =   96
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Re-Normalize"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "stretch picture back to 0-255 range"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CheckBox showwork 
      Caption         =   "Show Work (slower)"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Have the Terrain builder show what it is doing in the terrain box"
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Smoothen"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "If picture is too square, smoothen it a bit"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtDetail 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "32"
      ToolTipText     =   "Higher detail is more islands and more complex landscape."
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox smooth 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "2"
      ToolTipText     =   "Smoothen the picture.  Effects the original build and the smoothen command"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Build Terrain Map"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Build the Map"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label13 
      Caption         =   "Drag in window to change 3D model's position"
      Height          =   255
      Left            =   2520
      TabIndex        =   39
      ToolTipText     =   "Left button moves, right button rotates"
      Top             =   6960
      Width           =   3495
   End
   Begin VB.Label Label12 
      Caption         =   "Beach Height: 10"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Shallow Depth: 10"
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Threshold"
      Height          =   255
      Left            =   6840
      TabIndex        =   32
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Luminance: 4"
      Height          =   255
      Left            =   3960
      TabIndex        =   31
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Alan_Buzbee@hotmail.com"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Random Seed:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "z: 180"
      Height          =   255
      Left            =   4080
      TabIndex        =   19
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "y: 180"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "x: 180"
      Height          =   255
      Left            =   4320
      TabIndex        =   17
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Water Level: 96"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "detail:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "smoothness:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
End
Attribute VB_Name = "TerrainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim RM(0 To 3, 0 To 3) As Double 'rotation matrix
Dim PointList(0 To 4) As POINTAPI 'point list for Polygon call
Dim Max As Integer 'maximum height
Dim Min As Integer 'minimum height

Dim water As Integer 'water level
Dim zoom As Double 'zoom value



Private Sub chkLoRes_Click()
DrawRef
End Sub

Private Sub chkshallow_Click()
DrawRef
End Sub

Private Sub Command1_Click()
'build terrain button
Picture1.Cls 'clear out the box
Dim smoothness As Integer 'how many times to run the smoothening loop
Dim detail As Integer 'how detailed to make the picture (less than 128)
Dim Seed As Double 'for random number processing

TerrainForm.MousePointer = 11 'change cursor to wait
If txtDetail.Text > 127 Then txtDetail.Text = 127 'limits
If txtDetail.Text < 2 Then txtDetail.Text = 2
smoothness = smooth.Text 'get values
detail = txtDetail.Text
Seed = txtRandom.Text

' fill with random pixels
For Y = 0 To detail
For X = 0 To detail
'this is a manual random number function, so if you enter the same number
'you get the same map
Seed = (Seed * 10421 + 1) Mod 65536
 MapArray(X, Y) = Int(Seed / 256) 'set Map Position to the random value
Next X
Next Y

SmoothenMap detail, smoothness 'smoothen the picture

NormalizeMap detail, Min, Max 'normalize the picture

StretchMap detail 'stretch to 128x128

DrawMap 'draw it
TerrainForm.MousePointer = 0 'restore cursor
End Sub

Private Sub SmoothenMap(ByVal detail As Integer, ByVal smoothness As Integer)
Dim temp

' smoothen the pixels by averaging them with the surrounding pixels.  The higher
' the smoothness value, the more times this loop is repeated
For i = 0 To smoothness
Min = 255 'set min and max to defaults before finding them
Max = 0 'for instance, if Min is already 0 then it won't find something lower, of course
For Y = 0 To detail
For X = 0 To detail
temp = 0 'added up value of pixels around it
temp2 = 0 'how many pixels around it
If X > 0 Then ' get the average of the four pixels around it
temp = temp + MapArray(X - 1, Y)
temp2 = temp2 + 1
End If

If X < detail Then
temp = temp + MapArray(X + 1, Y)
temp2 = temp2 + 1
End If

If Y > 0 Then
temp = temp + MapArray(X, Y - 1)
temp2 = temp2 + 1
End If

If Y < detail Then
temp = temp + MapArray(X, Y + 1)
temp2 = temp2 + 1
End If

temp = temp / temp2 'average them
MapArray(X, Y) = temp 'set to the map array

If temp > Max Then Max = temp 'log maximum and minimum values
If temp < Min Then Min = temp

'if you want to see what it's doing
If showwork.Value = 1 Then SetPixel Picture1.hdc, X, Y, RGB(temp, temp, temp)

Next X
Next Y
Next i

End Sub

Private Sub NormalizeMap(ByVal detail As Integer, ByVal minimum As Integer, ByVal maximum As Integer)
'normalize values to 0 and 255 (for full range of height)
For Y = 0 To detail
For X = 0 To detail
'this makes Min = 0, Max = 255, and everything else spread evenly between
temp = MapArray(X, Y) 'get the value
temp = temp - minimum 'set the value where minimum is 0
temp = temp * 255 / (maximum - minimum) 'multiply by new maximum and divide range
MapArray(X, Y) = temp 'replace value
Next X
Next Y

End Sub

Private Sub StretchMap(ByVal detail As Integer)
ReDim tempArray(0 To 127, 0 To 127) As Integer
Dim Blocksize As Integer 'size of the blocks to copy, inversely proportional to detail

For Y = 0 To 127 'save to a temporary array
For X = 0 To 127
tempArray(X, Y) = MapArray(X, Y)
Next X
Next Y

' stretch to 128 x 128
Blocksize = 128 / detail 'get blocksize from detail
For Y = 0 To detail
For X = 0 To detail
    For cy = 0 To Blocksize
    For cx = 0 To Blocksize
        'convert array to 128x128
        MapArray(X * Blocksize + cx, Y * Blocksize + cy) = tempArray(X, Y)
    Next cx
    Next cy
Next X
Next Y

End Sub
Private Sub DrawMap()
' draw to screen
beach = scrbeachheight.Value
waterdiff = scrwaterdiff.Value
For Y = 0 To 127
For X = 0 To 127
    temp = MapArray(X, Y)
    If temp > water + beach Then 'above sea level
        SetPixel Picture1.hdc, X, Y, RGB(0, temp, 0) 'draw green
    ElseIf temp <= water Then 'below sea level
        If water - MapArray(X, Y) <= waterdiff Then SetPixel Picture1.hdc, X, Y, RGB(0, 64, 127)   'draw lite blue
        If water - MapArray(X, Y) > waterdiff Then SetPixel Picture1.hdc, X, Y, RGB(0, 0, 127)  'draw blue
    Else 'beach
        SetPixel Picture1.hdc, X, Y, RGB(temp, temp, temp / 2) 'draw brown
    End If
Next X
Next Y

End Sub

Private Sub Command2_Click()
'smoothen map
TerrainForm.MousePointer = 11 'wait cursor
smoothness = smooth.Text 'get smooth repetitions
SmoothenMap 127, smoothness 'call smooth sub
DrawMap 'redraw
TerrainForm.MousePointer = 0 'restore cursor
End Sub

Private Sub Command3_Click()
'Re-Normalize Map Button
TerrainForm.MousePointer = 11 'wait
SmoothenMap 127, 0 'smooth sub (to find a new Min and Max)
NormalizeMap 127, Min, Max 'normalize
DrawMap
TerrainForm.MousePointer = 0 'restore
End Sub

Private Sub Command4_Click()
'Redraw Map button
DrawMap
End Sub

Private Sub Command5_Click()
'Draw 3D button
Picture2.AutoRedraw = True 'have to do this to erase EVERYTHING
Picture2.Cls
Picture2.AutoRedraw = False

Draw3D
End Sub

Private Sub Command6_Click()
'Random button
Randomize
txtRandom.Text = Int(100000 * Rnd) 'get a random value for the seed text
End Sub

Private Sub Command7_Click()
'add a bit of randomness for a more rugged look
For Y = 0 To 127
For X = 0 To 127
MapArray(X, Y) = MapArray(X, Y) + (2 * Rnd - 1) 'add a random +or- 1
If MapArray(X, Y) > 255 Then MapArray(X, Y) = 255 'upper limit
If MapArray(X, Y) < 0 Then MapArray(X, Y) = 0 'lower limit
Next X
Next Y
DrawMap
End Sub



Private Sub Command8_Click()
'send MapArray to HeightArray (have to do this for water conflicts, it's easier)
For Y = 0 To 127
For X = 0 To 127
HeightArray(X, Y) = MapArray(X, Y)
If HeightArray(X, Y) <= water Then HeightArray(X, Y) = water
Next X
Next Y
End Sub

Private Sub Command9_Click()
options.Show 'show the options form
End Sub

Private Sub Form_Load()
Randomize 'initialization
txtRandom.Text = Int(100000 * Rnd)
water = 96
zoom = 2
ViewX = 255
ViewY = 255
ViewZ = -50
End Sub

Private Sub Form_Resize()

If Me.WindowState Mod 2 = 0 Then Call TIMEROUT(True): Exit Sub
If Me.WindowState = 1 Then Call TIMEROUT(False): Exit Sub

End Sub

Sub TIMEROUT(TF)
    On Error Resume Next
    For Each Form In Forms
        For Each Control In Controls
            If Control.Interval > 0 Then Control.Enabled = TF
        Next
    Next
End Sub


Private Sub HScroll1_Change()
water = HScroll1.Value 'water level
Label3.Caption = "Water Level: " & water
End Sub

Private Sub HScroll1_Scroll()
water = HScroll1.Value 'water level
Label3.Caption = "Water Level: " & water
End Sub

Private Sub MatrixBuild(ByVal X As Double, ByVal Y As Double, ByVal z As Double)
' this sub builds the rotation matrix with x, y and z as axis angles
Dim SinX, CosX, SinY, CosY, SinZ, CosZ, C1, C2

SinX = Sin(X)
CosX = Cos(X)
SinY = Sin(Y)
CosY = Cos(Y)
SinZ = Sin(z)
CosZ = Cos(z)
'fill out the rotation matrix.  Now we can multiply any point by this matrix
'to rotate it by the current set of angles!
RM(0, 0) = (CosZ * CosY)
RM(0, 1) = (CosZ * -SinY * -SinX + SinZ * CosX)
RM(0, 2) = (CosZ * -SinY * CosX + SinZ * SinX)
RM(1, 0) = (-SinZ * CosY)
RM(1, 1) = (-SinZ * -SinY * -SinX + CosZ * CosX)
RM(1, 2) = (-SinZ * -SinY * CosX + CosZ * SinX)
RM(2, 0) = SinY
RM(2, 1) = CosY * -SinX
RM(2, 2) = CosY * CosX
End Sub

Private Function RotatePoint(ByVal X As Double, ByVal Y As Double, ByVal z As Double) As TDPoint
'rotate the point by the current matrix
Dim ZPoint As TDPoint
ZPoint.X = (X * RM(0, 0)) + (Y * RM(0, 1)) + (z * RM(0, 2)) + RM(0, 3)
ZPoint.Y = (X * RM(1, 0)) + (Y * RM(1, 1)) + (z * RM(1, 2)) + RM(1, 3)
ZPoint.z = (X * RM(2, 0)) + (Y * RM(2, 1)) + (z * RM(2, 2)) + RM(2, 3)
RotatePoint = ZPoint
End Function

Private Sub Label2_Click()

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static OldX As Integer 'last location of the mouse
Static OldY As Integer
If Button = 1 Then
ViewX = ViewX + X - OldX 'change position depending on mouse movement
ViewY = ViewY + Y - OldY
DrawRef
End If
If Button = 2 Then
scrollX.Value = scrollX.Value + OldY - Y 'rotation
scrollY.Value = scrollY.Value + OldX - X
End If

OldX = X 'update old positions
OldY = Y
End Sub

Private Sub scrbeachheight_Change()
Label12.Caption = "Beach Height: " & scrbeachheight.Value
DrawRef
End Sub

Private Sub scrbeachheight_Scroll()
Label12.Caption = "Beach Height: " & scrbeachheight.Value
DrawRef
End Sub

Private Sub scrLum_Change()
Label9.Caption = "Luminance: " & scrLum.Value
DrawRef
End Sub

Private Sub scrLum_Scroll()
Label9.Caption = "Luminance: " & scrLum.Value
DrawRef
End Sub

Private Sub scrollX_Change()
Label4.Caption = "x: " & scrollX.Value 'x angle
DrawRef
End Sub

Private Sub scrollX_Scroll()
Label4.Caption = "x: " & scrollX.Value 'x angle
DrawRef
End Sub

Private Sub scrollY_Change()
Label5.Caption = "y: " & scrollY.Value 'y angle
DrawRef
End Sub

Private Sub scrollY_Scroll()
Label5.Caption = "y: " & scrollY.Value 'y angle
DrawRef
End Sub

Private Sub scrollZ_Change()
Label6.Caption = "z: " & scrollZ.Value 'z angle
DrawRef
End Sub

Private Sub scrollZ_Scroll()
Label6.Caption = "z: " & scrollZ.Value 'z angle
DrawRef
End Sub

Private Sub DrawRef()
Dim OldState As Integer 'old option state
Picture2.AutoRedraw = True 'have to do this to prevent nasty flicker
Picture2.Cls

If OptLoRes.Value = True Or chkLoRes.Value = 1 Then 'if low res, then it's fast enough to draw the whole thing
If OptHiRes.Value = True Then OldState = 1
If OptMedRes.Value = True Then OldState = 2
OptLoRes.Value = True
Draw3D
If OldState = 1 Then OptHiRes.Value = True
If OldState = 2 Then OptMedRes.Value = True
Else 'otherwise use reference lines
'build the rotation matrix then rotate a rectangle around at the proper angle
MatrixBuild scrollX.Value * pi / 180, scrollY.Value * pi / 180, scrollZ.Value * pi / 180
Dim LineStart As TDPoint
Dim LineEnd As TDPoint
LineStart.X = -64 * zoom
LineStart.Y = -64 * zoom
LineStart.z = 0

LineEnd.X = 64 * zoom
LineEnd.Y = -64 * zoom
LineEnd.z = 0

LineStart = RotatePoint(LineStart.X, LineStart.Y, LineStart.z)
LineEnd = RotatePoint(LineEnd.X, LineEnd.Y, LineEnd.z)

Picture2.Line (LineStart.X + 255, LineStart.Y + 255)-(LineEnd.X + 255, LineEnd.Y + 255), QBColor(15)

LineStart.X = 64 * zoom
LineStart.Y = -64 * zoom
LineStart.z = 0

LineEnd.X = 64 * zoom
LineEnd.Y = 64 * zoom
LineEnd.z = 0

LineStart = RotatePoint(LineStart.X, LineStart.Y, LineStart.z)
LineEnd = RotatePoint(LineEnd.X, LineEnd.Y, LineEnd.z)

Picture2.Line (LineStart.X + 255, LineStart.Y + 255)-(LineEnd.X + 255, LineEnd.Y + 255), QBColor(15)

LineStart.X = 64 * zoom
LineStart.Y = 64 * zoom
LineStart.z = 0

LineEnd.X = -64 * zoom
LineEnd.Y = 64 * zoom
LineEnd.z = 0

LineStart = RotatePoint(LineStart.X, LineStart.Y, LineStart.z)
LineEnd = RotatePoint(LineEnd.X, LineEnd.Y, LineEnd.z)

Picture2.Line (LineStart.X + 255, LineStart.Y + 255)-(LineEnd.X + 255, LineEnd.Y + 255), QBColor(15)

LineStart.X = -64 * zoom
LineStart.Y = 64 * zoom
LineStart.z = 0

LineEnd.X = -64 * zoom
LineEnd.Y = -64 * zoom
LineEnd.z = 0

LineStart = RotatePoint(LineStart.X, LineStart.Y, LineStart.z)
LineEnd = RotatePoint(LineEnd.X, LineEnd.Y, LineEnd.z)

Picture2.Line (LineStart.X + 255, LineStart.Y + 255)-(LineEnd.X + 255, LineEnd.Y + 255), QBColor(15)
End If
Picture2.AutoRedraw = False
End Sub

Private Sub scrwaterdiff_Change()
Label11.Caption = "Shallow Depth: " & scrwaterdiff.Value
DrawRef
End Sub

Private Sub scrwaterdiff_Scroll()
Label11.Caption = "Shallow Depth: " & scrwaterdiff.Value
DrawRef
End Sub

Private Sub ZoomBar_Change()
zoom = ZoomBar.Value / 10 'zoom value
DrawRef
End Sub

Private Sub ZoomBar_Scroll()
zoom = ZoomBar.Value / 10 'zoom value
DrawRef
End Sub

Private Sub Draw3D()
Dim XPoint As TDPoint 'for rotations
Dim Light As Integer 'how light the polygon is
Dim AddedLight As Integer 'overflow light
Dim threshold As Integer 'how tall a given value is in 3D scale
Dim ResBlock As Integer 'size of a polygon
Dim Luminance As Integer 'reflective value
Dim waterdiff As Integer 'difference between water level and underwater land
Dim beach As Integer 'beach length
Luminance = scrLum.Value
If OptHiRes.Value = True Then
    ResBlock = 1 '1 polygon per point (doesn't get much better)
ElseIf OptMedRes.Value = True Then
    ResBlock = 4 '1 polygon per 4 points
Else
    ResBlock = 8 '1 polygon per 8 points
End If
beach = scrbeachheight.Value
waterdiff = scrwaterdiff.Value
threshold = thresh.Text
'build the 3D rotation matrix
MatrixBuild scrollX.Value * pi / 180, scrollY.Value * pi / 180, scrollZ.Value * pi / 180


'we start at -64 so that the middle of the map is at 0,0
'this way the map rotates around the middle and not the upper-left corner
For Y = -64 To 63 - ResBlock Step ResBlock 'make it one less so as not to get a drop-off at the edge
For X = -64 To 63 - ResBlock Step ResBlock
'rotate the three corner points and put them in an array for the polygon DLL call
XPoint = RotatePoint(X, Y, -HeightArray(X + 64, Y + 64) / threshold)
PointList(0).X = XPoint.X * zoom + ViewX
PointList(0).Y = XPoint.Y * zoom + ViewY
XPoint = RotatePoint(X + ResBlock, Y, -HeightArray(X + 64 + ResBlock, Y + 64) / threshold)
PointList(1).X = XPoint.X * zoom + ViewX
PointList(1).Y = XPoint.Y * zoom + ViewY
XPoint = RotatePoint(X + ResBlock, Y + ResBlock, -HeightArray(X + 64 + ResBlock, Y + 64 + ResBlock) / threshold)
PointList(2).X = XPoint.X * zoom + ViewX
PointList(2).Y = XPoint.Y * zoom + ViewY
XPoint = RotatePoint(X, Y + ResBlock, -HeightArray(X + 64, Y + 64 + ResBlock) / threshold)
PointList(3).X = XPoint.X * zoom + ViewX
PointList(3).Y = XPoint.Y * zoom + ViewY

' a crude shading technique
'basically it's saying that if the left side is higher than the right side of
'the polygon then the polygon is tilted away from the source, and make it a negative
'correction factor.  Otherwise, it's positive, so the polygon is brighter
OverFlowLight = 0
Light = (HeightArray(X + 64, Y + 64) - HeightArray(X + 65, Y + 65)) * Luminance + 127
If Light < 0 Then Light = 0
If Light > 255 Then
OverFlowLight = Light - 255 'catch extra light for fading to white
Light = 255
End If
'draw for land
If HeightArray(X + 64, Y + 64) > water + beach Then 'above beach
Picture2.FillColor = RGB(OverFlowLight, Light, OverFlowLight)
ElseIf HeightArray(X + 64, Y + 64) <= water Then 'under water
If water - MapArray(X + 64, Y + 64) > waterdiff Then Picture2.FillColor = RGB(0, 0, 127)
If water - MapArray(X + 64, Y + 64) <= waterdiff Then
    'shallow water math
    'advanced
    If chkshallow.Value = 1 Then Picture2.FillColor = RGB(0, waterdiff - (water - MapArray(X + 64, Y + 64)), 127)
    'simple
    If chkshallow.Value = 0 Then Picture2.FillColor = RGB(0, 64, 127)
End If
Else 'beach
Picture2.FillColor = RGB(Light, Light, Light / 2)
End If
Picture2.ForeColor = Picture2.FillColor
'check for frontal clipping and then draw
If XPoint.z * zoom > -500 Then Polygon Picture2.hdc, PointList(0), 4
Next X
Next Y

End Sub

