Attribute VB_Name = "modDeclares"
Option Explicit

Global Const TYPE_MANDELBROT = 0
Global Const TYPE_MANDELZPOW = 1
Global Const TYPE_JULIA = 2
Global Const TYPE_JULIAZPOW = 3
Global Const TYPE_MANDEL3 = 4
Global Const TYPE_MANDEL4 = 5
Global Const TYPE_JULIA3 = 6
Global Const TYPE_JULIA4 = 7
Global Const TYPE_MANDELRINGS = 8
Global Const TYPE_MANDELZPOWRINGS = 9
Global Const TYPE_JULIARINGS = 10
Global Const TYPE_JULIAZPOWRINGS = 11
Global Const TYPE_BARNSLEYM1 = 12
Global Const TYPE_BARNSLEYM2 = 13
Global Const TYPE_BARNSLEYM3 = 14
Global Const TYPE_BARNSLEYM4 = 15
Global Const TYPE_BARNSLEYJ1 = 16
Global Const TYPE_BARNSLEYJ2 = 17
Global Const TYPE_BARNSLEYJ3 = 18
Global Const MARKSMANDEL = 19
Global Const MARKMANDELGENERAL = 20
Global Const MARKSJULIA = 21
Global Const MARKSJULIAGENERAL = 22
Global Const TYPE_BARNSLEYMANJUL = 23
Global Const TYPE_SPIDER = 24
Global Const TYPE_MANDELLAMBDA = 25
Global Const TYPE_SIERPINSKI = 26
Global Const TYPE_NEWTON = 27
Global Const TYPE_NEWTONJULIANOVA = 28

Global Const NUM_TYPES = 28

Global Const STATE_IDLE = 0
Global Const STATE_CALC = 1
Global Const STATE_SAVE = 2
Global Const STATE_ANIM = 3

Type tComplex
    R As Double
    i As Double
End Type

Type tRGBLong
    R As Long
    G As Long
    B As Long
    Dummy(0 To 3) As Long
End Type

Type tRGB
    R As Byte
    G As Byte
    B As Byte
    Dummy As Byte
End Type

Type tImgCell
    nI As Double
    X0 As Double
    Y0 As Double
    X1 As Double
    Y1 As Double
    State As Long   '0 = outside, 1 = inside
    rX As Double
    rY As Double
End Type

Type tFilterVars
    nType As Integer           'Filter type
    IsEnabled As Boolean    'Whether the filter is currently enabled
    Var(0 To 4) As Double  'Filter parameters
    nVars As Byte
End Type

Type tFilterEntry
    nComp As Byte           'Filter Compatibility (2D, 4D, Outside, Inside)
    szName As String        'Name displayed in filter list
    nVars As Byte
End Type

Type tZoom
    SX As Double
    SY As Double
    EX As Double
    EY As Double
    nType As Byte       ' 0 = zoom | 1 = move
    InUse As Boolean
End Type

Type tWin
    wWin As Form
    dw As Long
    dH As Long
    nFrac As Long
    InUse As Boolean
    mDown As Boolean
    mBtn As Byte
    nX As Long
    nY As Long
    nX2 As Long
    nY2 As Long
End Type

Type tFractal
    'Calculation Parameters:
    SX As Double
    SY As Double
    EX As Double
    EY As Double
    BailOut As Double
    MaxI As Double
    LinearContinuity As Boolean
    Param(0 To 4) As Double
    
    'Basic Info:
    nType As Byte
    nIndex As Integer
    Win As tWin
    LinesToDraw As Long
    Fullscreen As Boolean
    CalcPriority As Byte            '0 = idle 1 = normal 2 = high 3 = realtime|not in use!!!
    Flag_State As Byte
    Flag_Stop As Boolean
    InUse As Boolean
    Calc2File As Boolean
    CellsOnDisk As Boolean          'Flag to indicate whether the data has
    szDiskID As String              'been saved temporarily to disk because of mem load
                                    '+ file name (rnd generated)
    'Image Data:
    nW As Long
    nH As Long
    nDirectFileSize As Long         'The width & height when rendering to a file
    Cell() As tImgCell
    BMP As tBitmap
    ChunkBMP As tBitmap
    BB As tMemBMP
    PalDisp As tMemBMP
    
    'Palette & filter data:
    Pal() As tRGBLong
    nPalLen As Long
    nColPeriod As Long
    UseDispersionControl As Boolean
    FilterOut() As tFilterVars
    FilterIn() As tFilterVars
    FilterMapOut  As tFilterVars
    FilterMapIn  As tFilterVars
    nFiltersOut As Long
    nFiltersIn As Long
    szFilterName() As String
    bPalAnimNormalize As Boolean
    bPalAnimNormIntervalSpecified As Boolean
    nPalAnimNormInterval As Single
    
    'Palette animation vars:
    PAnFrame As Long
    PAnFrames As Long
    PAnPalOffset As Long
    
    'Zooming Info:
    Zoom() As tZoom
    nZooms As Long
    
    szPathCalcDirect As String
    szPathImage As String
    szPathPalette As String
    szPathFractal As String
    
    SavedCalcDirect As Boolean
    SavedImage As Boolean
    SavedPalette As Boolean
    SavedFractal As Boolean
End Type

Type tFractalParam
    nValue As Double
    szDescription As String
End Type

Type tFractalType
    szName As String
    szDescription As String
    nParams As Long
    Param(0 To 9) As tFractalParam
    Template As tFractal
End Type

Type tAppData
    nActive As Long
    nFractals As Long
    
    szAppPath As String
    szTmpPath As String
    szPathPalettes As String
    szPathFractals As String
    szPathTemplates As String
    Flag_Stop As Boolean
    Flag_State As Byte
    Initializing As Boolean
    
    FractType() As tFractalType
    Filters() As tFilterEntry
    nFilters As Long
    
    Col_Prg As Long
End Type


Global Frac() As tFractal
Global aData As tAppData

Public Sub Init_Types()
Dim Tmp As tFractalType

'Dimension Fractal Type array:
ReDim aData.FractType(0 To NUM_TYPES - 1)

End Sub

Public Sub ClearType(fType As tFractalType)
With fType
    .szName = ""
    .szDescription = ""
    .nParams = 0
End With
End Sub

