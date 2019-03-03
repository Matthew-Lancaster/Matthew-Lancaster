Attribute VB_Name = "DDSurface"
'---------------------------------------------------------------------------------------
' Module    : DDSurface
' DateTime  : 4/2/2005 12:27
' Author    : Jason
' Purpose   : Auxiliary form for loading surfaces
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

'For SetStretchBltMode
Const HALFTONE = 4

'Variables to get the 16-bit RGB Long
Private RedShiftLeft As Long
Private RedShiftRight As Long
Private GreenShiftLeft As Long
Private GreenShiftRight As Long
Private BlueShiftLeft As Long
Private BlueShiftRight As Long

'Type structure to set up a new surface...
Public Type tSurface
   Surface As DirectDrawSurface7  'The surface
   ddsd As DDSURFACEDESC2         'The surface description
   Transparency As Boolean        'Is it transparent?
End Type

'***BALL
Private Type UDTBall
   X As Single            'Current X
   Y As Single            'Current Y
   VelX As Single         'X Velocity
   VelY As Single         'Y Velocity
   BallRECT As RECT       'RECT for blitting
   CurrFrame As Single    'Current frame for animation
   InPlay As Boolean      'Is the ball in play?
   IsVisible As Boolean   'Is the ball visible?
   IsLost As Boolean      'Has the ball been lost?
   Count As Long          'how many balls the player has
End Type

'***PADDLE
Public Type UDTPaddle
   X As Single            'Current X
   Y As Single            'Current Y
   PaddleRECT As RECT     'RECT for blitting
   ARECT(3) As RECT       'RECT for animation
   Animate As Boolean     'Flag for animation
   CurrFrame As Single    'Current frame for animation
End Type

'***BLOCKS
Public Type UDTBlock
   X As Long              'Current X
   Y As Long              'Current Y
   BlockRECT As RECT      'RECT for blitting
   ARECT(3) As RECT       'RECT for animation
   Visible As Boolean     'Is it visible?
   Animate As Boolean     'Flag for animation
   Busted As Boolean      'Is it busted? (collision)
   PointValue As Long     'Points
   CurrFrame As Single    'Current frame for animation
End Type

'***TITLES/LABELS etc..
Public Type UDTTitle
   X As Long
   Y As Long
   TitleRECT As RECT
End Type

'DECLARE TYPES
Public Ball As UDTBall                  'RECT for the ball
Public PaddleA As UDTPaddle             'RECT for the paddle
Public Blocks(7, 9) As UDTBlock         'An 8 X 10 grid
Public BackGround As UDTTitle           'The background image
Public TitleLarge As UDTTitle           'The large opening title
Public TitleSmall As UDTTitle           'The small in-game title
Public Logo As UDTTitle                 'The noi_max logo
Public ScoreLabel As UDTTitle           'The label that says: SCORE
Public TopLabel As UDTTitle             'The label for the top score
Public BonusRECT As UDTTitle            'A bonus label for clearing the level
Public MeterRECT As UDTTitle            'For the in-game meter
Public GameOver As UDTTitle             'For the game over label
Public Menu As UDTTitle                 'The opening screen (click to begin)
Public MenuText As UDTTitle             'Text for the opening screen

'DECLARE SURFACES
Public BallSurface As tSurface          'The ball
Public PaddleSurface As tSurface        'The paddle
Public BlockSurface As tSurface         'The sheet of blocks
Public BackGroundSurface As tSurface    'The background image
Public TitleLargeSurface As tSurface    'The large opening title
Public TitleSmallSurface As tSurface    'The small in-game title
Public LogoSurface As tSurface          'The noi_max logo
Public ScoreLabelSurface As tSurface    'The label that says: SCORE
Public TopScoreSurface As tSurface      'The TOP label
Public LetterSurface As tSurface        'For the alphabet image
Public BonusLabelSurface As tSurface    'For the bonus label
Public ScoreSurface As tSurface         'For the Numbers.bmp surface
Public MeterSurface As tSurface         'For the wait meter
Public GameOverSurface As tSurface      'For the game over label
Public MenuSurface As tSurface          'The opening screen (click to begin)
Public MenuTextSurface As tSurface      'The text on the opening screen

Public PicFileArray(9) As String        'For the different pictures

Public Sub InitSurfaces()

   'Sub to initialize all of the game surfaces
   
   InitPictureArray
   LoadPictureArray 0
   
   BallSurface = CreateSurface(dd, BallSurface, App.Path & "\ImageFiles\BallSheet.bmp", False, , , lMagenta)
   PaddleSurface = CreateSurface(dd, PaddleSurface, App.Path & "\ImageFiles\PaddleA.bmp", False, , , lMagenta)
   BlockSurface = CreateSurface(dd, BlockSurface, App.Path & "\ImageFiles\BlocksA.bmp", False, , , lMagenta)
   TitleSmallSurface = CreateSurface(dd, TitleSmallSurface, App.Path & "\ImageFiles\Title2.bmp", False)
   TitleLargeSurface = CreateSurface(dd, TitleLargeSurface, App.Path & "\ImageFiles\title1.bmp", False)
   LogoSurface = CreateSurface(dd, LogoSurface, App.Path & "\ImageFiles\Logo.bmp", False)
   ScoreSurface = CreateSurface(dd, ScoreSurface, App.Path & "\ImageFiles\Numbers.bmp", False, , , lMagenta)
   TopScoreSurface = CreateSurface(dd, TopScoreSurface, App.Path & "\ImageFiles\Top.bmp", False, , , lMagenta)
   ScoreLabelSurface = CreateSurface(dd, ScoreLabelSurface, App.Path & "\ImageFiles\Score.bmp", False)
   LetterSurface = CreateSurface(dd, LetterSurface, App.Path & "\ImageFiles\Letters2.bmp", False, , , lMagenta)
   BonusLabelSurface = CreateSurface(dd, BonusLabelSurface, App.Path & "\ImageFiles\Clear.bmp", False, , , lMagenta)
   MeterSurface = CreateBlankSurface(dd, MeterSurface, 160, 10)
   GameOverSurface = CreateSurface(dd, GameOverSurface, App.Path & "\ImageFiles\GameOver.bmp", False, , , lMagenta)
   MenuSurface = CreateSurface(dd, MenuSurface, App.Path & "\ImageFiles\TitleBack.jpg", False)
   MenuTextSurface = CreateSurface(dd, MenuTextSurface, App.Path & "\ImageFiles\TitleText.gif", False, , , lMagenta)
       
End Sub

Public Sub LoadPictureArray(ByVal Index As Integer)

   If Index > 9 Then
      Index = 0
   End If
   
   Set BackGroundSurface.Surface = Nothing
   BackGroundSurface = CreateSurface(dd, BackGroundSurface, PicFileArray(Index), False, BACK_WIDTH, BACK_HEIGHT)
   
End Sub

Public Sub InitPictureArray()

   PicFileArray(0) = App.Path & "\ImageFiles\Landscape.jpg"
   PicFileArray(1) = App.Path & "\ImageFiles\Landscape2.jpg"
   PicFileArray(2) = App.Path & "\ImageFiles\Space.jpg"
   PicFileArray(3) = App.Path & "\ImageFiles\Pyramid.jpg"
   PicFileArray(4) = App.Path & "\ImageFiles\Rocks.jpg"
   PicFileArray(5) = App.Path & "\ImageFiles\Sunset.jpg"
   PicFileArray(6) = App.Path & "\ImageFiles\Water.jpg"
   PicFileArray(7) = App.Path & "\ImageFiles\River.jpg"
   PicFileArray(8) = App.Path & "\ImageFiles\Iceberg.jpg"
   PicFileArray(9) = App.Path & "\ImageFiles\Plasma.jpg"
   
End Sub

Public Function CreateSurface( _
   ByRef dd As DirectDraw7, _
   ByRef ddsurface As tSurface, _
   ByVal strFile As String, _
   Optional ByVal VideoMem As Boolean, _
   Optional ByVal StretchWidth As Long = 0, _
   Optional ByVal StretchHeight As Long = 0, _
   Optional ByVal ColorKey As Long = -1) As tSurface
    
    'This function creates a direct draw surface from any valid file format that
    'loadpicture uses, and returns the newly created surface in a tSurface type
    'structure.
    'The function has options for video memory or new (stretched) image dimensions.
    
    Dim hdcPicture As Long                 'Device context for picture
    Dim hdcSurface As Long                 'Device context for surface
    Dim Pic As StdPicture                  'stdole2 StdPicture object
    Dim picWidth As Long                   'To hold the image width
    Dim picHeight As Long                  'To hold the image height
    Dim cKey As DDCOLORKEY                 'For setting the optional color key
    
    Set Pic = LoadPicture(strFile)         'Load the bitmap
    
    'Convert the Picture object's HIMETRICS to PIXELS
    picWidth = Screen.ActiveForm.ScaleX(Pic.Width, vbHimetric, vbPixels)
    picHeight = Screen.ActiveForm.ScaleY(Pic.Height, vbHimetric, vbPixels)
    
    
    'Fill the surface description
    With ddsurface.ddsd
    
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        
        If VideoMem Then
            'Create the surface in video memory
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
        Else
            'Create the surface in system memory
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        End If
        
        If StretchWidth = 0 Then
           .lWidth = picWidth         'Use the image dimensions
        Else
           .lWidth = StretchWidth     'Use the user defined dimensions
        End If
                                                                            
        If StretchHeight = 0 Then
           .lHeight = picHeight
        Else
           .lHeight = StretchHeight
        End If
                                             
    End With
    
    'If the file provided is a bitmap...
    If Right$(strFile, 3) = "bmp" Then
    
       'Create the surface the traditional way, using the CreateSurfaceFromFile
       Set ddsurface.Surface = dd.CreateSurfaceFromFile(strFile, ddsurface.ddsd)
       
    Else
    
      'Its not a bitmap. Use the DC version
      Set ddsurface.Surface = dd.CreateSurface(ddsurface.ddsd)  'Create the surface
      hdcPicture = CreateCompatibleDC(ByVal 0&)  'Create a memory device context
      SelectObject hdcPicture, Pic.Handle        'Select the bitmap into this memory device
      ddsurface.Surface.restore                  'Restore the surface
      hdcSurface = ddsurface.Surface.GetDC       'Get the surface's DC
      SetStretchBltMode hdcSurface, HALFTONE     'Set to HALFTONE for optional image stretching
      
      'Copy from the memory device to the DirectDrawSurface
      StretchBlt hdcSurface, 0, 0, ddsurface.ddsd.lWidth, ddsurface.ddsd.lHeight, _
         hdcPicture, 0, 0, picWidth, picHeight, vbSrcCopy
         
      ddsurface.Surface.ReleaseDC hdcSurface  'Release the surface's DC
      DeleteDC hdcPicture                     'Release the memory device context
      Set Pic = Nothing                       'Release the picture object
      
    End If
       
    'set an optional color key
    If ColorKey > -1 Then
       cKey.low = ColorKey
       cKey.high = ColorKey
       ddsurface.Surface.SetColorKey DDCKEY_SRCBLT, cKey
       ddsurface.Transparency = True
    End If
    
    'Return tSurface
    CreateSurface = ddsurface
    
End Function

Public Function CreateBlankSurface(dd As DirectDraw7, ddsurface As tSurface, Width As Long, Height As Long) As tSurface

   With ddsurface.ddsd
      .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
      .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
      .lWidth = Width
      .lHeight = Height
   End With
   
   Set ddsurface.Surface = dd.CreateSurface(ddsurface.ddsd)
   
   CreateBlankSurface = ddsurface
   
End Function

Public Sub GetColorShiftValues(PrimarySurface As DirectDrawSurface7)
  
'******************************************************************************
'*****These next three procedures are used to obtain the 16-bit RGB long*******
'******************************************************************************
    Dim PixelFormat As DDPIXELFORMAT

    PrimarySurface.GetPixelFormat PixelFormat
    MaskToShiftValues PixelFormat.lRBitMask, RedShiftRight, RedShiftLeft
    MaskToShiftValues PixelFormat.lGBitMask, GreenShiftRight, GreenShiftLeft
    MaskToShiftValues PixelFormat.lBBitMask, BlueShiftRight, BlueShiftLeft
End Sub

Public Sub MaskToShiftValues(ByVal Mask As Long, ShiftRight As Long, ShiftLeft As Long)

Dim ZeroBitCount As Long
Dim OneBitCount As Long

' Count zero bits
ZeroBitCount = 0
Do While (Mask And 1) = 0
   ZeroBitCount = ZeroBitCount + 1
   Mask = Mask \ 2 ' Shift right
Loop

' Count one bits
OneBitCount = 0
Do While (Mask And 1) = 1
   OneBitCount = OneBitCount + 1
   Mask = Mask \ 2 ' Shift right
Loop

' Shift right 8-OneBitCount bits
ShiftRight = 2 ^ (8 - OneBitCount)
' Shift left ZeroBitCount bits
ShiftLeft = 2 ^ ZeroBitCount

End Sub

Public Function DDRGB(Red As Long, Green As Long, Blue As Long) As Long
   
   DDRGB = (Red \ RedShiftRight) * RedShiftLeft + _
   (Green \ GreenShiftRight) * GreenShiftLeft + _
   (Blue \ BlueShiftRight) * BlueShiftLeft
    
End Function

Public Sub KillSurfaces()
   
   Set BallSurface.Surface = Nothing
   Set PaddleSurface.Surface = Nothing
   Set BlockSurface.Surface = Nothing
   Set TitleSmallSurface.Surface = Nothing
   Set TitleLargeSurface.Surface = Nothing
   Set BackGroundSurface.Surface = Nothing
   Set LogoSurface.Surface = Nothing
   Set ScoreSurface.Surface = Nothing
   Set ScoreLabelSurface.Surface = Nothing
   Set TopScoreSurface.Surface = Nothing
   Set LetterSurface.Surface = Nothing
   Set BonusLabelSurface.Surface = Nothing
   Set MeterSurface.Surface = Nothing
   Set GameOverSurface.Surface = Nothing
   Set MenuSurface.Surface = Nothing
   Set MenuTextSurface.Surface = Nothing

End Sub

