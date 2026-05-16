Attribute VB_Name = "modPong"
'---------------------------------------------------------------------------------------
' Module    : modPong
' DateTime  : 4/2/2005 12:27
' Author    : Jason
' Purpose   : The huge Pong module that handles game loops and rendering
'---------------------------------------------------------------------------------------

Option Explicit

Dim F As Double

Public SCORE As Long                     'The running score for the player
Public TOP_SCORE As Long                 'The top score
Private BeatScore As Boolean             'Has the top score been beat?
Private BonusPoints As Long              'Keep track of bonus points
Private Multiplier As Long               'A multiplier for the bonus
Public ScoreText(9) As RECT              'For the Numbers.bmp                  (SCORE)
Public NumberPosition(5) As Long         'For the position of the numbers      (SCORE)
Public NumberValue(5) As String          'For the value of individual numbers  (SCORE)
Private LetterRECT(27) As RECT           'For letters
Public DrawingMeter As Boolean           'Are we refreshing the meter?
Private DrawingBonus As Boolean          'Are we drawing the bonus?
Public DrawingMenu As Boolean            'Are we busy drawing the opening screen?
Public GameIsOver As Boolean             'Is it time to declare Game Over?
Public bFPS As Boolean                   'Are we drawing the FPS?
Private Counter As Single                'A counter for waiting through title and blank screens
Public BallStick As Boolean              'Is the ball supposed to be sticking to the paddle?
Public FireStraightUp As Boolean         'Did we fire the ball from the paddle? (mouse click)
Public PaddleReceive As Boolean          'Are we waiting to capture the ball?
Private LevelClear As Boolean            'Set this to true when the level is cleared
Private BLocksRemaining As Long          'for debug
Private ExtraAwarded As Boolean          'For extra ball

Private Blocks_Tall As Long              'Keep track of how many blocks tall
                                         'This is used after the constant BLOCKS_HIGH
                                         'sets up the initial value
'Frame rate counter
Dim NumLoops As Integer
Dim FPS As Double

Public Const BALL_SQUARE As Long = 18    'Size of the ball in pixels (square
Public Const PADDLE_WIDTH As Long = 62   'Paddle width in pixels
Public Const PADDLE_HEIGHT As Long = 17  'Paddle height in pixels
Public Const BLOCK_WIDTH As Long = 64 '32 '63'64    'Individual block width in pixels
Public Const BLOCK_HEIGHT As Long = 16 '8 '16   'Individual block height in pixels
Private Const BLOCKS_WIDE As Long = 7 '14 '7    'Number of columns to start game with
Private Const BLOCKS_HIGH As Long = 8 '12 '8    'Number of rows to start game with
Public Const BLOCK_OFFSET As Long = 64   'Offset to begin drawing group of blocks
Public Const BACK_OFFSET As Long = 16    'Offset to draw the picture background
Public Const BACK_WIDTH As Long = 526    'Width of the picture background
Public Const BACK_HEIGHT As Long = 600 - BACK_OFFSET * 2   'Height of the picture background
Public Const TITLESMALL_X As Long = 560  'X-position of the in-game title graphic
Private Const LABEL_HEIGHT As Long = 16  'The SCORE label
Private Const LABEL_WIDTH As Long = 80   'The SCORE label

Private Const TOP_X As Long = 572
Private Const TOP_Y As Long = 300
Private Const TOP_SCORE_Y As Long = 316

Private Const LABEL_Y As Long = 140      'The SCORE label
Private Const LABEL_X As Long = 574      'THE SCORE label
Private Const CHAR_WH As Long = 16       'Character size for each letter/number
Public Const POINTS As Long = 100        'What each block is worth in points
Public Const BONUS_POINTS As Long = 4000 'Where the bonus points begin
Private Const EXTRA_BALL As Long = 80000 'Points needed for an extra ball
Public Const SCORE_STARTX As Long = 560
Public Const SCORE_STARTY As Long = 156
Public Const BALLSPEED As Single = -0.7  'Initial ball speed
Public Const FIREVEL As Single = -0.2
Public Const INCREASE As Single = 0.2 '0.01   'Speed increase for ball
Public Const MAXSPEED As Single = 5 '0.8    'Maximum ball speed  (if set too high can mess up collisions)
Public Const BLANK_TIME As Long = 200 '2000   'Time to blit blank screen
Public Const TITLE_TIME As Long = 500 '5000   'Time to blit title

'Skip title for dev mode
'Private Const TITLE_SKIP As Boolean = True
Private Const TITLE_SKIP As Boolean = False
Private Const BALL_LOSE As Boolean = False

'InitBallToPaddle

Public Sub RunGame()

   InitDX             'Start up DirectDraw
   InitSurfaces       'Load surfaces
   InitSound
   InitPositions      'Set initial positions
   ReadScore          'Read the top score
   
  'For timer function
   QueryPerformanceCounter TimeStart
   QueryPerformanceFrequency PerfFreq
   
   'A 'drag' loop before rendering
   Do While Elapsed > 30
      QTimer
      myDoEvents 0
   Loop
   
   'For dev mode. This will skip the opening title screens
   'if the constant TITLE_SKIP at the top of this module is
   'set to True
   
   If TITLE_SKIP = False Then
   
      'Each of these procedures falls under it's own
      'Do/Loop structure outside of the main game loop
      
      'Draws a blank screen
      DrawIntro 1, True, TitleLarge
      
      'Draws the noi_max logo
      DrawIntro 3, False, Logo, LogoSurface.Surface
      
      'Draws a blank screen
      DrawIntro 1, True, TitleLarge
      
      'Draws the BreakoutDX logo
      PlayVocoder
      DrawIntro 4, False, TitleLarge, TitleLargeSurface.Surface

   End If
   
   DrawMenu
    


End Sub

Private Sub GameLoop()

   'Start with a visible ball that's stuck to the paddle.
   'Ball count suggests how many balls the player starts with.
   
   Ball.InPlay = True     'Enables the MovePongObjects sub
   Ball.IsVisible = True  'The ball is visible... :-\
   Ball.Count = 3         'Start with 3 balls
   BallStick = True 'False 'True       'Stick the ball to the paddle
   Multiplier = 1         'Start the bonus multiplier at 1
   'Enable loop
   bRunning = True
   
   'Run the breakout game.
   
'   Call frmDX.Form_MouseMove(0, 0, 100, 0)
'   Call frmDX.Form_MouseDown(0, 0, 0, 0)
'   Call frmDX.Form_MouseUp(0, 0, 0, 0)
   
   Do While bRunning = True
   
      'Call timer to get elapsed time.
      QTimer
      
      If Ball.InPlay = True And LevelClear = False Then
      
         'The ball is in play. Do the MovePongObjects
         'sub to move the ball and check for collisions.
         MovePongObjects
         
      ElseIf LevelClear = True Then
      
         'The current level was cleared. Draw a new set of blocks
         'one row taller
         BallStick = True 'False 'True
         InitBallToPaddle
         PaddleA.Animate = True
         DrawBonusLabel
         
      ElseIf Ball.IsLost = True And Ball.Count > 0 Then
      
         'The ball has passed the paddle and gone off the
         'bottom of the screen. Regenerate the ball.
         DoBallLost
         
      ElseIf Ball.IsLost = True And Ball.Count = 0 Then
      
         'The ball has gone past the paddle and we're out
         GameIsOver = True
         DoGameOver
         
      End If
      
      'Handle the rendering.
      
      CheckForLostSurfaces     'Check for lost surfaces.
      FillBackGround           'Fill the backbuffer with the default color.
      DrawPongSurfaces         'Draw all of the blocks, paddle, background and ball(s).
      DoMeter                  'Check for and animate the meter
      DoFPS                    'Calculate FPS.
      'DrawText                 'Draw debug text.
      
      'Render FPS meter if the user hits the 'F' key
      If bFPS Then
         drawFPS
      End If
      
      Flip                     'Flip the whole mess to the primary surface.
      myDoEvents (0)           'Peekmessage/TranslateMessage/DispatchMessage
      
   Loop

End Sub

Public Sub InitPositions()

   'Set coordinates for everything using the X/Y and RECT structure
   
   With Ball
      .VelX = BALLSPEED
      .VelY = BALLSPEED
   End With
   
   With PaddleA
      .X = 0
      .Y = BACK_HEIGHT + BACK_OFFSET - (PaddleSurface.ddsd.lWidth * 2)
      .PaddleRECT.Top = 0
      .PaddleRECT.Left = 0
      .PaddleRECT.Bottom = PaddleSurface.ddsd.lHeight
      .PaddleRECT.Right = PaddleSurface.ddsd.lWidth
   End With
   
   With BackGround
      .X = BACK_OFFSET
      .Y = BACK_OFFSET
      .TitleRECT.Top = 0
      .TitleRECT.Left = 0
      .TitleRECT.Bottom = BACK_HEIGHT
      .TitleRECT.Right = BACK_WIDTH
   End With
   
   With TitleSmall
      .X = TITLESMALL_X
      .Y = BACK_OFFSET
      .TitleRECT.Top = 0
      .TitleRECT.Left = 0
      .TitleRECT.Bottom = TitleSmallSurface.ddsd.lHeight
      .TitleRECT.Right = TitleSmallSurface.ddsd.lWidth
   End With
   
   With TitleLarge
      .X = (DX_WIDTH / 2) - (TitleLargeSurface.ddsd.lWidth / 2)
      .Y = (DX_HEIGHT / 2) - (TitleLargeSurface.ddsd.lHeight / 2)
      .TitleRECT.Top = 0
      .TitleRECT.Left = 0
      .TitleRECT.Right = TitleLargeSurface.ddsd.lWidth
      .TitleRECT.Bottom = TitleLargeSurface.ddsd.lHeight
   End With
   
   With Logo
      .X = (DX_WIDTH / 2) - (LogoSurface.ddsd.lWidth / 2)
      .Y = (DX_HEIGHT / 2) - (LogoSurface.ddsd.lHeight / 2)
      .TitleRECT.Bottom = LogoSurface.ddsd.lHeight
      .TitleRECT.Right = LogoSurface.ddsd.lWidth
   End With
   
   With BonusRECT
      .X = (BACK_WIDTH / 2) - (BonusLabelSurface.ddsd.lWidth / 2) + BACK_OFFSET
      .Y = (BACK_HEIGHT / 2) - (BonusLabelSurface.ddsd.lHeight / 2) + BACK_OFFSET
      .TitleRECT.Bottom = BonusLabelSurface.ddsd.lHeight
      .TitleRECT.Right = BonusLabelSurface.ddsd.lWidth
   End With
      
   'PADDLE-A
   PaddleA.X = 0
   PaddleA.Y = BACK_HEIGHT + BACK_OFFSET - (PADDLE_HEIGHT * 2)
   
   Dim i As Integer
   
   For i = 0 To 3
      PaddleA.ARECT(i).Top = 0
      PaddleA.ARECT(i).Left = PADDLE_WIDTH * i
      PaddleA.ARECT(i).Bottom = PADDLE_HEIGHT
      PaddleA.ARECT(i).Right = PaddleA.ARECT(i).Left + PADDLE_WIDTH
   Next i
   
      'Set up the score number layout (number of digits)
   For i = 0 To 9
      ScoreText(i).Left = CHAR_WH * i
      ScoreText(i).Right = ScoreText(i).Left + CHAR_WH
      ScoreText(i).Top = 0
      ScoreText(i).Bottom = 17
   Next i
   
   'Set up the score label position
   With ScoreLabel
      .X = LABEL_X
      .Y = LABEL_Y
      .TitleRECT.Top = 0
      .TitleRECT.Left = 0
      .TitleRECT.Bottom = LABEL_HEIGHT
      .TitleRECT.Right = LABEL_WIDTH
   End With
   
   With TopLabel
      .X = TOP_X
      .Y = TOP_Y
      .TitleRECT.Right = TopScoreSurface.ddsd.lWidth
      .TitleRECT.Bottom = TopScoreSurface.ddsd.lHeight
   End With
   
   'Set up letters
   For i = 0 To UBound(LetterRECT)
      LetterRECT(i).Left = CHAR_WH * i
      LetterRECT(i).Right = LetterRECT(i).Left + CHAR_WH
      LetterRECT(i).Top = 0
      LetterRECT(i).Bottom = CHAR_WH
   Next i
   
   With MeterRECT
      .TitleRECT.Right = MeterSurface.ddsd.lWidth
      .TitleRECT.Bottom = MeterSurface.ddsd.lHeight
   End With
   
   With GameOver
      .X = (BACK_WIDTH / 2) - (GameOverSurface.ddsd.lWidth / 2) + BACK_OFFSET
      .Y = (BACK_HEIGHT / 2) - (GameOverSurface.ddsd.lHeight / 2) + BACK_OFFSET
      .TitleRECT.Bottom = GameOverSurface.ddsd.lHeight
      .TitleRECT.Right = GameOverSurface.ddsd.lWidth
   End With
   
   With Menu
      .X = (DX_WIDTH / 2) - (MenuSurface.ddsd.lWidth / 2)
      .Y = (DX_HEIGHT / 2) - (MenuSurface.ddsd.lHeight / 2)
      .TitleRECT.Top = 0
      .TitleRECT.Left = 0
      .TitleRECT.Right = MenuSurface.ddsd.lWidth
      .TitleRECT.Bottom = MenuSurface.ddsd.lHeight
   End With
   
   With MenuText
      .X = (DX_WIDTH / 2) - (MenuTextSurface.ddsd.lWidth / 2)
      .Y = (DX_HEIGHT / 2) - (MenuTextSurface.ddsd.lHeight / 2)
      .TitleRECT.Top = 0
      .TitleRECT.Left = 0
      .TitleRECT.Right = MenuTextSurface.ddsd.lWidth
      .TitleRECT.Bottom = MenuTextSurface.ddsd.lHeight
   End With
   
   'Set up the blocks
   Blocks_Tall = BLOCKS_HIGH
   InitBlocks Blocks_Tall
   
End Sub

Public Sub MovePongObjects()

   Dim VCollision As Boolean   'Local variables for collision
   Dim HCollision As Boolean
   Dim ColCounter As Long
   
   'Local loop variables
   Dim i As Integer, j As Integer

   'Skip this sub if there's a major delay
   If Elapsed > 30 Then Exit Sub
   
   'See if the ball is supposed to be stuck to the paddle
   If BallStick = True Then
   
      'Set the ball in the middle of the paddle
      InitBallToPaddle
      
      'Animate the paddle
      PaddleA.Animate = True
      
      'Skip the rest of this sub while the ball is in sticky mode
      Exit Sub
      
   End If

   'Move ball horizontally and check for collisions
   
   'If the ball is being fired from the paddle then there's no
   'X velocity, so check to see before moving.
   If FireStraightUp = False Then
      Ball.X = Ball.X + (Ball.VelX * Elapsed)
   End If
   
   'Loop through the blocks and check for a collision
   For i = 0 To BLOCKS_WIDE
      For j = 0 To Blocks_Tall '<-- variable, not a constant!
      
         'If the ball hit a block...
         If Ball.X < Blocks(i, j).X + BLOCK_WIDTH And Blocks(i, j).X < Ball.X + BALL_SQUARE Then
            If Ball.Y < Blocks(i, j).Y + BLOCK_HEIGHT And Blocks(i, j).Y < Ball.Y + BALL_SQUARE Then
            
               'Is the block already busted?
               If Blocks(i, j).Busted = False Then
               
                  'Collision!
                  HCollision = True
                  SCORE = SCORE + (POINTS * Abs(CInt(Ball.VelX * 10)))
                  
                  PlayGoalSound
                  
                  'If we fired straight up...
                  FireStraightUp = False
                  
                  'Animate the block (to see it break apart)
                  Blocks(i, j).Animate = True
                  
                  'Block is now busted!
                  Blocks(i, j).Busted = True
                  
               End If
            End If
         End If
      Next j
   Next i
         
   'Check edges of Background
   If Ball.X + BALL_SQUARE >= BACK_WIDTH + BACK_OFFSET Or Ball.X < 0 + BACK_OFFSET Then
      HCollision = True
      PLayBipSound
   End If
   
   'Check Paddle
   If Ball.X < PaddleA.X + PADDLE_WIDTH And PaddleA.X < Ball.X + BALL_SQUARE Then
      If Ball.Y < PaddleA.Y + PADDLE_HEIGHT And PaddleA.Y < Ball.Y + BALL_SQUARE Then
      
         PLayBipSound
         ColCounter = ColCounter + 1
         
         'Are we in receive mode (waiting for the ball to stick) ??
         If PaddleReceive = True Then
         
            'It hit the paddle... make it stick
            BallStick = True
            
            'Switch boolean
            PaddleReceive = False
            
            'Count remaining blocks
            If CountRemaining Then Exit Sub
            
            'Increase ball speed
            IncreaseSpeed
            
            'Skip the rest of this sub while in sticky mode
            Exit Sub
         End If
         
         If CountRemaining Then Exit Sub
         IncreaseSpeed
         HCollision = True
      End If
   End If
   
   If FireStraightUp = False Then
      Ball.Y = Ball.Y + (Ball.VelY * Elapsed)
   Else
      'Minus FIREVEL to the Ball.VelY (making a greater negative value)
      Ball.Y = Ball.Y + ((FIREVEL - Abs(Ball.VelY)) * Elapsed)
   End If
   
   For i = 0 To BLOCKS_WIDE
      For j = 0 To Blocks_Tall '<-- variable, not a constant!
      
         If Ball.X < Blocks(i, j).X + BLOCK_WIDTH And Blocks(i, j).X < Ball.X + BALL_SQUARE Then
            If Ball.Y < Blocks(i, j).Y + BLOCK_HEIGHT And Blocks(i, j).Y < Ball.Y + BALL_SQUARE Then
               If Blocks(i, j).Busted = False Then
                  'If counter < 1 Then
                     VCollision = True
                     Blocks(i, j).Busted = True
                     Blocks(i, j).Animate = True
                     SCORE = SCORE + (POINTS * Abs(CInt(Ball.VelX * 10)))
                     PlayGoalSound
                     FireStraightUp = False
                  'End If
               End If
            End If
         End If
      Next j
   Next i
   
   If Ball.Y < BACK_OFFSET Then
   
      PLayBipSound
      
      If FireStraightUp = True Then
         'Move the ball away from the collision point using the Abs(FIREVEL)
         Ball.Y = Ball.Y + ((Abs(FIREVEL) + Abs(Ball.VelY)) * Elapsed)
         FireStraightUp = False
      Else
         VCollision = True
      End If
   End If
   
   'If the ball tries to go beyond the bottom of the player's screen...
   If Ball.Y + BALL_SQUARE >= BACK_HEIGHT + BACK_OFFSET Then
   
      If Ball.Count > 1 Then
         'PlayMissSound
      Else
         PlayGameOverSound
      End If
      
      'If the dev constant calls for losing the ball...
      If BALL_LOSE = True Then
         Ball.InPlay = False
         Ball.IsVisible = False
         Ball.IsLost = True
         Ball.Count = Ball.Count - 1
      Else
         'The ball just bounces off the bottom...
         VCollision = True
         If CountRemaining Then Exit Sub
      End If
      
   End If
   
      'Check Paddle
   If Ball.X < PaddleA.X + PADDLE_WIDTH And PaddleA.X < Ball.X + BALL_SQUARE Then
      If Ball.Y < PaddleA.Y + PADDLE_HEIGHT And PaddleA.Y < Ball.Y + BALL_SQUARE Then
      
         PLayBipSound
         ColCounter = ColCounter + 1
         
         If PaddleReceive = True Then
            PaddleA.Animate = True
            BallStick = True
            PaddleReceive = False
            If CountRemaining Then Exit Sub
            IncreaseSpeed
            Exit Sub
         End If
         If CountRemaining Then Exit Sub
         IncreaseSpeed
         VCollision = True
      End If
   End If
   
   If HCollision Then
      Ball.VelX = -Ball.VelX
      Ball.X = Ball.X + (Ball.VelX * Elapsed)
   End If
   
   If VCollision Then
      Ball.VelY = -Ball.VelY
      Ball.Y = Ball.Y + (Ball.VelY * Elapsed)
   End If
   
End Sub

Public Sub DrawPongSurfaces()

   Dim i As Integer, j As Integer

   'Do the animation for the blocks and paddle
   AnimatePaddle
   AnimateBlocks
   
   'Blit the TitleSmall
   Render TitleSmall.X, TitleSmall.Y, TitleSmallSurface.Surface, TitleSmall.TitleRECT, False
   
   'Blit the background
   Render BackGround.X, BackGround.Y, BackGroundSurface.Surface, _
      BackGround.TitleRECT, False
      
   'Blit the paddle (using the current frame for animation)
   Render PaddleA.X, PaddleA.Y, PaddleSurface.Surface, _
      PaddleA.ARECT(CInt(PaddleA.CurrFrame)), True
      
   'Blit the ball************************
   
   'If the ball has just been fired, make it have a RED trail
   If Ball.IsVisible = True Then
   
      If FireStraightUp = True Then
      
         With Ball.BallRECT
            .Top = 0
            .Left = BALL_SQUARE * 2 '<- right ball image in the ballsheet
            .Bottom = BALL_SQUARE
            .Right = .Left + BALL_SQUARE
         End With
         
         For i = 8 To 1 Step -4
            Render Ball.X, (Ball.Y + i), BallSurface.Surface, Ball.BallRECT, True
         Next i
         
         'Render the top-most ball in the trail BLUE
         With Ball.BallRECT
            .Top = 0
            .Left = BALL_SQUARE
            .Bottom = BALL_SQUARE  '<- middle of the ballsheet
            .Right = .Left + BALL_SQUARE
         End With
         Render Ball.X, Ball.Y, BallSurface.Surface, Ball.BallRECT, True
         
      'If the ball is stuck to the paddle or in 'receive' mode make it BLUE
      ElseIf PaddleReceive = True Or BallStick = True Then
         With Ball.BallRECT
            .Top = 0
            .Left = BALL_SQUARE '<- Middle of the ballsheet
            .Bottom = BALL_SQUARE
            .Right = .Left + BALL_SQUARE
         End With
         Render Ball.X, Ball.Y, BallSurface.Surface, Ball.BallRECT, True
         
      'Otherwise it's GREEN
      Else
         With Ball.BallRECT
            .Top = 0
            .Left = 0  '<-left image in the ballsheet
            .Bottom = BALL_SQUARE
            .Right = BALL_SQUARE
         End With
         Render Ball.X, Ball.Y, BallSurface.Surface, Ball.BallRECT, True
      End If
   
   End If
      
   'Loop through the blocks and blit them.
   For i = 0 To 7
      For j = 0 To 9
         
         'Should the block be visible?
         If Blocks(i, j).Visible = True Then
         
            'Blit based on current frame for animation
            Render Blocks(i, j).X, Blocks(i, j).Y, _
               BlockSurface.Surface, Blocks(i, j).ARECT(CInt(Blocks(i, j).CurrFrame)), _
               True
               
         End If
         
      Next j
   Next i
   
   'If it's time to draw the bonus label, draw it on top of the game screen
   If DrawingBonus = True Then
       Render BonusRECT.X, BonusRECT.Y, BonusLabelSurface.Surface, BonusRECT.TitleRECT, True
       DrawScore CStr(BonusPoints), BonusRECT.X, BonusRECT.Y + BonusLabelSurface.ddsd.lHeight
   End If
   
   'If the player loses the last ball, render the Game Over label
   If GameIsOver Then
      Render GameOver.X, GameOver.Y, GameOverSurface.Surface, GameOver.TitleRECT, True
   End If
   
   'Draw the score
   DrawScore CStr(SCORE), SCORE_STARTX, SCORE_STARTY
   
   'Draw the label
   Render ScoreLabel.X, ScoreLabel.Y, ScoreLabelSurface.Surface, ScoreLabel.TitleRECT, False
   
   'Draw Top score label
   Render TopLabel.X, TopLabel.Y, TopScoreSurface.Surface, TopLabel.TitleRECT, True
   
   'Draw the Top score
   If SCORE > TOP_SCORE Then
      BeatScore = True
      DrawScore CStr(SCORE), SCORE_STARTX, TOP_SCORE_Y
   Else
      DrawScore CStr(TOP_SCORE), SCORE_STARTX, TOP_SCORE_Y
   End If
   
   'Draw remaining balls
   For i = 0 To Ball.Count - 2
      Render 560 + BALL_SQUARE * i, 200, BallSurface.Surface, Ball.BallRECT, True
   Next i
   
   'Draw the bar that represents the meter
   If DrawingMeter = False Then
      MeterSurface.Surface.BltColorFill MeterRECT.TitleRECT, DDRGB(133, 132, 206)
   Else
      MeterSurface.Surface.BltColorFill MeterRECT.TitleRECT, DDRGB(132, 206, 156)
   End If
   
   Render 560, 220, MeterSurface.Surface, MeterRECT.TitleRECT, False
        
End Sub

Private Sub AnimatePaddle()

   'Make the paddle flash when the ball hits it.

   Const FRAMES As Single = 0.02    '20 frames a second
   Const MAX_FRAME As Long = 4      'Don't go beyond 4 for frames
   Const FIRST_FRAME As Long = 0    'First frame is the normal paddle
   
   'If it's time to animate
   If PaddleA.Animate = True Then
   
      'Advance frame
      PaddleA.CurrFrame = PaddleA.CurrFrame + (FRAMES * Elapsed)
      
      'If we're at or beyond 4
      If CInt(PaddleA.CurrFrame) >= MAX_FRAME Then
      
         'Go back to the first frame (normal paddle)
         PaddleA.CurrFrame = FIRST_FRAME
         
         'We're done animating
         PaddleA.Animate = False
         
      End If
   End If
      
End Sub

Private Sub AnimateBlocks()

   Const FRAMES As Single = 0.03  '60 frames per second
   Const MAX_FRAME As Long = 4    'Dont go past 4 for frames
   Const FIRST_FRAME As Long = 0  'Start frame is 0
   
   Dim i As Integer, j As Integer
      
   For i = 0 To 7
      For j = 0 To 9
         
         'Is it time to animate the block?
         If Blocks(i, j).Animate = True Then
         
            'Advance frame
            Blocks(i, j).CurrFrame = Blocks(i, j).CurrFrame + (FRAMES * Elapsed)
            
            'If we're at or beyond the max frame
            If CInt(Blocks(i, j).CurrFrame) >= MAX_FRAME Then
            
               'Go back to the first frame
               Blocks(i, j).CurrFrame = FIRST_FRAME
               
               'We're done animating
               Blocks(i, j).Animate = False
               
               'The block is now invisible
               Blocks(i, j).Visible = False
               
            End If
         End If
      Next j
   Next i
         
End Sub

Private Sub DrawIntro(ByVal NumSeconds As Single, _
   ByVal Blank As Boolean, _
   ByRef Title As UDTTitle, _
   Optional ByVal Surface As DirectDrawSurface7)


   'Draw a 'Title' screen for a period of time
   Dim DrawingLogo As Boolean
   
   'switch boolean
   DrawingLogo = True
   
   Do While DrawingLogo = True
      
      'Call timer
      QTimer
      
      'Check for lost surfaces (alt-tab)
      CheckForLostSurfaces
      
      'Fill the backbuffer with it's default solid color
      FillBackGround
      
      If Blank = False Then
         'Render the 'Title' surface

         Render Title.X, Title.Y, Surface, Title.TitleRECT, False

      End If
      
      'Flip to screen
      Flip
      
      'Check the time elapsed
      If Elapsed < 30 Then
         Counter = Counter + (1 * Elapsed)
      End If
      
      'If time's up, switch the boolean and get out of the loop
      If Counter > NumSeconds * 1000 Then
         DrawingLogo = False
         Counter = 0
      End If
      myDoEvents 0
   Loop
   
End Sub

Private Sub DrawMenu()

   
   'Start the game loop if the user has clicked
'   GameLoop
'
'
'Exit Sub
'
   
   DrawingMenu = True
   'DrawingMenu =
   PlayBeatz
   
   Do While DrawingMenu
      
      'Check for lost surfaces (alt-tab)
      CheckForLostSurfaces
      
      'Fill the backbuffer with it's default solid color
      FillBackGround
      
      Render Menu.X, Menu.Y, MenuSurface.Surface, Menu.TitleRECT, False, True
      
      Render MenuText.X, 0, MenuTextSurface.Surface, MenuText.TitleRECT, True
      
      Render (DX_WIDTH / 2) - (TopScoreSurface.ddsd.lWidth / 2), 500, TopScoreSurface.Surface, TopLabel.TitleRECT, True
      DrawScore CStr(TOP_SCORE), (DX_WIDTH / 2) - (96 / 2), 532
      'Flip to screen
      Flip
      
      F = Now
      Do
      DoEvents
      Loop Until F < Now + TimeSerial(0, 0, 4)
      
      
      
      
        Exit Do
      
   Loop
   
   StopBeatz
   
   PlayFilter
   'Draw a blank screen
   DrawIntro 3, True, TitleLarge
   
   
   
   'Start the game loop if the user has clicked
   GameLoop
   
   
End Sub

Private Sub DrawText()

    'Draw debug text
    Call Backbuffer.DrawText(544, 400, _
       DX_WIDTH & "x" & DX_HEIGHT & "x" & PixFormat.lRGBBitCount & _
       "  Frames per Second " & Format$(FPS, "#.0"), False)
         
    Call Backbuffer.DrawText(544, 430, (POINTS * Abs(CInt(Ball.VelX * 10))), False)
    Call Backbuffer.DrawText(544, 460, "Multiplier: " & Multiplier, False)
    Call Backbuffer.DrawText(544, 490, "Elapsed: " & Elapsed, False)
    
End Sub

Private Sub DoFPS()

   'Calculate FPS
   NumLoops = NumLoops + 1
   If NumLoops = 9 Then
      QueryPerformanceCounter TimeEnd
      FPS = PerfFreq / (TimeEnd - TimeStart) * 10
      QueryPerformanceCounter TimeStart
      NumLoops = 0
   End If
      
End Sub

Private Sub drawFPS()

   DrawScore CStr(FPS), 500, 560
   
End Sub

Private Sub DoBallLost()

   'If the ball is lost by the player; i.e. it goes past the
   'paddle and beyond the bottom of the screen, do a delay
   'of about one second and then 'regenerate' the ball by resetting
   'some of it's variables.
   Static DelayTime As Long
    
   DelayTime = DelayTime + (1 * Elapsed)
   
   If DelayTime > 1000 Then    'About 1 second
      InitBallToPaddle         'Place the ball on top of the paddle
      BallStick = True         'Ball is stuck to the paddle, ready for firing
      Ball.InPlay = True       'Not only can you see the ball, but the movement/collision sub is also ready
      Ball.IsVisible = True
      Ball.IsLost = False
      DelayTime = 0            'Reset delay time
      'reset ball velocities
      Ball.VelX = BALLSPEED
      Ball.VelY = BALLSPEED
   End If

End Sub

Private Sub DoGameOver()

   'This allows the Game over label to render for about 3 seconds
   Static DelayTime As Long
   
   DelayTime = DelayTime + (1 * Elapsed)
   
   If DelayTime > 3000 Then
   
      bRunning = False          'Turn off the main game loop
      GameIsOver = False        'Switch the game over boolean
      
      'If the top score was beaten, write to file
      If BeatScore Then
         WriteScore SCORE
         TOP_SCORE = SCORE
         BeatScore = False
      End If
      
      DelayTime = 0             'Reset the delay time
      InitPositions             'Re-initialize positions
      SCORE = 0                 'Reset score
      Multiplier = 1            'Reset Multiplier
      ExtraAwarded = False
      LoadPictureArray (0)      'Go back to the first picture
      DrawMenu                  'Draw the "click to begin" screen
      
   End If
   
End Sub

Private Sub InitBallToPaddle()

      'Place the ball on top of the paddle in the center.
      With Ball
         .X = (PaddleA.X + PADDLE_WIDTH / 2) - BALL_SQUARE / 2
         .Y = PaddleA.Y - BALL_SQUARE
      End With
      
      BallStick = True
      BallStick = False
      
End Sub

Private Function CountRemaining() As Boolean

   'Sub to count how many blocks remain...
   Dim i As Long, j As Long
   Dim BlockCounter As Long
   
   'Loop through entire block array and see if any are visible
   For i = 0 To 7
      For j = 0 To 9
         If Blocks(i, j).Visible = True Then
            BlockCounter = BlockCounter + 1
         End If
      Next j
   Next i
   
   'For debug
   BLocksRemaining = BlockCounter
   
   'If the blockcounter is 0 (no blocks visible)
   If BlockCounter <= 0 Then
   
      CountRemaining = True
      
      InitBallToPaddle
      
      BallStick = True
      
      Ball.InPlay = False
      
      'Declare the level as clear
      LevelClear = True
      
      'Add one row to the blocks
      Blocks_Tall = Blocks_Tall + 1
      
      'If the Blocks_Tall exceeds the maximum
      If Blocks_Tall > 9 Then
      
         'Maximum reached... stay here.
         Blocks_Tall = 9
         
      End If
      
      'reset ball velocities
      Ball.VelX = BALLSPEED
      Ball.VelY = BALLSPEED
      
   End If
            
End Function

Private Sub InitBlocks(ByVal BlocksTall As Long)

   Dim i As Integer, j As Integer, n As Integer
 
   'Re-initialize blocks so that none are considered visible or broken
   For i = 0 To 7
      For j = 0 To 9
         Blocks(i, j).Visible = False
         Blocks(i, j).Busted = False
      Next j
   Next i
   
   For i = 0 To 7
      For j = 0 To BlocksTall
      
         'Set up X, y and BlockRECT coordinates
         'Use the BACK_OFFSET to place the blocks properly on the screen
         Blocks(i, j).X = ((BLOCK_WIDTH + 2) * i) + BACK_OFFSET
         Blocks(i, j).Y = ((BLOCK_HEIGHT + 2) * j) + BACK_OFFSET + BLOCK_OFFSET
         
         'Set up the BlockRECT coordinates
         Blocks(i, j).BlockRECT.Top = j * BLOCK_HEIGHT
         Blocks(i, j).BlockRECT.Bottom = Blocks(i, j).BlockRECT.Top + BLOCK_HEIGHT
         Blocks(i, j).BlockRECT.Right = Blocks(i, j).BlockRECT.Left + BLOCK_WIDTH
         
         'Make the block visible
         Blocks(i, j).Visible = True
         
         'The block are not broken yet
         Blocks(i, j).Busted = False
         
         'Set up the animation RECT (ARECT)
         For n = 0 To 3
            Blocks(i, j).ARECT(n).Top = j * BLOCK_HEIGHT
            Blocks(i, j).ARECT(n).Bottom = Blocks(i, j).ARECT(n).Top + BLOCK_HEIGHT
            Blocks(i, j).ARECT(n).Left = n * BLOCK_WIDTH
            Blocks(i, j).ARECT(n).Right = Blocks(i, j).ARECT(n).Left + BLOCK_WIDTH
         Next n
         
      Next j
   Next i

End Sub

Private Sub IncreaseSpeed()

   'Increase ball speed each time the ball hits the paddle.
   'If the INCREASE value is negative or positive,
   'add or subtract to the value accordingly until
   'a maximum value is reached.

   If Ball.VelX < 0 Then
      Ball.VelX = Ball.VelX - INCREASE
      If Ball.VelX < -MAXSPEED Then
         Ball.VelX = -MAXSPEED
      End If
   Else
      Ball.VelX = Ball.VelX + INCREASE
      If Ball.VelX > MAXSPEED Then
         Ball.VelX = MAXSPEED
      End If
   End If
   
   If Ball.VelY < 0 Then
      Ball.VelY = Ball.VelY - INCREASE
      If Ball.VelY < -MAXSPEED Then
         Ball.VelY = -MAXSPEED
      End If
   Else
      Ball.VelY = Ball.VelY + INCREASE
      If Ball.VelY > MAXSPEED Then
         Ball.VelY = MAXSPEED
      End If
   End If
   
End Sub

Private Sub DrawScore(theScore As String, X As Long, Y As Long)

   Dim t As String, i As Long
   
   'Clear any previous values
   Erase NumberValue
   
   t = Val(theScore)
   
   If t < 10 And t > 0 Then
      NumberValue(5) = t
   ElseIf t > 10 And t < 100 Then
      NumberValue(5) = t Mod 10
      NumberValue(4) = Left$(t, 1)
   ElseIf t >= 100 And t < 1000 Then
      NumberValue(5) = t Mod 10
      NumberValue(4) = Mid$(t, 2, 1)
      NumberValue(3) = Left$(t, 1)
   ElseIf t >= 1000 And t < 10000 Then
      NumberValue(5) = t Mod 10
      NumberValue(4) = Mid$(t, 3, 1)
      NumberValue(3) = Mid$(t, 2, 1)
      NumberValue(2) = Left$(t, 1)
   ElseIf t >= 10000 And t < 100000 Then
      NumberValue(5) = t Mod 10
      NumberValue(4) = Mid$(t, 4, 1)
      NumberValue(3) = Mid$(t, 3, 1)
      NumberValue(2) = Mid$(t, 2, 1)
      NumberValue(1) = Left$(t, 1)
   ElseIf t >= 100000 And t < 1000000 Then
      NumberValue(5) = t Mod 10
      NumberValue(4) = Mid$(t, 5, 1)
      NumberValue(3) = Mid$(t, 4, 1)
      NumberValue(2) = Mid$(t, 3, 1)
      NumberValue(1) = Mid$(t, 2, 1)
      NumberValue(0) = Left$(t, 1)
   ElseIf t = 1000000 Or t = 0 Then
      theScore = 0
      NumberValue(5) = 0
      NumberValue(4) = 0
      NumberValue(3) = ""
      NumberValue(2) = ""
      NumberValue(1) = ""
      NumberValue(0) = ""
   End If
   
   'Set up X positions
   For i = 0 To 5
      NumberPosition(i) = X + (i * 16)
   Next i
   
   'Draw the score
   For i = 5 To 0 Step -1
      If NumberValue(i) <> "" Then
         Render NumberPosition(i), Y, ScoreSurface.Surface, ScoreText(Val(NumberValue(i))), True
      End If
   Next i
   
   If SCORE > EXTRA_BALL And ExtraAwarded = False Then
      Ball.Count = Ball.Count + 1
      ExtraAwarded = True
      PlayExtraBall
   End If
         
End Sub

Private Sub DrawBonusLabel()

   Static DelayTime As Long

   DrawingBonus = True
   BallStick = True
   Ball.IsVisible = True
   PlayBzzz
   
   BonusPoints = BONUS_POINTS * Multiplier

   DelayTime = DelayTime + (1 * Elapsed)
   
   If DelayTime > 3000 Then
      DelayTime = 0
      DrawingBonus = False
      LoadPictureArray (Multiplier)
      Multiplier = Multiplier + 1
      SCORE = SCORE + BonusPoints
      BonusPoints = 0
      InitBlocks Blocks_Tall
      LevelClear = False
      Ball.InPlay = True
   End If

End Sub

Private Sub DoMeter()

   Const TIME_TO_WAIT As Long = 4000
   Static WaitTime As Long
   
   If DrawingMeter = True Then
      WaitTime = WaitTime + (1 * Elapsed)
      
      MeterRECT.TitleRECT.Right = MeterSurface.ddsd.lWidth / (TIME_TO_WAIT / WaitTime)
      
      If MeterRECT.TitleRECT.Right > MeterSurface.ddsd.lWidth Then
         MeterRECT.TitleRECT.Right = MeterSurface.ddsd.lWidth
      End If
      
      If WaitTime >= TIME_TO_WAIT Then
         WaitTime = 0
         DrawingMeter = False
         MeterSurface.Surface.BltColorFill MeterRECT.TitleRECT, RGB(255, 0, 0)
         PlayMeterSound
      End If
   End If
   
End Sub

Public Sub ReadScore()

   Dim ff As Integer
   ff = FreeFile
   
   Open App.Path & "\Game.hi" For Binary As ff Len = Len(TOP_SCORE)
   Get ff, 1, TOP_SCORE
   Close ff
   
End Sub

Public Sub WriteScore(ByVal theScore As Long)

   Dim ff As Integer
   ff = FreeFile
   
   Open App.Path & "\Game.hi" For Binary As ff Len = Len(theScore)
   Put ff, 1, theScore
   Close ff
   
End Sub

