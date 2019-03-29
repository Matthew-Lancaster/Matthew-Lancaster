Attribute VB_Name = "Module3"
'This module contains the heart of the visualization processing
'The FFT algorithms are here
'Special thanks to Murphy McCauley

Option Explicit

Public Const AngleNumerator# = 6.283185307179
Public Const NumSamples& = 1024
Public Const NumBits& = 10

Private ReversedBits&(0 To NumSamples - 1)

Sub FFTAudio(RealIn() As Stereo, RealOut() As FFTStereo)
Static ImagOut(0 To NumSamples - 1) As FFTStereo, iScale!(511)
Static i&, j&, k&, n&, BlockSize&, BlockEnd&, Fl&
Static TR!, TI!, AR!, AI!, DeltaAngle#, DeltaAr#, Alpha#, Beta#
  
  If Fl = 0 Then 'Precalc ReversedBits
    Fl = 1: Dim ii&, iii&, jj As Byte, Rev&
    For ii = LBound(ReversedBits) To UBound(ReversedBits)
      iii = ii: Rev = 0
      For jj = 0 To NumBits - 1
        Rev = (Rev * 2) Or (iii And 1): iii = iii \ 2
      Next jj
      ReversedBits(ii) = Rev
    Next ii
    For i = 0 To 511: iScale(i) = (1 + (i ^ 1.2) / 80): Next
  End If
  
  For i = 0 To (NumSamples - 1)
    j = ReversedBits(i)
    RealOut(j).L = RealIn(i).L
    RealOut(j).R = RealIn(i).R
  Next i
  Erase ImagOut: BlockEnd = 1: BlockSize = 2
    
  Do While BlockSize <= NumSamples
    DeltaAngle = AngleNumerator / BlockSize
    Alpha = Sin(0.5 * DeltaAngle)
    Alpha = 2# * Alpha * Alpha
    Beta = Sin(DeltaAngle)
    
    For i = 0 To NumSamples - 1 Step BlockSize
      AR = 1!: AI = 0!
      For j = i To i + BlockEnd - 1
        k = j + BlockEnd
        'Left
        TR = AR * RealOut(k).L - AI * ImagOut(k).L
        TI = AI * RealOut(k).L + AR * ImagOut(k).L
        RealOut(k).L = RealOut(j).L - TR: RealOut(j).L = RealOut(j).L + TR
        ImagOut(k).L = ImagOut(j).L - TI: ImagOut(j).L = ImagOut(j).L + TI
        'Right
        TR = AR * RealOut(k).R - AI * ImagOut(k).R
        TI = AI * RealOut(k).R + AR * ImagOut(k).R
        RealOut(k).R = RealOut(j).R - TR: RealOut(j).R = RealOut(j).R + TR
        ImagOut(k).R = ImagOut(j).R - TI: ImagOut(j).R = ImagOut(j).R + TI
        
        DeltaAr = Alpha * AR + Beta * AI
        AI = AI - (Alpha * AI - Beta * AR)
        AR = AR - DeltaAr
      Next j
    Next i
    
    BlockEnd = BlockSize
    BlockSize = BlockSize * 2
  Loop
  For i = 0 To 354
    With RealOut(i)
      .L = iScale(i) * (.L * .L + RealOut(1023 - i).L * RealOut(1023 - i).L)
      .R = iScale(i) * (.R * .R + RealOut(1023 - i).R * RealOut(1023 - i).R)
    End With
  Next i
End Sub

