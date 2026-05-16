Attribute VB_Name = "modFilters"
Option Explicit

'------- Compatibility Constants -------------
'---------------------------------------------
Global Const FILTER_COMP_2DOUT = 0
Global Const FILTER_COMP_2DIN = 1
Global Const FILTER_COMP_2D = 2

Global Const FILTER_COMP_4DOUT = 3
Global Const FILTER_COMP_4DIN = 4
Global Const FILTER_COMP_4D = 5

Global Const FILTER_COMP_2DMAP = 6
'---------------------------------------------

'------- Filter Constants --------------------
'---------------------------------------------
Global Const FILTER2DMAP_DIRECT_I = 0
Global Const FILTER2DMAP_CONTINOUS_I = 1
Global Const FILTER2DMAP_X02pY02 = 2
Global Const FILTER2D_ADD = 3
Global Const FILTER2D_MUL = 4
Global Const FILTER2D_X0 = 5
Global Const FILTER2D_Y0 = 6
Global Const FILTER2D_X1 = 7
Global Const FILTER2D_Y1 = 8
Global Const FILTER2D_COS_X0 = 9
Global Const FILTER2D_COS_Y0 = 10
Global Const FILTER2D_SIN_X0 = 11
Global Const FILTER2D_SIN_Y0 = 12
Global Const FILTER2D_COS2SIN2 = 13
Global Const FILTER2D_TRIG_SWIRL = 14
Global Const FILTER2DMAP_SUMCOLOR = 15
Global Const FILTER2DMAP_REALCOLOR = 16
Global Const FILTER2DMAP_IMAGCOLOR = 17
Global Const FILTER2DMAP_MULREALIMAG = 18
Global Const FILTER2D_MULRX = 19
Global Const FILTER2D_MULRY = 20
Global Const FILTER2D_MULRXRY = 21
Global Const FILTER2DMAP_TEST = 22
'---------------------------------------------

'---------------------------------------------
Global Const FILTER_MAX = 22
Global Const APPLY_OUT = 0
Global Const APPLY_IN = 1
'---------------------------------------------
'---------End of declares--------------

Public Sub Init_Filters()
Dim Cnt As Long

ReDim aData.Filters(0 To FILTER_MAX)
aData.nFilters = FILTER_MAX + 1

With aData
    'Mapping Filters:
    Add2FList FILTER_COMP_2DMAP, 0, FILTER2DMAP_DIRECT_I, "Iterations"
    Add2FList FILTER_COMP_2DMAP, 0, FILTER2DMAP_CONTINOUS_I, "Continous Iterations"
    Add2FList FILTER_COMP_2DMAP, 0, FILTER2DMAP_X02pY02, "X0^2 + Y0^2"
    Add2FList FILTER_COMP_2DMAP, 0, FILTER2DMAP_SUMCOLOR, "I + Abs(X0) + Abs(Y0)"
    Add2FList FILTER_COMP_2DMAP, 0, FILTER2DMAP_REALCOLOR, "I x Abs(X0)"
    Add2FList FILTER_COMP_2DMAP, 0, FILTER2DMAP_IMAGCOLOR, "I x Abs(Y0)"
    Add2FList FILTER_COMP_2DMAP, 0, FILTER2DMAP_MULREALIMAG, "I x Abs(X0) x Abs(Y0)"
    Add2FList FILTER_COMP_2DMAP, 0, FILTER2DMAP_TEST, "Test Mapping Filter"

    'Enhancement Filters:
    Add2FList FILTER_COMP_2D, 1, FILTER2D_ADD, "Add P1"
    Add2FList FILTER_COMP_2D, 1, FILTER2D_MUL, "Multiply with P1"
    Add2FList FILTER_COMP_2D, 1, FILTER2D_X0, "Add P1 x Abs(X0)"
    Add2FList FILTER_COMP_2D, 1, FILTER2D_Y0, "Add P1 x Abs(Y0)"
    Add2FList FILTER_COMP_2D, 1, FILTER2D_X1, "Add P1 x Abs(X1)"
    Add2FList FILTER_COMP_2D, 1, FILTER2D_Y1, "Add P1 x Abs(Y1)"
    Add2FList FILTER_COMP_2D, 1, FILTER2D_COS_X0, "Add P1 x Abs(Cos(X0))"
    Add2FList FILTER_COMP_2D, 1, FILTER2D_COS_Y0, "Add P1 x Abs(Cos(Y0))"
    Add2FList FILTER_COMP_2D, 1, FILTER2D_SIN_X0, "Add P1 x Abs(Sin(X0))"
    Add2FList FILTER_COMP_2D, 1, FILTER2D_SIN_Y0, "Add P1 x Abs(Sin(Y0))"
    Add2FList FILTER_COMP_2D, 2, FILTER2D_COS2SIN2, "Add P1 x Cos(X0^2) + P2 x Sin(Y0^2)"
    Add2FList FILTER_COMP_2D, 1, FILTER2D_TRIG_SWIRL, "Add P1 x < Swirl >"
    Add2FList FILTER_COMP_2D, 2, FILTER2D_MULRX, "Multiply with Xpos"
    Add2FList FILTER_COMP_2D, 0, FILTER2D_MULRY, "Multiply with Ypos"
    Add2FList FILTER_COMP_2D, 0, FILTER2D_MULRXRY, "Multiply with Xpos and Ypos"
End With

End Sub

Public Sub Add2FList(nComp As Byte, nVars As Byte, nIndex As Integer, szName As String)

With aData.Filters(nIndex)
    .nComp = nComp
    .nVars = nVars
    .szName = szName
End With

End Sub

Public Sub AddFilter(Fractal As tFractal, nType As Integer, OutSide As Boolean, Optional Var1 As Double = 1, Optional Var2 As Double = 1, Optional Var3 As Double = 1, Optional Var4 As Double = 1, Optional Var5 As Double = 1)

If OutSide Then
    ReDim Preserve Fractal.FilterOut(0 To Fractal.nFiltersOut)
    With Fractal.FilterOut(Fractal.nFiltersOut)
        .nType = nType
        .IsEnabled = True
        .Var(0) = Var1
        .Var(1) = Var2
        .Var(2) = Var3
        .Var(3) = Var4
        .Var(4) = Var5
        .nVars = aData.Filters(nType).nVars
    End With
    Fractal.nFiltersOut = Fractal.nFiltersOut + 1
Else
    ReDim Preserve Fractal.FilterIn(0 To Fractal.nFiltersIn)
    With Fractal.FilterIn(Fractal.nFiltersIn)
        .nType = nType
        .IsEnabled = True
        .Var(0) = Var1
        .Var(1) = Var2
        .Var(2) = Var3
        .Var(3) = Var4
        .Var(4) = Var5
        .nVars = aData.Filters(nType).nVars
    End With
    Fractal.nFiltersIn = Fractal.nFiltersIn + 1
End If

End Sub

Public Sub AddFilterIndirect(Fractal As tFractal, Filter As tFilterVars, OutSide As Boolean)

If OutSide Then
    ReDim Preserve Fractal.FilterOut(0 To Fractal.nFiltersOut)
    Fractal.FilterOut(Fractal.nFiltersOut) = Filter
    Fractal.nFiltersOut = Fractal.nFiltersOut + 1
Else
    ReDim Preserve Fractal.FilterIn(0 To Fractal.nFiltersIn)
    Fractal.FilterIn(Fractal.nFiltersIn) = Filter
    Fractal.nFiltersIn = Fractal.nFiltersIn + 1
End If

End Sub

Public Sub ApplyFilter(nFrac As Long, Filter As tFilterVars, StartLine As Long, EndLine As Long, ApplyTo As Byte)
Dim nX As Integer, nY As Integer
Dim Var0 As Double, Var1 As Double, Var2 As Double, Var3 As Double, Var4 As Double, Var5 As Double

Select Case Filter.nType

    '--------------------------------------------------------------------------------
    '------------------------------- MAPPING FILTERS --------------------------------
    '--------------------------------------------------------------------------------
    Case FILTER2DMAP_DIRECT_I
        'No code - keep the I values..
    
    Case FILTER2DMAP_CONTINOUS_I
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        Var0 = (.X0 * .X0) + (.Y0 * .Y0)
                        Var1 = (.X1 * .X1) + (.Y1 * .Y1)
                        If Var1 - Var0 <> 0 Then
                            .nI = .nI + ((Frac(nFrac).BailOut - Var0) / (Var1 - Var0))
                        End If
                    End If
                End With
            Next nX
        Next nY
    
    Case FILTER2DMAP_X02pY02
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = ((.X0 * .X0) + (.Y0 * .Y0))
                    End If
                End With
            Next nX
        Next nY
    Case FILTER2DMAP_SUMCOLOR
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + (.X0) + (.Y0)
                    End If
                End With
            Next nX
        Next nY
    Case FILTER2DMAP_REALCOLOR
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI * Abs(.X0)
                    End If
                End With
            Next nX
        Next nY
    Case FILTER2DMAP_IMAGCOLOR
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI * Abs(.Y0)
                    End If
                End With
            Next nX
        Next nY
    Case FILTER2DMAP_MULREALIMAG
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI * Abs(.X0) * Abs(.Y0)
                    End If
                End With
            Next nX
        Next nY
    Case FILTER2DMAP_TEST
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + 3 * Cos(Sin((.rX + .rY) * (.X1 - .X0)) - Cos(.Y1 - .Y0))
                    End If
                End With
            Next nX
        Next nY
    '--------------------------------------------------------------------------------
    '----------------------------- ENHANCEMENT FILTERS ------------------------------
    '--------------------------------------------------------------------------------
    Case FILTER2D_ADD
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + Filter.Var(0)
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_MUL
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI * Filter.Var(0)
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_X0
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + Filter.Var(0) * Abs(.X0)
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_Y0
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + Filter.Var(0) * Abs(.Y0)
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_X1
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + Filter.Var(0) * Abs(.X1)
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_Y1
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + Filter.Var(0) * Abs(.Y1)
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_COS_X0
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + Filter.Var(0) * Abs(Cos(.X0))
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_COS_Y0
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + Filter.Var(0) * Abs(Cos(.Y0))
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_SIN_X0
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + Filter.Var(0) * Abs(Sin(.X0))
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_SIN_Y0
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + Filter.Var(0) * Abs(Sin(.Y0))
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_COS2SIN2
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + Abs(Filter.Var(0) * Cos(.X0 * .X0) + Filter.Var(1) * Sin(.Y0 * .Y0))
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_TRIG_SWIRL
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        Var0 = ((.X0 * .X0) + (.Y0 * .Y0))
                        Var1 = ((.X1 * .X1) + (.Y1 * .Y1))
                        .nI = .nI + Filter.Var(0) * (Abs(Cos(Var0 * Var0 + 10 * ((.Y1 - .Y0) / (Var1 - Var0 + 0.01))) + Sin(.Y0 * .Y0)))
                    End If
                End With
            Next nX
        Next nY
    
    Case FILTER2D_MULRX
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI + Filter.Var(0) * ((.X1 + .Y1 + .X0 + .Y0) / (Abs(Filter.Var(1) * .rX) + 0.1))
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_MULRY
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI * .rY
                    End If
                End With
            Next nX
        Next nY
        
    Case FILTER2D_MULRXRY
        For nY = StartLine To EndLine
            For nX = 0 To Frac(nFrac).nW - 1
                With Frac(nFrac).Cell(nX, nY)
                    If .State = ApplyTo Then
                        .nI = .nI * .rX * .rY
                    End If
                End With
            Next nX
        Next nY
        
        
    '--------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------
End Select

End Sub

Public Sub PlotLinesToDIB(nFrac As Long, StartLine As Long, EndLine As Long, Optional bPlotToBB As Boolean = True, Optional bOnlyOneLine As Boolean = False, Optional Only24Bits As Boolean = False)
Dim nX As Long, nY As Long, nLine As Long, nIndexInDIB As Long, EndX As Long, EndY As Long
Dim nIndex As Single, nIntIndex As Single, nFraction As Single, pMax As Single
Dim SCol As tRGBLong, ECol As tRGBLong, nCol As tRGBLong
Dim n24BitLineSize As Long

If Frac(nFrac).Flag_Stop Or aData.Flag_Stop Then Exit Sub
EndX = Frac(nFrac).nW - 1: EndY = EndLine
'----------- CHECK FOR INVALID COLOR INDICES -----------'
pMax = Frac(nFrac).nPalLen - 1
For nY = StartLine To EndY
If bOnlyOneLine Then
    nLine = 0
Else
    nLine = nY
End If
    For nX = 0 To EndX
        nIndex = Frac(nFrac).Cell(nX, nLine).nI
        If nIndex < 0 Then
            nIndex = -nIndex
            Frac(nFrac).Cell(nX, nLine).nI = nIndex
        End If
        If nIndex > pMax Then
            If nIndex < 2000000000# Then
                nIntIndex = Int(nIndex)
                nFraction = nIndex - nIntIndex
                nIndex = pMax - ((nIntIndex Mod pMax) + nFraction)
                Frac(nFrac).Cell(nX, nLine).nI = nIndex
            Else
                Frac(nFrac).Cell(nX, nLine).nI = 0
            End If
        End If
    Next nX
Next nY

'---------- PLOT EITHER DIRECTLY OR BY INTERPOLATING BETWEEN COLORS-----------'
n24BitLineSize = ((Frac(nFrac).nW * 3) \ 4) * 4
If n24BitLineSize < Frac(nFrac).nW * 3 Then n24BitLineSize = n24BitLineSize + 4
If Frac(nFrac).LinearContinuity Then 'Interpolate
    For nY = StartLine To EndY
    If bOnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
        For nX = 0 To EndX
            nIndex = Frac(nFrac).Cell(nX, nLine).nI
            nIntIndex = Int(nIndex)
            
            nFraction = nIndex - nIntIndex
            
            SCol = Frac(nFrac).Pal(nIntIndex)
            If nFraction > 0 Then
                ECol = Frac(nFrac).Pal(nIntIndex + 1)
            Else
                ECol = Frac(nFrac).Pal(nIntIndex)
            End If
            
            nCol.R = SCol.R + ((ECol.R - SCol.R) * nFraction)
            nCol.G = SCol.G + ((ECol.G - SCol.G) * nFraction)
            nCol.B = SCol.B + ((ECol.B - SCol.B) * nFraction)
            
            With Frac(nFrac).ChunkBMP
                If Only24Bits Then
                    nIndexInDIB = ((.H - (nLine - StartLine) - 1) * n24BitLineSize) + (nX * 3)
                Else
                    nIndexInDIB = .tblOffsetY(nLine - StartLine) + .tblOffsetX(nX)
                End If
            End With
            With Frac(nFrac).ChunkBMP
                .Data(nIndexInDIB) = nCol.R
                .Data(nIndexInDIB + 1) = nCol.G
                .Data(nIndexInDIB + 2) = nCol.B
            End With
            
            If bPlotToBB Then
                With Frac(nFrac).BMP
                    If Only24Bits Then
                        nIndexInDIB = ((.H - (nLine - StartLine) - 1) * n24BitLineSize) + (nX * 3)
                    Else
                        nIndexInDIB = .tblOffsetY(nLine - StartLine) + .tblOffsetX(nX)
                    End If
                End With
                With Frac(nFrac).BMP
                    .Data(nIndexInDIB) = nCol.R
                    .Data(nIndexInDIB + 1) = nCol.G
                    .Data(nIndexInDIB + 2) = nCol.B
                End With
            End If
        Next nX
    Next nY
Else 'Plot directly
    For nY = StartLine To EndY
    If bOnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
        For nX = 0 To EndX
            nIntIndex = Int(Frac(nFrac).Cell(nX, nLine).nI)
            nIndexInDIB = Frac(nFrac).ChunkBMP.tblOffsetY(nLine - StartLine) + Frac(nFrac).ChunkBMP.tblOffsetX(nX)
            With Frac(nFrac).ChunkBMP
                .Data(nIndexInDIB) = Frac(nFrac).Pal(nIntIndex).R
                .Data(nIndexInDIB + 1) = Frac(nFrac).Pal(nIntIndex).G
                .Data(nIndexInDIB + 2) = Frac(nFrac).Pal(nIntIndex).B
            End With
            
            If bPlotToBB Then
                nIndexInDIB = Frac(nFrac).BMP.tblOffsetY(nLine) + Frac(nFrac).BMP.tblOffsetX(nX)
                With Frac(nFrac).BMP
                    .Data(nIndexInDIB) = Frac(nFrac).Pal(nIntIndex).R
                    .Data(nIndexInDIB + 1) = Frac(nFrac).Pal(nIntIndex).G
                    .Data(nIndexInDIB + 2) = Frac(nFrac).Pal(nIntIndex).B
                End With
            End If
        Next nX
    Next nY
End If
'-----------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------'
End Sub

Public Sub DoPalAnim(Fractal As tFractal, nFrames As Long, nInc As Long, Optional nStretch As Long = 1)

'VAR DECLARATION:
'--------------------------------------------------
Dim Cnt As Long, nX As Long, nY As Long, nW As Long, nH As Long, nEnd As Long
Dim T1 As Long, T2 As Long, nSum As Long, IdleCnt As Long, nVframeCnt As Long
Dim Cell() As Long, Data() As Long, Pal() As Long
Dim nOffset As Long, nColIndex As Long, PalMax As Long, nMaxI As Long, nMaxFrame As Long
Dim MyMBMP As tMemBMP, BMPInfo As BITMAPINFO
Dim MyInc As Long
Dim Col1 As tRGBLong, Col2 As tRGBLong, CurCol As tRGBLong

Dim nNormMin As Double, nNormMax As Double, nNormInterval As Double
'--------------------------------------------------

'Initialize some vars:
'--------------------------------------------------
Fractal.Flag_Stop = False
MyMBMP = CreateMemBMP(Fractal.BMP)
nW = Fractal.nW - 1
nH = Fractal.nH - 1
MyInc = nInc * nStretch
ReDim Cell(0 To (Fractal.nW * Fractal.nH) - 1)
'--------------------------------------------------

'Normalize frame:
'--------------------------------------------------
If Fractal.bPalAnimNormalize Then
    'Find maximum and minimum values of iterations:
    nNormMax = -1000000 'Just to be sure.. The user could have put on a filter that
    'subtracts a lot..
    For nY = 0 To Fractal.nH - 1
        For nX = 0 To Fractal.nW - 1
            If Fractal.Cell(nX, nY).nI > nNormMax Then nNormMax = Fractal.Cell(nX, nY).nI
        Next nX
    Next nY
    nNormMin = nNormMax    'The minimum values should be a lot lower than this, but again,
    'it may vary a lot depending on the filters..
    For nY = 0 To Fractal.nH - 1
        For nX = 0 To Fractal.nW - 1
            If Fractal.Cell(nX, nY).nI < nNormMin Then nNormMin = Fractal.Cell(nX, nY).nI
        Next nX
    Next nY
    
    'Now that the min and max values have been found, subtract the minimum value and
    'eventually scale to a specified interval size:
    If Fractal.bPalAnimNormIntervalSpecified Then
        nNormInterval = Fractal.nPalAnimNormInterval
        For nY = 0 To Fractal.nH - 1
            For nX = 0 To Fractal.nW - 1
                'Normalize:
                With Fractal.Cell(nW - nX, nH - nY)
                    Cell(nY * Fractal.nW + (nW - nX)) = ((.nI - nNormMin) / (nNormMax - nNormMin)) * nNormInterval
                End With
            Next nX
        Next nY
    Else
        For nY = 0 To Fractal.nH - 1
            For nX = 0 To Fractal.nW - 1
                'Normalize:
                With Fractal.Cell(nW - nX, nH - nY)
                    Cell(nY * Fractal.nW + (nW - nX)) = .nI - nNormMin
                End With
            Next nX
        Next nY
    End If
    nMaxI = (nNormMax - nNormMin) * nStretch
    
Else
    For nY = 0 To Fractal.nH - 1
        For nX = 0 To Fractal.nW - 1
            Cell(nY * Fractal.nW + (nW - nX)) = CLng(Fractal.Cell(nW - nX, nH - nY).nI * nStretch)
            If Fractal.Cell(nW - nX, nH - nY).nI > nMaxI Then nMaxI = CLng(Fractal.Cell(nW - nX, nH - nY).nI)
        Next nX
    Next nY
End If
'--------------------------------------------------

'Redim Data array (to contain the DIB section):
ReDim Data(0 To ((Fractal.nW * Fractal.nH) - 1))

'Transfer Palette array to Long array:
'--------------------------------------------------
If nStretch > 1 Then
    ReDim Pal(0 To (Fractal.nPalLen - 1) * nStretch)
    For Cnt = 0 To (Fractal.nPalLen - 2)
        Col1 = Fractal.Pal(Cnt)
        Col2 = Fractal.Pal(Cnt + 1)
        For nX = 0 To nStretch
            CurCol.R = Col1.R + (((Col2.R - Col1.R) * nX) \ nStretch)
            CurCol.G = Col1.G + (((Col2.G - Col1.G) * nX) \ nStretch)
            CurCol.B = Col1.B + (((Col2.B - Col1.B) * nX) \ nStretch)
            Pal((Cnt * nStretch) + nX) = RGB(CurCol.R, CurCol.G, CurCol.B)
        Next nX
    Next Cnt
Else
    ReDim Pal(0 To (Fractal.nPalLen) - 1)
    For Cnt = 0 To Fractal.nPalLen - 1
        With Fractal.Pal(Cnt)
            Pal(Cnt) = RGB(.R, .G, .B)
        End With
    Next Cnt
End If
'--------------------------------------------------

'Initialize some vars:
'--------------------------------------------------
nInc = nInc * nStretch
nEnd = (nW + 1) * (nH + 1) - 1
PalMax = (Fractal.nPalLen - 1) * nStretch
BMPInfo.bmiHeader = Fractal.BMP.InfoHeader
BMPInfo.bmiHeader.biBitCount = 32
'--------------------------------------------------

'Find the last frame number based on whether the I values will reference
'palette indexes greater than the palette size:
If nInc < 1 Then nInc = 1
If nMaxI >= PalMax Then
    MsgBox "No frames to draw - too high iterations values. Create a larger palette."
    Exit Sub
Else
    nMaxFrame = (PalMax - nMaxI) \ nInc
End If

'--------------------------------------------------
'---------------- DO STUFF -- ANIMATION -----------
'--------------------------------------------------
Cnt = 0
While True  'Go into a loop (until the Flag_Stop flag is set to True)
    'T1 = GetTickCount
    
    'Check for frame boundaries:
    If nVframeCnt >= nMaxFrame Then
        nVframeCnt = nMaxFrame
        MyInc = -Abs(MyInc)
    ElseIf nVframeCnt < 1 Then
        nVframeCnt = 0
        MyInc = Abs(MyInc)
    End If
    
    'Tight & fast palette animation loop:
    For nOffset = 0 To nEnd
        Data(nOffset) = Pal(Cell(nOffset) + nVframeCnt)
    Next
    
    'Insert DIB data into memory BMP:
    SetDIBits MyMBMP.hdc, MyMBMP.hBMP, 0, Fractal.nH, Data(0), BMPInfo, 0
    'Copy image to screen Device Context:
    BitBlt Fractal.Win.wWin.hdc, 0, 0, nW + 1, nH + 1, MyMBMP.hdc, 0, 0, SRCCOPY
    
    
    'This section is for preserving program functionality while looping:
    IdleCnt = IdleCnt + 1
    If IdleCnt = 20 Then
        IdleCnt = 0
        
        'FPS CALCULATION:
        'T2 = GetTickCount
        'Fractal.Win.wWin.Caption = Trim(Str(1000 \ (T2 - T1)))
        
        DoEvents 'do other stuff
        If Fractal.Flag_Stop Then GoTo DoPalAnim_End
    End If
    
    
    
    'Increase counters:
    nVframeCnt = nVframeCnt + MyInc
    Cnt = Cnt + 1
    
Wend
'--------------------------------------------------
'--------------------------------------------------
'--------------------------------------------------

DoPalAnim_End:

'Clean Up:
'---------------
FreeMemBMP MyMBMP
Erase Cell
Erase Pal
'---------------

'Set flag:
Fractal.Flag_Stop = True

End Sub

Public Sub PaintWin(nFractal As Long)

With Frac(nFractal)
    BitBlt .Win.wWin.hdc, 0, 0, .nW, .nH, .BB.hdc, 0, 0, SRCCOPY
End With

End Sub

Public Sub MakePalette(Fractal As tFractal, nLength As Long)
Dim Cnt As Long, CntP As Long, nStart As Long, nEnd As Long, nSteps As Long, nPeriod As Long
Dim SCol As tRGBLong, ECol As tRGBLong, ColOK As Boolean
Dim ColCnt As Byte

'------------------------------------
'---------CREATE PALETTE-------------
'------------------------------------

Randomize
With Fractal
    .nPalLen = nLength
    ReDim Preserve .Pal(0 To .nPalLen - 1)
    ColOK = False
    SCol.R = Rnd * 255
    SCol.G = Rnd * 255
    SCol.B = Rnd * 255
    
    nSteps = .nPalLen \ .nColPeriod
    If nSteps * .nColPeriod <> .nPalLen Then nSteps = nSteps + 1
    nSteps = nSteps - 1
    
    For Cnt = 0 To nSteps
        nStart = Cnt * .nColPeriod
        
        If Cnt < nSteps Then
            nEnd = nStart + .nColPeriod
            nPeriod = .nColPeriod
        Else
            nEnd = .nPalLen - 1
            nPeriod = nEnd - nStart + 1
        End If
        
        ColOK = False
        While Not ColOK
            ECol.R = Rnd * 255
            ECol.G = Rnd * 255
            ECol.B = Rnd * 255
            If .UseDispersionControl Then
                '--------------------
                If (Sqr(((ECol.R - SCol.R) ^ 2) + _
                    ((ECol.G - SCol.G) ^ 2) + ((ECol.B - SCol.B) ^ 2))) > 180 _
                    Then ColOK = True
                '--------------------
            Else
                ColOK = True
            End If
        Wend
        
        ColCnt = ColCnt + 1
        If ColCnt = 3 Then
            ECol.R = 0
            ECol.G = 0
            ECol.B = 0
        ElseIf ColCnt = 6 Then
            ECol.R = 255
            ECol.G = 255
            ECol.B = 255
        ElseIf ColCnt = 8 Then
            ECol.R = 255
            ECol.G = 255
            ECol.B = 0
        ElseIf ColCnt = 12 Then
            ECol.R = 255
            ECol.G = 0
            ECol.B = 0
        ElseIf ColCnt > 14 Then
            ColCnt = 0
        End If
        
        For CntP = nStart To nEnd
            With .Pal(CntP)
                .R = SCol.R + (((ECol.R - SCol.R) * (CntP - nStart)) / nPeriod)
                .G = SCol.G + (((ECol.G - SCol.G) * (CntP - nStart)) / nPeriod)
                .B = SCol.B + (((ECol.B - SCol.B) * (CntP - nStart)) / nPeriod)
            End With
        Next CntP
        SCol = ECol
    Next Cnt
End With

'------------------------------------
'------------------------------------
'------------------------------------

End Sub

Public Sub ClearFilters(nFrac As Long)

ReDim Frac(nFrac).FilterIn(0)
ReDim Frac(nFrac).FilterOut(0)
Frac(nFrac).nFiltersIn = 0
Frac(nFrac).nFiltersOut = 0

End Sub
