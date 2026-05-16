Attribute VB_Name = "ModMain"
Option Explicit

Public Sub Main()
Dim Cnt As Integer, bFirstTime As Long
Dim bOpen As Boolean
Dim sCmd As String

Load frmParent
Load frmParam
Load frmFilterParams

Init_App
Init_Filters

aData.Initializing = True
aData.nActive = 0
aData.nFractals = 1

If LCase(Right$(Command, 3)) = "fs1" And FileExist(Command) Then
    sCmd = Command
    bOpen = True
Else
    If FileExist(aData.szAppPath & "## 0 default.fs1") Then
        sCmd = aData.szAppPath & "## 0 default.fs1"
        bOpen = True
    Else
        bOpen = False
        Init_Fractal 0
    End If
End If

frmParent.Show

If bOpen Then LoadParamFile sCmd, False

With frmParam.lstFiltersOut
    .ColumnHeaders(2).Width = 450 '200
    For Cnt = 3 To 7
        .ColumnHeaders(Cnt).Width = 400 '150
    Next Cnt
    .ColumnHeaders(1).Width = 2970
End With
With frmParam.lstFiltersIn
    .ColumnHeaders(2).Width = 450
    For Cnt = 3 To 7
        .ColumnHeaders(Cnt).Width = 400
    Next Cnt
    .ColumnHeaders(1).Width = 2970
End With

If bOpen Then
    SizeFracWin Frac(aData.nActive)
End If
ArrangeWindows

GetNParam "firstuse", bFirstTime
If bFirstTime = -1 Then
    'Show welcome dialog:
    Load frmWelcome
    frmWelcome.Left = (Screen.Width \ 2) - (frmWelcome.Width \ 2)
    frmWelcome.Top = (Screen.Height \ 2) - (frmWelcome.Height \ 2)
    frmWelcome.Show vbModal
    SetNParam "firstuse", 1
End If

frmParam.Show
ShowParams Frac(aData.nActive)
Frac(aData.nActive).Win.wWin.Show
'DoEvents
'frmParam.chkCustomOutputSize.Value = vbUnchecked
'DoEvents

Calculate aData.nActive, True
frmParam.chkCustomOutputSize.Value = vbChecked
aData.Initializing = False


End Sub

Public Function AddFractal(Optional bInit As Boolean = True) As Long
Dim Cnt As Integer, nEmpty As Integer
Dim bBusy As Boolean

For Cnt = 0 To aData.nFractals - 1
    If Frac(Cnt).Flag_State <> STATE_IDLE Then bBusy = True
Next Cnt
If aData.Flag_State <> STATE_IDLE Then bBusy = True

If bBusy Then
    aData.Flag_Stop = True
    DoEvents
End If

nEmpty = -1
For Cnt = 0 To aData.nFractals - 1
    If Not Frac(Cnt).InUse Then
        nEmpty = Cnt
        Exit For
    End If
Next Cnt
If nEmpty = -1 Then
    'No fractal not in use found. Make new
    ReDim Preserve Frac(0 To aData.nFractals)
    nEmpty = aData.nFractals
    aData.nFractals = aData.nFractals + 1
End If

If bInit Then
    Init_Fractal CLng(nEmpty)
    Frac(nEmpty).Win.wWin.Show
End If
frmParam.Visible = True
Frac(nEmpty).InUse = True
If bInit Then
    ShowParams Frac(nEmpty)
    Calculate CLng(nEmpty), True
End If

AddFractal = nEmpty
End Function

Public Sub RemoveFractal(nFractal As Long)
Dim bFound As Boolean, Cnt As Long

aData.Flag_Stop = True
With Frac(nFractal)
    .Flag_Stop = True
    Erase .Cell
    Erase .FilterIn
    Erase .FilterOut
    .nFiltersIn = 0
    .nFiltersOut = 0
    FreeBitmapData .BMP
    FreeMemBMP .BB
    FreeMemBMP .PalDisp
    FreeBitmapData .ChunkBMP
    Unload .Win.wWin
    .InUse = False
End With

bFound = False
For Cnt = 0 To aData.nFractals - 1
    If Frac(Cnt).InUse Then bFound = True
Next Cnt

If Not bFound Then
    frmParam.Visible = False
    frmParent.mnuWindowShowparam.Visible = False
    frmParent.mnuWindowSep01.Visible = False
End If

End Sub

Public Sub InitWin(Fractal As tFractal)
Dim Cnt As Long, szCnt As String

Cnt = Fractal.nIndex + 1
If Cnt > 99 And Cnt < 1000 Then
    szCnt = Trim(Str(Cnt))
ElseIf Cnt > 9 And Cnt < 100 Then
    szCnt = "0" & Trim(Str(Cnt))
ElseIf Cnt >= 0 And Cnt < 10 Then
    szCnt = "00" & Trim(Str(Cnt))
Else
    szCnt = Trim(Str(Cnt))
End If

Set Fractal.Win.wWin = New frmMain
Fractal.Win.wWin.Tag = Cnt
Fractal.Win.wWin.Caption = "<New Fractal " & szCnt & ">"
Fractal.Win.InUse = True
Fractal.Win.dw = (Fractal.Win.wWin.Width \ Screen.TwipsPerPixelX) - Fractal.Win.wWin.ScaleWidth
Fractal.Win.dH = (Fractal.Win.wWin.Height \ Screen.TwipsPerPixelY) - Fractal.Win.wWin.ScaleHeight
Fractal.Win.wWin.Tag = Fractal.nIndex

End Sub

Public Sub Init_App()

ReDim Frac(aData.nActive)
aData.Col_Prg = CreateSolidBrush(RGB(0, 0, 128))
If Right(App.Path, 1) = "\" Then
    aData.szAppPath = App.Path
    aData.szTmpPath = App.Path & "Tmp\"
Else
    aData.szAppPath = App.Path & "\"
    aData.szTmpPath = App.Path & "\Tmp\"
End If
If Dir(aData.szTmpPath, vbDirectory) = "" Then
    MkDir aData.szTmpPath
End If

End Sub

Public Sub Init_Fractal(nFrac As Long)

With Frac(nFrac)
    .InUse = True
    .nType = TYPE_MANDELBROT
    .Param(0) = 0 '0.4
    .Param(1) = 0 '0.3
    .nIndex = nFrac
    .nW = 400
    .nH = 400
    .nDirectFileSize = 1000
    ReDim .Cell(0 To .nW - 1, 0 To .nH - 1)
    ReDim .Zoom(0)
    .nZooms = 0
    .LinesToDraw = 10
    .ChunkBMP = CreateBMP(.nW, .LinesToDraw)
    .PalDisp = CreateMemBMP(CreateBMP(frmParam.imgPal.ScaleWidth, frmParam.imgPal.ScaleHeight))
    .BMP = CreateBMP(.nW, .nH)
    .BB = CreateMemBMP(.BMP)
    .SX = -2
    .SY = -2
    .EX = 2
    .EY = 2
    .MaxI = 80
    .BailOut = 4
    .LinearContinuity = True
    
    .SavedFractal = False
    .SavedCalcDirect = False
    .SavedImage = False
    .SavedPalette = False
    .szPathFractal = ""
    .szPathCalcDirect = "c:\fractal001.bmp"
    .szPathImage = ""
    .szPathPalette = ""
    .nPalLen = 2000
    .nColPeriod = 20
    .nPalAnimNormInterval = 350
    .UseDispersionControl = True
    ReDim .Pal(0 To .nPalLen - 1)
    
    With Frac(nFrac).FilterMapOut
        .IsEnabled = True
        .nType = FILTER2DMAP_CONTINOUS_I
        .nVars = 0
    End With
    
    With Frac(nFrac).FilterMapIn
        .IsEnabled = True
        .nType = FILTER2DMAP_X02pY02
        .nVars = 0
    End With
    
    AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
    AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
    AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
    AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
    AddFilter Frac(nFrac), FILTER2D_MUL, False, 50, 0
    
    AddZoomCurrent nFrac, 0
    MakePalette Frac(nFrac), 2000
    InitWin Frac(nFrac)
    SizeFracWin Frac(nFrac)
    
End With

End Sub





Public Sub SavePalette(Fractal As tFractal, fPath As String)
Dim fNum As Integer, fPos As Long
Dim Cnt As Long
Dim fBuffer As String

fPos = 1
fBuffer = Space$((Fractal.nPalLen + 2) * 3)

For Cnt = 0 To Fractal.nPalLen - 1
    Mid$(fBuffer, fPos, 1) = Chr$(Fractal.Pal(Cnt).R)
    Mid$(fBuffer, fPos + 1, 1) = Chr$(Fractal.Pal(Cnt).G)
    Mid$(fBuffer, fPos + 2, 1) = Chr$(Fractal.Pal(Cnt).B)
    fPos = fPos + 3
Next Cnt

fNum = FreeFile
Open fPath For Binary As fNum
    Put #fNum, 1, fBuffer
Close fNum

End Sub

Public Sub OpenPalette(Fractal As tFractal, fPath As String)
Dim fNum As Long, fLen As Long

fLen = FileLen(fPath)
ReDim Pal(0 To (fLen \ 3) - 1)

fNum = FreeFile
Open fPath For Binary As fNum

Get #fNum, 1, Pal
Close fNum

'ReDim FractPalette(0 To FileLength \ 3)
'PalettePos = 0

'For FilePos = 1 To FileLength - 2 Step 3
'    FileChar = Mid$(PaletteBuffer, FilePos, 1)
'    NewR = Asc(FileChar)
'    FileChar = Mid$(PaletteBuffer, FilePos + 1, 1)
'    NewG = Asc(FileChar)
'    FileChar = Mid$(PaletteBuffer, FilePos + 2, 1)
'    NewB = Asc(FileChar)
    
'    FractPalette(PalettePos) = RGB(NewR, NewG, NewB)
'    PalettePos = PalettePos + 1
'Next FilePos

'BeginCalculate

End Sub

Public Sub SizeFracWin(Fractal As tFractal)

With Fractal.Win
    .wWin.Width = (Fractal.nW + (1 * .dw)) * Screen.TwipsPerPixelX
    .wWin.Height = (Fractal.nH + (1 * .dH)) * Screen.TwipsPerPixelY
End With

End Sub

Public Sub ArrangeWindows()
Dim Cnt As Long, dX As Long, dY As Long
Dim nOffset As Long

frmParam.Left = frmParent.ScaleWidth - frmParam.Width
frmParam.Top = 0

nOffset = 0
For Cnt = 0 To aData.nFractals - 1
    If Frac(Cnt).Win.wWin.WindowState = vbNormal Then
        Frac(Cnt).Win.wWin.Left = nOffset
        Frac(Cnt).Win.wWin.Top = nOffset
        nOffset = nOffset + (20 * Screen.TwipsPerPixelX)
    End If
Next Cnt

End Sub

Public Sub ShowParams(Fractal As tFractal)
Dim Cnt As Long, CntP As Long, Tmp As ListItem
Dim nX As Long, nY As Long

'General Options:
'------------------------------------------------------------------------------
With frmParam
    .txtMaxIterations.Text = Trim(Str(Fractal.MaxI))
    .txtBailout.Text = Trim(Str(Fractal.BailOut))
End With
With frmParam
    For Cnt = 0 To 4
        .txtCalcP(Cnt).Visible = False
        .lblCalcParam(Cnt).Visible = False
        .txtCalcP(Cnt).Enabled = True
        .lblCalcParam(Cnt).Enabled = True
    Next Cnt
    
    
    If Fractal.nType = TYPE_MANDELBROT Then
        .cmbFractType.Text = "Mandelbrot"
    ElseIf Fractal.nType = TYPE_MANDELZPOW Then
        .cmbFractType.Text = "MandelZPower"
    ElseIf Fractal.nType = TYPE_JULIA Then
        .cmbFractType.Text = "Julia"
    ElseIf Fractal.nType = TYPE_JULIAZPOW Then
        .cmbFractType.Text = "JuliaZPower"
    ElseIf Fractal.nType = TYPE_MANDEL3 Then
        .cmbFractType.Text = "Mandel^3"
    ElseIf Fractal.nType = TYPE_MANDEL4 Then
        .cmbFractType.Text = "Mandel^4"
    ElseIf Fractal.nType = TYPE_JULIA3 Then
        .cmbFractType.Text = "Julia^3"
    ElseIf Fractal.nType = TYPE_JULIA4 Then
        .cmbFractType.Text = "Julia^4"
    ElseIf Fractal.nType = TYPE_MANDELRINGS Then
        .cmbFractType.Text = "Mandel Rings"
    ElseIf Fractal.nType = TYPE_MANDELZPOWRINGS Then
        .cmbFractType.Text = "MandelZPower"
    ElseIf Fractal.nType = TYPE_JULIARINGS Then
        .cmbFractType.Text = "Mandel Rings"
    ElseIf Fractal.nType = TYPE_JULIAZPOWRINGS Then
        .cmbFractType.Text = "Mandel Rings"
    ElseIf Fractal.nType = TYPE_BARNSLEYM1 Then
        .cmbFractType.Text = "BarnsleyMandel 1"
    ElseIf Fractal.nType = TYPE_BARNSLEYM2 Then
        .cmbFractType.Text = "BarnsleyMandel 2"
    ElseIf Fractal.nType = TYPE_BARNSLEYM3 Then
        .cmbFractType.Text = "BarnsleyMandel 3"
    ElseIf Fractal.nType = TYPE_BARNSLEYM4 Then
        .cmbFractType.Text = "BarnsleyMandel 4"
    ElseIf Fractal.nType = TYPE_BARNSLEYJ1 Then
        .cmbFractType.Text = "BarnsleyJulia 1"
    ElseIf Fractal.nType = TYPE_BARNSLEYJ2 Then
        .cmbFractType.Text = "BarnsleyJulia 2"
    ElseIf Fractal.nType = TYPE_BARNSLEYJ3 Then
        .cmbFractType.Text = "BarnsleyJulia 3"
    ElseIf Fractal.nType = MARKSMANDEL Then
        .cmbFractType.Text = "Barnsley"
    ElseIf Fractal.nType = MARKMANDELGENERAL Then
        .cmbFractType.Text = "BarnsleyMandel"
    ElseIf Fractal.nType = MARKSJULIA Then
        .cmbFractType.Text = "Barnsley"
    ElseIf Fractal.nType = MARKSJULIAGENERAL Then
        .cmbFractType.Text = "JuliaZPower"
    ElseIf Fractal.nType = TYPE_BARNSLEYMANJUL Then
        .cmbFractType.Text = "BarnsleyMandelJulia"
    ElseIf Fractal.nType = TYPE_SPIDER Then
        .cmbFractType.Text = "Spider"
    ElseIf Fractal.nType = TYPE_MANDELLAMBDA Then
        .cmbFractType.Text = "MandelLambda"
    ElseIf Fractal.nType = TYPE_SIERPINSKI Then
        .cmbFractType.Text = "Sierpinski"
    ElseIf Fractal.nType = TYPE_NEWTON Then
        .cmbFractType.Text = "Newton"
    ElseIf Fractal.nType = TYPE_NEWTONJULIANOVA Then
        .cmbFractType.Text = "NewtonJuliaNova"
    End If
'Mandelbrot
'Julia
'Mandel^3
'Mandel^4
'Mandel Rings
'BarnsleyMandel 1
'BarnsleyMandel 4
'BarnsleyJulia1
'BarnsleyJulia 2
'BarnsleyJulia 3
'MandelZPower
'JuliaZPower
'BarnsleyMandelJulia
'Spider
'MandelLambda
'Sierpinski
'Newton
'NewtonJuliaNova
      
          
        

'Global Const TYPE_MANDELBROT = 0
'Global Const TYPE_MANDELZPOW = 1
'Global Const TYPE_JULIA = 2
'Global Const TYPE_JULIAZPOW = 3
'Global Const TYPE_MANDEL3 = 4
'Global Const TYPE_MANDEL4 = 5
'Global Const TYPE_JULIA3 = 6
'Global Const TYPE_JULIA4 = 7
'Global Const TYPE_MANDELRINGS = 8
'Global Const TYPE_MANDELZPOWRINGS = 9
'Global Const TYPE_JULIARINGS = 10
'Global Const TYPE_JULIAZPOWRINGS = 11
'Global Const TYPE_BARNSLEYM1 = 12
'Global Const TYPE_BARNSLEYM2 = 13
'Global Const TYPE_BARNSLEYM3 = 14
'Global Const TYPE_BARNSLEYM4 = 15
'Global Const TYPE_BARNSLEYJ1 = 16
'Global Const TYPE_BARNSLEYJ2 = 17
'Global Const TYPE_BARNSLEYJ3 = 18
'Global Const MARKSMANDEL = 19
'Global Const MARKMANDELGENERAL = 20
'Global Const MARKSJULIA = 21
'Global Const MARKSJULIAGENERAL = 22
'Global Const TYPE_BARNSLEYMANJUL = 23
'Global Const TYPE_SPIDER = 24
'Global Const TYPE_MANDELLAMBDA = 25
'Global Const TYPE_SIERPINSKI = 26
'Global Const TYPE_NEWTON = 27
'Global Const TYPE_NEWTONJULIANOVA = 28
    
    
    
    'If Fractal.nType = TYPE_MANDELBROT Or Fractal.nType = TYPE_JULIA Or _
    '    Fractal.nType = TYPE_MANDELRINGS Then
        
        .txtCalcP(0).Visible = True
        .txtCalcP(1).Visible = True
        .lblCalcParam(0).Visible = True
        .lblCalcParam(1).Visible = True
        .txtCalcP(0).Text = Trim(Str(Fractal.Param(0)))
        .txtCalcP(1).Text = Trim(Str(Fractal.Param(1)))
        .txtCalcP(2).Text = Trim(Str(Fractal.Param(2)))
        .lblCalcParam(0).Caption = "C Value 1"
        .lblCalcParam(1).Caption = "C Value 2"
        
        If Fractal.nType = TYPE_MANDELZPOW Or Fractal.nType = TYPE_JULIAZPOW Then
            .txtCalcP(2).Visible = True
            .lblCalcParam(2).Visible = True
            .lblCalcParam(2).Caption = "Power"
        Else
            .txtCalcP(2).Visible = False
            .lblCalcParam(2).Visible = False
        End If
        
'    End If
End With

'------------------------------------------------------------------------------
'Filters:
'------------------------------------------------------------------------------
frmParam.lstFiltersOut.ListItems.Clear
frmParam.lstFiltersIn.ListItems.Clear
frmParam.cmbMapOut.Clear
frmParam.cmbMapIn.Clear

For Cnt = 0 To FILTER_MAX
    If aData.Filters(Cnt).nComp = FILTER_COMP_2DMAP Then
        frmParam.cmbMapOut.AddItem aData.Filters(Cnt).szName
        frmParam.cmbMapIn.AddItem aData.Filters(Cnt).szName
        If Fractal.FilterMapOut.nType = Cnt Then
            frmParam.cmbMapOut.Text = aData.Filters(Cnt).szName
        End If
        If Fractal.FilterMapIn.nType = Cnt Then
            frmParam.cmbMapIn.Text = aData.Filters(Cnt).szName
        End If
    End If
Next Cnt

For Cnt = 0 To Fractal.nFiltersOut - 1
    Set Tmp = frmParam.lstFiltersOut.ListItems.Add(, , aData.Filters(Fractal.FilterOut(Cnt).nType).szName)
    If Fractal.FilterOut(Cnt).IsEnabled Then
        Tmp.SubItems(1) = "Yes"
    Else
        Tmp.SubItems(1) = "No"
    End If
    For CntP = 2 To 6
        Tmp.SubItems(CntP) = "-"
    Next CntP
    For CntP = 1 To Fractal.FilterOut(Cnt).nVars
        Tmp.SubItems(CntP + 1) = Trim(Str(Fractal.FilterOut(Cnt).Var(CntP - 1)))
    Next CntP
Next Cnt

For Cnt = 0 To Fractal.nFiltersIn - 1
    Set Tmp = frmParam.lstFiltersIn.ListItems.Add(, , aData.Filters(Fractal.FilterIn(Cnt).nType).szName)
    If Fractal.FilterIn(Cnt).IsEnabled Then
        Tmp.SubItems(1) = "Yes"
    Else
        Tmp.SubItems(1) = "No"
    End If
    For CntP = 2 To 6
        Tmp.SubItems(CntP) = "-"
    Next CntP
    For CntP = 1 To Fractal.FilterIn(Cnt).nVars
        Tmp.SubItems(CntP + 1) = Trim(Str(Fractal.FilterIn(Cnt).Var(CntP - 1)))
    Next CntP
Next Cnt

'------------------------------------------------------------------------------
'Zoom:
frmParam.lstZoom.ListItems.Clear
For Cnt = 0 To Fractal.nZooms - 1
    With frmParam.lstZoom
        If Cnt < (Fractal.nZooms - 1) Then
            .ListItems.Add 1, , Trim(Str(Cnt + 1)) & Space(20)
        Else
            .ListItems.Add 1, , Trim(Str(Cnt + 1)) & "   (Current)"
        End If
        If Fractal.Zoom(Cnt).nType = 0 Then
            .ListItems(1).SubItems(1) = "Zoom"
        Else
            .ListItems(1).SubItems(1) = "Move"
        End If
        .ListItems(1).SubItems(2) = "[" & Trim(Str(Fractal.Zoom(Cnt).SX)) & ", " & Trim(Str(Fractal.Zoom(Cnt).SY)) & "]-[" & Trim(Str(Fractal.Zoom(Cnt).EX)) & ", " & Trim(Str(Fractal.Zoom(Cnt).EY)) & "]"
    End With
Next Cnt
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'Position:
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'Palette:
DisplayPal Fractal
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'Animation:
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------

End Sub

Public Sub EditFilter(Fractal As tFractal, Map As Boolean, Out As Boolean, nIndex As Long)
Dim Filter As tFilterVars
Dim Cnt As Long

If Map Then
    If Out Then
        Filter = Fractal.FilterMapOut
    Else
        Filter = Fractal.FilterMapIn
    End If
Else
    If Out Then
        Filter = Fractal.FilterOut(nIndex)
    Else
        Filter = Fractal.FilterIn(nIndex)
    End If
End If

For Cnt = 0 To 4
    frmFilterParams.txtParam(Cnt).Text = "< Param" & Trim(Str(Cnt + 1)) & " >"
    frmFilterParams.lblParam(Cnt).ForeColor = vbButtonShadow
    frmFilterParams.txtParam(Cnt).Enabled = False
Next Cnt

For Cnt = 0 To Filter.nVars - 1
    frmFilterParams.txtParam(Cnt).Enabled = True
    frmFilterParams.txtParam(Cnt).Text = Trim(Str(Filter.Var(Cnt)))
    frmFilterParams.lblParam(Cnt).ForeColor = 0
Next Cnt

If Filter.IsEnabled Then
    frmFilterParams.chkEnabled.Value = 1
    frmFilterParams.chkEnabled.ForeColor = RGB(0, 127, 0)
Else
    frmFilterParams.chkEnabled.Value = 0
    frmFilterParams.chkEnabled.ForeColor = RGB(127, 0, 0)
End If

frmFilterParams.lblFilterName.Caption = aData.Filters(Filter.nType).szName
frmFilterParams.Tag = False
frmFilterParams.Show vbModal

If frmFilterParams.Tag = True Then
    'Change filter params
    For Cnt = 0 To Filter.nVars
        Filter.Var(Cnt) = Val(frmFilterParams.txtParam(Cnt).Text)
    Next Cnt
    If frmFilterParams.chkEnabled.Value = 1 Then
        Filter.IsEnabled = True
    Else
        Filter.IsEnabled = False
    End If
    '--------------------------------------------
    If Map Then
        If Out Then
            Fractal.FilterMapOut = Filter
        Else
            Fractal.FilterMapIn = Filter
        End If
    Else
        If Out Then
            Fractal.FilterOut(nIndex) = Filter
        Else
            Fractal.FilterIn(nIndex) = Filter
        End If
    End If
    '--------------------------------------------
Else
    Exit Sub
End If

End Sub

Public Sub ResizeFractal(Fractal As tFractal, nW As Long, nH As Long, CalculateIt As Boolean)
Dim TmpBMP As tBitmap

With Fractal
    .nW = nW
    .nH = nH
    ReDim Fractal.Cell(0 To nW - 1, 0 To nH - 1)
    TmpBMP = CreateBMP(nW, nH)
    FreeMemBMP .BB
    .BB = CreateMemBMP(TmpBMP)
    .BMP = TmpBMP
    If .LinesToDraw > .nH Then .LinesToDraw = .nH
    .ChunkBMP = CreateBMP(nW, .LinesToDraw)
End With
SizeFracWin Fractal
If CalculateIt Then Calculate aData.nActive, True

End Sub

Public Sub DisplayPal(Fractal As tFractal)
Dim nX As Long, nY As Long, Col As tRGBLong, nCol As Long
Dim TmpBMP As tMemBMP

TmpBMP = CreateMemBMP(CreateBMP(Fractal.PalDisp.w, Fractal.PalDisp.H))
frmParam.scrPalImg.Max = Fractal.nPalLen - frmParam.imgPal.ScaleWidth
With TmpBMP
    For nX = 0 To .w - 1
        Col = Fractal.Pal(frmParam.scrPalImg.Value + nX)
        nCol = RGB(Col.R, Col.G, Col.B)
        SetPixel .hdc, nX, 0, nCol
    Next nX
End With
With Fractal.PalDisp
    For nY = 0 To .H - 1
        BitBlt .hdc, 0, nY, .w, 1, TmpBMP.hdc, 0, 0, SRCCOPY
    Next nY
    BitBlt frmParam.imgPal.hdc, 0, 0, .w, .H, .hdc, 0, 0, SRCCOPY
End With

End Sub

Public Sub Terminate()
Dim Cnt As Long

'Clean up all resources used.
For Cnt = 0 To aData.nFractals - 1
    FreeMemBMP Frac(Cnt).BB
    FreeMemBMP Frac(Cnt).PalDisp
    FreeBitmapData Frac(Cnt).BMP
    FreeBitmapData Frac(Cnt).ChunkBMP
    Erase Frac(Cnt).Cell
    Unload Frac(Cnt).Win.wWin
    
    If Frac(Cnt).szDiskID <> "" Then
        If Dir(Frac(Cnt).szDiskID, vbHidden Or vbSystem Or vbReadOnly) <> "" Then
            Kill Frac(Cnt).szDiskID
        End If
    End If
Next Cnt
Unload frmParam
Unload frmAddFilter
Unload frmFilterParams
Unload frmMain

'End execution:
End

End Sub

Public Function GenerateID(nLength As Long) As String
Dim Cnt As Long, B() As Byte, ret As String

Randomize
ReDim B(0 To nLength - 1)
For Cnt = 0 To nLength - 1
    B(Cnt) = 0
    While (Not (B(Cnt) > 47 And B(Cnt) < 58)) And (Not (B(Cnt) > 96 And B(Cnt) < 123))
        B(Cnt) = Rnd * 255
    Wend
Next Cnt
ret = ""
For Cnt = 0 To nLength - 1
    ret = ret & Chr$(B(Cnt))
Next Cnt
GenerateID = ret

End Function

Public Sub Swap2Disk(Fractal As tFractal)
Dim fNum As Integer, szID As String, szPath As String

If Fractal.szDiskID = "" Then
    While szID = "" Or Dir(aData.szTmpPath & szID & ".tmp", vbHidden Or vbReadOnly Or vbSystem Or vbDirectory) <> ""
        szID = GenerateID(8) 'Generate a random ID with 8 alphanumeric, lowercase characters
    Wend
    Fractal.szDiskID = aData.szTmpPath & szID & ".tmp"
End If
szPath = Fractal.szDiskID

fNum = FreeFile 'Get a file number
Open szPath For Binary As fNum
    Put fNum, 1, Fractal.Cell() 'Write data
Close fNum
Erase Fractal.Cell
Erase Fractal.Cell
Fractal.CellsOnDisk = True

End Sub
Public Sub SwapFromDisk(Fractal As tFractal)
Dim fNum As Integer

If Not Fractal.CellsOnDisk Or Fractal.szDiskID = "" Then Exit Sub

'Redimension array to hold data:
ReDim Fractal.Cell(0 To Fractal.nW - 1, 0 To Fractal.nH - 1)

'Check whether the file exists. If the file doesn't exist, redim the array
'anyway to avoid IndexOutOfBounds (the file may have been deleted):
If Dir(Fractal.szDiskID, vbHidden Or vbReadOnly Or vbSystem) = "" Then
    MsgBox "Swap file has been deleted!"
    Exit Sub
End If

'Read from file:
fNum = FreeFile
Open Fractal.szDiskID For Binary As fNum
    Get fNum, 1, Fractal.Cell()
Close fNum
Fractal.CellsOnDisk = False

End Sub

Public Sub SetStatus(szStatus As String)
    If szStatus <> "" Then
        frmParent.status.Panels(1).Text = szStatus
    Else
        frmParent.status.Panels(1).Text = "Ready"
    End If
End Sub
Public Function GetOpenFilePath(ByVal szQuery As String, szFilter As String) As String
On Error GoTo Err_GetOpenFilePath

If szQuery = "" Then szQuery = "Open File"
With frmParam.CmD
    .CancelError = True
    .DialogTitle = szQuery
    .Filter = szFilter
    .FileName = ""
    .ShowOpen
    
    GetOpenFilePath = .FileName
End With
    
Exit Function
Err_GetOpenFilePath:
GetOpenFilePath = ""

End Function

Public Function GetSaveFilePath(ByVal szQuery As String, szFilter As String) As String
On Error GoTo Err_GetOpenFilePath

If szQuery = "" Then szQuery = "Save File"
With frmParam.CmD
    .CancelError = True
    .DialogTitle = szQuery
    .Filter = szFilter
    .FileName = ""
    .ShowSave
    
    GetSaveFilePath = .FileName
End With
    
Exit Function
Err_GetOpenFilePath:
GetSaveFilePath = ""

End Function

Public Sub AddZoomCurrent(nFrac As Long, nType As Byte)
Dim nIndex As Long

With Frac(nFrac)
    nIndex = .nZooms
    .Zoom(nIndex).SX = .SX
    .Zoom(nIndex).SY = .SY
    .Zoom(nIndex).EX = .EX
    .Zoom(nIndex).EY = .EY
    .Zoom(nIndex).nType = nType
    
    With frmParam.lstZoom
        .ListItems.Add 1, , Trim(Str(nIndex + 1)) & "   (Current)"
        
        If .ListItems.Count > 1 Then
            .ListItems(2).Text = Left(.ListItems(2).Text, Len(Trim(Str(.ListItems.Count - 1)))) & Space(20)
        End If
        
        If nType = 0 Then
            .ListItems(1).SubItems(1) = "Zoom"
        Else
            .ListItems(1).SubItems(1) = "Move"
        End If
        .ListItems(1).SubItems(2) = "[" & Trim(Str(Frac(nFrac).Zoom _
            (nIndex).SX)) & ", " & Trim(Str(Frac(nFrac).Zoom(nIndex).SY)) & _
            "]-[" & Trim(Str(Frac(nFrac).Zoom(nIndex).EX)) & ", " & _
            Trim(Str(Frac(nFrac).Zoom(nIndex).EY)) & "]"
    End With
        
    .nZooms = .nZooms + 1
    ReDim Preserve .Zoom(0 To .nZooms)
End With

End Sub

Public Function GetSParam(sParamName As String, ParamReceiver As String) As Boolean
    ParamReceiver = GetSetting(App.EXEName, "main", sParamName, "")
    GetSParam = (ParamReceiver <> "")
End Function
Public Function GetNParam(sParamName As String, ParamReceiver As Long) As Boolean
    ParamReceiver = CLng(Val(GetSetting(App.EXEName, "main", sParamName, "-1")))
    GetNParam = (ParamReceiver <> -1)
End Function
Public Sub SetSParam(sParamName As String, sValue As String)
    SaveSetting App.EXEName, "main", sParamName, sValue
End Sub
Public Sub SetNParam(sParamName As String, nValue As Long)
    SaveSetting App.EXEName, "main", sParamName, Trim(Str(nValue))
End Sub

