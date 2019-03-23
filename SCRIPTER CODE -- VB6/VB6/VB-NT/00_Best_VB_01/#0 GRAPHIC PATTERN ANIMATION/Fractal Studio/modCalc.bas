Attribute VB_Name = "modCalc"
Option Explicit



Public Sub Calculate(nFrac As Long, CalculateIt As Boolean)
Dim TmpBMP As tMemBMP, BMPInfo As BITMAPINFO
Dim nSteps As Long, Cnt_Step As Long, Cnt_Filter As Long
Dim nStart As Long, nEnd As Long
Dim PrgRect As RECT, hBrush As Long

If Frac(nFrac).CellsOnDisk Then
    SwapFromDisk Frac(nFrac)
    Frac(nFrac).Win.wWin.WindowState = vbNormal
End If

Frac(nFrac).Flag_State = STATE_CALC

If Frac(nFrac).Calc2File Then
    'Redirect to another sub when supposed to render directly to file:
    
    If Dir(Frac(nFrac).szPathCalcDirect, vbHidden Or vbSystem Or vbReadOnly) <> "" Then
        If MsgBox("File already exists. Overwrite?", vbYesNo, "Direct File Output") = vbNo Then
            Exit Sub
        Else
            Kill Frac(nFrac).szPathCalcDirect
        End If
    End If
    
    Load frmProgress
    frmProgress.Tag = nFrac
    frmProgress.Show vbModal
    'CalculateToFile nFrac, CalculateIt
    Exit Sub
End If

nSteps = Frac(nFrac).nH \ Frac(nFrac).LinesToDraw
TmpBMP = CreateMemBMP(Frac(nFrac).ChunkBMP)

'---------------------------------------------------------------------------
'---------------------RENDERING LOOP----------------------------------------
'---------------------------------------------------------------------------
Frac(nFrac).Flag_Stop = False
aData.Flag_Stop = False
For Cnt_Step = 0 To nSteps
    nStart = Cnt_Step * Frac(nFrac).LinesToDraw
    nEnd = (Cnt_Step + 1) * Frac(nFrac).LinesToDraw - 1
    If nEnd > (Frac(nFrac).nH - 1) Then nEnd = Frac(nFrac).nH - 1
    
    '--------------------------------------------------------'
    '--------  CALCULATION (RENDER TO ON-SREEN IMAGE) -------'
    '--------------------------------------------------------'
    If CalculateIt Then
        'Calculate some lines:
        CalcLines nFrac, nStart, nEnd
    
        'Apply Mapping filters:
        If Frac(nFrac).FilterMapOut.IsEnabled Then ApplyFilter nFrac, Frac(nFrac).FilterMapOut, nStart, nEnd, APPLY_OUT
        If Frac(nFrac).FilterMapIn.IsEnabled Then ApplyFilter nFrac, Frac(nFrac).FilterMapIn, nStart, nEnd, APPLY_IN
    
        'Apply Outside Filters:
        For Cnt_Filter = 0 To Frac(nFrac).nFiltersOut - 1
            If Frac(nFrac).FilterOut(Cnt_Filter).IsEnabled Then
                ApplyFilter nFrac, Frac(nFrac).FilterOut(Cnt_Filter), nStart, nEnd, APPLY_OUT
            End If
        Next Cnt_Filter
    
        'Apply Inside Filters:
        For Cnt_Filter = 0 To Frac(nFrac).nFiltersIn - 1
            If Frac(nFrac).FilterIn(Cnt_Filter).IsEnabled Then
                ApplyFilter nFrac, Frac(nFrac).FilterIn(Cnt_Filter), nStart, nEnd, APPLY_IN
            End If
        Next Cnt_Filter
    End If
    
    'Do other things:
    DoEvents
    If Frac(nFrac).Flag_Stop Or aData.Flag_Stop Then
        GoTo Calc_Finished
    End If
    
    '--------------------------------------------------------'
    '------------------  PLOTTING  --------------------------'
    '--------------------------------------------------------'
    PlotLinesToDIB nFrac, nStart, nEnd
    With Frac(nFrac).ChunkBMP
        BMPInfo.bmiHeader = .InfoHeader
        SetDIBits TmpBMP.hdc, TmpBMP.hBMP, 0, Frac(nFrac).LinesToDraw, .Data(0), BMPInfo, DIB_RGB_COLORS
    End With
    BitBlt Frac(nFrac).BB.hdc, 0, nStart, Frac(nFrac).nW, (nEnd - nStart) + 1, TmpBMP.hdc, 0, 0, SRCCOPY
    
    If Not Frac(nFrac).Win.mDown Then
        BitBlt Frac(nFrac).Win.wWin.hdc, 0, nStart, Frac(nFrac).nW, (nEnd - nStart) + 1, TmpBMP.hdc, 0, 0, SRCCOPY
    Else
        Frac(nFrac).Win.wWin.TriggerMouseMove CInt(Frac(nFrac).Win.mBtn), 0, CSng(Frac(nFrac).Win.nX2), CSng(Frac(nFrac).Win.nY2)
    End If
    '--------------------------------------------------------'
    
    'Show Progress:
    '---------------------------'
    With frmParam.picPrg
        PrgRect.Left = 0
        PrgRect.Top = 0
        PrgRect.Right = (nEnd * .ScaleWidth) \ Frac(nFrac).nH - 1
        PrgRect.Bottom = .ScaleHeight
        .Cls
        hBrush = (nEnd * 255) \ Frac(nFrac).nH - 1
        hBrush = CreateSolidBrush(RGB(hBrush, hBrush, hBrush))
        FillRect .hdc, PrgRect, hBrush 'aData.Col_Prg
        DeleteObject hBrush
        .Refresh
    End With
    '---------------------------'
    
    'Do other things:
    DoEvents
    If Frac(nFrac).Flag_Stop Then Exit Sub
    Frac(nFrac).Flag_State = STATE_CALC
    
Next Cnt_Step

Calc_Finished:

FreeMemBMP TmpBMP           'Delete tmp mem BMP
frmParam.picPrg.Cls         'Clear progress Picturebox
Frac(nFrac).Flag_Stop = True    'Make other procs stop
Frac(nFrac).Flag_State = STATE_IDLE

Call frmMain.SAVE_PARAM_DETAIL

End Sub


Public Sub CalculateToFile(nFrac As Long, CalculateIt As Boolean)
Dim fNum As Integer, nLine As Long, nOffset As Long, Cnt_Filter As Long, nIdleCnt As Long
Dim nChunkSize As Long, nOldWidth As Long, nOldHeight As Long, PrgRect As RECT
Dim nHeaderLen As Long, nLineLen As Long
Dim ImgW32 As Long

Frac(nFrac).Flag_Stop = False
Frac(nFrac).Flag_State = STATE_CALC
aData.Flag_Stop = False

If Val(frmParam.txtCalc2File_Size) < 50 Or Val(frmParam.txtCalc2File_Size) > 32767 Then
    MsgBox "Invalid image size. Valid range: 50 - 32767", vbOKOnly, "Error"
    Exit Sub
Else
    Frac(nFrac).nDirectFileSize = Val(frmParam.txtCalc2File_Size)
End If

ImgW32 = ((Frac(nFrac).nDirectFileSize * 3) \ 4) * 4
If ImgW32 < Frac(nFrac).nDirectFileSize * 3 Then ImgW32 = ImgW32 + 4

Frac(nFrac).BMP = CreateBMP(Frac(nFrac).nDirectFileSize, Frac(nFrac).nDirectFileSize, False)
Erase Frac(nFrac).BMP.Data
With Frac(nFrac).BMP
    .FileHeader.bfSize = Len(.FileHeader) + Len(.InfoHeader) + (ImgW32 * .H)
    .InfoHeader.biBitCount = 24
    .InfoHeader.biSizeImage = ImgW32 * .H
End With

fNum = FreeFile
Open Frac(nFrac).szPathCalcDirect For Binary As fNum

'Put a byte in the end of the file, just to allocate it.
'This will increase the speed of subsequent write oprations
Put fNum, Frac(nFrac).BMP.FileHeader.bfSize - 1, 1

'Write the file headers:
Put fNum, 1, Frac(nFrac).BMP.FileHeader
Put fNum, , Frac(nFrac).BMP.InfoHeader

'To find out where to start dumping the image data:
nHeaderLen = Len(Frac(nFrac).BMP.FileHeader) + Len(Frac(nFrac).BMP.InfoHeader)
nLineLen = ImgW32

'Change some properties temporarily:
nChunkSize = Frac(nFrac).LinesToDraw
Frac(nFrac).LinesToDraw = 1
nOldWidth = Frac(nFrac).nW: nOldHeight = Frac(nFrac).nH
Frac(nFrac).nW = Frac(nFrac).nDirectFileSize
Frac(nFrac).nH = Frac(nFrac).nDirectFileSize
Frac(nFrac).ChunkBMP = CreateBMP(Frac(nFrac).nW, 1)
ReDim Frac(nFrac).ChunkBMP.Data(0 To ImgW32 - 1)

'Move the 'image cell' data to a disk file
Swap2Disk Frac(nFrac)
ReDim Frac(nFrac).Cell(0 To Frac(nFrac).nDirectFileSize - 1, 1)

'Process the fractal one line at a time:
For nLine = 0 To Frac(nFrac).nH - 1
With Frac(nFrac)
    If CalculateIt Then
        'Calculate one line:
        CalcLines nFrac, nLine, nLine, True
    
        'Apply Mapping filters:
        If Frac(nFrac).FilterMapOut.IsEnabled Then ApplyFilter nFrac, Frac(nFrac).FilterMapOut, 0, 0, APPLY_OUT
        If Frac(nFrac).FilterMapIn.IsEnabled Then ApplyFilter nFrac, Frac(nFrac).FilterMapIn, 0, 0, APPLY_IN
    
        'Apply Outside Filters:
        For Cnt_Filter = 0 To Frac(nFrac).nFiltersOut - 1
            If Frac(nFrac).FilterOut(Cnt_Filter).IsEnabled Then
                ApplyFilter nFrac, Frac(nFrac).FilterOut(Cnt_Filter), 0, 0, APPLY_OUT
            End If
        Next Cnt_Filter
    
        'Apply Inside Filters:
        For Cnt_Filter = 0 To Frac(nFrac).nFiltersIn - 1
            If Frac(nFrac).FilterIn(Cnt_Filter).IsEnabled Then
                ApplyFilter nFrac, Frac(nFrac).FilterIn(Cnt_Filter), 0, 0, APPLY_IN
            End If
        Next Cnt_Filter
    End If
    
    PlotLinesToDIB nFrac, 0, 0, False, True, True
    Put fNum, nHeaderLen + ((.nH - nLine - 1) * nLineLen) + 1, Frac(nFrac).ChunkBMP.Data
    
    nIdleCnt = nIdleCnt + 1
    If nIdleCnt > 4 Then
        nIdleCnt = 0
        'Show Progress:
        '---------------------------'
        With frmProgress.picProg
            PrgRect.Left = 0
            PrgRect.Top = 0
            PrgRect.Right = (nLine * .ScaleWidth) \ Frac(nFrac).nH - 1
            PrgRect.Bottom = .ScaleHeight
            .Cls
            FillRect .hdc, PrgRect, aData.Col_Prg
            .Refresh
        End With
        '---------------------------'
        DoEvents
        
        If Frac(nFrac).Flag_Stop Then
            Close fNum
            Kill Frac(nFrac).szPathCalcDirect
            Frac(nFrac).Flag_Stop = False
            Exit For
        End If
        
    End If
End With
Next nLine
Close fNum

'Set the properties back again:
Frac(nFrac).LinesToDraw = nChunkSize
Frac(nFrac).nW = nOldWidth
Frac(nFrac).nH = nOldHeight
Frac(nFrac).ChunkBMP = CreateBMP(Frac(nFrac).nW, Frac(nFrac).LinesToDraw)
Frac(nFrac).BMP = CreateBMP(Frac(nFrac).nW, Frac(nFrac).nH)
SwapFromDisk Frac(nFrac)
frmParam.picPrg.Cls

Frac(nFrac).Flag_State = STATE_IDLE

End Sub


