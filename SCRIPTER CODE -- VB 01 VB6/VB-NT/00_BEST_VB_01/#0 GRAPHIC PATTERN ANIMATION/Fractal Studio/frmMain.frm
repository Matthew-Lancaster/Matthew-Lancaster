VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fractal Window Template"
   ClientHeight    =   1395
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   4035
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MouseIcon       =   "frmMain.frx":0A02
   MousePointer    =   2  'Cross
   ScaleHeight     =   93
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   269
   ShowInTaskbar   =   0   'False
   Begin VB.Menu MNU_LOGGF 
      Caption         =   "LOGG FOLDER"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim TmpBMP_Move As tMemBMP

Public Sub TriggerMouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Form_MouseMove Button, Shift, x, Y
End Sub

Private Sub Form_Activate()

If Me.Tag <> aData.nActive Then
    aData.nActive = Me.Tag
    ShowParams Frac(Me.Tag)
End If

End Sub

'FRAC TYPE NEWTON
'ITER=1000
'BAIL VAL .0005
'FRAC PARAM C1=8
'FRAC PARAM C2=8
'COLOR INPERPOLE LINE CHUNK 15
'SIZE 800*540
'FILE OUT SIZE 9000

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = 1 Or Button = 2 Then
    Frac(Me.Tag).Win.mDown = True
    Frac(Me.Tag).Win.mBtn = Button
    Frac(Me.Tag).Win.nX = x
    Frac(Me.Tag).Win.nY = Y
    Frac(Me.Tag).Win.nX2 = x
    Frac(Me.Tag).Win.nY2 = Y
    If Button = 2 Then
        TmpBMP_Move = CreateMemBMP(CreateBMP(Frac(Me.Tag).nW, Frac(Me.Tag).nH, True))
        Me.MousePointer = vbCustom
    End If
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
Dim TmpBMP As tMemBMP, MyRect As RECT, MyBrush As Long

With Frac(aData.nActive)
    If .Win.mDown And .Win.mBtn = 1 Then
        .Win.nX2 = x
        .Win.nY2 = Y
        If .Win.nX > x Then
            X1 = x
            X2 = .Win.nX
        Else
            X1 = .Win.nX
            X2 = x
        End If
        If .Win.nY > Y Then
            Y1 = Y
            Y2 = .Win.nY
        Else
            Y1 = .Win.nY
            Y2 = Y
        End If
    
        MyBrush = CreateSolidBrush(16777215)
        With MyRect
            .Left = X1
            .Right = X2
            .Top = Y1
            .Bottom = Y2
        End With
        TmpBMP = CreateMemBMP(.BMP)
        BitBlt TmpBMP.hdc, 0, 0, .nW, .nH, .BB.hdc, 0, 0, SRCCOPY
        FillRect TmpBMP.hdc, MyRect, MyBrush
        DeleteObject MyBrush
        BitBlt TmpBMP.hdc, X1, Y1, X2 - X1, Y2 - Y1, .BB.hdc, X1, Y1, SRCINVERT
        BitBlt .Win.wWin.hdc, 0, 0, .nW, .nH, TmpBMP.hdc, 0, 0, SRCCOPY
        FreeMemBMP TmpBMP
    ElseIf .Win.mDown And .Win.mBtn = 2 Then
        'Change position
        MyBrush = CreateSolidBrush(0)
        .Win.nX2 = x
        .Win.nY2 = Y
        With MyRect
            .Left = 0
            .Top = 0
            .Right = Frac(Me.Tag).nW
            .Bottom = Frac(Me.Tag).nH
        End With
        FillRect TmpBMP_Move.hdc, MyRect, MyBrush
        BitBlt TmpBMP_Move.hdc, x - .Win.nX, Y - .Win.nY, Frac(Me.Tag).nW, Frac(Me.Tag).nH, Frac(Me.Tag).BB.hdc, 0, 0, SRCCOPY
        BitBlt Me.hdc, 0, 0, Frac(Me.Tag).nW, Frac(Me.Tag).nH, TmpBMP_Move.hdc, 0, 0, SRCCOPY
    End If
End With

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim X1 As Long, X2 As Long, Y1 As Long, Y2 As Long
Dim nX1 As Long, nX2 As Long, nY1 As Long, nY2 As Long
Dim nSX As Double, nSY As Double, nEX As Double, nEY As Double
Dim OldDX As Double, TmpBMP As tMemBMP

Dim Tmp As Long

With Frac(Me.Tag).Win
    If .mDown And .mBtn = 1 Then
        .mDown = False
        X1 = .nX
        Y1 = .nY
        X2 = x
        Y2 = Y
        
        'Zoom
        '-------------------------------
        
        OldDX = Frac(Me.Tag).EX - Frac(Me.Tag).SX
        
        If X2 < X1 Then
            Tmp = X2
            X2 = X1
            X1 = Tmp
        End If
        If Y2 < Y1 Then
            Tmp = Y2
            Y2 = Y1
            Y1 = Tmp
        End If
        
        If (Y2 - Y1) > (X2 - X1) Then
            nY1 = Y1
            nY2 = Y2
            nX1 = ((X1 + X2) \ 2) - ((Y2 - Y1) \ 2)
            nX2 = nX1 + (Y2 - Y1)
        Else
            nX1 = X1
            nX2 = X2
            nY1 = ((Y1 + Y2) \ 2) - ((X2 - X1) \ 2)
            nY2 = nY1 + (X2 - X1)
        End If
        
        'Calculate new coordinates:
        With Frac(Me.Tag)
            nSX = .SX + ((nX1 * (.EX - .SX)) / .nW)
            nSY = .SY + ((nY1 * (.EY - .SY)) / .nH)
            
            nEX = .SX + ((nX2 * (.EX - .SX)) / .nW)
            nEY = .SY + ((nY2 * (.EY - .SY)) / .nH)
            
            .SX = nSX
            .SY = nSY
            .EX = nEX
            .EY = nEY
            
            If .EX - .SX < OldDX Then
                'Stretchblt to screen
                TmpBMP = CreateMemBMP(CreateBMP(.nW, .nH))
                If StretchBlt(TmpBMP.hdc, 0, 0, .nW, .nH, .BB.hdc, nX1, nY1, nX2 - nX1, nY2 - nY1, SRCCOPY) > 0 Then
                    BitBlt .BB.hdc, 0, 0, .nW, .nH, TmpBMP.hdc, 0, 0, SRCCOPY
                    BitBlt .Win.wWin.hdc, 0, 0, .nW, .nH, TmpBMP.hdc, 0, 0, SRCCOPY
                End If
                FreeMemBMP TmpBMP
            End If
            AddZoomCurrent Me.Tag, 0
        End With
        '--------------------------------
        
        Calculate aData.nActive, True
        
    ElseIf .mDown And .mBtn = 2 Then
        .mDown = False
        With Frac(Me.Tag)
            nSX = .SX + ((CDbl(.Win.nX - x) * (.EX - .SX)) / CDbl(.nW))
            nEX = nSX + (.EX - .SX)
            nSY = .SY + ((CDbl(.Win.nY - Y) * (.EY - .SY)) / CDbl(.nH))
            nEY = nSY + (.EX - .SX)
            .SX = nSX
            .EX = nEX
            .SY = nSY
            .EY = nEY
            FreeMemBMP TmpBMP_Move
            
            AddZoomCurrent Me.Tag, 1
            
            Me.MousePointer = 2
            DoEvents
            Calculate Me.Tag, True
        End With
    End If
End With

End Sub

Private Sub Form_Paint()

With Frac(aData.nActive)
    If .Win.mDown Then
        Form_MouseMove CInt(.Win.mBtn), 0, CSng(.Win.nX2), CSng(.Win.nY2)
    Else
        PaintWin Me.Tag
    End If
End With

End Sub



Private Sub Form_Resize()

If Me.WindowState Mod 2 = 0 Then Call TIMEROUT(True): Exit Sub
If Me.WindowState = 1 Then Call TIMEROUT(False): Exit Sub






'MODDIFY BY MATT THU 08-09-11
Exit Sub
    If Me.WindowState = vbMinimized Then
        'Save the data to a tmp file:
        Frac(Me.Tag).Flag_Stop = True
        SetStatus "Swapping to disk.."
        DoEvents
        Swap2Disk Frac(Me.Tag)
        SetStatus ""
    Else
        If Frac(Me.Tag).CellsOnDisk Then
            'Get cells from tmp file..
            SetStatus "Swapping from disk.."
            SwapFromDisk Frac(Me.Tag)
            SetStatus ""
        End If
    End If
End Sub
Sub TIMEROUT(TF)
    On Error Resume Next
    For Each Form In Forms
        For Each Control In Controls
            If Control.Interval > 0 Then Control.Enabled = TF
        Next
    Next
End Sub


Private Sub Form_Unload(Cancel As Integer)



If Me.Tag <> "" Then
    SAVE_PARAM_DETAIL
    Frac(Me.Tag).Flag_Stop = True
    RemoveFractal Me.Tag
End If

End Sub

Sub SAVE_PARAM_DETAIL()
    Dim sFile1 As String
    Dim sFile2 As String
    
    'Save
    XD = Format(Now, "YYYY-MM-DD HH-MM-SS") + " -- "
'    sFile1 = GetSaveFilePath("Save Fractal Param File", "Fstudio Param File|*.fs1")
    sFile1 = App.Path + "\## 0 default.fs1"
    sFile2 = App.Path + "\## " + XD + "Fstudio Param File.fs1"
'    If FileExist(sFile) And sFile <> "" Then
'        If MsgBox("The file already exists. Do you want to overwrite it?", _
'            vbYesNo) = vbNo Then
'            Exit Sub
'        End If
'    End If
'    If sFile1 <> "" Then
        SaveParamFile aData.nActive, sFile1
        SaveParamFile aData.nActive, sFile2
'    End If
End Sub

Private Sub MNU_LOGGF_Click()
    TR = 0
    Do
        XD = Format(Now - TS, "YYYY-MM-DD HH-MM-SS") + " -- "
        sFile2 = App.Path + "\## " + XD + "Fstudio Param File.fs1"
        TS = TimeSerial(0, 0, TR)
        TR = TR + 1
        DoEvents
    Loop Until Dir(sFile2) <> ""
    Shell "Explorer.exe /SELECT, " + sFile2, vbNormalFocus
End Sub
