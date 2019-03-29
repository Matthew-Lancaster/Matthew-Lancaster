Attribute VB_Name = "modUtils"
Option Explicit
Option Base 0

Public Function PtrToString(ByVal lParam As Long) As String

    Dim i As Long, b As Byte
    
    i = lParam
    
    CopyMemory b, ByVal i, 1
    
    Do While b <> 0
    
        PtrToString = PtrToString + Chr(b)
            
        i = i + 1
        CopyMemory b, ByVal i, 1
        
    Loop

End Function

Public Function PtrToStringW(ByVal lParam As Long) As String

    Dim i As Long, b As Integer
    
    i = lParam
    
    CopyMemory b, ByVal i, 2
    
    Do While b <> 0
    
        PtrToStringW = PtrToStringW + ChrW(b)
            
        i = i + 2
        CopyMemory b, ByVal i, 2
        
    Loop

End Function

Public Function CStrToString(bArray() As Byte) As String
    Dim i As Long
    
    On Error Resume Next
    
    i = -1
    i = LBound(bArray)
    If i < 0 Then Exit Function
    
    For i = LBound(bArray) To UBound(bArray)
    
        If bArray(i) = 0 Then Exit For
        CStrToString = CStrToString + Chr(bArray(i))
    
    Next i
    
End Function

Public Function StringToCStr(sData As String, bArray() As Byte, Optional ByVal Fixed As Boolean = True)
    
    Dim i As Long, _
        l As Long
    
    If (Not Fixed) Then
        ReDim bArray(0 To Len(sData))
    Else
        l = -1&
        l = LBound(bArray)
    End If

    For i = 1 To Len(sData)
    
        bArray(l) = Asc(Mid(sData, i, 1))
        l = l + 1
    Next i
    
    bArray(l) = 0&

End Function

Public Function CStrToStringW(iArray() As Integer) As String
    Dim i As Long
    
    On Error Resume Next
    
    i = -1
    i = LBound(iArray)
    If i < 0 Then Exit Function
    
    For i = LBound(iArray) To UBound(iArray)
    
        If iArray(i) = 0 Then Exit For
        CStrToStringW = CStrToStringW + ChrW(iArray(i))
    
    Next i
    
End Function

Public Function StringToCStrW(sData As String, iArray() As Integer, Optional ByVal Fixed As Boolean = True)
    
    Dim i As Long, _
        l As Long
    
    If (Not Fixed) Then
        ReDim iArray(0 To Len(sData))
    Else
        l = -1&
        l = LBound(iArray)
    End If

    For i = 1 To Len(sData)
    
        iArray(l) = Asc(Mid(sData, i, 1))
        l = l + 1
    Next i
    
    iArray(l) = 0&

End Function

Public Function ClipNulls(ByVal varString As String) As String

    If InStr(1, varString, Chr(0)) Then
        ClipNulls = Mid(varString, 1, InStr(1, varString, Chr(0)) - 1)
    Else
        ClipNulls = varString
    End If
    
End Function

Public Function ClipNullsW(ByVal varStringW As String) As String

    If InStr(1, varStringW, ChrW(0)) Then
        ClipNullsW = Mid(varStringW, 1, InStr(1, varStringW, ChrW(0)) - 1)
    Else
        ClipNullsW = varStringW
    End If
    
End Function

Public Function GetOleFont(lfFont As LOGFONT) As StdFont

    Dim riid As GUID
    Dim vData As FONTDESC
    
    Dim vFont As StdFont
    Dim vName As String
    Dim hDC As Long
    
    With riid
    
        .Data1 = &HBEF6E003
        .Data2 = &HA874
        .Data3 = &H101A
        
        .Data4(0) = &H8B
        .Data4(1) = &HBA
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
        
    End With
    
    With vData
    
        .cbSize = Len(vData)
        
        hDC = GetDC(0&)
        .cySize = Abs(MulDiv(lfFont.lfHeight, 72, GetDeviceCaps(hDC, LOGPIXELSY)))
        ReleaseDC 0&, hDC
        
        .sWeight = lfFont.lfWeight
        
        .fStrikethrough = CLng(lfFont.lfStrikeOut)
        .fUnderline = CLng(lfFont.lfUnderline)
        .fItalic = CLng(lfFont.lfItalic)
        vName = CStrToString(lfFont.lfFaceName)
        .szName = StrPtr(vName)
        .sCharset = lfFont.lfCharSet
                
    End With
    
    OleCreateFontIndirect vData, riid, vFont

    ' Sometimes the OS has a difficult time with the size
    
    If Round(vFont.Size, 0) <> Round(vData.cySize, 0) Then
        vFont.Size = vData.cySize
    End If

    Set GetOleFont = vFont
    
End Function

Public Sub GetLogFont(OleFont As StdFont, lfData As LOGFONT, hDC As Long)

    With lfData
        .lfCharSet = OleFont.Charset
        .lfItalic = OleFont.Italic
        .lfWeight = OleFont.Weight
        
        If OleFont.Bold Then
            .lfWeight = 700
        End If
        
        .lfUnderline = OleFont.Underline
        StringToCStr OleFont.Name, .lfFaceName
        .lfStrikeOut = OleFont.Strikethrough
        .lfHeight = -(MulDiv(OleFont.Size, GetDeviceCaps(hDC, LOGPIXELSY), 72))
            
    End With

End Sub

Public Function GetOleFontW(lfFont As LOGFONTW) As StdFont

    Dim riid As GUID
    Dim vData As FONTDESC
    
    Dim vFont As StdFont
    Dim vName As String
    Dim hDC As Long
    
    With riid
    
        .Data1 = &HBEF6E003
        .Data2 = &HA874
        .Data3 = &H101A
        
        .Data4(0) = &H8B
        .Data4(1) = &HBA
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
        
    End With
    
    With vData
    
        .cbSize = Len(vData)
        
        hDC = GetDC(0&)
        .cySize = Abs(MulDiv(lfFont.lfHeight, 72, GetDeviceCaps(hDC, LOGPIXELSY)))
        ReleaseDC 0&, hDC
        
        .sWeight = lfFont.lfWeight
        
        .fStrikethrough = CLng(lfFont.lfStrikeOut)
        .fUnderline = CLng(lfFont.lfUnderline)
        .fItalic = CLng(lfFont.lfItalic)
                
        vName = String(32, 0)
        CopyMemory ByVal StrPtr(vName), lfFont.lfFaceName(1), 64
        
        .szName = StrPtr(vName)
        .sCharset = lfFont.lfCharSet
                
    End With
    
    OleCreateFontIndirect vData, riid, vFont

    ' Sometimes the OS has a difficult time with the size
    
    If Round(vFont.Size, 0) <> Round(vData.cySize, 0) Then
        vFont.Size = vData.cySize
    End If

    Set GetOleFontW = vFont
    
End Function

Public Sub GetLogFontW(OleFont As StdFont, lfData As LOGFONTW, hDC As Long)

    With lfData
        .lfCharSet = OleFont.Charset
        .lfItalic = OleFont.Italic
        .lfWeight = OleFont.Weight
        
        If OleFont.Bold Then
            .lfWeight = 700
        End If
        
        .lfUnderline = OleFont.Underline
        CopyMemory .lfFaceName(1), ByVal StrPtr(OleFont.Name), LenB(OleFont.Name)
        
        .lfStrikeOut = OleFont.Strikethrough
        .lfHeight = -(MulDiv(OleFont.Size, GetDeviceCaps(hDC, LOGPIXELSY), 72))
            
    End With

End Sub


Public Function CDToArray(ByVal lpComma As String) As String()
    Dim lpStr As String, _
        lpCurrent As String
        
    Dim varNew() As String, _
        j As Long
        
    lpStr = lpComma
    
    lpCurrent = NextString(lpStr)
    ReDim Preserve varNew(0 To j)
    varNew(j) = lpCurrent
    j = j + 1
    
    Do While lpStr <> ""
        lpCurrent = NextString(lpStr)
        
        ReDim Preserve varNew(0 To j)
        varNew(j) = lpCurrent
        j = j + 1
    Loop
    
    CDToArray = varNew
    Erase varNew
    
End Function


Public Function NextString(vData As String) As String

    Dim i As Long, _
        j As Long
    
    i = InStr(1, vData, Chr(34))
    
    If i = 0& Then
        NextString = vData
        vData = ""
    Else
        j = InStr(i + 1, vData, Chr(34))
        If (j = 0&) Then
            NextString = Mid(vData, i + 1)
            vData = ""
        Else
            NextString = Mid(vData, i + 1, (j - 1) - (i + 1))
            vData = Mid(vData, j + 1)
        End If
        
    End If
    
End Function


