Attribute VB_Name = "modTextMan"
''' GUInerd Standard Menu System
''' Version 5.0

'''

''' Common Text Manipulation Routines


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

Public Const DefaultSeparators = " _-\%"
Public Const WrapLimit_Minimum = 20&

Public g_Headers(0 To 4) As String
Public g_WinVersion As WindowsVersionConstants

Public Function FileByteStr(Bytes() As Byte) As String

    Dim iUnicode As Integer, _
        i As Long, _
        j As Long
        
    Dim sUni As String
        
    On Error Resume Next
    
    i = -1&
    i = UBound(Bytes) + 1
    
    If (i = 0&) Then Exit Function
    
    If (i < 2) Then
        FileByteStr = StrConv(Bytes, vbUnicode)
        Exit Function
    End If
    
    CopyMemory iUnicode, Bytes(0), 2
    
    Select Case iUnicode
        
        Case &HFFFE, &HFEFF
            If (i Mod 2) Then
                i = i - 1
            End If
            
            sUni = String(i / 2, 0)
            CopyMemory ByVal StrPtr(sUni), Bytes(0), i
            
            FileByteStr = sUni
            
        Case Else
            FileByteStr = StrConv(Bytes, vbUnicode)
        
    End Select
                    
End Function

Public Function NoLines(ByVal varData As String, Optional ByVal LnChr As String) As String

    Dim i As Long, _
        j As Long
        
    Dim s As String, _
        t As String
    
    Dim x As Long
    
    j = 1
    i = InStr(1, varData, vbCrLf)
    
    If i = 0 Then
        NoLines = varData
        Exit Function
    End If
    
    x = Len(varData)
    
    Do While i <> 0
            
        t = Mid(varData, j, i - j)
        
        If (Mid(s, Len(s), 1) <> " ") And (Mid(t, 1, 1) <> " ") Then
            s = s + " "
        End If
        
        s = s + LnChr + t
        t = ""
        j = i + Len(vbCrLf)
        If (j > x) Then Exit Do
        
        i = InStr(j, varData, vbCrLf)
        If i = 0 Then
            s = s + Mid(varData, j)
            Exit Do
        End If
        
    Loop
    
    NoLines = s
End Function

Public Function RemoveNulls(ByVal szText As String) As String
    Dim i As Long, _
        j As Long
        
    Dim c As String, _
        d As String
    
    i = InStr(1, szText, Chr(0))
    
    If (i = 0&) Then
        RemoveNulls = szText
        Exit Function
    End If
    
    i = Len(szText)
    
    For j = 1 To i
        c = Mid(szText, j, 1)
        If (c <> Chr(0)) And (c <> ChrW(0)) Then
            d = d + c
        End If
    Next j
    
    RemoveNulls = d
    
End Function

'' Determine whether character in buffer is a separator character or not.
Public Function IsSepChar(ByVal sChr As String, Optional ByVal sSeparators As String = DefaultSeparators) As Boolean
    Dim i As Long, _
        j As Long
        
    i = Len(sChr)
    If (i <> 1) Or (sSeparators = "") Then Exit Function
    i = Len(sSeparators)
    
    For j = 1 To i
        If (sChr = Mid(sSeparators, j, 1)) Then
            IsSepChar = True
            Exit Function
        End If
    
    Next j
        
    IsSepChar = False
    
End Function

Public Function MakeSep(ByVal vData As String) As String

    Dim i As Long, _
        j As Long
        
    i = Len(vData)
    
    For j = 1 To i
    
        If (MakeSep <> "") Then MakeSep = MakeSep + Chr(0)
        MakeSep = MakeSep + Mid(vData, j, 1)
        
    Next j
    
End Function

'' Find the next word/separator(or crlf) pair
Public Function NextWordSep(ByVal vData As String, ByVal uStartPos As Long, lpszText As String, lpszSeparator As String, Optional ByVal szPattern As String, Optional ByVal sCrlf As String = vbCrLf) As Long
    Dim i As Long, _
        j As Long
    
    Dim c As Long, _
        z As Long
    
    Dim uFind As String, _
        uTxt As String
    
    Dim s1 As String, _
        p As Long
    
    Dim s() As String
    
    If (szPattern = "") Then
        szPattern = MakeSep(DefaultSeparators)
    End If
    
    s = BatchParse(szPattern, Chr(0))
    
    i = uStartPos
    If i = 0& Then i = 1&
    
    j = UBound(s)
        
    For c = 0 To j
        s1 = s(c)
        z = InStr(i, vData, s1)
        
        If (z <> 0&) Then
            If (z < p) Or (p = 0&) Then
                p = z
                uFind = s1
            End If
        End If
    Next c
    
    If (sCrlf <> "") Then
        z = InStr(i, vData, sCrlf)
        
        If (z <> 0&) Then
            If (z < p) Or (p = 0&) Then
                p = z + (Len(sCrlf) - 1)
                uFind = sCrlf
            End If
        End If
    End If
    
    If (p <> 0&) Then
        If (p <> i) Then
            uTxt = Mid(vData, i, p - i)
        Else
            uTxt = ""
        End If
        
        lpszText = uTxt
        lpszSeparator = uFind
        
        If ((p + 1) <= Len(vData)) Then
            NextWordSep = p + 1
        Else
            NextWordSep = -1&
        End If
        
        Exit Function
    End If
    
    lpszText = Mid(vData, i)
    lpszSeparator = ""
    NextWordSep = -1&
    
End Function

'' Parse out all words in a text buffer using the default/argument separators and new-line sequence

Public Function Words(ByVal vData As String, Optional ByVal szPattern As String) As String()
    Dim i As Long, _
        j As Long
    
    Dim vOut() As String, _
        v1 As String, _
        v2 As String
    
    i = 1&
    Do While i <> -1&
        i = NextWordSep(vData, i, v1, v2, szPattern)
        
        If v1 <> "" Then
            ReDim Preserve vOut(0& To j)
            vOut(j) = RemoveNulls(v1)
            j = j + 1
        End If
        
        If v2 <> "" Then
            ReDim Preserve vOut(0& To j)
            vOut(j) = RemoveNulls(v2)
            j = j + 1
        End If
    
    Loop
    
    Words = vOut
    
End Function

'' Wrap Text and Parse Lines (w/ Cr+Lf)

Public Function WrapLines(ByVal hDC As Long, cbLimit As Long, ByVal szText As String, _
                          Optional ByVal sSeparators As String = DefaultSeparators, Optional ByVal sCrlf As String = vbCrLf, _
                          Optional ByVal fHasAccel As Boolean) As String()

    Dim sWords() As String, _
        i As Long, _
        j As Long, _
        lWords As Long, _
        lHeight As Long
    
    Dim sCurrent As String, _
        sPrev As String, _
        sLines() As String, _
        nLines As Long
    
    Dim lpSize As SIZEAPI, _
        lpWordSize As SIZEAPI, _
        bNoInc As Boolean
    
    Dim lWidth As Long
    
    On Error Resume Next
    
    If szText = "" Then Exit Function
    sWords = Words(szText)
    
    i = UBound(sWords)
    
    sCurrent = sWords(0&)
    sPrev = sCurrent
    
    ReDim sLines(0&)
    lWords = 1
    
    Do While sCurrent <> ""
    
        MeasureText_API hDC, sWords(j), lpWordSize
        MeasureText_API hDC, sCurrent, lpSize
                
        If (lWords = 1) And (cbLimit >= WrapLimit_Minimum) Then
            If lpSize.cx > cbLimit Then
                cbLimit = lpSize.cx
            End If
        
            If lpWordSize.cx > cbLimit Then
                cbLimit = lpWordSize.cx
            End If
        End If
        
        If (lpSize.cx > lWidth) Then lWidth = lpSize.cx
        
        If ((cbLimit >= WrapLimit_Minimum) And (lpSize.cx > cbLimit)) Or ((sCrlf <> "") And (sWords(j) = sCrlf)) Then
            sLines(nLines) = sPrev
            
            nLines = nLines + 1
            
            ReDim Preserve sLines(0& To nLines)
            
            lWords = 1
            
            If (IsSepChar(sWords(j), sSeparators) = True) Or ((sCrlf <> "") And (sWords(j) = sCrlf)) Then
                sLines(nLines - 1) = sLines(nLines - 1) + sWords(j)
                
                If (j < i) Then
                    sPrev = sWords(j + 1)
                Else
                    nLines = nLines + 1
                    ReDim Preserve sLines(0& To nLines)
                    sPrev = ""
                    bNoInc = False
                End If
                
                j = j + 1
            Else
                sPrev = sWords(j)
            End If
            
            sCurrent = sPrev
            bNoInc = True
        End If
        
        If (j < i) And (bNoInc = False) Then
            sPrev = sCurrent
            
            j = j + 1
            sCurrent = sCurrent + sWords(j)
            lWords = lWords + 1
        Else
            If bNoInc = True Then
                bNoInc = False
            ElseIf (j >= i) Then
                Exit Do
            End If
        End If
    Loop

    sLines(nLines) = sCurrent
    If sLines(nLines) = "" Then
        ReDim Preserve sLines(0& To (nLines - 1))
    Else
        nLines = nLines + 1
    End If
    
    If (cbLimit < WrapLimit_Minimum) Then
        cbLimit = lWidth
    End If
    
    WrapLines = sLines
    
End Function

Public Function StreamToLines(ByVal Stream As String) As String()
    
    Dim i As Long, _
        b As Long, _
        d As Long, _
        l As Long
        
    Dim s() As String
    
    On Error Resume Next
    
    l = 1
    i = InStr(l, Stream, vbCrLf)
    
    Do While i <> 0
        ReDim Preserve s(0 To d)
        
        If ((i - l) + 1) > 0 Then
            s(d) = Mid(Stream, l, (i - l) + 1)
        End If
        
        d = d + 1
        l = i + 2
        
        i = InStr(l, Stream, vbCrLf)
    Loop
    
    If (l < Len(Stream)) Then
        ReDim Preserve s(0 To d)
        s(d) = Mid(Stream, l)
    End If
    
    StreamToLines = s

End Function

'' Parse a string and optionally measure/retrieve accelerator info (for use with menus)

Public Function Parse_String(ByVal hDC As Long, ByVal lpszText As String, Optional lpdwLines As Long, Optional lpdwLimit As Long, Optional lpdwAccel As Long, _
                             Optional ByVal lpSize As Long, Optional ByVal sCrlf As String = vbCrLf) As String()
     
    Dim x As String, _
        y As String
        
    Dim wrapStr() As String, _
        OutStr() As String
    
    Dim i As Long, _
        j As Long
        
    Dim a As Long, _
        b As Long
        
    Dim nCount As Long, _
        cbText As SIZEAPI
        
    Dim dwLimit As Long, _
        n As Long
                
    Dim fHasAccel As Boolean
                
    On Error Resume Next
    
    Erase OutStr
    
    If (IsMissing(lpdwAccel) = False) Then
        fHasAccel = True
    End If
    
    If (IsMissing(lpdwLimit) = False) Then
        dwLimit = lpdwLimit
    End If
    
    OutStr = WrapLines(hDC, dwLimit, lpszText, , , fHasAccel)

    b = -1&
    b = UBound(OutStr)
    
    nCount = b + 1
        
    If (nCount > 0&) Then
        Parse_String = OutStr
        OutStr = Empty
    End If
    
End Function


'' re-assemble parsed-out lines

Public Function Reassemble(lpLines() As String) As String
    Dim s1 As String, _
        i As Long, _
        j As Long
            
    On Error Resume Next
    
    i = -1&
    i = UBound(lpLines)
    
    For j = 0& To i
        s1 = s1 + lpLines(j)
    Next j
    
    Reassemble = s1
    
End Function



Public Function NoSpace(ByVal varData As String) As String
    Dim i As Long, _
        j As Long
        
    Dim s As String, _
        u As String
    
    j = Len(varData)
    
    For i = 1 To j
    
        u = Mid(varData, i, 1)
        
        If (u <> " ") And (u <> Chr(9)) Then
            s = s + u
        End If
    Next i
    
    NoSpace = s
End Function


Public Function NoQuote(ByVal varData As String) As String
    Dim i As Long, _
        j As Long
        
    Dim s As String
    
    j = Len(varData)
    
    For i = 1 To j
        Select Case Mid(varData, i, 1)
                            
            Case Is <> Chr(34)
                s = s + Mid(varData, i, 1)
                
        End Select
    Next i
    
    NoQuote = s
End Function

Public Function NoPrefixSpace(ByVal varData As String) As String
    Dim i As Long, _
        j As Long
        
    Dim s As String, _
        sw As Boolean
    
    j = Len(varData)
    
    For i = 1 To j
        If ((Mid(varData, i, 1) <> " ") And (Mid(varData, i, 1) <> Chr(9))) Or (sw = True) Then
                            
            s = s + Mid(varData, i, 1)
            If (sw = False) Then sw = True
            
        End If
    Next i
    
    NoPrefixSpace = s
End Function

Public Function OneSpace(ByVal varData As String) As String

    Dim i As Long, _
        j As Long
        
    Dim sStr As String, _
        sp As Boolean
    
    j = Len(varData)
    
    For i = 1 To j
        
        
        If Mid(varData, i, 1) = " " Then
            If (sp = False) Then
                sStr = sStr + " "
                sp = True
            End If
        Else
            sStr = sStr + Mid(varData, i, 1)
            sp = False
        End If
    
    Next i
    
    OneSpace = sStr
    
End Function

Public Sub NoEof(varData As String)
    Dim i As Long, _
        j As Long
        
    Dim s As String
    
    j = Len(varData)
    
    For i = 1 To j
        Select Case Mid(varData, i, 1)
            
            Case Chr(26), vbCr, vbLf, Chr(0)
                s = s + ""
                
            Case Else
                s = s + Mid(varData, i, 1)
                
        End Select
    Next i
    
    varData = s
    
End Sub

Public Function USpace(ByVal varData As String) As String
    Dim i As Long, _
        j As Long
        
    Dim s As String
    
    j = Len(varData)
    
    For i = 1 To j
        Select Case Mid(varData, i, 1)
            
            Case " "
                s = s + "_"
                
            Case Else
                s = s + Mid(varData, i, 1)
                
        End Select
    Next i
    
    USpace = s
End Function

Public Function SUnder(ByVal varData As String) As String
    Dim i As Long, _
        j As Long
        
    Dim s As String
    
    j = Len(varData)
    
    For i = 1 To j
        Select Case Mid(varData, i, 1)
            
            Case "_"
                s = s + " "
                
            Case Else
                s = s + Mid(varData, i, 1)
                
        End Select
    Next i
    
    SUnder = s

End Function

Public Function GetPlainCaption(ByVal vData As String) As String

    Dim i As Long, _
        j As Long, _
        s As String, _
        t As String

    j = Len(vData)
    
    For i = 1 To j
        If (i < (j - 1)) Then
            s = Mid(vData, i, 3)
            If (s = (vbCr + vbCrLf)) Then
                i = i + 3
            End If
        End If
        
        If (i < j) Then
            s = Mid(vData, i, 2)
            
            If Mid(s, 1, 1) = "&" Then
                t = t + Mid(s, 2)
                i = i + 1
            ElseIf s <> vbCrLf Then
                t = t + Mid(s, 1, 1)
            Else
                i = i + 1
            End If
        Else
            t = t + Mid(vData, i)
        End If
    Next i
    
    GetPlainCaption = t
            
End Function

Public Function GetAccelPos(ByVal vData As String) As Long

    Dim i As Long, s As String, t As String
    
    For i = 1 To Len(vData)
        s = Mid(vData, i, 1)
        If s = "&" Then
            If Len(vData) > i Then
                If Mid(vData, i + 1, 1) <> "&" Then
                    GetAccelPos = Len(t) + 1
                End If
            End If
        Else
            t = t + s
        End If
    Next i
    
End Function

Public Function ExtractString(ByVal lpsz As Long, Optional fIsUnicode As Boolean) As String
    Dim xUni As Boolean

    Dim lpStrNew As String, _
        a As Long, _
        b As Long

    Dim tIn As Byte, _
        wIn As Integer

    If IsMissing(fIsUnicode) Then
        If (g_WinVersion = Windows2000) Or (g_WinVersion = WindowsNT) Then
            xUni = True
        Else
            xUni = False
        End If
    Else
        xUni = fIsUnicode
    End If
    
    
    a = lpsz
    
    If (xUni = True) Then
        CopyMemory wIn, ByVal a, 2&
        b = wIn
    Else
        CopyMemory tIn, ByVal a, 1&
        b = tIn
    End If
    
    Do While b <> 0&
        If xUni = True Then
            lpStrNew = lpStrNew + ChrW(b)
            a = a + 2
        Else
            lpStrNew = lpStrNew + Chr(b)
            a = a + 1
        End If
    
        If (xUni = True) Then
            CopyMemory wIn, ByVal a, 2&
            b = wIn
        Else
            CopyMemory tIn, ByVal a, 1&
            b = tIn
        End If
            
    Loop
    
    ExtractString = lpStrNew
End Function


''' Copyright (C) 2001 Nathan Moschkin

''' ****************** NOT FOR COMMERCIAL USE *****************
''' Inquire if you would like to use this code commercially.
''' Unauthorized recompilation and/or re-release for commercial
''' use is strictly prohibited.
'''
''' please send changes made to code to me at the address, below,
''' if you plan on making those changes publicly available.

''' e-mail questions or comments to logeist@tampabay.rr.com


