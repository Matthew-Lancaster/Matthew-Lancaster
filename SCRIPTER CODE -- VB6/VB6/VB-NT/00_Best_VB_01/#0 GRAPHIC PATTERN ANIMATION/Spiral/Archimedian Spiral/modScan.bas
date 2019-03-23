Attribute VB_Name = "modScan"
Option Explicit

Private Enum ScanTypeConstants
    scnInt = 0
    scnStr = 1
    scnFlt = 2
    scnFmt = 3
    scnByt = 4
    scnChr = 5
    scnHex = 6
    scnOct = 7
End Enum

Private Type SCAN_STRUCT

    Record As Boolean
    LongScan As Boolean
    
    ScanType As ScanTypeConstants
    Range As String
    
    Length As Long
    
    Inline_Pos As Long
    Inline_Len As Long
    
End Type


Public Function TrimWSpace(ByVal varData As String) As String

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
    
    TrimWSpace = sStr
    
End Function


'' String Token Function (works like in C/C++)
Public Function StrTok(Optional ByVal ScanString As String, Optional ByVal Token As String, Optional ByVal SyncTok As Boolean, Optional ByVal SkipQuotes As Boolean, Optional ByVal QuoteChar As String, Optional ByVal EscapeChar As String = "\") As String
    Static Stored As String
        
    Dim i As Long, _
        l As Long, _
        vCh As String, _
        OutStr As String
        
    Dim inQ As Boolean, _
        skipQ As Boolean, _
        qChar As String, _
        eChar As String
        
    If (QuoteChar = "") Then qChar = Chr(34) Else qChar = QuoteChar
    
    If (SyncTok = True) Then
        If (ScanString = "") Then
            StrTok = Stored
        Else
            Stored = ScanString
        End If
        
        Exit Function
    End If
    
    If (Token = "") Then Exit Function
        
    If ScanString <> vbNullString Then
        Stored = ScanString
    End If
    
    If Stored = vbNullString Then Exit Function
    
    If Mid(Stored, 1, Len(Token)) = Token Then
        Stored = Mid(Stored, Len(Token) + 1)
        StrTok = ""
        Exit Function
    End If
    
    skipQ = SkipQuotes
    
    For i = 1 To Len(Stored)
        
        vCh = Mid(Stored, i, Len(Token))
        If (vCh = Token) And ((inQ = False) Or (skipQ = False)) Then
            
            If ((i + Len(Token)) <= Len(Stored)) Then
                StrTok = OutStr
                Stored = Mid(Stored, i + Len(Token))
                Exit Function
            Else
                Exit For
            End If
            
        Else
            vCh = Mid(Stored, i, 1)
            OutStr = OutStr + vCh
            
            If (vCh = qChar) And (skipQ = True) Then
                inQ = (Not inQ)
                
                If (i > 1) Then
                    vCh = Mid(Stored, i - 1, 1)
                    If (vCh = eChar) Then
                        inQ = (Not inQ)
                    End If
                End If
            End If
        End If
    Next i
    
    StrTok = OutStr
    Stored = vbNullString
    
End Function


'' Function to retrieve a quote from a string of data.  The quote character
'' must be: exactly one before, exactly at, or anywhere after the location specified by
'' 'Pos'.

Public Function QuoteFromHere(ByVal StrData As String, ByVal Pos As Long, Optional ByVal QuoteChar As String, Optional ByVal EscapeChar As String = "\", Optional ByVal WithQuotes As Boolean) As String

    Dim i As Long, _
        l As Long, _
        vCh As String, _
        OutStr As String
        
    Dim inQ As Boolean, _
        qChar As String, _
        eChar As String
        
    If (QuoteChar = "") Then qChar = Chr(34) Else qChar = QuoteChar
    eChar = EscapeChar
    
    If (eChar = "") Then eChar = "\"
    
    If Pos > 1 Then l = Pos - 1 Else l = Pos
    l = InStr(l, StrData, qChar)
    
    If (l = 0) Then Exit Function
    
    For i = l To Len(StrData)
        
        vCh = Mid(StrData, i, 1)
        
        If (vCh = qChar) Then
            inQ = (Not inQ)
            
            If (i > 1) Then
                If (i > 2) Then
                    vCh = Mid(StrData, i - 2, 2)
                Else
                    vCh = Mid(StrData, i - 1, 1)
                End If
                
                If ((Mid(vCh, 2) = eChar) And (InStr(1, vCh, eChar + eChar) = 0)) Then
                    OutStr = OutStr + qChar
                    inQ = (Not inQ)
                End If
            End If
            
            If (inQ = False) Then Exit For
        
        ElseIf (vCh = eChar) Then
            
            If (i > 1) Then
                vCh = Mid(StrData, i - 1, 1)
                If (vCh = eChar) Then
                    OutStr = OutStr + eChar
                End If
            End If
        
        ElseIf (vCh <> eChar) Then
            OutStr = OutStr + vCh
        
        End If
    
    Next i
    
    If (WithQuotes = True) Then
        OutStr = qChar + OutStr + qChar
    End If
    
    QuoteFromHere = OutStr
    
End Function

'' Uses SyncTok so it can be used between searches
Public Function BatchParse(ByVal Scan As String, ByVal Separator As String, Optional ByVal SkipQuote As Boolean) As String()
    Dim sOut() As String
    
    Dim i As Long, _
        j As Long
        
    Dim s As String, _
        t As String, _
        u As String
        
    Dim tSave As String
        
    tSave = StrTok(, , True)
    
    
    s = Scan
    t = StrTok(s, Separator, , SkipQuote)
    
    Do While t <> ""
        ReDim Preserve sOut(0 To i)
        sOut(i) = t
        i = i + 1
    
        t = StrTok(, Separator, , SkipQuote)
    Loop
    
    StrTok tSave, , True
        
    BatchParse = sOut
End Function


Private Function PreScan(ByVal Format As String, Scans() As SCAN_STRUCT, Strings() As String) As Long

    
    Dim iStr As String, _
        iFmt As String
        
    Dim s As String, _
        t As String, _
        u As String, _
        v As String, _
        w As String
        
    Dim Fmts() As SCAN_STRUCT, _
        z() As String, _
        Record As Boolean, _
        ScanLen As Long, _
        IsLong As Boolean
                
    Dim i As Long, _
        j As Long, _
        g As Long
        
    Dim a As Long, _
        b As Long, _
        f As Long
        
    Dim c As Long, _
        e(0 To 1) As String
        
    Dim sC As Long, _
        mRange As Boolean
        
    On Error Resume Next
    
    iFmt = TrimWSpace(Format)
    
    Erase Strings
    
    j = 1&
    i = InStr(j, iFmt, "%")
    
    If (i = 0) Then
        PreScan = 1
        ReDim Strings(0)
        Strings(0) = Format
        
        Exit Function
    End If
    
    Do While i <> 0
    
        ReDim Preserve Strings(0 To sC)
        
        If (i <> 1) And ((i - j) > 0) Then
        
            Strings(sC) = Mid(iFmt, j, (i - j))
            
            sC = sC + 1
            ReDim Preserve Strings(0 To sC)
            
        End If
        
        Record = True
        ScanLen = -1&
        
        IsLong = False
        
ReScan:
        If Mid(iFmt, i + 1, 1) = "[" Then
            
            j = InStr(i + 1, iFmt, "]")
            If (j = 0) Then Exit Function
            
            If (Mid(iFmt, j - 1, 1) = "\") Then
            
                Do While Mid(iFmt, j - 1, 1) = "\"
                
                    j = InStr(i + 1, iFmt, "]")
                    If (j = 0) Then Exit Function
                    
                Loop
                
            End If
            
            t = Mid(iFmt, i, (j - i) + 1)
                
            Strings(sC) = t
            sC = sC + 1
            
            ReDim Preserve Fmts(0 To b)
            
            Fmts(b).Inline_Len = (j - i) + 1
            Fmts(b).Inline_Pos = i
            Fmts(b).Record = Record
            Record = True
            
            Fmts(b).ScanType = scnFmt
            Fmts(b).Length = ScanLen
            
            
            t = Mid(t, 3, Len(t) - 3)
            s = Mid(t, 1, 1)
            
            f = Len(t)
            
            For a = 1 To f Step 0
            
                s = Mid(t, a, 1)
                
                If (s = "-") And ((a <> f) Or (a <> 1)) Then
                    ReDim Preserve z(0 To c)
                    z(c) = s
                    
                    c = c + 1
                    a = a + 1
                    
                Else
                
                    If (s = "\") Then
                    
                        v = Mid(t, a, 2)
                        s = Mid(t, a + 1, 1)
                        
                        Select Case s
                            Case "0"
                                u = "&O"
                            
                            Case "x"
                                u = "&H"
                            
                        End Select
                        
                        a = a + 2
                        
                        s = Mid(t, a, 1)
                        
                        g = 1&
                        Do While _
                            (InStr(1, "0123456789", s) And ((u = "") And (g <= 3))) Or _
                            (InStr(1, "0123456789ABCDEF", UCase(s)) And ((u = "&H") And (g <= 2))) Or _
                            (InStr(1, "01234567", s) And ((u = "&O") And (g <= 3)))
                                                        
                            v = v + s
                            a = a + 1
                            g = g + 1
                            
                            If (a >= f) Then Exit Do
                            
                            s = Mid(t, a, 1)
                        Loop
                        
                        s = v
                    Else
                        a = a + 1
                    End If
                    
                    If (mRange = True) Then
                        ReDim Preserve z(0 To c)
                        z(c) = w + s
                        mRange = False
                        c = c + 1
                    Else
                        If (Mid(t, a, 1) = "-") And (a <> f) Then
                            w = s + "-"
                            mRange = True
                            a = a + 1
                        Else
                            ReDim Preserve z(0 To c)
                            z(c) = s
                            c = c + 1
                        End If
                    End If
                End If
            
            Next a
            
            c = c - 1
            For a = 0 To c
            
                If (InStr(1, z(a), "-")) Then
                    e(0) = StrTok(z(a), "-")
                    e(1) = StrTok(, "-")
                Else
                    e(0) = z(a)
                    e(1) = ""
                End If
                    
                For f = 0 To 1
                
                    s = e(f)
                    
                    If (Len(s) >= 2) Then
                        Select Case Mid(s, 1, 2)
                                                    
                            Case "\x"
                                s = "&H" + Mid(s, 3)
                                s = Chr(Val(s))
                            
                            Case "\0"
                                s = "&O" + Mid(s, 3)
                                s = Chr(Val(s))
                            
                            Case "\1" To "\9"
                            
                                s = Mid(s, 2)
                                s = Chr(Val(s))
                            
                            Case "\\", "\'", ("\" + Chr(34)), "\[", "\]"
                                s = Mid(s, 2)
                            
                            Case "\r"
                                s = vbCr
                            
                            Case "\n"
                                s = vbLf
                            
                            Case "\t"
                                s = vbTab
                            
                        End Select
                    End If
                    
                    e(f) = s
                    If (e(1) = "") Then Exit For
                    
                Next f
                                                
                If (e(1) <> "") Then
                    For f = Asc(e(0)) To Asc(e(1))
                        Fmts(b).Range = Fmts(b).Range + Chr(f)
                    Next f
                Else
                    Fmts(b).Range = Fmts(b).Range + e(0)
                End If
                            
            Next a
            
            b = b + 1
            j = j + 1
        
        Else
            Select Case Mid(iFmt, i + 1, 1)
            
                Case "%"
                                        
                    j = i + 2
                    
                    Strings(sC) = Mid(iFmt, i, 2)
                    sC = sC + 1
                    
                    GoTo PctLoop
                    
                Case "s", "S", "b", "B", "c", "C", _
                     "d", "f", "x", "o", "X", "O"
                    
                    ReDim Preserve Fmts(0 To b)
                    
                    j = i + 1
                    
                    Fmts(b).Inline_Len = (j - i) + 1
                    Fmts(b).Inline_Pos = i
                    
                    Fmts(b).Record = Record
                    Record = True
                    
                    Fmts(b).Length = ScanLen
                    
                    Select Case (Mid(iFmt, i + 1, 1))
                    
                        Case "s", "S"
                            Fmts(b).ScanType = scnStr
                            
                        Case "C", "c"
                            Fmts(b).ScanType = scnChr
                            
                        Case "B", "b"
                            Fmts(b).ScanType = scnByt
                            
                        Case "d"
                            Fmts(b).ScanType = scnInt
                            Fmts(b).LongScan = IsLong
                        
                        Case "x", "X"
                            Fmts(b).ScanType = scnHex
                            Fmts(b).LongScan = IsLong
                        
                        Case "o", "O"
                            Fmts(b).ScanType = scnOct
                            Fmts(b).LongScan = IsLong
                        
                        Case "f"
                            Fmts(b).ScanType = scnFlt
                            Fmts(b).LongScan = IsLong
                            
                    End Select
                    
                    If (IsLong) Then
                        i = i - 1
                        IsLong = False
                    End If
                    
                Case "l"
                
                    IsLong = True
                    i = i + 1
                    
                    GoTo ReScan
                    
                Case "*"
                
                    Record = False
                    i = i + 1
                    
                    GoTo ReScan
                    
                Case "0" To "9"
            
                    a = i + 1
                    Do While (Mid(iFmt, a, 1) >= "0") And (Mid(iFmt, a, 1) <= "9")
                        a = a + 1
                    Loop
                    
                    ScanLen = Val(Mid(iFmt, i + 1, (a - i) - 1))
                    j = a
                    
                    GoTo ReScan
                    
                Case Else
                    
                    PreScan = -1&
                    Exit Function
                    
            End Select
        
            b = b + 1
            j = j + 1
    
            Strings(sC) = Mid(iFmt, i, (j - i))
            sC = sC + 1
            
        End If

PctLoop:
        i = InStr(j, iFmt, "%")
    Loop
    
    If (j <= Len(iFmt)) Then
    
        ReDim Preserve Strings(0 To sC)
        Strings(sC) = Mid(iFmt, j)
        sC = sC + 1
        
    End If
    
    For i = 0 To (sC - 1)
        If (i > (sC - 1)) Then Exit For
        
        b = 1
        j = InStr(b, Strings(i), "%%")
        
        Do While (j <> 0)
            
            s = Mid(Strings(i), b, (j - b) + 1) + Mid(Strings(i), j + 2)
            Strings(i) = s
            
            b = j + 1
            j = InStr(b, Strings(i), "%%")
            
        Loop
        
        If (i > 0) Then
            If (((Mid(Strings(i), 1, 1) <> "%") And (Len(Strings(i)) > 1)) Or ((Strings(i) = "%") Or (Strings(i) = "%%"))) And _
               (((Mid(Strings(i - 1), 1, 1) <> "%") And (Len(Strings(i - 1)) > 1)) Or ((Strings(i - 1) = "%") Or (Strings(i - 1) = "%%"))) Then
        
                Strings(i - 1) = Strings(i - 1) + Strings(i)
                
                For b = i To (sC - 2)
                    Strings(b) = Strings(b + 1)
                Next b
                
                ReDim Preserve Strings(0 To (sC - 2))
                sC = sC - 1
                i = i - 1
                
            End If
        End If
        
    Next i
    
    Scans = Fmts
    PreScan = 1
    
End Function

Public Function ScanString(ByVal InputStr As String, ByVal Format As String, ParamArray OutVar()) As Long
    On Error Resume Next
    
    Dim ActualOut() As Variant, _
        pOut() As Variant
        
    
    '' Emulates the functionality of C's sscanf() function,
    '' adapted for VB programming
    
    '' Following are allowed:
    ''
    ''
    '' Numeric
    ''
    '' d, ld, x,X,lx,lX, o, lo, f, df
    ''
    '' c byte
    '' s string (no spaces)
    ''
    '' [\x##-\x##\0##-\0##\##-\##A-Z etc...]
    ''
    '' escape sequences translate to characters,
    '' as do characters in single quotes.  The
    '' hyphen symbol is synonymous with VB's To directive.
    '' numeric ranges evaluate implicitly (octal and hexidecimal formatting
    '' is accepted à la C.)
    
    '' Other escape sequences = " \[, \], \\, \', \n, \r, \t "
    ''
    '' A literal percent character is %%
    '' Width and no-reads are processed the same, the whole format is:
    ''
    '' %*[Width][l][d,f,o][s,c][...]
    
    '' sscanf "Jan 14, 2001 <format input>", "%s %d, %ld <%[\x20A-Za-z0-9]>", s, d, ld, fmt
    

    Dim Scans() As SCAN_STRUCT
    
    On Error Resume Next
    
    Dim i As Long, _
        j As Long
    
    Dim sT As Long
    
    Dim a As Long, _
        b As Long, _
        c As Long, _
        d As Long
    
    Dim strParts() As String, _
        s As String, _
        t As String, _
        v As String, _
        cSize As Long
    
    Dim cVal As Long, _
        fVal As Double, _
        sVal As String
    
    Dim rCnt As Long
    
    i = -1&
    i = UBound(OutVar)
    
    ActualOut = OutVar
    t = TrimWSpace(InputStr)
    
    If (i = 0) Then
        If IsArray(OutVar(0)) = True Then
            ActualOut = OutVar(0)
        End If
    End If
        
    i = -1&
    i = UBound(ActualOut)
    
    ReDim pOut(0 To i)
    
    i = PreScan(Format, Scans, strParts)
    
    If (i <= 0) Then
        ScanString = -1&
        Exit Function
    End If
    
    i = -1&
    i = UBound(Scans)
    
    c = 1
                            
    If (Scans(0).Inline_Pos = 1) Then d = 1 Else d = 0
    
    For j = 0 To i
        
        a = Scans(j).Inline_Pos
        If (a <> 1) Then
            
            If (InStr(c, t, strParts(d)) <> c) Then
                ScanString = -1&
                Exit Function
            End If
            
            c = c + Len(strParts(d))
            d = d + 2
            
        End If
                
        sT = Scans(j).ScanType
        
        Select Case sT
        
            Case scnStr, scnByt, scnChr
                '' String scan
                
                If (sT = scnStr) Then
                    cVal = InStr(c, t, " ")
                Else
                    cVal = c + 1
                End If
                
                If (cVal = 0) Then
                    v = Mid(t, c)
                    c = Len(t) + 1
                Else
                    v = Mid(t, c, cVal - c)
                    c = cVal
                End If
        
                If (sT = scnByt) Then
                    If (v > Chr(255)) Then
                        ScanString = -1&
                        Exit Function
                    End If
                End If
                
                If (Scans(j).Record = True) Then
                
                    pOut(rCnt) = v
                    rCnt = rCnt + 1&
                    
                End If
            
            Case scnFmt
                '' Range scan
                
                cVal = 0
                s = Mid(t, c, 1)
                
                v = ""
                
                Do While InStr(1, Scans(j).Range, s)
                    
                    v = v + s
                    cSize = cSize + 1
                    
                    c = c + 1
                    If (cSize = Scans(j).Length) Then Exit Do
                    s = Mid(t, c, 1)
                    
                Loop

                If (Scans(j).Length <> -1&) And (cSize <> Scans(j).Length) Then
                    ScanString = -1&
                    Exit Function
                End If
                
                If (Scans(j).Record = True) Then
                    pOut(rCnt) = v
                    rCnt = rCnt + 1
                End If
                
            Case scnInt, scnFlt, scnHex, scnOct
                '' Numeric scan
                
                cVal = 0
                s = Mid(t, c, 1)
                
                If (sT = scnOct) Then
                    v = "&O"
                ElseIf (sT = scnHex) Then
                    v = "&H"
                Else
                    v = ""
                End If
                
                Do While _
                    (InStr(1, "0123456789", s) And ((sT = scnInt))) Or _
                    (InStr(1, "0123456789.Ee-+", s) And ((sT = scnFlt))) Or _
                    (InStr(1, "0123456789ABCDEF", UCase(s)) And ((sT = scnHex))) Or _
                    (InStr(1, "01234567", s) And ((sT = scnOct)))

                    v = v + s
                    cSize = cSize + 1
                    
                    c = c + 1
                    
                    If (cSize = Scans(j).Length) Then Exit Do
                    s = Mid(t, c, 1)
                    
                Loop


                If (Scans(j).Length <> -1&) And (cSize <> Scans(j).Length) Then
                    ScanString = -1&
                    Exit Function
                End If
                                
                Select Case sT
                    
                    Case scnInt, scnHex, scnOct
                        cVal = CLng(Val(v))
                            
                        If (Scans(j).LongScan = True) Then
                            If (Scans(j).Record = True) Then
                                pOut(rCnt) = cVal
                                rCnt = rCnt + 1
                                
                            End If
                        ElseIf (cVal >= &HFFFF) And (cVal <= &H7FFF) Then
                            
                            If (Scans(j).Record = True) Then
                                pOut(rCnt) = CInt(cVal)
                                rCnt = rCnt + 1
                            End If
                        Else
                            ScanString = -1&
                            Exit Function
                        End If
                    
                    Case scnFlt
                        fVal = Val(v)
                        
                        If (Scans(j).LongScan = True) Then
                            If (Scans(j).Record = True) Then
                                pOut(rCnt) = fVal
                                rCnt = rCnt + 1
                            End If
                        ElseIf (fVal >= -3.402823E+38) And (fVal <= 3.402823E+38) Then
                            If (Scans(j).Record = True) Then
                                pOut(rCnt) = CSng(fVal)
                                rCnt = rCnt + 1
                            End If
                        Else
                            ScanString = -1&
                            Exit Function
                        End If
                        
                End Select
                
        End Select
        
    Next j
    
    If (c <= Len(t)) Then
        
        s = Mid(t, c)
        b = 1
        j = InStr(b, s, "%%")
        
        Do While (j <> 0)
            
            s = Mid(s, b, (j - b) + 1) + Mid(s, j + 2)
            s = s
            
            b = j + 1
            j = InStr(b, s, "%%")
            
        Loop
        
        If (s <> strParts(d)) Then
            ScanString = -1&
            Exit Function
        End If
        
    End If
                
    If (IsArray(OutVar(0))) Then
        For j = 0 To (rCnt - 1)
            ActualOut(j) = pOut(j)
        Next j
        
        OutVar(0) = ActualOut
    Else
        For j = 0 To (rCnt - 1)
            OutVar(j) = pOut(j)
        Next j
    End If
    
    ScanString = rCnt
        
End Function


