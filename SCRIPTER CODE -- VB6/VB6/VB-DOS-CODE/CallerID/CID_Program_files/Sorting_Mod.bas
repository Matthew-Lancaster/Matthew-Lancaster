Attribute VB_Name = "Sorting_Mod"
'Public Number3$
'Public Linea$


Public Sub Sort_Ci_FastFind()
'Fast Find Used in Hall Of fame

Number$ = Mid$(Linea$, 1, 16)
Code$ = Mid$(Linea$, CodePos, 2)
If Number$ = "WITHHELD        " Then Code$ = "01"
If Mid$(Number$, 7, 2) = "  " Then Code$ = "02"

aq = -1
keek = 1
Do
    weds = InStr(keek, FastFind2String$, "***v*" + Number$)
    If weds > 0 And Mid$(FastFind2String$, weds + 46, 2) = Code$ Then aq = 1
    If weds = 0 Then aq = 0
    keek = weds + 1
Loop Until aq > -1

If aq = 0 Then
    weds = InStr(keek, FastFind2String$, "***v*" + Number$)
    If weds > 0 Then aq = 1
End If

If aq = 1 Then YuBt = Val(Mid$(FastFind2String$, weds - 5, 5))
If aq = 0 Then YuBt = -1

If YuBt > 0 Or aq = 1 Then
    Find$ = Trim(Mid$(Ciform2.List2.List(YuBt), FindPos))
    Code$ = Mid$(Ciform2.List2.List(YuBt), CodePos, 2)
    Counter$ = Mid$(Ciform2.List2.List(YuBt), CountPos, 4)
End If

Word$ = "    "

If Mid$(Code$, 1, 2) = "01" Then
    LSet Word$ = Mid$(Find$, 1, 4)
    If Mid$(Number$, 1, 2) = "00" Then Word$ = "INTL"

End If

If Mid$(Code$, 1, 2) = "02" Then
    LSet Word$ = Mid$(Find$, 1, 4)
End If

If YuBt = -1 Then
    Find$ = "NEW NUMBER"
End If


End Sub





Public Sub FastFind()
'Fast Find() Used From CI & CiEdit

'FastFind2String$ = "?"

'For rt = 0 To Ciform2.List2.ListCount
'    FastFind2String$ = FastFind2String$ + Format$(rt, "00000") + "***v*" + Ciform2.List2.List(rt)
'Next

If LTrim(Code$) = "" Then Code$ = "00"

aq = -1
keek = 1
Do
    weds = InStr(keek, FastFind2String$, "***v*" + Number$)
    If weds > 0 And Mid$(FastFind2String$, weds + 46, 2) = Code$ Then aq = 1
    If weds = 0 Then aq = 0
    keek = weds + 1
Loop Until aq > -1

If aq = 1 Then YuBt = Val(Mid$(FastFind2String$, weds - 5, 5)) + 1
If aq = 0 Then YuBt = -1

If aq > 0 Then Foundfind = 1 Else Foundfind = 0

aq = 0

If YuBt <> -1 Then
    If InStr(Ciform2.List2.List(YuBt), Number$) = 0 Then
        For r1 = (YuBt - 3) To (YuBt + 3)
            If InStr(Ciform2.List2.List(r1), Number$) > 0 Then YuBt = r1
            'Exit For
        Next
    End If
End If

If YuBt < Ciform2.List2.ListCount Then
If YuBt > -1 Then
'If InStr(Ciform2.List2.List(YuBt), Number$) = 0 Then YuBt = YuBt + 1
If InStr(Ciform2.List2.List(YuBt), Number$) = 0 Then Stop
Find$ = Mid$(Ciform2.List2.List(YuBt), FindPos)
Code$ = Mid$(Ciform2.List2.List(YuBt), CodePos, 2)
Counter$ = Mid$(Ciform2.List2.List(YuBt), CountPos, 4)
End If
End If

If YuBt = -1 Then

If LTrim(Word$) <> "" Then Find$ = RTrim(LTrim(Word$)) + " NEW NUMBER"
If LTrim(Word$) = "" Then Find$ = "NEW NUMBER"
End If

End Sub




Public Sub Sort_Ci_Find()

'This is used in the startup Build of Lists

If LTrim(Code$) = "" Then Code$ = "00"
If LTrim(Code$) = "xx22" Then Code$ = "": chic = 1
aq = -1
keek = 1
Do
    weds = InStr(keek, FastFind2String$, "***v*" + Number$)
    If weds > 0 And Mid$(FastFind2String$, weds + 46, 2) = Code$ Then aq = 1
    If weds > 0 And chic = 1 Then aq = 1
    If weds = 0 Then aq = 0
    keek = weds + 1
Loop Until aq > -1

If aq = 1 Then YuBt = Val(Mid$(FastFind2String$, weds - 5, 5))
If aq = 0 Then YuBt = -1

If YuBt > -1 Then
    Find$ = Mid$(Ciform2.List2.List(YuBt), FindPos)
    Code$ = Mid$(Ciform2.List2.List(YuBt), CodePos, 2)
    Counter$ = Mid$(Ciform2.List2.List(YuBt), CountPos, 4)
    Tdate$ = Mid$(Ciform2.List2.List(YuBt), TdatePos, 6) + Mid$(Ciform2.List2.List(YuBt), TdatePos2, 2)
    Ttime$ = Mid$(Ciform2.List2.List(YuBt), TTimePos, 8)
    Xdate$ = Mid$(Ciform2.List2.List(YuBt), TdatePos, 6) + Mid$(Ciform2.List2.List(YuBt), TdatePos2 - 2, 4)
End If

RSet Counter$ = LTrim(Str(Val(Counter$)))
Word$ = "    "

If Mid$(Code$, 1, 2) = "01" Then
    LSet Word$ = Mid$(Find$, 1, 4)
    If Mid$(Number$, 1, 2) = "00" Then Word$ = "INTL"

End If

If Mid$(Code$, 1, 2) = "02" Then
    LSet Word$ = Mid$(Find$, 1, 4)
End If

If InStr(Find$, RTrim(Number$)) = 1 Then LSet Word$ = Mid$(Find$, 1, 4)
End Sub



Public Sub Sort_CiEdit_Find()

Number$ = Mid$(Ciform2.List2.List(YuBt), 1, 16)
Find$ = Mid$(Ciform2.List2.List(YuBt), FindPos)
Code$ = Mid$(Ciform2.List2.List(YuBt), CodePos, 2)
Counter$ = Mid$(Ciform2.List2.List(YuBt), CountPos, 4)
Tdate$ = Mid$(Ciform2.List2.List(YuBt), TdatePos, 6) + Mid$(Ciform2.List2.List(YuBt), TdatePos2, 2)
Ttime$ = Mid$(Ciform2.List2.List(YuBt), TTimePos, 8)

RSet Counter$ = LTrim(Str(Val(Counter$)))
Word$ = "    "

If Mid$(Code$, 1, 2) = "01" Then
    LSet Word$ = Mid$(Find$, 1, 4)
    If Mid$(Number$, 1, 2) = "00" Then Word$ = "INTL"
End If

If Mid$(Code$, 1, 2) = "02" Then
    LSet Word$ = Mid$(Find$, 1, 4)
End If

If InStr(Find$, RTrim(Number$)) = 1 Then LSet Word$ = Mid$(Find$, 1, 4)
End Sub

Public Sub Sort_CiStats_Find(Number3$)

If Number3$ = "11223344556677889900112233445566778899" Then
    Number3$ = "[Spam Block]"
    Exit Sub
End If

'If Number3$ = "157133" Then Stop
If Mid$(Number3$, 1, 4) = "1571" Then
Number3$ = "1571"
End If
aq = -1
keek = 1
Do
    weds = InStr(keek, FastFind2String$, "***v*" + Number3$)
    If weds > 0 Then aq = 1
    If weds = 0 Then aq = 0
    keek = weds + 1
Loop Until aq > -1

If aq = 1 Then YuBt = Val(Mid$(FastFind2String$, weds - 5, 5))
If aq = 0 Then YuBt = -1

If aq > 0 Then Foundfind = 1 Else Foundfind = 0

aq = 0

If YuBt <> -1 Then
    If InStr(Ciform2.List2.List(YuBt), Number3$) = 0 Then
        For r1 = (YuBt - 3) To (YuBt + 3)
            If InStr(Ciform2.List2.List(r1), Number3$) > 0 Then YuBt = r1
            'Exit For
        Next
    End If
End If



Find$ = ""
If YuBt < Ciform2.List2.ListCount Then
    If YuBt > -1 Then
        Find$ = Mid$(Ciform2.List2.List(YuBt), FindPos)
        Code$ = Mid$(Ciform2.List2.List(YuBt), CodePos, 2)
    End If
End If




If YuBt = -1 Then
    If LTrim(Word$) <> "" Then Find$ = RTrim(LTrim(Word$)) + " NEW NUMBER"
'    If LTrim(Word$) = "" Then Find$ = "NEW NUMBER"
    Find$ = "NEW NUMBER"
End If

Number3$ = Find$

End Sub





Public Sub Sort_Dialer_find()

'Dialer Has More

            For rtyu9 = 0 To Ciform2.List2.ListCount - 1
          
                If Number$ = Mid$(Ciform2.List2.List(rtyu9), 1, 16) Then
                    DIALER.List1.AddItem "--------Description ==== " + Mid$(Ciform2.List2.List(rtyu9), 45)
                    DIALER.List1.ListIndex = DIALER.List1.ListCount - 1
                    find5$ = " . " + Mid$(Ciform2.List2.List(rtyu9), 45)
          
                    DIALER.Label1.Caption = Number$ + " -:-" + find5$
                    
                    
                    speak$ = Mid$(Ciform2.List2.List(rtyu9), 45)
                    spe2 = InStr(speak$, "***")
                    If spe2 > 0 Then speak$ = Mid$(speak$, 1, spe2 - 1)
                    If VoiceOnOff = 0 Then
                        If Mid$(LCase$(speak$), 1, 3) = "mum" Then speak$ = Mid$(speak$, 5)
                        Call Ci.SpeakV(speak$)
                    End If
          
                    
                End If
            Next
          
End Sub

