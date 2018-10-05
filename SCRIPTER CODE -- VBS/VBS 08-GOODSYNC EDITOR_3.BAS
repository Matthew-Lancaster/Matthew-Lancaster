Sub GOODSYNC_EDITOR()

FR1 = FreeFile
Open "C:\SCRIPTER\jobs-groups-options - Copy.tic" For Input As FR1
TT_EDITOR = StrConv(Input(LOF(FR1), FR1), vbUnicode)
Debug.Print "----"
Debug.Print LOF(FR1)
Debug.Print Len(TT_EDITOR)
Close #FR1

X = 1
Do
    T1 = InStr(X, TT_EDITOR, StrConv("wait", vbUnicode))
    
    If T1 > 0 Then
        T3 = InStrRev(TT_EDITOR, StrConv("|M", vbUnicode), T1)
        
        T2 = InStr(T3, TT_EDITOR, StrConv("|i=", vbUnicode))
        N3 = Mid(TT_EDITOR, T3, (T2 - T3) + Len(StrConv("|i=", vbUnicode)))
        
        REASSEMBLE_VAR = StrConv("|M:0:|U=0|B:0:|J:0:|i=", vbUnicode)
        TT_EDITOR = Replace(TT_EDITOR, N3, REASSEMBLE_VAR, , 1)
        
        Counter_VAR = Counter_VAR + 1
    End If
    
    X = T1 + 8
Loop Until T1 = 0

FR1 = FreeFile
Open "C:\SCRIPTER\jobs-groups-options.tic" For Output As FR1
Print #FR1, StrConv(TT_EDITOR, vbFromUnicode);
Close #FR1


MsgBox "Done Total " & Counter_VAR & " Proecesses"

End Sub