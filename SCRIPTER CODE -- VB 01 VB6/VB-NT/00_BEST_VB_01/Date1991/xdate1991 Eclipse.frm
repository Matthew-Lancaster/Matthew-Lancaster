VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tagger()
Dim Tagge2()
Dim XvRs1()
Dim XvRs2()
Dim XvRs4()
Dim XvRs5()
Dim MInfo2$(10)







Private Sub Form_Load()


'------------------------------------------------------------
ReDim Tagger(50)
ReDim Tagge2(50)
ReDim XvRs1(50)
ReDim XvRs2(50)
ReDim XvRs4(50)
ReDim XvRs5(50)


'MInfo2$(6) = "Lunar Eclipse"
MInfo2$(1) = "Solar Eclipse"
MInfo2$(2) = "T -Total"
MInfo2$(3) = "A -Annular"
MInfo2$(4) = "H -Hybrid(Annular / Total)"
MInfo2$(5) = "P -Partial"
MInfo2$(6) = "Lunar Eclipse"
MInfo2$(7) = "t -Total(Umbral)"
MInfo2$(8) = "p -Partial(Umbral)"
MInfo2$(9) = "n -Penumbral"


Dim wfh As Date
questart = R

bfree = FreeFile
Open App.Path + "\00 Data\Eclipse.txt" For Input As #bfree
bfree2 = FreeFile
Open App.Path + "\00 Data\Eclipse DONE.txt" For Output As #bfree2
Do
Line Input #bfree, as2$
If Trim$(as2$) = "" Then valyear = 0: valyear2 = 0
If valyear = 0 Then valyear = Val(Mid$(as2$, 1, 8))
If valyear >= 2000 Then valyear2 = valyear
If valyear2 >= Year(Now) - 3 Then
For ti = 0 To 3
If ti = 0 Then cg = 0
If ti = 1 Then cg = 18
If ti = 2 Then cg = 36
If ti = 3 Then cg = 53
as3$ = Mid$(as2$, 9 + cg, 15)
as4$ = Mid$(as2$, 23 + cg, 1)
If Trim$(as4$) <> "" Then
wfh = DateValue(Mid$(as3$, 1, 6) + "," + Str$(valyear2))
wfh = wfh + TimeValue(Mid$(as3$, 9, 5))
wfh2 = DateValue(Mid$(as3$, 1, 6) + "," + Str$(valyear2))
wfh2 = wfh2 + TimeValue(Mid$(as3$, 9, 5))
For ry7 = 2 To 9
If ry7 = 6 Then ry7 = 7
If Mid$(MInfo2$(ry7), 1, 1) = as4$ Then
egg = 0
If valyear2 < Year(Now) + 8 Then egg = 1
If as4$ = "T" Or as4 = "t" Then egg = 1
If egg = 1 Then
master$ = ""
If ry7 > 6 Then master$ = MInfo2$(6) Else master$ = MInfo2$(1)
master$ = master$ + Mid$(MInfo2$(ry7), 2)
master$ = Format$(wfh, " DD-MMM-yyYY HH:MMa/p ") + master$ ' + " GMT"
iy = 0
Print #bfree2, master$
'R = R + 3
'abc = R Mod 3: If abc <> 1 Then MsgBox "R out of sequence ": Stop
Mls$ = Format$(wfh, "DD-MM-yyyy")

'MInfo2$(6) = "Lunar Eclipse"
'MInfo2$(1) = "Solar Eclipse"
'
'dizzy$ = "# " + master$
'iy = iy + 1
'If InStr(dizzy$, "Lunar Eclipse") > 0 Then
'    If wfh > Now And Tagger(iy) = 0 Then
'        Mls$(R + 1) = "--- Next " + dizzy$
'        Tagger(iy) = 1
'        Mls$(XvRs1(iy)) = "--- Last " + dizzy2$
'    Else
'        XvRs1(iy) = R + 1: XvRs2(iy) = wfh: dizzy2$ = dizzy$
'        Mls$(R + 1) = "--- " + dizzy$
'    End If
'End If
'
'
'If InStr(dizzy$, "Solar Eclipse") > 0 Then
'    If wfh > Now And Tagge2(iy) = 0 Then
'        Mls$(R + 1) = "--- Next " + dizzy$
'        Tagge2(iy) = 1
'        Mls$(XvRs4(iy)) = "--- Last " + dizzy3$
'    Else
'        XvRs4(iy) = R + 1: XvRs5(iy) = wfh: dizzy3$ = dizzy$
'        Mls$(R + 1) = "--- " + dizzy$
'    End If
'End If
'

End If
End If
Next
End If
Next
End If
Loop Until EOF(bfree)
Close #bfree, #bfree2



End Sub
