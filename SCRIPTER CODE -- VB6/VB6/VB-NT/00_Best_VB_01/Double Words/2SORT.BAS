Attribute VB_Name = "Module1"
Public PeJing
Public Count2
Public Dlab

Const Dir_Doc$ = "D:\MY WEBS\"

Public XFreeW1 As Integer
Public XFreeW2 As Integer
Public XFreeW3 As Integer
Public XFreeW4 As Integer
Public XFreeW5 As Integer
Public XFreeW6 As Integer
Public XFreeW7 As Integer
Public XFreeW8 As Integer
Public XFreeW9 As Integer
Public Peel$
Public TQs$
Public TQs2$
Public Fake
Public SortFrmPop
Public SelectedEighty() As Integer


Sub Main_Sort_1st()


Count1 = 0
Count2 = 0

XFreeW6 = FreeFile
Open "c:\callerid\2sort22.txt" For Output As #XFreeW6


XFreeW1 = FreeFile
Open "c:\callerid\2double.txt" For Input As #XFreeW1
XFreeW2 = FreeFile
Open "c:\callerid\2sort-a.txt" For Output As #XFreeW2
Do
Line Input #XFreeW1, a$
If InStr(a$, "ANY OTHER TYPE OF COMBINATION WORDS") Then r = 1
If Mid$(a$, 1, 4) <> "----" And r = 1 Then c = 1
If Mid$(a$, 1, 4) = "----" And c = 1 Then r = 0
If r = 1 Then
Print #XFreeW2, a$ ': PRINT a$: SLEEP 1
Count1 = Count1 + 1
End If
Loop Until EOF(XFreeW1)
Count1 = Count1 - 1
Close #XFreeW1, #XFreeW2
r = 0
c = 0


XFreeW1 = FreeFile
Open "c:\callerid\2double.txt" For Input As #XFreeW1
XFreeW2 = FreeFile
Open "c:\callerid\2sort-b.txt" For Output As #XFreeW2
Do
Line Input #XFreeW1, a$

If InStr(a$, "DOUBLE WORD SAME FIRST LETTER") Then r = 1

If Mid$(a$, 1, 4) <> "----" And r = 1 Then c = 1
If Mid$(a$, 1, 4) = "----" And c = 1 Then r = 0


If r = 1 Then
Print #XFreeW2, a$
Count2 = Count2 + 1
End If
Loop Until EOF(XFreeW1)
Count2 = Count2 - 1
Close #XFreeW1, #XFreeW2

'Shell "sort <c:\callerid\2sort-b.txt >c:\callerid\2sort-d.txt", vbMinimizedNoFocus
'Shell "sort <c:\callerid\2sort-a.txt >>c:\callerid\2sort-d.txt", vbMinimizedNoFocus

'Kill "c:\callerid\2sort-a.txt"
'Kill "c:\callerid\2sort-b.txt"

'SHELL "t c:\callerid\2sort-d.txt"


'GoTo jump1
'-----------------------------------------------------------------


fc5 = FreeFile
Open "c:\callerid\2sort-b.txt" For Input As #fc5
Form1.Sort_List1.Clear
ajax = 0
Do
DoEvents
ajax = ajax + 1
Line Input #fc5, tdy$
Form1.Sort_List1.AddItem Trim(tdy$)
Loop Until EOF(fc5)
Close #fc5

Open "c:\callerid\2sort-d.txt" For Output As #fc5
'For r = 0 To Form1.Sort_List1.ListCount - 1

ajax2 = 0
Do
DoEvents
ajax2 = ajax2 + 1
Print #fc5, Form1.Sort_List1.List(0)
Form1.Sort_List1.RemoveItem (0)
Loop Until Form1.Sort_List1.ListCount = 0
'Next
Close #fc5

Kill "c:\callerid\2sort-b.txt"

'c:\callerid\2sort-d.txt

Dim xcompt1$
Dim xCompt2$


'----------------------------------------------------

GoTo jump1

XFreeW2 = FreeFile
Open "c:\callerid\2sort-d.txt" For Input As #XFreeW2


xcompt1$ = ""
IdUp = 0
SortFrmPop = 0
Do
    Line Input #XFreeW2, xCompt2$

    Click = 0
    
'    If InStr(RTrim$(xCompt2$), " ") And xcompt1$ <> "" Then
    If InStr(RTrim$(xCompt2$), xcompt1$ + " ") And xcompt1$ <> "" Then
        IdUp = 1
        SortFrm.List1.AddItem xcompt1$
        SortFrm.List1.AddItem xCompt2$
        SortFrm.List1.AddItem "#Thats a Pair Above"
        SortFrmPop = 1
        
        Click = 1
    End If
    If Click = 0 And RTrim$(xCompt2$) = RTrim$(xcompt1$) Then
        
        If xcompt1$ <> "" Then
            SortFrm.List1.AddItem xcompt1$
            SortFrmPop = 1
        End If
        IdUp = 1
    
    End If
    xcompt1$ = RTrim$(xCompt2$)
Loop Until EOF(XFreeW2)
Close #XFreeW2

jump1:

Call UpDating

End Sub


Public Sub UpDating()

PeJing = PeJing = 0
Dlab = 0

XFreeW1 = FreeFile
Open "c:\callerid\2double.txt" For Input As #XFreeW1
XFreeW2 = FreeFile
Open "c:\callerid\2sort-d.txt" For Input As #XFreeW2
XFreeW3 = FreeFile
Open "c:\callerid\2sort-h.txt" For Output As #XFreeW3

c = 0
r = 0
Do
Line Input #XFreeW1, a$
'X$ = UCASE$(X$)
'a$ = UCASE$(a$)
'PRINT a$
bg$ = "ANY OTHER TYPE OF COMBINATION WORDS"
bg$ = "DOUBLE WORD SAME FIRST LETTER"
If InStr(a$, bg$) Then
a$ = a$ + Str$(Count2)
r = 1: lm = 1
End If
If r = 0 Then Print #XFreeW3, a$
g = 0
x$ = ""
If Mid$(a$, 1, 4) = "----" And r = 1 And k = 0 Then
Print #XFreeW3, a$
If lm = 1 Then
Print #XFreeW3, "---------------------------------------------------------------------"
lm = 0
End If

If Not EOF(XFreeW2) Then
Line Input #XFreeW2, x$
'X$ = UCASE$(X$)
'a$ = UCASE$(a$)
End If

g = 1
bg$ = "DOUBLE WORD SAME FIRST LETTER"
bg$ = "ANY OTHER TYPE OF COMBINATION WORDS"
If InStr(x$, bg$) Then
k = 1: g = 0: x$ = ""
End If
End If

bg$ = "DOUBLE WORD SAME FIRST LETTER"
bg$ = "ANY OTHER TYPE OF COMBINATION WORDS"

If InStr(a$, bg$) Then
a$ = a$ + Str$(Count1)

Print #XFreeW3, a$
Print #XFreeW3, "----------------------------------------------------------------"
k = 0
Line Input #XFreeW1, a$
'X$ = UCASE$(X$)
'a$ = UCASE$(a$)
End If

'IF MID$(a$, 1, 4) = "----" AND c = 1 THEN r = 0
If r = 1 And g = 0 Then
'IF INSTR(x$, "DOUBLE WORD SAME FIRST LETTER") THEN k = 1
If k = 0 And g = 0 And Not EOF(XFreeW2) Then Line Input #XFreeW2, x$
'X$ = UCASE$(X$)
'a$ = UCASE$(a$)

If Len(RTrim$(a$)) > 39 Or g = 1 Then
Print #XFreeW3, a$
If RTrim$(x$) <> "" Then
Print #XFreeW3, Space$(40); "<font color=""#3399ff"">"; x$; "</font>"
End If
End If
If g = 0 And Len(RTrim$(a$)) < 40 Then
Print #XFreeW3, RTrim$(a$) + Space$(40 - Len(RTrim$(a$))); "<font color=""#3399ff"">"; x$; "</font>"
'PRINT RTRIM$(a$) + SPACE$(40 - LEN(RTRIM$(a$))); x$
End If
End If
Loop Until EOF(XFreeW1)
Close #XFreeW1, #XFreeW2, #XFreeW3
'SHELL "t c:\callerid\2double.txt"

'SHELL "t c:\callerid\2sort-d.txt"

'XFreeW1 = FreeFile
'bg$ = "DOUBLE WORD SAME FIRST LETTER"
'Open "C:\callerid\2sort-ls.html" For Input As #XFreeW1
'Do
'Line Input #XFreeW1, a$
'Loop Until InStr(a$, bg$) Or EOF(XFreeW1)
'Close #XFreeW1
'count5 = Val(Mid$(a$, 70))


Set m_CRC = New clsCRC

m_CRC.Algorithm = CRC32
    
'WxHex$ = Hex(m_CRC.CalculateString(Form3.Text1.Text))
WXHex1$ = Hex(m_CRC.CalculateFile("C:\callerid\2Sort-H.txt"))



'sr = FreeFile
'Open "C:\callerid\2sort-H.txt" For Input As #sr
''Seek sr, 5000
'wapt$ = Input$(LOF(sr), sr)
'Close sr

'WXHex1$ = Hex(m_CRC.CalculateString(wapt$))


'sr = FreeFile
'Open App.Path + "\test3.txt" For Output As #sr
'Print #sr, wapt$;
'Close #sr

If Dir$("C:\callerid\2Sort-HCRC.txt") = "" Then
    fj6 = FreeFile
    Open "C:\callerid\2Sort-HCRC.txt" For Output As #fj6
    Print #fj6, "----------------------"
    Close #fj6
Else
    fj6 = FreeFile
    Open "C:\callerid\2Sort-HCRC.txt" For Input As #fj6
    Line Input #fj6, wxhex2$
    Close fj6
End If

WXHex1$ = Hex(m_CRC.CalculateFile("C:\callerid\2Sort-H.txt"))

If WXHex1$ = wxhex2$ Then
    pip = 1
Else
    pip = 0
    gr5 = FreeFile
    Open "C:\callerid\2sort-HCRC.txt" For Output As #gr5
    Print #gr5, WXHex1$;
    Close gr5
End If

If pip = 1 Then
    SortFrm.List1.AddItem "#Not Updated.............."
End If

If Command$ <> "" Then
    SortFrm.List1.AddItem "#Force Updated.............."
End If

If pip = 0 Or Fake = True Then
updatedflag = 1
XFreeW1 = FreeFile
Open "c:\callerid\2sort-h.txt" For Input As #XFreeW1
SortFrm.List1.AddItem "#Updated.............."
XFreeW2 = FreeFile
Open "C:\callerid\2sort-ls.html" For Output As #XFreeW2
XFreeW4 = FreeFile
Open "C:\callerid\2sort-lz.html" For Output As #XFreeW4
XFreeW7 = FreeFile
Open "C:\callerid\2sort-lxx.txt" For Output As #XFreeW7
XFreeW8 = FreeFile
Open "C:\callerid\2sort-lxx2.txt" For Output As #XFreeW8
XFreeW9 = FreeFile
Open "C:\callerid\2sort-lxx3.txt" For Output As #XFreeW9

Do
Line Input #XFreeW1, a$
Loop Until InStr(a$, "-- UN SORTED COLUMN ---------------------------- SORTED COLUMN --")

XFreeW3 = FreeFile
Open "c:\callerid\2sort-h1.txt" For Input As #XFreeW3

Print #XFreeW2, "Last Updated:-"; Format$(Now, "dddd dd mmmm yyyy TTTTTam/pm"); "<br>"
Print #XFreeW4, "Last Updated:-"; Format$(Now, "dddd dd mmmm yyyy TTTTTam/pm"); "<br>"

Do
Line Input #XFreeW3, a$
If InStr(a$, "<META name=""Description") Then
 a$ = "<META name=""Description"" content=""Double Words With The Same First Letter - Matthewlan.com - Last updated " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + """>"
End If

Print #XFreeW2, a$
Print #XFreeW4, a$
'Print #XFreeW6, a$

Loop Until EOF(XFreeW3)
Close #XFreeW3

'Print #XFreeW2, "<pre>Last Updated:-"; Format$(Now, "dddd dd mmmm yyyy TTTTTam/pm")
'Print #XFreeW4, "<pre>Last Updated:-"; Format$(Now, "dddd dd mmmm yyyy TTTTTam/pm")


'Print #XFreeW7, "Last Updated:-"; Format$(Now, "dddd dd mmmm yyyy TTTTTam/pm")
'Print #XFreeW8, "Last Updated:-"; Format$(Now, "dddd dd mmmm yyyy TTTTTam/pm")
'Print #XFreeW9, "Last Updated:-"; Format$(Now, "dddd dd mmmm yyyy TTTTTam/pm")



TQs$ = "Last Updated:-" + Format$(Now, "dddd dd mmmm yyyy TTTTTam/pm") + vbCrLf

'-------------------------------------
'-------------------------------------
'-------------------------------------
'-------------------------------------
Call Skillion("Sports Shop", "14 Feb 2003")
Call Skillion("Speed Search", "08 Dec 2003")
Call Skillion("Hare Hunter", "20 May 2004")
Call Skillion("Frozen Foetuses", "17 Feb 2005")
Call Skillion("Clear Complexion", "17 Feb 2005")
Call Skillion("Super Scorpion Sub", "10 Aug 2005")
Call Skillion("Free Flow", "01 Sep 2005")
Call Skillion("Trivial Troll", "01 Oct 2005")
Call Skillion("Frog Frenzy", "19 Oct 2005")
Call Skillion("Verbal Volumes", "13 Jan 2006")
Call Skillion("Aching Arm", "01 Mar 2006")
Call Skillion("Fake Fiver", "02 Mar 2006")
Call Skillion("Skull Surgery", "13 Apr 2006")
Call Skillion("Yo Yo", "14 Apr 2006")
Call Skillion("Rare Rubies", "16 Aug 2006")

'-------------------------------
'-------------------------------
'-------------------------------
'-------------------------------

Do
Line Input #XFreeW1, a$
Print #XFreeW2, a$
Print #XFreeW4, a$

If InStr(a$, "<font color=") > 0 And Trim(Mid$(a$, 1, 40)) <> "" Then
    bsd$ = Mid$(a$, 1, 40) + Mid$(a$, 63, Len(Mid$(a$, 63)) - 7) '+ vbCrLf
    bsd2$ = Trim(Mid$(a$, 1, 40))
    bsd3$ = Trim(Mid$(a$, 63, Len(Mid$(a$, 63)) - 7))
End If
If InStr(a$, "<font color=") > 0 And Trim(Mid$(a$, 1, 40)) = "" Then
    bsd$ = Mid$(a$, 1, 40) + Mid$(a$, 63, Len(Mid$(a$, 63)) - 7) '+ vbCrLf
    bsd2$ = ""
    bsd3$ = Trim(Mid$(a$, 63, Len(Mid$(a$, 63)) - 7))
End If
If InStr(a$, "<font color=") < 1 And Trim(Mid$(a$, 1, 40)) <> "" Then
    bsd$ = Trim(a$)
    bsd2$ = Trim(a$)
    bsd3$ = ""
End If



Print #XFreeW7, bsd$
If bsd2$ <> "" Then Print #XFreeW8, bsd2$
If bsd3$ <> "" Then Print #XFreeW9, bsd3$



If RTrim$(Mid$(a$, 1, 2)) <> "--" Then
    If RTrim$(Mid$(a$, 1, 2)) <> "" Then
        PeJing = PeJing + 1
        tip = 0
    End If
End If


Loop Until EOF(XFreeW1) '''''''''''

XFreeW3 = FreeFile
Open "c:\callerid\2sort-h3.txt" For Input As #XFreeW3
Do
Line Input #XFreeW3, a$
If RTrim$(a$) <> "" Then Print #XFreeW2, a$
Loop Until EOF(XFreeW3)

baa$ = " #1=" + Str$(PeJing) + " #2=" + Str$(PeJing - Dlab)
Dlab = PeJing
Print #XFreeW2, "<font color=""#00FF00"">";
Print #XFreeW4, "<font color=""#00FF00"">";
Print #XFreeW2, String$(40, "-")
Print #XFreeW4, String$(40, "-")
Print #XFreeW7, String$(40, "-")
Print #XFreeW8, String$(40, "-")

Print #XFreeW2, "Totals For Previous"; baa$
Print #XFreeW4, "Totals For Previous"; baa$
Print #XFreeW7, "Totals For Previous"; baa$
Print #XFreeW8, "Totals For Previous"; baa$
Print #XFreeW2, String$(40, "-");
Print #XFreeW4, String$(40, "-");
Print #XFreeW7, String$(40, "-")
Print #XFreeW8, String$(40, "-")
Print #XFreeW2, "<a name=""#skip to bottom""></a>"
Print #XFreeW4, "<a name=""#skip to bottom""></a>"
Print #XFreeW2, "</font>"
Print #XFreeW4, "</font>"

XFreeW5 = FreeFile
Open "c:\callerid\2sort-h3.txt" For Input As #XFreeW5
Do
Line Input #XFreeW5, a$
If RTrim$(a$) <> "" Then Print #XFreeW4, a$
If RTrim$(a$) <> "" Then Print #XFreeW6, a$
Loop Until EOF(XFreeW5)

Close #XFreeW9, #XFreeW8, #XFreeW7, #XFreeW6, #XFreeW5, #XFreeW4, #XFreeW3, #XFreeW2, #XFreeW1

XFreeW7 = FreeFile
Open "C:\callerid\2sort-lxx.txt" For Input As #XFreeW7
XFreeW8 = FreeFile
Open "C:\callerid\2sort-lxx2.txt" For Input As #XFreeW8
XFreeW9 = FreeFile
Open "C:\callerid\2sort-lxx3.txt" For Input As #XFreeW9
bit3$ = Input(LOF(XFreeW7), XFreeW7)
bit4$ = Input(LOF(XFreeW8), XFreeW8)
bit5$ = Input(LOF(XFreeW9), XFreeW9)

bit3$ = TQs$ + bit3$
bit4$ = TQs$ + bit4$
bit5$ = TQs$ + bit5$
Close #XFreeW7, #XFreeW8, #XFreeW9
tty$ = "C:\callerid\"

XFreeW7 = FreeFile
Open tty$ + "2sort-lxx.txt" For Output As #XFreeW7
XFreeW8 = FreeFile
Open tty$ + "2sort-lxx2.txt" For Output As #XFreeW8
XFreeW9 = FreeFile
Open tty$ + "2sort-lxx3.txt" For Output As #XFreeW9

Print #XFreeW7, bit3$;
Print #XFreeW8, bit4$;
Print #XFreeW9, bit5$;

Close #XFreeW8, #XFreeW9, #XFreeW7


'MYDOCS
t$ = Special_Folder.GetSpecialFolderPath(&H5)


tty$ = t$ + "\01 Palm HotSync\"

XFreeW7 = FreeFile
Open tty$ + "2sort-lxx.txt" For Output As #XFreeW7
XFreeW8 = FreeFile
Open tty$ + "2sort-lxx2.txt" For Output As #XFreeW8
XFreeW9 = FreeFile
Open tty$ + "2sort-lxx3.txt" For Output As #XFreeW9

Print #XFreeW7, bit3$;
Print #XFreeW8, bit4$;
Print #XFreeW9, bit5$;

Close #XFreeW8, #XFreeW9, #XFreeW7


XFreeW7 = FreeFile
Open "c:\callerid\2sort-lx.html" For Output As #XFreeW7
XFreeW6 = FreeFile
Open "c:\callerid\2sort22.txt" For Input As #XFreeW6
XFreeW4 = FreeFile
Open "C:\callerid\2sort-lz.html" For Input As #XFreeW4
Do
    Line Input #XFreeW4, a$
    Print #XFreeW7, a$
Loop Until InStr(a$, "<!-- Skip -->")
Do
    Line Input #XFreeW6, a$
    Print #XFreeW7, a$
Loop Until EOF(XFreeW6)
Close XFreeW6
Print #XFreeW7, "</font><pre>";
Do
    Line Input #XFreeW4, a$
    Print #XFreeW7, a$
Loop Until EOF(XFreeW4)
Close #XFreeW7, #XFreeW4

Kill "C:\callerid\2sort-lz.html"
Name "C:\callerid\2sort-lx.html" As "C:\callerid\2sort-lz.html"


End If


'Print "Totals For Previous"; baa$:

'SHELL "t c:\callerid\2sort-h.txt"


XFreeW1 = FreeFile
Open "c:\callerid\2doubz1.txt" For Input As XFreeW1
whj1 = LOF(XFreeW1)
Close XFreeW1


XFreeW2 = FreeFile
Open "c:\callerid\2double.txt" For Input As XFreeW2
XFreeW3 = FreeFile
Open "c:\callerid\2doubz2.txt" For Output As XFreeW3



Do
    Line Input #XFreeW2, a$
Loop Until InStr(a$, "-----Triple Same First Letter-----")


Do
    Line Input #XFreeW2, a$
    Print #XFreeW3, a$
Loop Until InStr(a$, "-------------End")

Close XFreeW2, XFreeW3, XFreeW6

XFreeW1 = FreeFile
Open "c:\callerid\2doubz2.txt" For Input As XFreeW1
whj2 = LOF(XFreeW1)
Close XFreeW1

If whj2 <> whj1 Or Command$ <> "" Then
    SortFrm.List1.AddItem "#Tripple's Updated"

    XFreeW2 = FreeFile
    Open "c:\callerid\2doubz2.txt" For Input As #XFreeW2
    XFreeW3 = FreeFile
    Open "c:\callerid\2doubz1.txt" For Output As #XFreeW3
    Do
        Line Input #XFreeW2, a$
        Print #XFreeW3, a$
    Loop Until EOF(XFreeW2)
    Close XFreeW2, XFreeW3
    Call SortFrm_StartUp("hyp")
End If

If updatedflag = 1 Then Call SortFrm_StartUp("")



End Sub


Sub Skillion(In1$, In2$)

Do
    Line Input #XFreeW1, a$
    Print #XFreeW2, a$
    Print #XFreeW4, a$
    If InStr(a$, "<font color") > 0 Then
        bsd$ = Mid$(a$, 1, 40) + Mid$(a$, 63, Len(Mid$(a$, 63)) - 7)
        bsd2$ = Trim(Mid$(a$, 1, 40))
        bsd3$ = Trim(Mid$(a$, 63, Len(Mid$(a$, 63)) - 7))
    Else
        bsd$ = a$
        bsd2$ = Trim(a$)
        bsd3$ = Trim(a$)
    End If
    Print #XFreeW7, bsd$
    If bsd2$ <> "" Then Print #XFreeW8, bsd2$
    If bsd3$ <> "" Then Print #XFreeW9, bsd3$
    If RTrim$(Mid$(a$, 1, 2)) <> "--" Then
        If RTrim$(Mid$(a$, 1, 2)) <> "" Then
            PeJing = PeJing + 1
        End If
    End If
Loop Until InStr(a$, In1$) = 1

baa$ = " #1=" + Str$(PeJing) + " #2=" + Str$(PeJing - Dlab)
Dlab = PeJing
Print #XFreeW2, "<font color=""#30FF30"">";
Print #XFreeW4, "<font color=""#30FF30"">";

If InStr(a$, "Sports Shop") Then
    bty$ = "Left Hospital Here 4th Stay 14 Feb 2003 --"
    Call Million(bty$)
End If

If InStr(a$, "Frog Frenzy") Then
    bty$ = "Started Hospital Here 5th Stay 19 Oct 2005 --"
    Call Million(bty$)
End If

If InStr(a$, "Verbal Volumes") Then
    bty$ = "Left Hospital Here 5th Stay 11 Jan 2006 --"
    Call Million(bty$)
End If

If InStr(a$, "Frozen Foetuses") Then
    bty$ = "My Mate Roy Gave the following"
    Call Million(bty$)
End If

If InStr(a$, "Clear Complexion") Then
    bty$ = "Roy's Stopped Here"
    Call Million(bty$)
End If

If InStr(a$, "Fake Fiver") Then
    bty$ = "Decide To Add All The ones I had been Keeping Secret"
    Call Million(bty$)
End If

If InStr(a$, "Skull Surgery") Then
    bty$ = "Found a Web Site With a Few"
    Call Million(bty$)
End If

If InStr(a$, "Yo Yo") Then
    bty$ = "Continued With Mine"
    Call Million(bty$)
End If

If InStr(a$, "Rare Rubies") Then
    bty$ = "Mum's Birthday - Now Were Blasting"
    Call Million(bty$)
End If


Print #XFreeW2, String$(40, "-")
Print #XFreeW4, String$(40, "-")
Print #XFreeW7, String$(40, "-")
Print #XFreeW8, String$(40, "-")
Print #XFreeW2, "Totals For Previous"; baa$
Print #XFreeW4, "Totals For Previous"; baa$
Print #XFreeW7, "Totals For Previous"; baa$
Print #XFreeW8, "Totals For Previous"; baa$
Print #XFreeW2, String$(40, "-")
Print #XFreeW4, String$(40, "-")
Print #XFreeW7, String$(40, "-")
Print #XFreeW8, String$(40, "-")

Print #XFreeW2, String$(40, "-")
Print #XFreeW4, String$(40, "-")
Print #XFreeW7, String$(40, "-")
Print #XFreeW8, String$(40, "-")
a$ = "Latest additions -- Since 17 Feb 2005---"
Mid$(a$, 27, 11) = In2$

bo2$ = "<a name=""" + a$ + """></a><a href=""#top"">" + a$ + "</a>"


Print #XFreeW6, Peel$ + "<a href=""#" + a$ + """>" + "A" + Mid$(a$, 9) + " " + bty$ + "</a>"
TQs$ = TQs$ + a$ + " - " + bty$ + vbCrLf
TQs2$ = TQs2$ + a$ + " - " + bty$ + vbCrLf


If Peel$ = "" Then Peel$ = "<br>"


Print #XFreeW2, bo2$
Print #XFreeW4, bo2$
Print #XFreeW7, a$
Print #XFreeW8, a$
Print #XFreeW2, String$(40, "-");
Print #XFreeW4, String$(40, "-");
Print #XFreeW7, String$(40, "-")
Print #XFreeW8, String$(40, "-")
Print #XFreeW2, "</font>"
Print #XFreeW4, "</font>"


End Sub

Sub Million(bty$)

    Print #XFreeW2, String$(40, "-")
    Print #XFreeW4, String$(40, "-")
    Print #XFreeW7, String$(40, "-")
    Print #XFreeW8, String$(40, "-")
    Print #XFreeW2, bty$
    Print #XFreeW4, bty$
    Print #XFreeW7, bty$
    Print #XFreeW8, bty$
    Print #XFreeW2, String$(40, "-")
    Print #XFreeW4, String$(40, "-")
    Print #XFreeW7, String$(40, "-")
    Print #XFreeW8, String$(40, "-")


End Sub

Public Sub SortFrm_StartUp(ancient$)

'ancient$ = LCase$(Command$)

'ancient$ = "hyp"

If ancient$ = "hyp" Then Call egg22: End

'Wlfn1$ = Dir_Doc$ + "My Webs\zz Matt.Lan BT Web\0DOUBLE.html"
Wlfn2$ = Dir_Doc$ + "My Webs\MatthewLan.Com Web\LoveFolder\HTML\0DOUBLE.html"


df1 = FreeFile
Open "C:\CallerId\2Sort-LS.html" For Input As #df1
df2 = FreeFile
Open "C:\CallerId\2Sort-LZ.html" For Input As #df2
'df3 = FreeFile
'Open Wlfn1$ For Output As #df3
df4 = FreeFile
Open Wlfn2$ For Output As #df4
'Do
'Line Input #df1, a$
'Print #df3, a$
'Loop Until EOF(df1)
Do
Line Input #df2, a$
Print #df4, a$
Loop Until EOF(df2)

Close #df1, #df2, #df3, #df4




iky = FreeFile

Open "C:\CallerID\2SORT-LZ.html" For Input As #iky

Do
    Line Input #iky, a$
Loop Until InStr(a$, "DOUBLE WORD SAME FIRST LETTER")

B$ = Mid$(a$, 50)

Do
    Line Input #iky, a$
    If InStr(a$, "Latest additions") Then
        B$ = Mid$(a$, 36, 11)
    End If
    If InStr(a$, "Totals For Previous") Then
        c1$ = Mid$(a$, 25)
        c1$ = Mid$(c1$, 1, InStr(c1$, " "))
        c2$ = Mid$(a$, 33)
    End If
Loop Until EOF(iky)



Close #iky
'Totals For Previous
'Latest additions
'DOUBLE WORD SAME FIRST LETTER
iky = FreeFile

Open Dir_Doc$ + "01 Email Settings\Signatures\Signature-Stat.txt" For Input As #iky
Line Input #iky, lpo$
Close #iky

iky = FreeFile
Open Dir_Doc$ + "01 Email Settings\Signatures\Signature-Stat.txt" For Output As #iky
Print #iky, lpo$
Print #iky, "Double Word Same First Letter Count #" + c1$
Print #iky, "Additions Since " + B$ + " #" + Trim(c2$)
Close #iky

Prert = False

'wsfn1$ = GetShortName(WLFN1$)
'wsfn2$ = GetShortName(WLFN2$)

'freme = FreeFile
'Open "C:\CallerId\2Sort-LS.txt" For Output As #freme
'Print #freme, wsfn1$
'Print #freme, wsfn2$
'Print #freme, WLFN1$
'Print #freme, WLFN2$
'Close #freme

End Sub

Sub egg22()

Dim juh As Date


Wlfn5$ = Dir_Doc$ + "My Webs\MatthewLan.Com Web\LoveFolder\HTML\0past_txt.html"
Wlfn6$ = Dir_Doc$ + "My Webs\MatthewLan.Com Web\LoveFolder\HTML\0past_txt2.html"


Open "c:\callerid\2doubz2.txt" For Input As #3
Open Wlfn6$ For Output As #1
Open Wlfn5$ For Input As #2
Do
    Line Input #2, WFD$
    Print #1, WFD$
Loop Until InStr(WFD$, "TRIPLE same first letter")
Do
    Line Input #3, WFD$
    Print #1, WFD$
Loop Until EOF(3)
Do
    Line Input #2, WFD$
'Loop Until wfd$ = "DOUBLE---With a Single Or Double Word In Middle -----"
Loop Until InStr(WFD$, "---------------------End")
Do
    Line Input #2, WFD$
    Print #1, WFD$
Loop Until EOF(2)
Close #1, #2, #3

Kill Wlfn5$
Name Wlfn6$ As Wlfn5$

juh = Now + TimeSerial(0, 0, 1)
Do
DoEvents
Loop Until juh < Now


End Sub

