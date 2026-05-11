Attribute VB_Name = "Module1"
Dim dfwd$(1800)
Dim Start As Integer
Dim kjh As Integer
Dim JKd As Integer
Dim qw As Integer
Dim lm As Integer
Dim xx As Integer
Dim plonk As Integer
Dim indy As Integer
Dim sort As Integer
Dim ads As Integer
Dim tr1$
Dim tr2$
Dim TR3$
Dim TR4$
Dim t1$
Dim t2$
Dim t3$
Dim t4$
'Dim hollow3$

Dim tp$
Dim KG As Long


Sub Main()

hollow3$ = "----------------------------"

sort = 0

JKd = 0

Start = 1: indy = 0 'indy = 11

For r = 1 To 26
    Hollow2$ = Chr$(r + 64) + Chr$(r + 64) + "--"
    HOllow1$ = "--------" + Hollow2$ + "--------"
    Call Tube
Next


End Sub

Sub mysub()

If Start = 1 Then indy = 11 Else indy = 0


If kjh = 1 Then
For oi = 1 To qw
    kj$ = dfwd$(oi)
    For r = 1 To qw
        If dfwd$(r) = kj$ And r <> oi Then poiu = 1: 'Form1.list1.AddItem kj$: poiu = 1
    Next
Next


'If poiu = 0 Then Print "No Bad Matches..........."
'If poiu = 1 Then Print "A Bad Match..........."

If poiu = 1 Then End
    kjh = 0
End If

For r = 1 To qw
    If dfwd$(r) = Hollow2$ Then lm = r
Next

If sort = 0 Then GoTo kli

If plonk = 1 Then
    hollow4$ = Mid$(Hollow2$, 1, 7)
    Open "C:\callerid\Double\4" + hollow4$ + ".txt" For Input As #1
Do
Line Input #1, w$
ws = ws + 1
Loop Until EOF(1)
Print Format$(ws, "000000"); "."; Format$(LOF(1), "0000000"); , Hollow2$; " --"; indy; "--"; HOllow1$
Print #5, Format$(ws, "000000"); "."; Format$(LOF(1), "0000000"); , Hollow2$; " --"; indy; "--"; HOllow1$
Close #1
End If

GoTo ende




kli:
If Start = 0 Then GoTo ende

xx = 1
KG = 0
tr1$ = Mid$(Hollow2$, 1, 1)
tr2$ = Mid$(Hollow2$, 2, 1)
TR3$ = Mid$(Hollow2$, 3, 1)
TR4$ = Mid$(Hollow2$, 4, 1)

GoTo skip1
DER$ = "cSR1"
DET$ = "cSR2"

Open DER$ + ".log" For Output As #2
'Print #2, "#**************----------------------------------************"
'Print #2, "#---CHECKED monster1.log    LARGE BBS LIST"
'Print #2, "#---CHECKED onefile.txt     DASH'S CDROM LIST"
'Print #2, "#---CHECKED pcbfiles.lst    US ROBOTICS BBS LIST"
'Print #2, "#---CHECKED list.inx        PUBLIC DOMAIN SW LIB BBS LIST"
'Print #2, "#---CHECKED text.lst        OLD DOS BASED CDROM LIST"
'Print #2, "#---CHECKED tube.txt        14.5 MEGS OF EMAILS"
'Print #2, "#---CHECKED netscap2.txt    7.5  MEGS OF internet EMAILS"
'Print #2, "#---CHECKED textfile.txt    24.6 MEGS OF all MY COMPUTER DOCUMENTS"
'Print #2, "#---CHECKED bashtml.zip     9.9  MEGS OF all MY HTML & BASIC files"
'Print #2, "#---CHECKED acrobdoc.txt    2.9  MEGS OF acrobat manual files"
'Print #2, "#@**************---------------------------------************"
hollow4$ = Mid$(Hollow2$, 1, 7)
GoTo killoff
If Dir$("4" + hollow4$ + ".txt") <> "" Then
Open "4" + hollow4$ + ".txt" For Input As #7
For r = 1 To 12
Line Input #7, sdg$
Next
Do
If Not EOF(7) Then
Line Input #7, sdg$
'PRINT #2, sdg$
End If
Loop Until EOF(7)
End If
Close #7

killoff:
'GOTO her
'GOTO here2
'a$ = DIR$("E:\tube.txt")
'IF a$ <> "" THEN bd$ = "E:\" ELSE bd$ = "J:\"
bd$ = ".\"
tp$ = " 1-10."
ty$ = bd$ + "monster1.log"
Call Tube
tp$ = " 2-10."
ty$ = bd$ + "onefile.txt"
Call Tube
tp$ = " 3-10."
ty$ = bd$ + "pcbfiles.lst"
Call Tube
tp$ = " 4-10."
ty$ = bd$ + "list.inx"
Call Tube
tp$ = " 5-10."
ty$ = bd$ + "text.lst"
Call Tube
tp$ = " 6-10."
ty$ = bd$ + "tube.txt"
Call Tube
tp$ = " 7-10."
ty$ = bd$ + "netscape.txt"
Call Tube
tp$ = " 8-10."
ty$ = bd$ + "textfile.txt"
Call Tube
tp$ = " 9-10."
ty$ = bd$ + "bashtml.zip"
Call Tube
tp$ = "10-10."
ty$ = bd$ + "acrobdoc.txt"
Call Tube

here2:

'tp$ = " 11-11."
'ty$ = "4" + hollow4$ + ".txt"
'IF DIR$(ty$) <> "" THEN CALL tube


Close #2

Color 15
Print "Sorting........."
Shell "sort " + DER$ + ".log >" + DER$ + ".txt"
Color 14
Print "Filtering......."

Open DER$ + ".txt" For Input As #1
Open DET$ + ".txt" For Output As #2
b$ = ""
Print #2, "#@***-------------------"; hollow3$; "------***"
Print #2, "#@***-------**-"; HOllow1$; "-**------------***"
Print #2, "#@***-------------------"; hollow3$; "------***"
Do
Line Input #1, A$
If UCase$(A$) <> UCase$(b$) Then Print #2, RTrim$(A$)
b$ = A$
Loop Until EOF(1)
Close #1, #2
Kill DER$ + ".txt"
Kill DER$ + ".log"
hollow4$ = Mid$(Hollow2$, 1, 7)
A$ = Dir$("4" + hollow4$ + ".txt")
If A$ <> "" Then Kill "4" + hollow4$ + ".txt"
Name DET$ + ".txt" As "4" + hollow4$ + ".txt"

'SHELL "t"
'SHELL "t 4" + holLow4$ + ".txt"


ende:
End Sub

DefInt A-Z
Sub ploink()

Shell "sort /R <#taggot4.txt >#taggot5.txt"

Shell "t #taggot5.txt"

Open "#taggot5.txt" For Input As #1
Open "#taggot6.txt" For Output As #2
Do
Line Input #1, A$
A = A + 1
b$ = Mid$(A$, 29, 5)

If Val(Mid$(A$, 1, 6)) <= 15 And xds = 0 Then
xds = 1
Print #2,
Print
Print #2, "..........................No Matches Found for these......................"
Print "..........................No Matches Found for these......................"
b = 0
End If

Print #2, b$;
Print b$;

b = b + 1
If b > 14 Then Print #2,: b = 0
Loop Until EOF(1)
Print
Print A
Print #2,
Print #2, A
Close #1, #2
Shell "t #taggot6.txt"


'END



Open "#taggot6.txt" For Input As #1
Do
Line Input #1, xaz$
Loop Until EOF(1)
Close #1


Open "c:\callerid\2double.txt" For Input As #1
Open "c:\callerid\2doubl1.txt" For Output As #2

Do
wx = wx + 1
Line Input #1, A$
If InStr(A$, "MOST COMMON 2 to") Then Exit Do
'PRINT A$
Print #2, A$
Loop Until EOF(1)
'PRINT A$
Mid$(A$, 56, Len(xaz$)) = xaz$
'PRINT A$
Print #2, A$
Line Input #1, A$
Print #2, A$

Open "#taggot6.txt" For Input As #3
Do
Line Input #3, xaz$
If Val(xaz$) < 100 And xaz$ <> "" Then Print #2, xaz$
Loop Until EOF(3)


Do
Line Input #1, A$
Loop Until InStr(A$, "-------------------")

Print #2, A$

Do
Line Input #1, A$
Print #2, A$
Loop Until EOF(1)

Close #1, #2, #3
If Dir$("c:\callerid\2double.bak") <> "" Then Kill "c:\callerid\2double.bak"
Name "c:\callerid\2double.txt" As "c:\callerid\2double.bak"
Name "c:\callerid\2doubl1.txt" As "c:\callerid\2double.txt"

Shell "u c:\callerid\2double.txt /#" + LTrim$(Str$(wx))

End Sub

DefDbl L


