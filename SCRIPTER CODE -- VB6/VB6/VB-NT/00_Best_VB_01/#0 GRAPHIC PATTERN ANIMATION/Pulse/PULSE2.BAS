Attribute VB_Name = "Module1"
If Command$ = "1" Then
da$ = Mid$(Date$, 4, 2) + "-" + Mid$(Date$, 1, 2) + "-" + Mid$(Date$, 7, 4)
ti$ = Mid$(Time$, 1, 5)
Open "D:\VB6\vb-nt\pulse\pulse.txt" For Append As #1
Print #1, da$; "-"; ti$; "-     -      -"
Close #1
End
End If

If Command$ = "4" Then
da$ = Mid$(Date$, 4, 2) + "-" + Mid$(Date$, 1, 2) + "-" + Mid$(Date$, 7, 4)
ti$ = Mid$(Time$, 1, 5)
Open "D:\VB6\vb-nt\pulse\pulsem.txt" For Append As #1
Print #1, da$; "-"; ti$; "-     -      -"
Close #1
End
End If

If Command$ = "3" Or Command$ = "2" Or Command$ = "" Then
If Command$ = "2" Or Command$ = "" Then fe$ = "D:\VB6\vb-nt\pulse\pulse.txt": fe2$ = "D:\VB6\vb-nt\pulse\Pulseavr.txt": fe3$ = "D:\VB6\vb-nt\pulse\Pulseav2.txt"
If Command$ = "3" Then fe$ = "D:\VB6\vb-nt\pulse\pulsem.txt": fe2$ = "D:\VB6\vb-nt\pulse\Pulsemvr.txt": fe3$ = "D:\VB6\vb-nt\pulse\Pulsemv2.txt"
Dim op(3)
Dim qp(3)
Screen 12

Cls
start:
DefDbl A-Z
Line (0, 0)-(639, 479), 1, B
VIEW (1, 1)-(638, 478), 7
'LINE (2, 2)-(637, 477), 7, BF
'CLS
LOCATE 15, 2
'PALETTE 0, 7
Print "Y"
LOCATE 28, 40
Print "X Time."
LOCATE 2, 24
Color 2
Print "PULSE"
LOCATE 2, 38
Color 6
Print "SYS"
LOCATE 2, 50
Color 5
Print "DIA"
For r = 400 To 50 Step -50
Line (75, r)-(80, r), 3
Next
For r = 400 To 50 Step -50
Line (80, r)-(600, r), 4
Next
Line (80, 400)-(600, 400), 3 'x Along
Line (80, 400)-(80, 50), 3  'y Up
a = 0
Color 14
For r = 400 To 50 Step -50
LOCATE (r / 17) + 1, 4: Print a
a = a + 25
Next
Open fe$ For Input As #1
Line Input #1, a$
B = 0
Line Input #1, a$
a = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
a = a + TimeSerial(Val(Mid$(a$, 12, 2)), Val(Mid$(a$, 15, 2)), 0)
B = B + 1
If Not EOF(1) Then
Do
Line Input #1, a$
B = B + 1
Loop Until EOF(1)
c = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
c = c + TimeSerial(Val(Mid$(a$, 12, 2)), Val(Mid$(a$, 15, 2)), 0)
End If
LOCATE 27, 10: Print Format$(a, "DD/MM/YY")
LOCATE 27, 70: Print Format$(c, "DD/MM/YY")
Close #1

If c = 0 Then
Pulse = Val(Mid$(a$, 36, 4))
dia = Val(Mid$(a$, 27, 4))
sys = Val(Mid$(a$, 18, 4))
Line (80, 400)-(600, 400 - (Pulse * 2)), 2
Line (80, 400)-(600, 400 - (dia * 2)), 5
Line (80, 400)-(600, 400 - (sys * 2)), 6
End If
If c > 0 Then
qp(1) = 400
qp(2) = 400
qp(3) = 400
Open fe$ For Input As #1
Line Input #1, a$
l = (c - a) / 520
q = 80
po = 1
For r = 1 To B
Line Input #1, a$
e = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
e = e + TimeSerial(Val(Mid$(a$, 12, 2)), Val(Mid$(a$, 15, 2)), 0)
t = 0
For o = 0 To (c - a) Step l
t = t + 1
If (c - e) < o Then Exit For
Next
'PRINT t
Pulse = Val(Mid$(a$, 36, 4))
dia = Val(Mid$(a$, 27, 4))
sys = Val(Mid$(a$, 18, 4))

p = 80 + (520 - t)

op(1) = 400 - (Pulse * 2)
op(2) = 400 - (dia * 2)
op(3) = 400 - (sys * 2)
If po = 0 Then
Line (q, qp(1))-(p, op(1)), 2
Line (q, qp(2))-(p, op(2)), 5
Line (q, qp(3))-(p, op(3)), 6
End If
If po = 1 Then
po = 0
Line (p, op(1))-(p, op(1)), 2
Line (p, op(2))-(p, op(2)), 5
Line (p, op(3))-(p, op(3)), 6
End If
qp(1) = op(1)
qp(2) = op(2)
qp(3) = op(3)
q = p

Next
Close #1
End If

a$ = " "
'PAINT (3, 3), 7

Do
Loop Until INKEY$ <> ""




aver2 = 0
averp1 = 0
averp2 = 0
averd1 = 0
averd2 = 0
avers1 = 0
avers2 = 0
avera1 = 0
avera2 = 0



'SCREEN 12
Line (0, 0)-(639, 479), 1, B
VIEW (1, 1)-(638, 478), 7
'LINE (2, 2)-(637, 477), 7, BF
'CLS
LOCATE 15, 2
Print "Y"
LOCATE 28, 40
Print "X Time."
LOCATE 2, 24
Color 2
Print "PULSE"
LOCATE 2, 38
Color 6
Print "SYS"
LOCATE 2, 50
Color 5
Print "DIA"
For r = 400 To 50 Step -50
Line (75, r)-(80, r), 3
Next
For r = 400 To 50 Step -50
Line (80, r)-(600, r), 4
Next
Line (80, 400)-(600, 400), 3 'x Along
Line (80, 400)-(80, 50), 3  'y Up
a = 0
Color 14
For r = 400 To 50 Step -50
LOCATE (r / 17) + 1, 4: Print a
a = a + 25
Next
Open fe$ For Input As #1
Line Input #1, a$
B = 0
Line Input #1, a$
a = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
a = a + TimeSerial(Val(Mid$(a$, 12, 2)), Val(Mid$(a$, 15, 2)), 0)
B = B + 1
If Not EOF(1) Then
Do
Line Input #1, a$
B = B + 1
Loop Until EOF(1)
c = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
c = c + TimeSerial(Val(Mid$(a$, 12, 2)), Val(Mid$(a$, 15, 2)), 0)
End If
Close #1

LOCATE 27, 10: Print Format$(a, "DD/MM/YY")
LOCATE 27, 70: Print Format$(c, "DD/MM/YY")

LOCATE 29, 33: Print "....Average Reading...";

If c = 0 Then
Pulse = Val(Mid$(a$, 36, 4))
dia = Val(Mid$(a$, 27, 4))
sys = Val(Mid$(a$, 18, 4))
Line (80, 400)-(600, 400 - (Pulse * 2)), 2
Line (80, 400)-(600, 400 - (dia * 2)), 5
Line (80, 400)-(600, 400 - (sys * 2)), 6
End If
If c > 0 Then
'DIM op(3)
'DIM qp(3)
qp(1) = 400
qp(2) = 400
qp(3) = 400
Open fe2$ For Output As #2
Print #2, "DATE------ TIME- SYS mmHg-- DIA mmHg- PULSE/min Average--- Activity---------"
Open fe$ For Input As #1
Line Input #1, a$
l = (c - a) / 520
q = 80
po = 1
For r = 1 To B
Line Input #1, a$
e = DateSerial(Val(Mid$(a$, 7, 4)), Val(Mid$(a$, 4, 2)), Val(Mid$(a$, 1, 2)))
e = e + TimeSerial(Val(Mid$(a$, 12, 2)), Val(Mid$(a$, 15, 2)), 0)
t = 0
For o = 0 To (c - a) Step l
t = t + 1
If (c - e) < o Then Exit For
Next

Pulse = Val(Mid$(a$, 36, 4))
dia = Val(Mid$(a$, 27, 4))
sys = Val(Mid$(a$, 18, 4))

aver2 = aver2 + 1
averp1 = averp1 + Pulse
averp2 = averp1 / aver2
averd1 = averd1 + dia
averd2 = averd1 / aver2
avers1 = avers1 + sys
avers2 = avers1 / aver2
avera1 = avera1 + ((sys + dia + Pulse) / 3)
avera2 = avera1 / aver2

Print #2, Mid$(a$, 1, 17); Format$(avers2, "#00.000000"); "-"; Format$(averd2, "#00.000000"); "-"; Format$(averp2, "#00.000000"); "-"; Format$(avera2, "#00.000000"); "-"; Mid$(a$, 50)

p = 80 + (520 - t)

op(1) = 400 - (averp2 * 2)
op(2) = 400 - (averd2 * 2)
op(3) = 400 - (avers2 * 2)

If po = 0 Then
Line (q, qp(1))-(p, op(1)), 2
Line (q, qp(2))-(p, op(2)), 5
Line (q, qp(3))-(p, op(3)), 6
End If
If po = 1 Then
po = 0
Line (p, op(1))-(p, op(1)), 2
Line (p, op(2))-(p, op(2)), 5
Line (p, op(3))-(p, op(3)), 6
End If
qp(1) = op(1)
qp(2) = op(2)
qp(3) = op(3)
q = p

Next
Close #1
Close #2
End If

Do
a$ = INKEY$
Loop Until a$ <> ""
If a$ = "1" Then GoTo start
Screen 0
Open fe2$ For Input As #1
Do
Line Input #1, a$
Print a$
Loop Until EOF(1)
Close #1

'PRINT "DATE------ TIME- SYS mmHg-- DIA mmHg- PULSE/min Average--- Activity---------"

Open fe2$ For Input As #1
lines = 0
Do
Line Input #1, a$
lines = lines + 1
Loop Until EOF(1)
Close #1
'ytr = lines - 51''lines to produce 60 in pulseav2.txt
ytr = lines
Open fe2$ For Input As #1
lines = 0
Do
Line Input #1, a$
lines = lines + 1
Loop Until lines = ytr Or EOF(1)
If EOF(1) Then Close #1: Open fe$ For Input As #1
Open fe3$ For Output As #2
Print #2, "DATE------ TIME- SYS mmHg-- DIA mmHg- PULSE/min Average--- Activity---------"

Do
Line Input #1, a$
Print #2, a$
Loop Until EOF(1)
Close #1, #2

'Shell "c:\utils\u " + fe2$ + " /#:" + LTrim$(Str$(ytr - 21))


''SHELL "t " + fe3$

End If

