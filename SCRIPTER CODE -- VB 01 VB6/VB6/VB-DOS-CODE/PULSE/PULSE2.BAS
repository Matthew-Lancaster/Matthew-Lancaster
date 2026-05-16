IF COMMAND$ = "1" THEN
da$ = MID$(DATE$, 4, 2) + "-" + MID$(DATE$, 1, 2) + "-" + MID$(DATE$, 7, 4)
ti$ = MID$(TIME$, 1, 5)
OPEN "c:\vb\pulse\pulse.txt" FOR APPEND AS #1
PRINT #1, da$; "-"; ti$; "-     -      -"
CLOSE #1
END
END IF

IF COMMAND$ = "4" THEN
da$ = MID$(DATE$, 4, 2) + "-" + MID$(DATE$, 1, 2) + "-" + MID$(DATE$, 7, 4)
ti$ = MID$(TIME$, 1, 5)
OPEN "c:\vb\pulse\pulsem.txt" FOR APPEND AS #1
PRINT #1, da$; "-"; ti$; "-     -      -"
CLOSE #1
END
END IF

IF COMMAND$ = "3" OR COMMAND$ = "2" OR COMMAND$ = "" THEN
IF COMMAND$ = "2" OR COMMAND$ = "" THEN fe$ = "c:\vb\pulse\pulse.txt": fe2$ = "c:\vb\pulse\Pulseavr.txt": fe3$ = "c:\vb\pulse\Pulseav2.txt"
IF COMMAND$ = "3" THEN fe$ = "c:\vb\pulse\pulsem.txt": fe2$ = "c:\vb\pulse\Pulsemvr.txt": fe3$ = "c:\vb\pulse\Pulsemv2.txt"
DIM op(3)
DIM qp(3)
SCREEN 12

CLS
start:
DEFDBL A-Z
LINE (0, 0)-(639, 479), 1, B
VIEW (1, 1)-(638, 478), 7
'LINE (2, 2)-(637, 477), 7, BF
'CLS
LOCATE 15, 2
'PALETTE 0, 7
PRINT "Y"
LOCATE 28, 40
PRINT "X Time."
LOCATE 2, 24
COLOR 2
PRINT "PULSE"
LOCATE 2, 38
COLOR 6
PRINT "SYS"
LOCATE 2, 50
COLOR 5
PRINT "DIA"
FOR r = 400 TO 50 STEP -50
LINE (75, r)-(80, r), 3
NEXT
FOR r = 400 TO 50 STEP -50
LINE (80, r)-(600, r), 4
NEXT
LINE (80, 400)-(600, 400), 3'x Along
LINE (80, 400)-(80, 50), 3  'y Up
a = 0
COLOR 14
FOR r = 400 TO 50 STEP -50
LOCATE (r / 17) + 1, 4: PRINT a
a = a + 25
NEXT
OPEN fe$ FOR INPUT AS #1
LINE INPUT #1, a$
B = 0
LINE INPUT #1, a$
a = DATESERIAL(VAL(MID$(a$, 7, 4)), VAL(MID$(a$, 4, 2)), VAL(MID$(a$, 1, 2)))
a = a + TIMESERIAL(VAL(MID$(a$, 12, 2)), VAL(MID$(a$, 15, 2)), 0)
B = B + 1
IF NOT EOF(1) THEN
DO
LINE INPUT #1, a$
B = B + 1
LOOP UNTIL EOF(1)
c = DATESERIAL(VAL(MID$(a$, 7, 4)), VAL(MID$(a$, 4, 2)), VAL(MID$(a$, 1, 2)))
c = c + TIMESERIAL(VAL(MID$(a$, 12, 2)), VAL(MID$(a$, 15, 2)), 0)
END IF
LOCATE 27, 10: PRINT FORMAT$(a, "DD/MM/YY")
LOCATE 27, 70: PRINT FORMAT$(c, "DD/MM/YY")
CLOSE #1

IF c = 0 THEN
pulse = VAL(MID$(a$, 36, 4))
dia = VAL(MID$(a$, 27, 4))
sys = VAL(MID$(a$, 18, 4))
LINE (80, 400)-(600, 400 - (pulse * 2)), 2
LINE (80, 400)-(600, 400 - (dia * 2)), 5
LINE (80, 400)-(600, 400 - (sys * 2)), 6
END IF
IF c > 0 THEN
qp(1) = 400
qp(2) = 400
qp(3) = 400
OPEN fe$ FOR INPUT AS #1
LINE INPUT #1, a$
l = (c - a) / 520
q = 80
po = 1
FOR r = 1 TO B
LINE INPUT #1, a$
e = DATESERIAL(VAL(MID$(a$, 7, 4)), VAL(MID$(a$, 4, 2)), VAL(MID$(a$, 1, 2)))
e = e + TIMESERIAL(VAL(MID$(a$, 12, 2)), VAL(MID$(a$, 15, 2)), 0)
t = 0
FOR o = 0 TO (c - a) STEP l
t = t + 1
IF (c - e) < o THEN EXIT FOR
NEXT
'PRINT t
pulse = VAL(MID$(a$, 36, 4))
dia = VAL(MID$(a$, 27, 4))
sys = VAL(MID$(a$, 18, 4))

p = 80 + (520 - t)

op(1) = 400 - (pulse * 2)
op(2) = 400 - (dia * 2)
op(3) = 400 - (sys * 2)
IF po = 0 THEN
LINE (q, qp(1))-(p, op(1)), 2
LINE (q, qp(2))-(p, op(2)), 5
LINE (q, qp(3))-(p, op(3)), 6
END IF
IF po = 1 THEN
po = 0
LINE (p, op(1))-(p, op(1)), 2
LINE (p, op(2))-(p, op(2)), 5
LINE (p, op(3))-(p, op(3)), 6
END IF
qp(1) = op(1)
qp(2) = op(2)
qp(3) = op(3)
q = p

NEXT
CLOSE #1
END IF

a$ = " "
'PAINT (3, 3), 7

DO
LOOP UNTIL INKEY$ <> ""




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
LINE (0, 0)-(639, 479), 1, B
VIEW (1, 1)-(638, 478), 7
'LINE (2, 2)-(637, 477), 7, BF
'CLS
LOCATE 15, 2
PRINT "Y"
LOCATE 28, 40
PRINT "X Time."
LOCATE 2, 24
COLOR 2
PRINT "PULSE"
LOCATE 2, 38
COLOR 6
PRINT "SYS"
LOCATE 2, 50
COLOR 5
PRINT "DIA"
FOR r = 400 TO 50 STEP -50
LINE (75, r)-(80, r), 3
NEXT
FOR r = 400 TO 50 STEP -50
LINE (80, r)-(600, r), 4
NEXT
LINE (80, 400)-(600, 400), 3'x Along
LINE (80, 400)-(80, 50), 3  'y Up
a = 0
COLOR 14
FOR r = 400 TO 50 STEP -50
LOCATE (r / 17) + 1, 4: PRINT a
a = a + 25
NEXT
OPEN fe$ FOR INPUT AS #1
LINE INPUT #1, a$
B = 0
LINE INPUT #1, a$
a = DATESERIAL(VAL(MID$(a$, 7, 4)), VAL(MID$(a$, 4, 2)), VAL(MID$(a$, 1, 2)))
a = a + TIMESERIAL(VAL(MID$(a$, 12, 2)), VAL(MID$(a$, 15, 2)), 0)
B = B + 1
IF NOT EOF(1) THEN
DO
LINE INPUT #1, a$
B = B + 1
LOOP UNTIL EOF(1)
c = DATESERIAL(VAL(MID$(a$, 7, 4)), VAL(MID$(a$, 4, 2)), VAL(MID$(a$, 1, 2)))
c = c + TIMESERIAL(VAL(MID$(a$, 12, 2)), VAL(MID$(a$, 15, 2)), 0)
END IF
CLOSE #1

LOCATE 27, 10: PRINT FORMAT$(a, "DD/MM/YY")
LOCATE 27, 70: PRINT FORMAT$(c, "DD/MM/YY")

LOCATE 29, 33: PRINT "....Average Reading...";

IF c = 0 THEN
pulse = VAL(MID$(a$, 36, 4))
dia = VAL(MID$(a$, 27, 4))
sys = VAL(MID$(a$, 18, 4))
LINE (80, 400)-(600, 400 - (pulse * 2)), 2
LINE (80, 400)-(600, 400 - (dia * 2)), 5
LINE (80, 400)-(600, 400 - (sys * 2)), 6
END IF
IF c > 0 THEN
'DIM op(3)
'DIM qp(3)
qp(1) = 400
qp(2) = 400
qp(3) = 400
OPEN fe2$ FOR OUTPUT AS #2
PRINT #2, "DATE------ TIME- SYS mmHg-- DIA mmHg- PULSE/min Average--- Activity---------"
OPEN fe$ FOR INPUT AS #1
LINE INPUT #1, a$
l = (c - a) / 520
q = 80
po = 1
FOR r = 1 TO B
LINE INPUT #1, a$
e = DATESERIAL(VAL(MID$(a$, 7, 4)), VAL(MID$(a$, 4, 2)), VAL(MID$(a$, 1, 2)))
e = e + TIMESERIAL(VAL(MID$(a$, 12, 2)), VAL(MID$(a$, 15, 2)), 0)
t = 0
FOR o = 0 TO (c - a) STEP l
t = t + 1
IF (c - e) < o THEN EXIT FOR
NEXT

pulse = VAL(MID$(a$, 36, 4))
dia = VAL(MID$(a$, 27, 4))
sys = VAL(MID$(a$, 18, 4))

aver2 = aver2 + 1
averp1 = averp1 + pulse
averp2 = averp1 / aver2
averd1 = averd1 + dia
averd2 = averd1 / aver2
avers1 = avers1 + sys
avers2 = avers1 / aver2
avera1 = avera1 + ((sys + dia + pulse) / 3)
avera2 = avera1 / aver2

PRINT #2, MID$(a$, 1, 17); FORMAT$(avers2, "#00.000000"); "-"; FORMAT$(averd2, "#00.000000"); "-"; FORMAT$(averp2, "#00.000000"); "-"; FORMAT$(avera2, "#00.000000"); "-"; MID$(a$, 50)

p = 80 + (520 - t)

op(1) = 400 - (averp2 * 2)
op(2) = 400 - (averd2 * 2)
op(3) = 400 - (avers2 * 2)

IF po = 0 THEN
LINE (q, qp(1))-(p, op(1)), 2
LINE (q, qp(2))-(p, op(2)), 5
LINE (q, qp(3))-(p, op(3)), 6
END IF
IF po = 1 THEN
po = 0
LINE (p, op(1))-(p, op(1)), 2
LINE (p, op(2))-(p, op(2)), 5
LINE (p, op(3))-(p, op(3)), 6
END IF
qp(1) = op(1)
qp(2) = op(2)
qp(3) = op(3)
q = p

NEXT
CLOSE #1
CLOSE #2
END IF

DO
a$ = INKEY$
LOOP UNTIL a$ <> ""
IF a$ = "1" THEN GOTO start
SCREEN 0
OPEN fe2$ FOR INPUT AS #1
DO
LINE INPUT #1, a$
PRINT a$
LOOP UNTIL EOF(1)
CLOSE #1

'PRINT "DATE------ TIME- SYS mmHg-- DIA mmHg- PULSE/min Average--- Activity---------"

OPEN fe2$ FOR INPUT AS #1
lines = 0
DO
LINE INPUT #1, a$
lines = lines + 1
LOOP UNTIL EOF(1)
CLOSE #1
'ytr = lines - 51''lines to produce 60 in pulseav2.txt
ytr = lines
OPEN fe2$ FOR INPUT AS #1
lines = 0
DO
LINE INPUT #1, a$
lines = lines + 1
LOOP UNTIL lines = ytr OR EOF(1)
IF EOF(1) THEN CLOSE #1: OPEN fe$ FOR INPUT AS #1
OPEN fe3$ FOR OUTPUT AS #2
PRINT #2, "DATE------ TIME- SYS mmHg-- DIA mmHg- PULSE/min Average--- Activity---------"

DO
LINE INPUT #1, a$
PRINT #2, a$
LOOP UNTIL EOF(1)
CLOSE #1, #2

SHELL "c:\utils\u " + fe2$ + " /#:" + LTRIM$(STR$(ytr - 21))
'SHELL "t " + fe3$

END IF

