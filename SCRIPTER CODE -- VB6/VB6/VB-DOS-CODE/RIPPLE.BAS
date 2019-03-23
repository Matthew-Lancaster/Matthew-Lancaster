DECLARE SUB plusminushigh ()
DECLARE SUB plusminuslow ()
DEFDBL A-Z
DECLARE SUB palette2 ()
DECLARE SUB lowres ()
DECLARE SUB highres ()
COMMON SHARED awat AS DOUBLE
COLOR 14
PRINT "Keep Pressing key in the running of the program to Exit "
PRINT "Press 'l' for low res show more colours"
PRINT "Press 'h' for high res show"
PRINT "-------New --- just press a key for instant exit from these ones"
PRINT "Press '1' for plus minus low res show more colours"
PRINT "Press '2' for plus minus high res show"
PRINT "Default option 2 for plus minus high res show will be picked in 5 SEcs"
'CALL lowres
'END
w = NOW + TIMESERIAL(0, 0, 5)
DO
a$ = UCASE$(INKEY$)
LOOP UNTIL a$ = "1" OR a$ = "2" OR a$ = "L" OR a$ = "H" OR w < NOW
IF a$ = "" THEN a$ = "2"
IF a$ = "L" THEN CALL lowres
IF a$ = "H" THEN CALL highres
IF a$ = "1" THEN CALL plusminuslow
IF a$ = "2" THEN CALL plusminushigh

DEFSNG A-Z
SUB highres ()
DEFDBL A-Z
SCREEN 12
pi = 4 * ATN(1)
lk = 150
cv = .0000001
cv = .00002
'FOR h = 0 TO lk STEP cv
FOR h = 1 TO lk STEP cv
IF INKEY$ <> "" THEN h = h + 1
'IF h > 1 THEN h = h + .0001: GOTO jump
'IF h <= 1 THEN h = h + .0002: GOTO jump
jump:
'h = h
LOCATE 1, 1: PRINT FORMAT$(h, "0000.000000000000000"); " Of"; lk; " Keep Pressing key to Exit "
q = h
tyu = tyu + .003
r = tyu
'r = 0
w = 1
c = cb / 20'c / 10000
c2 = cb'c2 / 10000

cb = cb + .01

cb2 = (100000 + (SIN((cb)))) / 2



c = c + cb2
c2 = c2 + cb2

DO
'c2 = c2 + (cv)
'gh = 300 * (SIN(c2)) + 100
c = c + .001'(cv * 1000) + ((COS(c2)) / gh)
'IF c > 15 THEN c = 0
c1 = 14 * (ABS(COS(c)))
r = r + q
w = ((r / (pi)) / ((h)))  '(h * 2))'h * 1.5))
x = (w) * SIN(r * 5) + 320
y = (w / 1.333333) * COS(r * 5) + 240





PSET (x, y), c1
'CIRCLE (x, y), 2, c1
'LINE STEP(0, 0)-(x, y), c1
s1 = 0
'wl = 100
'IF x < wl OR y < wl THEN s1 = 1
'IF x > 640 - wl OR y > 480 - wl THEN s1 = 1
'IF x > 590 OR y > 430 THEN s1 = 1

IF w > 200 THEN s = s + 1 ELSE s = 0
LOOP UNTIL s > 100
s = 0
NEXT
END SUB

DEFSNG A-Z
SUB lowres ()
DEFDBL A-Z
SCREEN 13

CALL palette2

tyu = 1
pi = 4 * ATN(1)
lk = 150
cv = .00002
FOR h = 1 TO lk STEP cv
h = h + cv
IF awat < NOW THEN CALL palette2
IF INKEY$ <> "" THEN h = h + 1
'IF h > 1 THEN h = h + .0001: GOTO jump2
'IF h <= 1 THEN h = h + .0002: GOTO jump2
jump2:
'h = h
COLOR 1
LOCATE 1, 1: PRINT FORMAT$(h, "0000.000000000000000"); " Of"; lk
q = h * 3





tyu = tyu + .003
r = tyu
w = 1
c = (cb * 5)'c / 10000
c2 = cb'c2 / 10000

cb = cb + .01

cb2 = (10000 + (SIN((cb) + SIN(cb - pi / 2)))) / 2

c = c + cb2
c2 = c2 + cb2

DO
'c2 = c2 + cv
'gh = (2000 * (SIN(c2 * 500)) + 4000)
c = c - .0001'(cv) + (((COS(c2) / 1000)))
c1 = 255 * (ABS(COS(c)))
r = r + q
w = (((r - tyu) / (pi))) / (tyu * 90)' * ((1 + (h * 5))))'(h * 2))'h * 1.5))
x = (w) * SIN(r) + 160
y = (w / 1.6) * COS(r) + 100





PSET (x, y), c1
'CIRCLE (x, y), 2, c1
'LINE STEP(0, 0)-(x, y), c1
s1 = 0
'wl = 50
'IF x < wl OR y < wl THEN s1 = 1
'IF x > 320 - wl OR y > 200 - wl THEN s1 = 1
IF w > 150 THEN s1 = 1
IF s1 = 1 THEN s = s + 1 ELSE s = 0
LOOP UNTIL s > 100
s = 0
NEXT
END SUB

SUB palette2 ()
'4145400
lk = 4144959'976
backabit:
'pk = INT(RND(1) * 20000)
'pk = INT(RND(1) * 64) * INT(RND(1) * 200)
pk = INT((RND(1) * 64) * RND(1) * 2000)
'pk = 62 * 100
FOR r = 1 TO 252 STEP 4
ON LOCAL ERROR GOTO ers
'po = po + 1
pk = pk + 1
'IF po MOD 3 = 0 THEN SLEEP
'pk = INT((lk / r)) + x
'pk = lk * ABS(SIN(r / 162.8)) + x
PALETTE r, pk
PALETTE r + 1, pk
PALETTE r + 2, pk
PALETTE r + 3, pk

ON LOCAL ERROR GOTO 0
IF w = 1 AND r = 1 THEN
w = 0
pk = INT((RND(1) * 64) * RND(1) * 2000)
END IF
x = 0
'COLOR r
'PRINT r, pk
NEXT
'SLEEP
IF 1 = 2 THEN
ers:
w = 1
RESUME NEXT
END IF

awat = NOW + TIMESERIAL(0, 0, 10)

END SUB

SUB plusminushigh ()
DEFDBL A-Z
SCREEN 12
ab = 1
CALL palette2

tyu = 1
pi = 4 * ATN(1)
lk = 150
cv = .000002
FOR h = 150 TO lk STEP cv
'h = h + cv
IF awat < NOW THEN CALL palette2
IF INKEY$ <> "" THEN h = h + 1
COLOR 1
LOCATE 1, 1: PRINT FORMAT$(h, "0000.000000000000000"); " Of"; lk
q = h





IF ab = -1 THEN tyu = tyu - .0003
IF ab = 1 THEN tyu = tyu + .0003
r = tyu
w = 1
c = (cb * 5)'c / 10000
c2 = cb'c2 / 10000

cb = cb + .01

cb2 = (1000 + (SIN((cb) + SIN(cb - pi / 2)))) / 2

IF ab = -1 THEN c = c - cb2
IF ab = 1 THEN c = c + cb2
IF ab = -1 THEN c2 = c2 - cb2
IF ab = 1 THEN c2 = c2 + cb2

DO
q1 = q1 + .00000001#
q = 100 + (q1 * 20)' + SIN(q1)
'q = q / 2
'c2 = c2 + cv
'gh = (2000 * (SIN(c2 * 500)) + 4000)
c = c + .00001'(cv) + (((COS(c2) / 1000)))
c1 = 255 * (ABS(COS(c)))
c1 = c
IF c > 15 THEN c = 0
IF ab = -1 THEN r = r - q
IF ab = 1 THEN r = r + q
w = (((r - tyu) / (pi))) / (tyu * 90)' * ((1 + (h * 5))))'(h * 2))'h * 1.5))
x = (w) * SIN(r) + 320
y = (w / 1.333333) * COS(r) + 240





PSET (x, y), c1
'CIRCLE (x, y), 2, c1
'LINE STEP(0, 0)-(x, y), c1
s1 = 0
'wl = 50
'IF x < wl OR y < wl THEN s1 = 1
'IF x > 320 - wl OR y > 200 - wl THEN s1 = 1
IF w > 280 OR w < -280 THEN s1 = 1
IF s1 = 1 THEN s = s + 1 ELSE s = 0
IF s1 = 1 THEN s = 0: IF ab = 1 THEN ab = -1 ELSE ab = 1: IF INKEY$ <> "" THEN END
LOOP UNTIL 1 = 2
'LOOP UNTIL s>100
s = 0
NEXT
END SUB

SUB plusminuslow ()
DEFDBL A-Z
SCREEN 13
ab = 1
CALL palette2

tyu = 1
pi = 4 * ATN(1)
lk = 150
cv = .000002
FOR h = 150 TO lk STEP cv
'h = h + cv
IF awat < NOW THEN CALL palette2
IF INKEY$ <> "" THEN h = h + 1
'IF h > 1 THEN h = h + .0001: GOTO jump3
'IF h <= 1 THEN h = h + .0002: GOTO jump3
jump3:
'h = h
COLOR 1
LOCATE 1, 1: PRINT FORMAT$(h, "0000.000000000000000"); " Of"; lk
q = h





IF ab = -1 THEN tyu = tyu - .003
IF ab = 1 THEN tyu = tyu + .003
r = tyu
w = 1
c = (cb * 5)'c / 10000
c2 = cb'c2 / 10000

cb = cb + .01

cb2 = (10000 + (SIN((cb) + SIN(cb - pi / 2)))) / 2

IF ab = -1 THEN c = c - cb2
IF ab = 1 THEN c = c + cb2
IF ab = -1 THEN c2 = c2 - cb2
IF ab = 1 THEN c2 = c2 + cb2

DO
q1 = q1 + .00000001#
q = 100 + (q1 * 25) + SIN(q1) * 2
'q = q / 2
'c2 = c2 + cv
'gh = (2000 * (SIN(c2 * 500)) + 4000)
c = c + .001'(cv) + (((COS(c2) / 1000)))
c1 = 255 * (ABS(COS(c)))
c1 = c
IF c > 255 THEN c = 0
IF ab = -1 THEN r = r - q
IF ab = 1 THEN r = r + q
w = (((r - tyu) / (pi))) / (tyu * 90)' * ((1 + (h * 5))))'(h * 2))'h * 1.5))
x = (w) * SIN(r) + 160
y = (w / 1.6) * COS(r) + 100





PSET (x, y), c1
'CIRCLE (x, y), 2, c1
'LINE STEP(0, 0)-(x, y), c1
s1 = 0
'wl = 50
'IF x < wl OR y < wl THEN s1 = 1
'IF x > 320 - wl OR y > 200 - wl THEN s1 = 1
IF w > 150 OR w < -150 THEN s1 = 1
IF s1 = 1 THEN s = s + 1 ELSE s = 0
IF s1 = 1 THEN s = 0: IF ab = 1 THEN ab = -1 ELSE ab = 1: IF INKEY$ <> "" THEN END
LOOP UNTIL 1 = 2
'LOOP UNTIL s>100
s = 0
NEXT
END SUB

