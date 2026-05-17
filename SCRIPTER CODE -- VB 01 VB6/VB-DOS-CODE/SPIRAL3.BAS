DEFDBL A-Z
DEFCUR R, X-Y
'PRINT SQR((20 ^ 2) + (5 ^ 2))
'END
SCREEN 12
CLS
'LINE (200, 200)-(200, 200)
f = .1
p = .000001
t = .000001
c = 0
RANDOMIZE
tr = TIMER + 160
uh = (.6545 * .03)
uh = .95441 * 1
'uh = .5198
'uh = .372
'uh = (1 + RND * 9999) / 10000
uh = .5198 * .2339

LOCATE CSRLIN, 1: PRINT "IRON FILLINGS"; uh;
yg = 0

DO
uo = 1
ui = uo * 2
wsa = .0035
YT = YT + 1
IF (YT MOD ui) = 0 THEN tyg = tyg + wsa
IF (YT MOD ui) = uo THEN tyg = tyg + wsa
IF YT MOD ui = 0 THEN uh = (uh + ((SIN(tyg)) * 5))
IF YT MOD ui = uo THEN uh = (uh - ((SIN(tyg)) * 5))
' Arcsin# = ATN(X# / SQR(1 - X# * X#))   ' Calculates ArcSin

'g = g + 1
'IF g MOD 11 = 0 THEN t = t + .000001
r = r + uh'+ ((tr - TIMER) / 5999999))'.2
p = p + .32
o = p / (f)
X = (o * ((640 / 480) * 1.3)) * SIN(r) + 320
Y = o * COS(r) + 240
z = ABS(320 - X) ^ 2
z2 = SQR(z + (ABS(240 - Y)) ^ 2)
f = (1 + ((z2 * 2.000003)))
'LOCATE 1, 1
'PRINT r
'c = c + 15: IF c = 30 THEN c = 0
c = c + 1: IF c = 16 THEN c = 0
'c = c + 5: IF c >= 16 THEN c = 1
'LINE STEP(0, 0)-(x, y), c

PSET (X, Y), c
a$ = INKEY$
IF X > 0 AND Y > 0 AND X < 640 AND Y < 480 THEN tr = TIMER + 4
LOOP UNTIL tr < TIMER OR a$ <> ""
LOCATE CSRLIN, 1: PRINT "IRON FILLINGS"; uh;
IF tr < TIMER THEN SOUND 1000, 3
DO
a$ = INKEY$
LOOP UNTIL a$ <> ""
IF UCASE$(a$) = "R" THEN RUN


