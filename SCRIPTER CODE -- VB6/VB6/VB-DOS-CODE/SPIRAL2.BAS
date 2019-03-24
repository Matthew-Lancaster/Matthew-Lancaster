DEFDBL A-Z
'PRINT SQR((20 ^ 2) + (5 ^ 2))
'END
SCREEN 12
CLS
'LINE (200, 200)-(200, 200)
f = .1000001
p = .000001
t = .000001
c = 0
RANDOMIZE
tr = TIMER + 160
uh = .6545 * 10
uh = .95441
uh = .5198
uh = .372
uh = (1 + RND * 9999) / 10000
LOCATE CSRLIN, 1: PRINT "PRESS A KEY TO END 'R' TO RERUN IRON FILLINGS"; uh;
DO
'g = g + 1
'IF g MOD 11 = 0 THEN t = t + .000001
r = (r + uh) '+ ((tr - TIMER) / 5999999))'.2
p = p + .07
o = p / (f)
x = (o * ((640 / 480) * 1.3)) * SIN(r) + 320
y = o * COS(r) + 240
z = ABS(320 - x) ^ 2
z2 = SQR(z + (ABS(240 - y)) ^ 2)
f = (1 + ((z2 * 2.000003)))
'LOCATE 1, 1
'PRINT r
'c = c + 15: IF c = 30 THEN c = 0
c = c + 1: IF c = 16 THEN c = 0
'c = c + 5: IF c >= 16 THEN c = 1
'LINE STEP(0, 0)-(x, y), c
PSET (x, y), c
a$ = INKEY$
LOOP UNTIL tr < TIMER OR a$ <> ""
LOCATE CSRLIN, 1: PRINT "IRON FILLINGS"; uh;
SOUND 1000, 3
DO
a$ = INKEY$
LOOP UNTIL a$ <> ""
IF UCASE$(a$) = "R" THEN RUN


