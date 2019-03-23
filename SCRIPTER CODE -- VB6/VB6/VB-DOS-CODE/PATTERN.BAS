SCREEN 12
DIM VERT%(2000)
DIM v AS DOUBLE
'KEY OFF
'CLS
'INPUT "enter radius (range 10-170)"; r
'INPUT "enter number of vertices (range 3-50)"; v
'CLS
r = 200
FOR v = 1 TO 500 STEP 5
IF v < 40 THEN v = v - 4.98
IF v > 230 THEN v = v + 5
IF v > 320 THEN v = v + 10
IF v > 400 THEN v = v + 20
X0 = 360 + r * 1.5: Y0 = 170
PSET (X0, Y0)
THETA = 6.28319 / v: N = v - 1
'THETA = 5 / v: N = v - 1
FOR I = 1 TO N
VERT%(2 * I - 1) = r * 1.5 * COS(THETA * I) + 360
VERT%(2 * I) = r * SIN(THETA * I) + 170
PSET (VERT%(2 * I - 1) - 40, 50 + VERT%(2 * I)), 1
NEXT I
N0 = N
FOR J = 1 TO N
ws = ws + .3
IF ws > 8 THEN ws = 0
COLOR ws

FOR I = 1 TO N0
XN = VERT%(2 * I - 1): YN = VERT%(2 * I)
LINE (X0 - 40, 50 + Y0)-(XN - 40, 50 + YN)
NEXT I
X0 = XN: Y0 = YN: N0 = N0 - 1
NEXT J
NEXT

'LOCATE 1, 1
'INPUT "Try another (Y/N)"; A$
'IF A$ = "Y" OR A$ = "y" GOTO 40
'CLS
'END

