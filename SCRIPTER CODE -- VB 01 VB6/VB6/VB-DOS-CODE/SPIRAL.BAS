DEFDBL A-Z
SCREEN 12
LINE (200, 200)-(200, 200)
f = .1000001
p = .000001#
t = .00001
DO
'g = g + 1
'IF g MOD 1 = 0 THEN t = t + .000001
r = r + .002
p = p + .01
o = p / (f)
x = o * SIN(r) + 200
y = o * COS(r) + 200
z = ABS(200 - x) / 2
z2 = z + (ABS(200 - y)) / 2
f = (1 + ((z2 * 1.00003)))
'LOCATE 1, 1
'PRINT f
c = c + 1: IF c = 16 THEN c = 1
LINE STEP(0, 0)-(x, y), c
LOOP UNTIL r > 37832

