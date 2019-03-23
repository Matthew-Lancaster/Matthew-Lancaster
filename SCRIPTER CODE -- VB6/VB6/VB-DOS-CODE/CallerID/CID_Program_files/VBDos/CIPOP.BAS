DECLARE SUB average_it_all_out ()
DECLARE SUB oldest_history ()
DECLARE SUB busy ()
DECLARE SUB rings ()
DECLARE SUB calls_in ()
DIM SHARED asax(7) AS INTEGER
DIM SHARED bsax(7) AS INTEGER
DIM SHARED atax(12) AS INTEGER
DIM SHARED btax(12) AS INTEGER
DIM SHARED ztax(12) AS INTEGER
DIM SHARED we(14) AS DOUBLE
DIM SHARED cc(7) AS INTEGER
DIM SHARED cctax(12) AS INTEGER
DIM SHARED cmc AS DOUBLE
DIM SHARED cdc AS DOUBLE
DIM SHARED ide AS CURRENCY
DIM SHARED idx AS CURRENCY
DIM SHARED ide1 AS CURRENCY
DIM SHARED idc1 AS CURRENCY
DIM SHARED idx1 AS CURRENCY
DIM SHARED mo1 AS INTEGER
DIM SHARED tu1 AS INTEGER
DIM SHARED we1 AS INTEGER
DIM SHARED th1 AS INTEGER
DIM SHARED fr1 AS INTEGER
DIM SHARED sa1 AS INTEGER
DIM SHARED su1 AS INTEGER
DIM SHARED mo AS DOUBLE
DIM SHARED tu AS DOUBLE
DIM SHARED we AS DOUBLE
DIM SHARED th AS DOUBLE
DIM SHARED fr AS DOUBLE
DIM SHARED sa AS DOUBLE
DIM SHARED su AS DOUBLE
DIM asd2 AS DOUBLE

IF DIR$("C:\callerid\CILOGS~1\cipopti.txt") <> "" THEN
OPEN "C:\callerid\CILOGS~1\cipopti.txt" FOR INPUT AS #1
LINE INPUT #1, asd$
CLOSE #1
asd2 = VAL(asd$)
END IF

IF asd2 < NOW THEN
OPEN "C:\callerid\CILOGS~1\cipopti.txt" FOR OUTPUT AS #1
asd2 = NOW + 14
PRINT #1, STR$(asd2)
CLOSE #1
asd3 = 1
END IF



wq = 1994
FOR r = 1 TO 14
we(r) = DATESERIAL(wq + r, 1, 1)
NEXT


COLOR 12
PRINT "--------------------------------------------------------------------"
COLOR 12
PRINT "--------------------------------------------------------------------"
COLOR 9
PRINT "In order of most Common First the most Common Days of the week"

CALL rings
CALL calls_in
CALL busy
CALL oldest_history
CALL average_it_all_out
COLOR 12
PRINT "--------------------------------------------------------------------"
IF asd3 = 1 THEN SLEEP 20

KILL "c:\callerid\CILOGS~1\ci-aver.tmp"

SUB average_it_all_out ()
DEFDBL A-Z
DIM kl$(6)
OPEN "c:\callerid\CILOGS~1\ci-aver.txt" FOR INPUT AS #2
DO
kl$(1) = kl$(2)
kl$(2) = kl$(3)
kl$(3) = kl$(4)
kl$(4) = kl$(5)
kl$(5) = kl$(6)
LINE INPUT #2, kl$(6)
LOOP UNTIL EOF(2)
CLOSE #2

OPEN "c:\callerid\CILOGS~1\ci-aver.tmp" FOR OUTPUT AS #3
s = 0
FOR r = 1 TO 12
btax(r) = 0
NEXT
FOR r = 1 TO 7
bsax(r) = 0
asax(r) = cc(r)
NEXT
COLOR 12
PRINT "--------------------------------------------------------------------"
COLOR 15
PRINT "And the Averag of all the above results is :-:-"
'PRINT asax(1); asax(2); asax(3); asax(4); asax(5); asax(6); asax(7)
FOR i = 1 TO 7
w = 0
l = 1
FOR r = 1 TO 7
IF asax(r) > w AND bsax(r) = 0 THEN l = r: w = asax(l)
NEXT
bsax(l) = 1

COLOR 11
IF l = 1 THEN tr = mo: tr1 = mo1'g$ = "Mon"
IF l = 2 THEN tr = tu: tr1 = tu1'g$ = "Tue"
IF l = 3 THEN tr = we: tr1 = we1'g$ = "Wed"
IF l = 4 THEN tr = th: tr1 = th1'g$ = "Thu"
IF l = 5 THEN tr = fr: tr1 = fr1'g$ = "Fri"
IF l = 6 THEN tr = sa: tr1 = sa1'g$ = "Sat"
IF l = 7 THEN tr = su: tr1 = su1'g$ = "Sun"
'PRINT #3, g$; "-";
'PRINT g$; "-";
IF asax(l) <> h THEN s = s + 1
h = asax(l)
IF s MOD 2 = 0 THEN COLOR 14
IF s MOD 2 = 1 THEN COLOR 10
PRINT FORMAT$(tr / tr1, "HH:MM:SS"); " ";
PRINT #3, FORMAT$(tr / tr1, "HH:MM:SS"); " ";
'PRINT FORMAT$(asax(l), "0000"); " ";
'PRINT #3, FORMAT$(asax(l), "0000"); " ";
NEXT
PRINT
PRINT #3,
FOR r = 1 TO 7
bsax(r) = 0
NEXT
      

FOR i = 1 TO 7
w = 0
l = 1
FOR r = 1 TO 7
IF asax(r) > w AND bsax(r) = 0 THEN l = r: w = asax(l)
NEXT
bsax(l) = 1

COLOR 11
IF l = 1 THEN g$ = "Mon"
IF l = 2 THEN g$ = "Tue"
IF l = 3 THEN g$ = "Wed"
IF l = 4 THEN g$ = "Thu"
IF l = 5 THEN g$ = "Fri"
IF l = 6 THEN g$ = "Sat"
IF l = 7 THEN g$ = "Sun"
PRINT #3, g$; "-";
PRINT g$; "-";
IF asax(l) <> h THEN s = s + 1
h = asax(l)
IF s MOD 2 = 0 THEN COLOR 14
IF s MOD 2 = 1 THEN COLOR 10
PRINT FORMAT$(asax(l), "0000"); " ";
PRINT #3, FORMAT$(asax(l), "0000"); " ";
NEXT

PRINT #3, FORMAT$(cmc / 4, "HH:MM:SS")
PRINT FORMAT$(cmc / 4, "HH:MM:SS")

'-------------------------------------------
s = 0
FOR i = 1 TO 12
w = 0
l = 1
FOR r = 1 TO 12
IF ztax(r) > w AND btax(r) = 0 THEN l = r: w = ztax(r)
NEXT
btax(l) = 1
s = s + 1
IF s MOD 2 = 0 THEN COLOR 14
IF s MOD 2 = 1 THEN COLOR 10
IF s = 9 THEN PRINT : PRINT #3,
sw = DATESERIAL(2000, l, 1)
PRINT FORMAT$(sw, "MMM"); "-"; FORMAT$(ztax(l), "0000"); " ";
PRINT #3, FORMAT$(sw, "MMM"); "-"; FORMAT$(ztax(l), "0000"); " ";
NEXT

COLOR 11
PRINT #3, FORMAT$(cdc / 4, "00.000000"); " ";
PRINT FORMAT$(cdc / 4, "00.000000"); " ";

COLOR 14
PRINT #3, FORMAT$(ide / idx, "DD/MM/YYYY"); " ";
PRINT FORMAT$(ide / idx, "DD/MM/YYYY"); " ";

COLOR 15
tr = (ide1 / idx) + we(9)
PRINT FORMAT$(tr, "-DD/MM/YYYY-")
PRINT #3, FORMAT$(tr, "-DD/MM/YYYY-")
COLOR 14
PRINT DATE$; "/"; TIME$
PRINT #3, DATE$; "/"; TIME$
PRINT #3, "------------------"

CLOSE #3



OPEN "c:\callerid\CILOGS~1\ci-aver.tmp" FOR INPUT AS #1
lp = 0
LINE INPUT #1, a1$
IF a1$ <> kl$(1) THEN lp = 1: '
LINE INPUT #1, a2$
IF a2$ <> kl$(2) THEN lp = 1: '
LINE INPUT #1, a3$
IF a3$ <> kl$(3) THEN lp = 1: '
LINE INPUT #1, a4$
IF a4$ <> kl$(4) THEN lp = 1: '
LINE INPUT #1, a5$
'IF a5$ <> kl$(5) THEN lp = 1: '
LINE INPUT #1, a6$
IF a6$ <> kl$(6) THEN lp = 1: '
CLOSE #1

IF lp = 1 THEN


OPEN "c:\callerid\CILOGS~1\ci-aver.txt" FOR APPEND AS #1
PRINT #1, a1$
PRINT #1, a2$
PRINT #1, a3$
PRINT #1, a4$
PRINT #1, a5$
PRINT #1, a6$
PRINT "Updated"
CLOSE #1

END IF
END SUB

DEFSNG A-Z
SUB busy ()
DEFDBL A-Z
idc1 = 0: idx1 = 0
FOR i = 1 TO 12
atax(i) = 0: btax(i) = 0
NEXT
FOR i = 1 TO 7
asax(i) = 0: bsax(i) = 0
NEXT
OPEN "c:\callerid\CILOGS~1\cidud5od.txt" FOR INPUT AS #1
DO
LINE INPUT #1, a$
IF INSTR(a$, "NO DIAL") THEN
e = INSTR(a$, "-20")

b$ = MID$(a$, e - 5, 10)
g$ = MID$(a$, e + 6, 8)

IF MID$(b$, 6, 2) = "-2" AND MID$(b$, 3, 1) = "-" THEN
yhg = TIMESERIAL(VAL(MID$(g$, 1, 2)), VAL(MID$(g$, 4, 2)), VAL(MID$(g$, 7, 2)))
ks = ks + yhg
lk = lk + 1
kds = kds + VAL(MID$(b$, 1, 2))
ldk = ldk + 1
k = DATESERIAL(VAL(MID$(b$, 7, 4)), VAL(MID$(b$, 1, 2)), VAL(MID$(b$, 4, 2)))
FOR re = 1 TO 14
IF k < we(re) THEN EXIT FOR
NEXT
ide1 = ide1 + (k - we(re - 1))
idc1 = idc1 + (k - we(re - 1))
idx1 = idx1 + 1
ide = ide + k
idx = idx + 1
g$ = FORMAT$(k, "ddd")
atax(VAL(MID$(b$, 1, 2))) = atax(VAL(MID$(b$, 1, 2))) + 1
'PRINT g$
IF g$ = "Mon" THEN asax(1) = asax(1) + 1: mo = mo + yhg: mo1 = mo1 + 1
IF g$ = "Tue" THEN asax(2) = asax(2) + 1: tu = tu + yhg: tu1 = tu1 + 1
IF g$ = "Wed" THEN asax(3) = asax(3) + 1: we = we + yhg: we1 = we1 + 1
IF g$ = "Thu" THEN asax(4) = asax(4) + 1: th = th + yhg: th1 = th1 + 1
IF g$ = "Fri" THEN asax(5) = asax(5) + 1: fr = fr + yhg: fr1 = fr1 + 1
IF g$ = "Sat" THEN asax(6) = asax(6) + 1: sa = sa + yhg: sa1 = sa1 + 1
IF g$ = "Sun" THEN asax(7) = asax(7) + 1: su = su + yhg: su1 = su1 + 1
END IF
END IF  'nodial
LOOP UNTIL EOF(1)
CLOSE #1


COLOR 12
PRINT "--------------------------------------------------------------------"
COLOR 15
PRINT "When The phones Been Busy "
FOR i = 1 TO 7
w = 0
l = 1
FOR r = 1 TO 7
IF asax(r) > w AND bsax(r) = 0 THEN l = r: w = asax(l)
NEXT
bsax(l) = 1

IF asax(l) <> h THEN s = s + 1
h = asax(l)
IF s MOD 2 = 0 THEN COLOR 14
IF s MOD 2 = 1 THEN COLOR 10
PRINT FORMAT$(asax(l), "0000");
COLOR 11
IF l = 1 THEN g$ = "Mon"
IF l = 2 THEN g$ = "Tue"
IF l = 3 THEN g$ = "Wed"
IF l = 4 THEN g$ = "Thu"
IF l = 5 THEN g$ = "Fri"
IF l = 6 THEN g$ = "Sat"
IF l = 7 THEN g$ = "Sun"
PRINT "-"; g$; " ";
NEXT

PRINT

'-------------------------------------------
FOR i = 1 TO 12
w = 0
l = 1
FOR r = 1 TO 12
IF atax(r) > w AND btax(r) = 0 THEN l = r: w = atax(r)
NEXT
btax(l) = 1
s = s + 1
IF s MOD 2 = 0 THEN COLOR 14
IF s MOD 2 = 1 THEN COLOR 10

sw = DATESERIAL(2000, l, 1)
PRINT FORMAT$(sw, "MMM"); " ";
NEXT


COLOR 11
PRINT FORMAT$(ks / lk, "HH:MM:SS  ");
cmc = cmc + (ks / lk)

PRINT FORMAT$(kds / ldk, "00.000000");
cdc = cdc + (kds / ldk)

COLOR 15
tr = (idc1 / idx1) + we(9)
PRINT FORMAT$(tr, "-DD/MM-")


FOR r = 1 TO 7
cc(r) = cc(r) + asax(r)
NEXT

FOR r = 1 TO 12
ztax(r) = ztax(r) + atax(r)
NEXT


END SUB

DEFSNG A-Z
SUB calls_in ()
DEFDBL A-Z
idc1 = 0: idx1 = 0
FOR i = 1 TO 12
atax(i) = 0: btax(i) = 0
NEXT
FOR i = 1 TO 7
asax(i) = 0: bsax(i) = 0
NEXT
OPEN "c:\callerid\CILOGS~1\cilog.txt" FOR INPUT AS #1
DO
LINE INPUT #1, a$
b$ = MID$(a$, 17, 10)
g$ = MID$(a$, 28, 8)

IF MID$(b$, 6, 2) = "/2" AND MID$(b$, 3, 1) = "/" THEN
yhg = TIMESERIAL(VAL(MID$(g$, 1, 2)), VAL(MID$(g$, 4, 2)), VAL(MID$(g$, 7, 2)))
ks = ks + yhg
lk = lk + 1
kds = kds + VAL(MID$(b$, 1, 2))
ldk = ldk + 1

k = DATESERIAL(VAL(MID$(b$, 7, 4)), VAL(MID$(b$, 4, 2)), VAL(MID$(b$, 1, 2)))
FOR re = 1 TO 14
IF k < we(re) THEN EXIT FOR
NEXT
ide1 = ide1 + (k - we(re - 1))
idc1 = idc1 + (k - we(re - 1))
idx1 = idx1 + 1
ide = ide + k
idx = idx + 1
g$ = FORMAT$(k, "ddd")
atax(VAL(MID$(b$, 4, 2))) = atax(VAL(MID$(b$, 4, 2))) + 1
'PRINT VAL(MID$(b$, 1, 2))
IF g$ = "Mon" THEN asax(1) = asax(1) + 1: mo = mo + yhg: mo1 = mo1 + 1
IF g$ = "Tue" THEN asax(2) = asax(2) + 1: tu = tu + yhg: tu1 = tu1 + 1
IF g$ = "Wed" THEN asax(3) = asax(3) + 1: we = we + yhg: we1 = we1 + 1
IF g$ = "Thu" THEN asax(4) = asax(4) + 1: th = th + yhg: th1 = th1 + 1
IF g$ = "Fri" THEN asax(5) = asax(5) + 1: fr = fr + yhg: fr1 = fr1 + 1
IF g$ = "Sat" THEN asax(6) = asax(6) + 1: sa = sa + yhg: sa1 = sa1 + 1
IF g$ = "Sun" THEN asax(7) = asax(7) + 1: su = su + yhg: su1 = su1 + 1
END IF
LOOP UNTIL EOF(1)
CLOSE #1
'PRINT asax(1); "-"; asax(2); "-"; asax(3); "-"; asax(4); "-"; asax(5); "-"; asax(6); "-"; asax(7)
COLOR 12
PRINT "--------------------------------------------------------------------"
COLOR 15
PRINT "The A Call comes in from since the permanant records file started"
FOR i = 1 TO 7
w = 0
l = 1
FOR r = 1 TO 7
IF asax(r) > w AND bsax(r) = 0 THEN l = r: w = asax(l)
NEXT
bsax(l) = 1

IF asax(l) <> h THEN s = s + 1
h = asax(l)
IF s MOD 2 = 0 THEN COLOR 14
IF s MOD 2 = 1 THEN COLOR 10
PRINT FORMAT$(asax(l), "0000");
COLOR 11
IF l = 1 THEN g$ = "Mon"
IF l = 2 THEN g$ = "Tue"
IF l = 3 THEN g$ = "Wed"
IF l = 4 THEN g$ = "Thu"
IF l = 5 THEN g$ = "Fri"
IF l = 6 THEN g$ = "Sat"
IF l = 7 THEN g$ = "Sun"
PRINT "-"; g$; " ";
NEXT
PRINT

'-------------------------------------------
FOR i = 1 TO 12
w = 0
l = 1
FOR r = 1 TO 12
IF atax(r) > w AND btax(r) = 0 THEN l = r: w = atax(r)
NEXT
btax(l) = 1
s = s + 1
IF s MOD 2 = 0 THEN COLOR 14
IF s MOD 2 = 1 THEN COLOR 10

sw = DATESERIAL(2000, l, 1)
PRINT FORMAT$(sw, "MMM"); " ";
NEXT



COLOR 11
PRINT FORMAT$(ks / lk, "HH:MM:SS  ");
cmc = cmc + (ks / lk)

PRINT FORMAT$(kds / ldk, "00.000000");
cdc = cdc + (kds / ldk)

COLOR 15
tr = (idc1 / idx1) + we(9)
PRINT FORMAT$(tr, "-DD/MM-")

FOR r = 1 TO 7
cc(r) = cc(r) + asax(r)
NEXT

FOR r = 1 TO 12
ztax(r) = ztax(r) + atax(r)
NEXT


END SUB

DEFSNG A-Z
SUB oldest_history ()
DEFDBL A-Z
idc1 = 0: idx1 = 0
FOR i = 1 TO 12
atax(i) = 0: btax(i) = 0
NEXT
FOR i = 1 TO 7
asax(i) = 0: bsax(i) = 0
NEXT
OPEN "c:\callerid\CILOGS~1\telv2.txt" FOR INPUT AS #1
DO
LINE INPUT #1, a$
LOOP UNTIL MID$(a$, 17, 12) = "L I S T    1"

DO
LINE INPUT #1, a$

b$ = MID$(a$, 17, 10)
g$ = MID$(a$, 28, 8)

IF MID$(b$, 6, 1) = "/" AND MID$(b$, 3, 1) = "/" THEN
yhg = TIMESERIAL(VAL(MID$(g$, 1, 2)), VAL(MID$(g$, 4, 2)), VAL(MID$(g$, 7, 2)))
ks = ks + yhg
lk = lk + 1
kds = kds + VAL(MID$(b$, 1, 2))
ldk = ldk + 1

k = DATESERIAL(VAL(MID$(b$, 7, 4)), VAL(MID$(b$, 4, 2)), VAL(MID$(b$, 1, 2)))
FOR re = 1 TO 14
IF k < we(re) THEN EXIT FOR
NEXT
ide1 = ide1 + (k - we(re - 1))
idc1 = idc1 + (k - we(re - 1))
idx1 = idx1 + 1
ide = ide + k
idx = idx + 1
g$ = FORMAT$(k, "ddd")
atax(VAL(MID$(b$, 4, 2))) = atax(VAL(MID$(b$, 4, 2))) + 1
IF g$ = "Mon" THEN asax(1) = asax(1) + 1: mo = mo + yhg: mo1 = mo1 + 1
IF g$ = "Tue" THEN asax(2) = asax(2) + 1: tu = tu + yhg: tu1 = tu1 + 1
IF g$ = "Wed" THEN asax(3) = asax(3) + 1: we = we + yhg: we1 = we1 + 1
IF g$ = "Thu" THEN asax(4) = asax(4) + 1: th = th + yhg: th1 = th1 + 1
IF g$ = "Fri" THEN asax(5) = asax(5) + 1: fr = fr + yhg: fr1 = fr1 + 1
IF g$ = "Sat" THEN asax(6) = asax(6) + 1: sa = sa + yhg: sa1 = sa1 + 1
IF g$ = "Sun" THEN asax(7) = asax(7) + 1: su = su + yhg: su1 = su1 + 1
END IF

LOOP UNTIL MID$(a$, 17, 12) = "L I S T    3" OR EOF(1)

CLOSE #1


COLOR 12
PRINT "--------------------------------------------------------------------"
COLOR 15
PRINT "When The calls have came in using the longest history"
FOR i = 1 TO 7
w = 0
l = 1
FOR r = 1 TO 7
IF asax(r) > w AND bsax(r) = 0 THEN l = r: w = asax(l)
NEXT
bsax(l) = 1

IF asax(l) <> h THEN s = s + 1
h = asax(l)
IF s MOD 2 = 0 THEN COLOR 14
IF s MOD 2 = 1 THEN COLOR 10
PRINT FORMAT$(asax(l), "0000");
COLOR 11
IF l = 1 THEN g$ = "Mon"
IF l = 2 THEN g$ = "Tue"
IF l = 3 THEN g$ = "Wed"
IF l = 4 THEN g$ = "Thu"
IF l = 5 THEN g$ = "Fri"
IF l = 6 THEN g$ = "Sat"
IF l = 7 THEN g$ = "Sun"
PRINT "-"; g$; " ";
NEXT
PRINT

'-------------------------------------------
FOR i = 1 TO 12
w = 0
l = 1
FOR r = 1 TO 12
IF atax(r) > w AND btax(r) = 0 THEN l = r: w = atax(r)
NEXT
btax(l) = 1
s = s + 1
IF s MOD 2 = 0 THEN COLOR 14
IF s MOD 2 = 1 THEN COLOR 10

sw = DATESERIAL(2000, l, 1)
PRINT FORMAT$(sw, "MMM"); " ";
NEXT


COLOR 11
PRINT FORMAT$(ks / lk, "HH:MM:SS  ");

cmc = cmc + (ks / lk)

PRINT FORMAT$(kds / ldk, "00.000000");
cdc = cdc + (kds / ldk)

COLOR 15
tr = (idc1 / idx1) + we(9)
PRINT FORMAT$(tr, "-DD/MM-")

FOR r = 1 TO 7
cc(r) = cc(r) + asax(r)
NEXT

FOR r = 1 TO 12
ztax(r) = ztax(r) + atax(r)
NEXT


END SUB

DEFSNG A-Z
SUB rings ()
DEFDBL A-Z
OPEN "c:\callerid\CILOGS~1\cidud6.txt" FOR INPUT AS #1
DO
LINE INPUT #1, a$
b$ = MID$(a$, 26, 10)

g$ = MID$(a$, 37, 8)

IF MID$(b$, 6, 2) = "-2" AND MID$(b$, 3, 1) = "-" THEN
yhg = TIMESERIAL(VAL(MID$(g$, 1, 2)), VAL(MID$(g$, 4, 2)), VAL(MID$(g$, 7, 2)))
ks = ks + yhg
lk = lk + 1
kds = kds + VAL(MID$(b$, 1, 2))
ldk = ldk + 1
k = DATESERIAL(VAL(MID$(b$, 7, 4)), VAL(MID$(b$, 1, 2)), VAL(MID$(b$, 4, 2)))
FOR re = 1 TO 14
IF k < we(re) THEN EXIT FOR
NEXT
ide1 = ide1 + (k - we(re - 1))
idc1 = idc1 + (k - we(re - 1))
idx1 = idx1 + 1

ide = ide + k
idx = idx + 1
g$ = FORMAT$(k, "ddd")
atax(VAL(MID$(b$, 1, 2))) = atax(VAL(MID$(b$, 1, 2))) + 1

IF g$ = "Mon" THEN asax(1) = asax(1) + 1: mo = mo + yhg: mo1 = mo1 + 1
IF g$ = "Tue" THEN asax(2) = asax(2) + 1: tu = tu + yhg: tu1 = tu1 + 1
IF g$ = "Wed" THEN asax(3) = asax(3) + 1: we = we + yhg: we1 = we1 + 1
IF g$ = "Thu" THEN asax(4) = asax(4) + 1: th = th + yhg: th1 = th1 + 1
IF g$ = "Fri" THEN asax(5) = asax(5) + 1: fr = fr + yhg: fr1 = fr1 + 1
IF g$ = "Sat" THEN asax(6) = asax(6) + 1: sa = sa + yhg: sa1 = sa1 + 1
IF g$ = "Sun" THEN asax(7) = asax(7) + 1: su = su + yhg: su1 = su1 + 1
END IF
LOOP UNTIL EOF(1)
CLOSE #1
'PRINT atax(1); "-"; atax(2); "-"; atax(3); "-"; atax(4); "-"; atax(5); "-"; atax(6); "-"; atax(7); "-"; atax(8); "-"; atax(9); "-"; atax(10); "-"; atax(11); "-"; atax(12)
COLOR 12
PRINT "--------------------------------------------------------------------"
COLOR 15
PRINT "The Phone Rings and for how long it does......................"
FOR i = 1 TO 7
w = 0
l = 1
FOR r = 1 TO 7
IF asax(r) > w AND bsax(r) = 0 THEN l = r: w = asax(l)
NEXT
 bsax(l) = 1
IF asax(l) <> h THEN s = s + 1
h = asax(l)
IF s MOD 2 = 0 THEN COLOR 14
IF s MOD 2 = 1 THEN COLOR 10
PRINT FORMAT$(asax(l), "0000");
COLOR 11
IF l = 1 THEN g$ = "Mon"
IF l = 2 THEN g$ = "Tue"
IF l = 3 THEN g$ = "Wed"
IF l = 4 THEN g$ = "Thu"
IF l = 5 THEN g$ = "Fri"
IF l = 6 THEN g$ = "Sat"
IF l = 7 THEN g$ = "Sun"
PRINT "-"; g$; " ";
NEXT
PRINT

'-------------------------------------------
FOR i = 1 TO 12
w = 0
l = 1
FOR r = 1 TO 12
IF atax(r) > w AND btax(r) = 0 THEN l = r: w = atax(r)
NEXT
btax(l) = 1
s = s + 1
IF s MOD 2 = 0 THEN COLOR 14
IF s MOD 2 = 1 THEN COLOR 10

sw = DATESERIAL(2000, l, 1)
PRINT FORMAT$(sw, "MMM"); " ";
NEXT


COLOR 11
PRINT FORMAT$(ks / lk, "HH:MM:SS  ");
cmc = cmc + (ks / lk)

PRINT FORMAT$(kds / ldk, "00.000000");
cdc = cdc + (kds / ldk)

COLOR 15
tr = (ide1 / idx) + we(9)
PRINT FORMAT$(tr, "-DD/MM-")

FOR r = 1 TO 7
cc(r) = cc(r) + asax(r)
NEXT
FOR r = 1 TO 12
ztax(r) = ztax(r) + atax(r)
NEXT

END SUB

