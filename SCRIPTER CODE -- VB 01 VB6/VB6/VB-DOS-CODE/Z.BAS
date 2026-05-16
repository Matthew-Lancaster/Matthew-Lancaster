'IF RTRIM$(COMMAND$) = "NOTSTAYING" THEN END
DEFINT A-Z

DIM zef    AS LONG
DIM szip   AS LONG
DIM ca     AS LONG
DIM poi    AS LONG
DIM poi2   AS LONG
DIM zpos   AS LONG
DIM zstlar AS LONG
DIM zipid  AS INTEGER
DIM tap1   AS DOUBLE


TYPE gar
 etr  AS STRING * 67
 cr   AS STRING * 2
' etr2 AS STRING * 67
' cr2  AS STRING * 2
END TYPE
DIM gar AS gar

TYPE zip
pk    AS INTEGER
pkid  AS INTEGER
whit  AS STRING * 20
big   AS LONG
flen  AS INTEGER
whit3 AS STRING * 8
dorf  AS INTEGER
whit4 AS STRING * 6
END TYPE

DIM zip AS zip
DEF SEG = 0
absol = PEEK(&H484)'''''50 line mode
WIDTH (80), 50
dis$ = "°°°°°°°°°°°°   0% of All Files in Zip have been Read.°°°M.LANCASTER APR 1994°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°"
dis2$ = "°°°°°°°°°   0% of All Files in Zip have been Deleted.°°°M.LANCASTER APR 1994°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°"

1 ON ERROR GOTO ers
rur$ = COMMAND$
rur2$ = rur$
DIM pkdd AS STRING * 3
DIM dir$(1000)
DIM fil$(5000)

dell = 7
DIM two$(dell, 2)
two$(1, 1) = "pkzipfix.exe"
two$(2, 1) = "pkunzip.exe"
two$(3, 1) = "pkzip.exe"
two$(4, 1) = "ncd.exe"
two$(5, 1) = "z.exe"
two$(6, 1) = "ncd.exe"
two$(7, 1) = "zgargant.com"
DIM twos(dell)
DIM a$(60)
d$ = "zip\packet.zip ": 'must be a space at this end
e$ = "zip\*.*"
a = 0
a = a + 1: a$(a) = "A) Add    Max           -a+ex -whs -rp -x" + e$ + " " + rur2$
a = a + 1: a$(a) = "B) Add    Norm          -a+en -whs -rp -x" + e$ + " " + rur2$
a = a + 1: a$(a) = "F) Add    Super fast    -a+es -whs -rp -x" + e$ + " " + rur2$
a = a + 1: a$(a) = "Z) Add    Zero          -a+e0 -whs -rp -x" + e$ + " " + rur2$
a = a + 1: a$(a) = "I) Insert modified loot -i-ex -whs -rp -x" + e$ + " " + rur2$
a = a + 1: a$(a) = "U) Update               -a-ex -whs -rp -x" + e$ + " " + rur2$
a = a + 1: a$(a) = "R) Repair damage        " + d$
a = a + 1: a$(a) = "P) Purge File(s)        Exculding " + d$
a = a + 1: a$(a) = "T) Test                 -t " + d$
a = a + 1: a$(a) = "L) List (View)          -v " + d$
a = a + 1: a$(a) = "X) Extract              -$ -d -n " + d$ + " " + rur2$
a = a + 1: a$(a) = "S) Span to floppys      -& -e0 A:PACK1 " + d$
t = a
d = 0
e = 0
eg = 0

GOSUB enviro
DEF SEG = 0

'absol = PEEK(&H484)'''''50 line mode
'tdri = PEEK(&H475)'''''two drives or not

red = 0


GOSUB show

dir$ = "ncd"


jump1:





FOR r = 1 TO t
a$(15 + r) = MID$(a$(r), 1, 1)
a$(30 + r) = MID$(a$(r), 25, LEN(a$(r)))
NEXT

d = 1

DO
b$ = UCASE$(INKEY$)
IF b$ = CHR$(13) AND twos(4) = 1 THEN SHELL DIR$: RUN'd = 0: GOSUB show: b$ = ""
IF cdri$ = "" THEN END
IF b$ <> "" THEN
e = 0 'menu option to key press converter....
FOR r = 1 TO l
IF r = 7 AND twos(5) = 0 THEN r = r + 1
IF b$ = a$(15 + r) THEN e = r: d = 1: GOSUB show: e = 0
NEXT
END IF
LOOP UNTIL b$ = " "


IF eg = 1 THEN GOSUB enviro


END
chomp: ''''''''''''''''''''''''''''''''''''''''''''''
f = 0
ON ERROR GOTO ers:
COLOR 14
ty = FREEFILE
OPEN "c:\current" FOR OUTPUT AS ty
PRINT #ty, cdri$
PRINT #ty, driv$
CLOSE #ty

gos = 0
IF e < 6 AND l = t THEN
GOSUB someexist
IF gos = 1 THEN COLOR 12: PRINT : PRINT "No update needed..Check Your Files.."
END IF
IF e < 6 AND gos = 1 THEN GOTO jump

er = 0
IF e < 6 THEN MKDIR "zip"
IF er > 0 THEN PRINT "Make dir    not nes on zip"
IF er = 0 THEN
MKDIR "dummy"
CHDIR "dummy"
SHELL "ncd rd . >nul"
END IF

IF a$(15 + e) = "I" THEN
IF tdri = 2 THEN
PRINT "Second drive present........Use IT (Y/N)";
DO
key$ = INKEY$
IF key$ <> "" THEN IF UCASE$(key$) = "Y" THEN PRINT "Yes..": ddriv$ = drives$ ELSE PRINT "No...": ddriv$ = ""
LOOP UNTIL key$ <> ""
END IF
END IF


IF a$(15 + e) = "P" THEN
GOSUB purge

GOSUB someexist

IF gos = 0 THEN
COLOR 14
PRINT "Some files Still Exist. I'll try and remove R/O Attribs and Re-Delete"
SHELL "attrib -r -h -s /s >nul"
GOSUB purge
END IF

GOSUB someexist

IF gos = 0 THEN COLOR 10: PRINT "Their still in Here, Update."

IF zdr > 0 THEN MKDIR "dummy": CHDIR "dummy": SHELL "ncd rd . >nul"
END IF

IF a$(15 + e) = "P" THEN GOTO jump

f = 1'''''''''''

IF a$(15 + e) = "R" THEN
f = 0
CHDIR "zip"
SHELL "pkzipfix packet.zip"
e5 = 0
OPEN "pkfixed.zip" FOR INPUT AS #1: CLOSE #1
IF e5 > 0 THEN PRINT "Fixed.Zip   does not exist"
IF e5 = 0 THEN
KILL "packet.zip"
NAME "pkfixed.zip " AS "packet.zip"
CHDIR ".."
GOSUB more

END IF
END IF

IF a$(15 + e) = "X" THEN
'---------------------------
WIDTH (80), absol + 1
SHELL "pkunzip " + a$(30 + e)
'GOSUB ziperrors
IF zdr > 0 THEN MKDIR "dummy": CHDIR "dummy": SHELL "ncd rd . >nul"
f = 0
PRINT
PRINT "Clear attribs:";
IF l = t THEN
IF MID$(a$, 3, 1) = "A" THEN arc = 1 ELSE arc = 0
END IF
SHELL "attrib *.* -a /s >nul"

END
'******************************
END IF

IF a$(15 + e) = "T" THEN
SHELL "pkunzip " + a$(30 + e)
'GOSUB ziperrors
f = 0
END IF

IF a$(15 + e) = "L" THEN
SHELL "pkunzip " + a$(30 + e)
'GOSUB ziperrors
f = 0
END IF

IF a$(15 + e) = "S" THEN
f = 0
SHELL "pkzip " + a$(30 + e)
'GOSUB ziperrors
END IF
'PRINT




IF f = 1 THEN
'-------------------------------

WIDTH (80), absol + 1
gr$ = CHR$(13) + CHR$(0)
gar.cr = gr$': gar.cr2 = gr$
er = 0
IF a$(15 + e) <> "I" AND a$(15 + e) <> "U" THEN KILL "zip\packet.zip"
IF er > 0 THEN PRINT "Killer      not nes on zip"''''''''from erl 21
gar.etr = "PKZIP" + CHR$(0) + "= " + RTRIM$(ddriv$ + d$ + a$(30 + e))
'gar.etr2 = "z    " + CHR$(0) + CHR$(&H3D) + " notstaying"
ON ERROR GOTO 0
OPEN two$(7, 2) + "zgargant.com" FOR BINARY AS #1
'OPEN "zgargant.com" FOR BINARY AS #1

PUT #1, &H4DD, gar'etr
'SHELL "t c:\zgargant.com "
CHAIN two$(7, 2) + "zgargant.com"
END
END IF

'****************************
ON ERROR GOTO ers:
'IF f = 1 THEN f = 0: GOSUB ziperrors: f = 1
IF f = 1 THEN GOSUB zipexist
IF f = 1 AND l = t - 4 THEN RMDIR "zip"
jump:
COLOR 14
IF CSRLIN = 50 THEN PRINT
PRINT "\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ Finished /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"

'SOUND 300, 1
RETURN 'from chomp

setpath:
CHDIR cdri$
IF tdri = 2 THEN CHDIR driv$
RETURN

show:
'SCREEN 6
rt = 0: GOSUB part2: rt = 1
IF e > 0 THEN GOSUB chomp: GOSUB part2

COLOR 15
PRINT "Space to exit......."

RETURN

part2:

IF rt = 0 THEN CLS
IF rt = 0 AND e = 0 THEN
GOSUB zipexist: GOSUB getdr: GOSUB more
ELSE
xln = CSRLIN
GOSUB reading
END IF

red = 0
COLOR 11: PRINT "Current Path";
IF tdri = 2 THEN COLOR 14: PRINT "(s). ";  ELSE PRINT ". ";
IF da$ = "C" THEN COLOR 11: path$ = cdri$: of = 81 ELSE COLOR 9
PRINT cdri$;
IF tdri = 2 THEN COLOR 14: PRINT " .&. "; :
IF da$ = "D" THEN COLOR 11: path$ = driv$: of = 118 ELSE COLOR 9
PRINT driv$
'IF tdri = 1 THEN PRINT




COLOR 11: PRINT "Size of Unpacked Zip  Is   :";
COLOR 14: PRINT zstlar;
zxc = 100 - INT((zef / zstlar) * 100)
COLOR 11: PRINT "Ratio = ";
COLOR 14: PRINT zxc; "% ";
COLOR 11: PRINT "Squashed \ Files = ";
COLOR 14: PRINT zt


COLOR 11: PRINT "Size  Of  This   Zip  Is   :";
COLOR 14: PRINT zips$; " ";
COLOR 11: PRINT "DIRectory's = ";
COLOR 14: PRINT zdr; " ";
COLOR 11: PRINT "Unpacked = ";
COLOR 14: PRINT CLNG(100 * (unpac / zt)); "%"
COLOR 11: PRINT "Size  Of  Free  Space Is   : ";
COLOR 14: PRINT size$; " ";
COLOR 11: PRINT "Zip Status Is : =";
COLOR 14: PRINT stat$

IF VAL(size$) < zstlar THEN red = 1

FOR r = 1 TO t
IF red = 1 AND r = t - 1 THEN COLOR 12: PRINT a$(r): r = r + 1
IF e <> r THEN
IF l >= r THEN COLOR 11
IF l < r THEN COLOR 8
IF r < 5 THEN COLOR 10
IF twos(5) = 0 AND r = 7 THEN COLOR 8
PRINT a$(r)
END IF

IF e = r THEN
COLOR 12
PRINT a$(15 + r);
COLOR 15
PRINT MID$(a$(r), 2, 1);
COLOR 10
PRINT MID$(a$(r), 3, 70)
END IF

NEXT
RETURN






zipexist:
l = t
8 OPEN d$ FOR INPUT AS #1
9 zips$ = STR$(LOF(1))
IF LOF(1) = 0 THEN zips$ = " Nothingness": zt = 0: zdr = 0
CLOSE #1
'GOSUB more
RETURN

someexist:
gos = 0
SCREEN 0, 1, 2, 0
CLS
FILES
w = SCREEN(2, 37)
w2 = SCREEN(2, 55)
PRINT ".............."
SCREEN 0, 1, 0, 0
COLOR 14
IF w2 = 32 AND l = t THEN gos = 1
IF w2 = 32 AND w = 32 AND l < t THEN gos = 1
RETURN


ers:
er = ERR
fuck = ERL
IF ERL = 1 THEN a = ERR: RESUME NEXT
IF ERL = 6 THEN fline = 5: RESUME NEXT
IF ERL = 8 THEN l = t - 6: RESUME NEXT
IF ERL = 9 THEN a = 1: RESUME NEXT
IF ERL = 19 THEN er = 1: RESUME NEXT

END


enviro: ''''''''''''''''''''''''''''''''''''''''''''''''''''
IF eg = 0 THEN
rt$ = ENVIRON$("PATH")
oi = 1
COLOR 12
uio = 0
yes = 0
IF INSTR(oi, rt$, ";") <> 0 THEN
DO
er = 0
w = INSTR(oi, rt$, ";")
IF w = 0 THEN r$ = MID$(rt$, oi)
yes = 1
IF w > 0 THEN r$ = MID$(rt$, oi, w - oi)
FOR r = 1 TO dell
er = 0'''''''''''''''''''''''''''line 8
IF RIGHT$(r$, 1) <> "\" THEN r$ = r$ + "\"
IF twos(r) = 0 THEN OPEN r$ + two$(r, 1) FOR INPUT AS #1: CLOSE #1
IF er = 0 AND twos(r) = 0 THEN twos(r) = 1: two$(r, 2) = r$
NEXT

oi = w + 1
LOOP UNTIL w = 0
END IF

IF yes = 0 THEN PRINT "I cant allow this program to run as you have no Path.": END


FOR r = 1 TO dell
IF twos(r) = 0 AND eg = 0 THEN eg = 1
NEXT
END IF

IF eg = 1 THEN COLOR 12: PRINT "Some files cannot be located. sort it out..": PRINT

FOR r = 1 TO dell
IF twos(r) = 0 THEN PRINT "I cant find "; : COLOR 14: PRINT two$(r, 1); : COLOR 12: PRINT " On your system "
NEXT
RETURN

getdr: '''''''''''''''''''''''''''''''''''''''''''''''
'GOSUB zipexist:
that$ = ""
size$ = ""
that2$ = ""
SCREEN 0, 1, 2, 0
CLS
FILES "nul"
FOR r = 1 TO 50
the$ = CHR$(SCREEN(1, r))
IF the$ = " " THEN EXIT FOR
that$ = that$ + the$
NEXT

FOR r = 2 TO 50
sis$ = CHR$(SCREEN(3, r))
IF sis$ = " " THEN EXIT FOR
size$ = size$ + sis$
NEXT

va = VAL(size$)
path$ = that$
cdri$ = that$
da$ = LEFT$(that$, 1)
IF da$ = "C" THEN thier$ = "D" ELSE thier$ = "C"
that$ = ""
CLS

er = 0
FILES thier$ + ":nul"''''''''''eln=2
drive$ = ""
IF er = 0 THEN
FOR r = 1 TO 50
the$ = CHR$(SCREEN(1, r))
IF the$ = " " THEN EXIT FOR
that$ = that$ + the$
NEXT
driv$ = that$: tdri = 2
drives$ = "-bd:\ "
END IF

SCREEN 0, 1, 0, 0

RETURN







more:

unpac = 0
stat$ = "None"
xln = CSRLIN

IF l <> t THEN GOTO theend
stat$ = "O.K."


OPEN d$ FOR BINARY AS #1

zef = LOF(1)
IF zef = 0 THEN CLOSE #1: KILL d$: GOTO theend
GET 1, zef - 5, szip
ca = SEEK(1)
szip = szip + 1
GET 1, szip + 2, zipid
stat$ = "O.K.": w = 1

IF zipid <> &H201 THEN
stat$ = "O.K. + Header"
SEEK 1, zef - 3000: hn$ = INPUT$(3000, #1)
w = INSTR(hn$, "PK" + CHR$(5) + CHR$(6))
IF w = 0 THEN stat$ = "Corupted **"
GET 1, zef - 3000 + w + 15, szip
ca = SEEK(1)
szip = szip + 1
GET 1, szip + 2, zipid
IF zipid <> &H201 THEN w = 0: stat$ = "Corupted .."

END IF

IF w = 0 THEN GOTO theend

poi = ca - szip
zpos = 20
SEEK 1, szip



zxc = 0
zstlar = 0          'counter for unpacked size
zdr = 0              'counter for directory
zt = 0                'counter for files
er = 0
hid = 0
COLOR 15
cent = 0
DO
GOSUB reading
GET 1, , zip
hn$ = INPUT$(zip.flen, #1)
zstlar = zip.big + zstlar
IF zip.pkid = &H605 THEN EXIT DO
IF zip.pk <> &H4B50 THEN w = 0: stat$ = "Corupted.": EXIT DO'''????w=0
zpos = zpos + 46 + zip.flen

IF zip.dorf = &H30 THEN
zdr = zdr + 1: dir$(zdr) = LEFT$(hn$, zip.flen - 1): pop = 1

ELSE
zt = zt + 1: fil$(zt) = hn$
er = 0: OPEN fil$(zt) FOR INPUT AS #2: CLOSE #2
IF er = 0 THEN unpac = unpac + 1
END IF

LOOP UNTIL 1 = 2' OR zt>1000
CLOSE #1

sdf$ = STR$(stlar)

rezt = zt
theend:
GOSUB reading

RETURN


reading:
LOCATE xln, 1

IF szip > 0 THEN tap1 = zpos / poi
IF zef = 0 THEN PRINT "°"; MID$(dis$, 1, 79): GOTO exit1

RSET pkdd = LTRIM$(STR$(CLNG(tap1 * 100)))
MID$(dis$, 14) = pkdd
PRINT "°"; MID$(dis$, 1, CLNG(tap1 * 79))
exit1:
RETURN


purging:
RSET pkdd = LTRIM$(STR$(CLNG((cent / rezt) * 100)))
MID$(dis2$, 11) = pkdd
LOCATE xln, 1
PRINT MID$(dis2$, 1, CLNG((zr / rezt) * 80))
RETURN

purge:
xln = CSRLIN
hid75 = 0
hid53 = 0
ON TIMER(1) GOSUB purging
TIMER ON
cent = 0
COLOR 15

FOR zr = 1 TO zt
er = 0
KILL fil$(zr)

IF er = 53 THEN hid53 = hid53 + 1
IF er = 75 THEN hid75 = hid75 + 1
IF er = 0 THEN cent = cent + 1
NEXT
TIMER OFF
zr = zr - 1
GOSUB purging
PRINT
IF hid75 > 0 THEN COLOR 12: PRINT "° There is "; hid75; " read only file(s)."
IF hid53 > 0 THEN COLOR 12: PRINT "° "; hid53; " file(s) Did Not exist."


IF zdr > 0 THEN
FOR rr = 1 TO 12
er = 0
FOR zr = 1 TO zdr
'PRINT dir$(zr)
RMDIR DIR$(zr)
IF er = 76 THEN hid76 = hid76 + 1: er = 0
NEXT
IF er = 0 THEN EXIT FOR
NEXT
COLOR 14
DIM lie(4) AS STRING * 2: lie(1) = "st": lie(2) = "nd": lie(3) = "rd": lie(4) = "Th":
li = VAL(RIGHT$(STR$(rr), 1))
IF li = 0 OR li > 3 THEN li = 4
IF hid76 = 0 AND er = 0 THEN PRINT "° All Directory's in Zip have been Deleted. In "; rr; " Pass(es)"
IF hid76 > 0 THEN PRINT "² On the "; rr; lie(li); " Pass of Removing Directrys "; hid76; " Did Not exist"
END IF
RETURN



















