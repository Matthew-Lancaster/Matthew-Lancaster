DEFDBL A-Z
IF 1 = 2 THEN
SHELL "c:\Progra~1\Window~1\MPLAYER2 E:\WIN98SE\cdsample\videos\barneysw.avi"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\combat.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\encrs.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\golf99.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\hompub99.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\keyboard.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\mcmi5.avi"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\mon99.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\mouseall.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\mpt2.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\mtm2.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\open.avi"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\phone.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\photo.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\pinball.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\rome.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\swfree.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\swwheel.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\varcade.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\WIN98SE\cdsample\videos\works.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\videos\goodtime.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\videos\weezer.mpg"
END IF
'OPEN "E:\videos\vcum.avi" FOR INPUT AS #1
't = LOF(1)
'CLOSE #1

SHELL "c:\progra~1\window~1\MPLAYEr2 E:\videos\vcum.avi"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\videos\vnasty.avi"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\videos\kinky\kinkyc~8.mpg"
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\videos\vcum.avi"
t = TIMER + (60 * 7 + 6)
IF t > 86398.95 THEN t = 0
DO
IF INKEY$ <> "" THEN EXIT DO
LOOP UNTIL t < TIMER
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\videos\vnasty.avi"
t = TIMER + (60 * 6 + 56)
IF t > 86398.95 THEN t = 0
DO
IF INKEY$ <> "" THEN EXIT DO
LOOP UNTIL t < TIMER
SHELL "c:\progra~1\window~1\MPLAYEr2 E:\videos\kinky\kinkyc~8.mpg"
t = TIMER + (60 * 38 + 34)
IF t > 86398.95 THEN t = 0
DO
y$ = INKEY$
IF y$ <> "" THEN EXIT DO
LOOP UNTIL t < TIMER
IF y$ <> " " THEN END
RUN

