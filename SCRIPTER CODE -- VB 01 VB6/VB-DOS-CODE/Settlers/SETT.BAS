DEFDBL A-Z
DIM s1 AS STRING * 1
DIM s2 AS STRING * 1
DIM s3 AS STRING * 1
OPEN "settle.exe" FOR BINARY AS #4
OPEN "settle~1.exe" FOR BINARY AS #1
OPEN "settle~2.exe" FOR BINARY AS #2
OPEN "settle~3.exe" FOR BINARY AS #3
q = NOW + TIMESERIAL(0, 1, 0)
FOR r = 1 TO LOF(1)
'IF q < NOW THEN q = NOW + TIMESERIAL(0, 1, 0): PRINT r
IF r MOD 100000 = 0 THEN PRINT r

GET #1, r, s1
GET #2, r, s2
GET #3, r, s3
ok = 0
IF s1 = s2 AND s1 = s3 AND s2 = s3 THEN ok = 1: PUT #4, r, s1
IF ok = 0 THEN
IF s1 = s2 THEN ok = 1: PRINT "1", r: PUT #4, r, s1
IF s1 = s3 THEN ok = 1: PRINT "2", r: PUT #4, r, s1
IF s2 = s3 THEN ok = 1: PRINT "3", r: PUT #4, r, s2
END IF
IF ok = 0 THEN PRINT "error"; r: ers = 1
NEXT
CLOSE #1, #2, #3, #4
IF ers = 1 THEN PRINT "error"

