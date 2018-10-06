

SET v=%DATE:~3,2%
SET map=01-Jan;02-Feb;03-Mar;04-Apr;05-May;06-Jun;07-Jul;08-Aug;09-Sep;10-Oct;11-Nov;12-Dec
CALL SET v=%%map:*%v%-=%%
SET v=%v:;=&rem.%
ECHO %DATE:~6,4%-%v%-%DATE:~0,2% -- %TIME:~0,8%

PAUSE
