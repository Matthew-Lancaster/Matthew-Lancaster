

REM  ----------------------------------------------------------------
REM -- EXAMPLE 
REM  ----------------------------------------------------------------
REM -- AppCache132361150295058860.txt
REM -- AppCache*.txt
REM -- NOT DELETE ALL GET DATE
REM -- LEAVE RECENT THERE
REM  ----------------------------------------------------------------
REM  ----------------------------------------------------------------


CD /D C:\Users\MATT 04\AppData\Local\Packages\Microsoft.Windows.Cortana_cw5n1h2txyewy\LocalState\DeviceSearchCache\




GOTO ENDER



"C:\PStart\# NOT INSTALL REQUIRED\FFMPEG\ffmpeg.exe"

for %%i in (*.mp4) do ECHO -err_detect ignore_err -i %%i -c copy %%~ni_fixed.mp4
pause

:ENDER

PAUSE

EXIT
