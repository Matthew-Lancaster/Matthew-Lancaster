REM COPY THIS TO WINDOWS\SYSTEM32 XP 32BIT AND RUN THIS BAT FILE
REM http://www.fmsinc.com/microsoftaccess/controls/mscomctl/
REM 17 OCT 2015

Regsvr32.exe /u C:\Windows\System32\MSCOMCTL.OCX
COPY C:\Windows\System32\MSCOMCTL.OCX.LEGACY C:\Windows\System32\MSCOMCTL.OCX
Regsvr32.exe C:\Windows\System32\MSCOMCTL.OCX
Regsvr32.exe /u C:\Windows\System32\MSCOMCTL.OCX
DEL C:\Windows\System32\MSCOMCTL.OCX
COPY C:\Windows\System32\MSCOMCTL.OCX.LEGACY  C:\Windows\System32\MSCOMCTL.OCX
Regsvr32.exe C:\Windows\System32\MSCOMCTL.OCX
PAUSE
EXIT
