@echo off
cd c:\
ECHO -----------------------------------------------
ECHO STARTING TO COPY   c:\ drive to d: 
ECHO ----------------------------------------------- 

xCOPY C:\.         c:\root2         /d /f /h /k /r /c /x /y >c:\rootback.txt

