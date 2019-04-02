@ECHO OFF
W32TM /REGISTER

REM -- SEVER NOT WORKING -- "ntp2c.mcc.ac.uk" 
REM -- SEVER NOT WORKING -- "ntp.exnet.com"

REM -- 'Sun 14 August 2016 17:00:00
REM -- '----------------------------------------------------------------------------------------------------
REM -- 'ALL AS DESCRIBED HERE 4 STAGE
REM -- '----
REM -- 'Synchronize time throughout your entire Windows network - Page 6040425 - TechRepublic
REM -- 'http://www.techrepublic.com/article/synchronize-time-throughout-your-entire-windows-network/6040425/
REM -- '----
REM -- '----------------------------------------------------------------------------------------------------

REM -- ----
REM -- UK Time Server | Time servers and atomic clocks | Galleon Systems Ltd
REM -- http://www.galsys.co.uk/time-server/uk-time-server.html
REM -- ----

W32tm /config /manualpeerlist:"time.nist.gov","time.windows.com" /syncfromflags:manual

REM W32tm /config /manualpeerlist:"0.uk.pool.ntp.org","1.uk.pool.ntp.org" /syncfromflags:manual

REM W32tm /config /manualpeerlist:"0.uk.pool.ntp.org","1.uk.pool.ntp.org","ntp2a.mcc.ac.uk","ntp.exnet.com" /syncfromflags:manual

REM W32tm /config /manualpeerlist:"time.windows.com" /syncfromflags:manual
REM W32tm /config /manualpeerlist:"time.windows.com","ntp2c.mcc.ac.uk","ntp.exnet.com" /syncfromflags:manual


REM -- NOT TO USE RESET TIME HAPPEN TO ABOUT 2015
W32tm /config /update

REM -- NOT TO USE RESET TIME HAPPEN TO ABOUT 2015
W32TM /RESYNC
pause