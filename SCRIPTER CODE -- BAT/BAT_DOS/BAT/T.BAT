@echo off
rem if exist \edrive.txt c:\utils\u %1
rem if exist \fdrive.txt c:\utils\u %1
rem if exist \gdrive.txt c:\utils\u %1
rem if exist \hdrive.txt c:\utils\u %1
rem if exist \edrive.txt goto end
rem if exist \fdrive.txt goto end
rem if exist \gdrive.txt goto end
rem if exist \hdrive.txt goto end
rem c:\utils\original\u %1
rem c:\utils\u /v %1
c:\utils\u %1 /V
rem :end
