@echo off
if exist \edrive.txt c:\utils\u %1
if exist \fdrive.txt c:\utils\u %1
if exist \gdrive.txt c:\utils\u %1
if exist \hdrive.txt c:\utils\u %1
if exist \edrive.txt goto end
if exist \fdrive.txt goto end
if exist \gdrive.txt goto end
if exist \hdrive.txt goto end
rem c:\utils\original\u %1
c:\utils\u %1
:end
