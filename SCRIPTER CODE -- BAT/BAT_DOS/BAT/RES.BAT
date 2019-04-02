@echo off
if "%config%!"=="ramdisk!" goto loadg
if "%config%!"=="!" Echo set your system up better terminating.
if "%config%!"=="!" goto end
echo loading for C: Main drive
c:
cd\
md coolcid
cd\coolcid
c:
goto loader
:loadg
echo loading for G: Ram drive.
g:
cd\
md coolcid
cd\coolcid
:loader
'if exist a:allzips1.zip xcopy a:allzips1.zip
'if ERRORLEVEL 1 if exist a:allzips2.zip xcopy a:allzips2.zip
'if ERRORLEVEL 1 if exist a:allzips3.zip xcopy a:allzips3.zip
'if ERRORLEVEL 1 echo restore terminated on allzip check files.
'if ERRORLEVEL 1 goto end
'if exist allzips2.zip rename allzips2.zip allzips1.zip
'if exist allzips3.zip rename allzips3.zip allzips1.zip
'if exist a:passmat1.zip xcopy a:passmat1.zip
'if ERRORLEVEL 1 if exist a:passmat2.zip xcopy a:passmat2.zip
'if ERRORLEVEL 1 if exist a:passmat3.zip xcopy a:passmat3.zip
'if ERRORLEVEL 1 echo restore terminated on passZip check files.
'if ERRORLEVEL 1 goto end
'if exist passmat2.zip rename passmat2.zip passmat1.zip
'if exist passmat3.zip rename passmat3.zip passmat1.zip
rem --------------
if exist allzips1.zip pkunzip -o -d mats.zip errors.zip cisub.zip ci0.zip telzips.bat bigjunk.zip passzip.zip
if ERRORLEVEL 1 echo check Your Files Errors Occured.
if ERRORLEVEL 1 goto end
if exist allzips1.zip goto endofstory
rem -----
rem -----
if exist passmat1.zip pkunzip -o -d -s%1 allzips1.zip mats.zip errors.zip cisub.zip ci0.zip telzips.bat bigjunk.zip passzip.zip
if ERRORLEVEL 1 echo Unzip terminated on Zip Coz password may be incorrect.
if ERRORLEVEL 1 goto end
'if exist passmat1.zip del passmat1.zip
:endofstory
if exist allzips1.zip call \coolcid\bclink
:end
