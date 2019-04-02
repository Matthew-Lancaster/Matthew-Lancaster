@echo off
H:
cd H:\winnt\system32\config
attrib -h *.log
g:
cd g:\winnt\system32\config
attrib -h *.log
H:
xcopy . g: <c:\bat\win2kbak.tmp
xcopy . g:d <c:\bat\win2kbak.tmp
xcopy . H:d <c:\bat\win2kbak.tmp
attrib +h *.log
g:
attrib +h *.log
g:
cd g:\
attrib -h -s pagefile.sys
H:
cd H:\
attrib -h -s pagefile.sys
xcopy pagefile.sys g: 
attrib +h +s pagefile.sys
rem --------------------
g:
cd g:\docume~1\matthe~1
attrib -s -h ntuser.dat
attrib -s -h ntuser.ini
h:
cd H:\docume~1\matthe~1
attrib -s -h ntuser.dat
attrib -s -h ntuser.ini
xcopy ntuser.dat g:
xcopy ntuser.ini g:
attrib +h ntuser.dat
attrib +s +h ntuser.ini
rem -------------------
rem --------------------
g:
cd g:\docume~1\defaul~1
attrib -s -h ntuser.dat
h:
cd H:\docume~1\defaul~1
attrib -s -h ntuser.dat
xcopy ntuser.dat g:
attrib +h ntuser.dat
rem -------------------
rem --------------------
g:
cd g:\docume~1\alluse~1
attrib -s -h -r ntuser.pol
h:
cd H:\docume~1\alluse~1
attrib -s -h -r ntuser.pol
xcopy ntuser.pol g:
attrib +h ntuser.pol
rem -------------------
g:
cd g:\docume~1\matthe~1\locals~1\applic~1\micros~1\windows
attrib -s -h -r usrclass.dat
h:
cd h:\docume~1\matthe~1\locals~1\applic~1\micros~1\windows
attrib -s -h -r usrclass.dat
xcopy usrclass.dat g:
attrib +h usrclass.dat
rem -------------------

superman

