@echo off
echo Erase all files and sub directory Files
erase *.* /f /ah /s /q
echo remove all directorys and sub directorys
rd . /s /q
rdd 1
cd ..
rdd 2
rem vbdos /run c:\vb\rdd.bas
rem @c:\utils\kill\kill
