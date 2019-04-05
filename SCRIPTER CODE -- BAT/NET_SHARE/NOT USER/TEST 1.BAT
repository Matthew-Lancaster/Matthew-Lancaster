@echo off
setlocal enabledelayedexpansion

:: User Defined Variables
set FilterOU=TRUE
set OU=Environmental
if %FilterOU%==TRUE (
set qComp=wmic /NAMESPACE:\\root\directory\ldap PATH ds_computer Where "ds_distinguishedname like '%%%OU%%%'" get ds_name
) ELSE (
set qComp=wmic /NAMESPACE:\\root\directory\ldap PATH ds_computer get ds_name
) 

Echo.
Echo Getting list of computers.
Echo.
for /f "skip=1" %%c in ('%qComp%') do ( 
  For /f %%d in ("%%c") do (
    REM ping -n 1 %%d | find "TTL=">nul && set r="up"||set r="down"
    
	REM if /i !r!=="up" (
       REM psexec -c -f -s -h /accepteula \\%%d "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\NET_SHARE_ALL.BAT" 
	   ECHO \\%%d
	REM )  
  )
)

PAUSE

exit /b

