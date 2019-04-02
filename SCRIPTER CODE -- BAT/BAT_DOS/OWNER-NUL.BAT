Icacls * /setowner "Everyone" /T /C /L >NUL

@ECHO NEXT 02 OF 02 -- GRANT EVERYONE
@PAUSE

Icacls * /t  /grant Everyone:F >NUL

@PAUSE