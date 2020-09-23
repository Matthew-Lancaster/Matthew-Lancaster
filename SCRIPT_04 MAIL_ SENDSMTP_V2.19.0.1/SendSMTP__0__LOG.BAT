@ECHO.
CD\

ECHO ____ TITLE.BAT ____ 
TITLE %~f0

TYPE C:\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP.log >> C:\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP__01_%COMPUTERNAME%.log

type nul >C:\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP.log


TYPE C:\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP_SMTP.log >> C:\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP__02_%COMPUTERNAME%.log

type nul >C:\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP_SMTP.log

