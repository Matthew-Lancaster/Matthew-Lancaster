' -------------------------------------------------------------------
' Mon 18-Mar-2019 02:19:46 P
' C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 41-GET DATE AND TIME FORMATTER.VBS
' -------------------------------------------------------------------
' END WRITE
' Mon 18-Mar-2019 15:10:00 P
' [ Monday 15:10:00 Pm_18 March 2019 ]
' -------------------------------------------------------------------
' CREDIT
' ----
' Echo VBScript results in .bat file - Stack Overflow
' https://stackoverflow.com/questions/23918779/echo-vbscript-results-in-bat-file
' ----
' ----
' Get the Date and Time - VBScript - SS64.com
' https://ss64.com/vb/syntax-getdatetime.html
' ----
' -------------------------------------------------------------------


' -------------------------------------------------------------------
' USAGE IN BATCH FILE
' SET FILENAME_VBS="C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 41-GET DATE AND TIME FORMATER.VBS"
' for /f "tokens=*" %%a in ('cscript //nologo %FILENAME_VBS%') do (set "strVBScriptResults=%%a")
' SET DATE_TIME_FORMATER=%strVBScriptResults%
' ECHO %DATE_TIME_FORMATER%
' -------------------------------------------------------------------


VAR_NOW=NOW
YY=Year(VAR_NOW)
MM=MONTH(VAR_NOW)
MMM = MonthName(MONTH(VAR_NOW),TRUE)
DD=Day(VAR_NOW)
HH24=Hour(VAR_NOW)
MM=Minute(VAR_NOW)
SS=Second(VAR_NOW)
HH12=Hour(VAR_NOW)
IF HH12>12 THEN 
	HH12=HH12-12
	AM_PM="PM"
ELSE
	AM_PM="AM"
END IF 
IF HH12=12 THEN 
	AM_PM="PM"
END IF 


DDD=WeekdayName(Weekday(VAR_NOW),TRUE)

i_TIME=""
i_TIME=i_TIME+LPad(YY, "0", 4)+"-"
i_TIME=i_TIME+MMM+"-"
i_TIME=i_TIME+LPad(DD, "0", 2)
i_TIME=i_TIME+" -- "
i_TIME=i_TIME+LPad(HH24, "0", 2)+":"
i_TIME=i_TIME+LPad(MM, "0", 2)+":"
i_TIME=i_TIME+LPad(SS, "0", 2)
i_TIME=i_TIME+" - "
i_TIME=i_TIME+LPad(HH12, "0", 2)+":"
i_TIME=i_TIME+LPad(MM, "0", 2)+":"
i_TIME=i_TIME+LPad(SS, "0", 2)+"-"+AM_PM
i_TIME=i_TIME+" - "
i_TIME=i_TIME+DDD


Wscript.Echo i_TIME


WScript.Quit(0)

'--------------------------------------------------------------------
'END STOP -----------------------------------------------------------
'--------------------------------------------------------------------

Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function



' MsgBox i_TIME
' MsgBox Format(Now, "YYYY-MMM-DD -- HH:MM:SS - HH:MM:SS-AM/PM DDD")
' 2019-Mar-18 -- 14:00:40 - 02:00:40-PM - Sun
