'8788 GOODSYNC PID

for each Process in GetObject("winmgmts:{impersonationLevel=impersonate}//localhost").ExecQuery("Select PercentProcessorTime,IDProcess from Win32_PerfFormattedData_PerfProc_Process where IDProcess=8788")
    WScript.Echo(Process.PercentProcessorTime)
    WScript.quit  
next