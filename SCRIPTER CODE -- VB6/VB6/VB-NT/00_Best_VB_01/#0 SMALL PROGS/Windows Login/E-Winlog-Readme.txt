E-WinLog - Windows95 Log System 

This tiny app keeps track of when Windows95 is started and when
it is shut down and who was logged on in the period. 
It creates a log file which by default is named
E-WinLog.sys and placed in the Windows System directory unless anything
else is specified on the command line.

The log file is written in standard Windows text (.txt) format.
The log file contains one line per login/logout, which looks
something like this:


User            Log in               Log out              Time logged on         IP used         Host computer
--------------- -------------------- -------------------- ---------------------- --------------- ----------------
Bill98          04.05.00 - 18:02:38  04.05.00 - 19:42:18  D 0, H 1, M 40 ,S 25   211.032.132.005 Matrix

E-Winlog runs in the background at all times, but
it uses very little system resources. It is invisible in the tasklist,
so the user has no way to detect it.

E-Winlog can be startet from System Polycies, a loginscript or by placing it
in the windows startup folder.

From a script, it is started like this: e-winlog [commandline parameters] 

If started without parameters, it starts in default mode.

The commandline parameters are:
/?		Shows a helpwindow with all parameters.
/r path		Redirects the logfile to somwhere else than the system dir. *
/q		Kills a already running instance of the program.
/i		Shows information about existing running instance of program or not.
/l		Shows info about the location of the current logfile.
/v		Open the current logfile in Notepad. **

*  Note! Only path, like "e-winlog.exe /r c:\temp" or "e-winlog.exe /r c:\temp\". 
         The logfilename will be the same.

** Note! If the program is running and the logfile has been redirected with the /r parameter,
	 the current logfile will be opened. 
	 If the program is not running, the logfile in default location will be tried opened.

INSTALLATION:
Run Setup.exe
The program will be installed in the windows/command directory. It is therefore not
nessecery to give full path to the program when starting it, as long as the command
directory is in the search path (wich it is by default).

UNINSTALLATION:
Simply access the Windows95 installation program in the Control Panel,
where E-Winlog leaves an uninstall string.

About E-WinLog
- Programmed by Trond Sørensen, 2000
- Written and Compiled in Visual Basic 6.0

KNOWN LIMITATIONS:
E-Winlog calculates the time between logon and logof using only the last to numbers in the
year so there is bound to be a "millenium bug" here, in 2100 that is ;-) 

DISTRIBUTION:
This software is FREEWARE.
It may be freely copied and distributed on any 
NON-PROFIT FREEWARE collection.

No guarantee.

