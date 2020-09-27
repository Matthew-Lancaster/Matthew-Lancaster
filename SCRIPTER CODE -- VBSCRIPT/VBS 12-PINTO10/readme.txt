: Release v1.1 - Stuart Pearson 14th Nov 2015
:
: Command line tool to pin and unpin exe / lnk files to the Windows 10 taskbar and start menu.
:
: The exe needs to be run with two switches specified for each function - Pin / Unpin to Taskbar / Start Menu...
:
: To pin an application or shortcut to the taskbar you must include the two following switches...
: /PTFOL: followed directly by the folder containing the exe you want to pin (always surround in quotes).
: /PTFILE: followed directly by the name of the executable you want to pin (always surround in quotes).
:
: To unpin an application or shortcut to the taskbar you must include the two following switches...
: /UTFOL: followed directly by the folder containing the exe you want to unpin (always surround in quotes).
: /UTFILE: followed directly by the name of the executable you want to unpin (always surround in quotes).
:
: To pin an application or shortcut to the start menu you must include the two following switches...
: /PSFOL: followed directly by the folder containing the exe you want to pin (always surround in quotes).
: /PSFILE: followed directly by the name of the executable you want to pin (always surround in quotes).
:
: To unpin an application or shortcut to the start menu you must include the two following switches...
: /USFOL: followed directly by the folder containing the exe you want to unpin (always surround in quotes).
: /USFILE: followed directly by the name of the executable you want to unpin (always surround in quotes).

: Example for pinning a file to the taskbar...
PinTo10.exe /PTFOL:"C:\Windows" /PTFILE:"notepad.exe"

: Example for unpinning a file to the taskbar...
PinTo10.exe /UTFOL:"C:\Windows" /UTFILE:"notepad.exe"

: Example for pinning a file to the start menu...
PinTo10.exe /PSFOL:"C:\Windows" /PSFILE:"notepad.exe"

: Example for unpinning a file from the start menu...
PinTo10.exe /USFOL:"%USERPROFILE%\Desktop" /USFILE:"Word 2016.lnk"