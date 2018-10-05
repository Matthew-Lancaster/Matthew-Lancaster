'Launches Registry Editor with the chosen branch open automatically
'Author  : Ramesh Srinivasan
'Website: http://windowsxp.mvps.org
'http://stackoverflow.com/questions/137182/how-to-launch-windows-regedit-with-certain-path
'-----

Set WshShell = CreateObject("WScript.Shell")
Dim MyKey
MyKey = Inputbox("Type the Registry path")
MyKey = "My Computer\" & MyKey
WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit\Lastkey",MyKey,"REG_SZ"
WshShell.Run "regedit", 1,True
Set WshShell = Nothing




EXIT

From http://windowsxp.mvps.org/jumpreg.htm (I have not tried any of these):

When you start Regedit, it automatically opens the last key that was viewed. (Registry Editor in Windows XP saves the last viewed registry key in a separate location). If you wish to jump to a particular registry key directly without navigating the paths manually, you may use any of these methods / tools.

Option 1
Using a VBScript: Copy these lines to a Notepad document as save as registry.vbs

'Launches Registry Editor with the chosen branch open automatically
'Author  : Ramesh Srinivasan
'Website: http://windowsxp.mvps.org

Set WshShell = CreateObject("WScript.Shell")
Dim MyKey
MyKey = Inputbox("Type the Registry path")
MyKey = "My Computer\" & MyKey
WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit\Lastkey",MyKey,"REG_SZ"
WshShell.Run "regedit", 1,True
Set WshShell = Nothing
Double-click Registry.vbs and then type the full registry path which you want to open.

Example: HKEY_CLASSES_ROOT\.MP3

Limitation: The above method does not help if Regedit is already open.

Note: For Windows 7, you need to replace the line MyKey = "My Computer\" & MyKey with MyKey = "Computer\" & MyKey (remove the string My). For a German Windows XP the string "My Computer\" must be replaced by "Arbeitsplatz\".

Option 2
Regjump from Sysinternals.com

This little command-line applet takes a registry path and makes Regedit open to that path. It accepts root keys in standard (e.g. HKEY_LOCAL_MACHINE) and abbreviated form (e.g. HKLM).

Usage: regjump [path]

Example: C:\Regjump HKEY_CLASSES_ROOT\.mp3

Option 3
12Ghosts JumpReg from 12ghosts.com

Jump to registry keys from a tray icon! This is a surprisingly useful tool. You can manage and directly jump to frequently accessed registry keys. Unlimited list size, jump to keys and values, get current key with one click, jump to key in clipboard, jump to same in key in HKCU or HKLM. Manage and sort keys with comments in an easy-to-use tray icon menu. Create shortcuts for registry keys.

shareimprove this answer
edited Mar 14 at 14:06

Mofi
7,50631025	
answered Sep 26 '08 at 1:10

The Hoss
8712
