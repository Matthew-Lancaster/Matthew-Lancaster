@REM __ WINDOWS UPDATE 
@REM __ ERROR SERACH
@REM __ 17-May
@REM __ 0xc1900208

@REM __ URL
@REM __ https://answers.microsoft.com/en-us/windows/forum/windows_10-windows_install-winpc/windows-10-install-not-working-0xc1900208/bb09f53a-5d96-4946-94a2-0d40a13b02c6

@REM __ URL
@REM __ https://answers.microsoft.com/en-us/windows/forum/windows_10-update/how-can-i-fix-windows-update-error-cumulative/c4372ddc-224c-46dd-856c-53b128b26a6b?auth=1

net stop bits

net stop wuauserv

net stop appidsvc

net stop cryptsvc

@REM Rename the software distribution folders backup copies. To do this, at a command prompt, type the following commands. Make sure that you press Enter after you type each command.

Ren %systemroot%\SoftwareDistribution SoftwareDistribution.bak

Ren %systemroot%\system32\catroot2 catroot2.bak

@REM Restart the BITS service, the Windows Update service, and the Cryptographic service. To do this, at a command prompt, type the following commands. Make sure that you press Enter after you type each command.

net start bits

net start wuauserv

net start appidsvc

net start cryptsvc 

PAUSE