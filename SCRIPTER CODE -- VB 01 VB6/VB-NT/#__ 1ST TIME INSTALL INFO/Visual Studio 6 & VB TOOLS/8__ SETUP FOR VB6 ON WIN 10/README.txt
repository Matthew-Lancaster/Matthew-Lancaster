NOTICE:
* VB6-generic.iss, SP6-generic.iss, and VS6-generic.issare Inno setup Script files.
* Use VB6-generic.iss for porting Visual Basic 6.0 installer CD.
* Use SP6-generic.iss for porting Visual Studio 6.0 SP6 installer.
* Use VB6-VS6-generic.iss for porting Visual Studio 6.0 installer CD (Creates VB6 installer only).
* Use the output installer for personal use only. Always keep Visual Studio 6.0 CD for future use.

INSTRUCTIONS:
* Download and install Inno Setup 5.4.3 (Download Link: http://files.jrsoftware.org/is/5/isetup-5.4.3.exe)
* Copy all the contents of VB6/VS6 installer CD into a directory
* Place the inno setup script beside SETUP.EXE
* Open the script with Inno Setup.
* On menu bar click BUILD->COMPILE or press CTRL+F9
* Wait until it finish compiling.
* When compiling is finished. The created installer was on a folder named OUTPUT beside SETUP.EXE (For example if the path of setup.exe is C:\VB60, then the created installer is located at C:\VB60\OUTPUT)

INSTRUCTIONS FOR PORTING SP6:
* Extract VS6 SP6 installer using WinRAR or by just running the installer
* Extract VS6sp61.cab (make sure that VS6sp62.cab, VS6sp63.cab, and VS6sp64.cab are also exists beside VS6sp61.cab) in a folder
* Place SP6-generic.iss on the folder where VS6sp61.cab is extracted
* Open the script with Inno Setup.
* On menu bar click BUILD->COMPILE or press CTRL+F9
* Wait until it finish compiling.