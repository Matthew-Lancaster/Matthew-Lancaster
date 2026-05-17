[]
  [] Dalsoft Solutions Www.Dalsoft.co.uk
[]


Windows 2000/XP Run As without having to Type a password
_________________________________________________________


Platform Windows 2000, Windows XP
Support at support@dalsoft.co.uk
Suggestions to RunAsSuggestions@dalsoft.co.uk

This class enhances the brillant Run as command which ships
with Windows 2000/XP Run As allows you to Run a program as any user 
without the need to log on and off. 

My RunAs class allows Network Admins to use Run as without having to type in a password this
means you can run a batch file using run as with no user input.

This class uses the API code from my forthcoming RunApp product the only different
is my final product will allow admins to create encrypted files which can be passed into the 
excuteable so the username and password won't be exposed in a batch or VBScript file. 


To use this class compile it on the server or client you wish to use it 
on. Register it using regsvr32 "Path to compiled runas.dll" at the command prompt.
 
regsvr32 C:\WINDOZE\system32\runas.dll

Example of VBScript or VB code you would use.

Set oRunAs = CreateObject("runas.clsrunas","Your Server")
with oRunAs
	.sDomain = "Your Domain Here"
	.sUserName = "Username you want to Run the Program  "
	.sPassword = "The password to the the above Username"
	.sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember
		     you must explictly use the relevant program for example if you wanted
		     to open a text document you can't use c:\text.txt or "notepad c:\text.txt" 
		     you would have to use 
		     "c:\winnt\notepad.exe c:\text.txt"
	.RunAs 'Call the Run As method
End With
set oRunas = Nothing

Please Note the disclaimer below before using this Class.

Disclaimer: You may install copy, modify, use and distribute the included code/files in 
RunAs.zip providing you agree to the following: 
     
To the maximum extent permitted by applicable law, Darran Jones and Dalsoft hereby 
disclaim with respect to the files included in RunAs.zip and any technical support provided
all warranties and conditions, whether express, implied or statutory. 
In no event shall Darrran Jones or Dalsoft be liable for any direct, special, incidental, punitive, consequential indirect, or 
any other damages whatsoever (including, but not limited to, damages for: 
loss of confidential or other information, business interruption, loss of profits, 
loss of privacy, personal injury, failure to meet any duty (including of good faith or 
of reasonable care), negligence, and any other pecuniary or other loss whatsoever) 
related to the use of or inability to use the files 
included in RunAs.zip and any technical support provided (or any failure to provide 
any technical support).
