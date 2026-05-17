SHELTARG.DLL has one fundtion to retrieve the target path to a 
ShellLink (Windows shortcut)

'Handel to the calling window
ByVal hWnd As Long
'Full path the a ShellLing
ByVal lpszLinkFile As String
'String var to retrieve target path
ByVal lpszPath As String
'String var to retrieve description
ByVal lpszDescription As String

Public Declare Function GetLinkInfo Lib "sheltarg.dll" (ByVal hWnd As Long, ByVal lpszLinkFile As String, ByVal lpszPath As String, ByVal lpszDescription As String) As Long


I was in a position were I needed to be able to shell to a shortcut and
pass parameters. This did not always work well with the Shell() or 
ShellExecute() functions. I needed to get the target of the shortcut
and then shell to it. Another problem was that if I shelled to a shortcut 
that is a link to a Folder I would need to copy files to that folder and 
not pass them as a parameter to a shelled function.

I found a C++ routine on the MS Knowledge Base that retrieved the target 
path and a file description. After I compiled it into a DLL I found that 
it did not always return a description. This may be because the program that 
created the shortcut did not supply it. I'm not sure. I incorporated a 
GetFileType function to get around this.

The DLL is freeware and I will not be supporting it. If you have problems 
with it I won't be able to help you. You can use it in your programs and 
distribute it freely with your programs. You can pass along this sample 
project and DLL but do not separate them.

Thank you,

Greg DeBacker
http://windsweptsoftware.com




