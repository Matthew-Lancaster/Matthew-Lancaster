Set objShell = CreateObject("WScript.Shell")
 
objShell.RegWrite "HKCR\Folder\Shell\MenuText\Command\", _
    "cmd.exe /K chdir /D " & chr(34) & "%1" & chr(34)
objShell.RegWrite "HKCR\Folder\Shell\MenuText\", "Command Prompt Here"

'----
'Script Windows 7 Command Prompt Here
'https://gallery.technet.microsoft.com/scriptcenter/bbf46677-ebc9-4c0e-b5d7-ef061a605002
'----