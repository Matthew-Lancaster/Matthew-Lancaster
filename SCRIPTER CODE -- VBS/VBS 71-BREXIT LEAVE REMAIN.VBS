
' [ Sunday 00:30:00 Pm_22 September 2019 ]
' [ Sunday 15:51:30 Pm_22 September 2019 ]
' -------------------------------------------------------------------


' THEY COMING TO TAKE YOU AWAY HA HA

' **
' Matthew Lancaster
' HTTP://TINYURL.COM/Y4SJBLRN
' **



DIM A(7)
A(1)="LEAVE"
A(2)="EXIT"
A(3)="REMAIN"
A(4)="HIGHER"
A(5)="LOWER"
A(6)="STAY"
A(7)="TAKEAWAY"
FOR R1=1 TO 7
FOR R2=1 TO 7
	T=T+A(R1)+" "+A(R2)+VBCRLF
NEXT
NEXT

' PUT IT TO THE CLIPBOARD
' -----------------------
Set WshShell = CreateObject("WScript.Shell")
Set oExec = WshShell.Exec("clip")

Set oIn = oExec.stdIn

oIn.WriteLine T
oIn.Close

Wscript.QUIT
	
END IF
' ----
' How can I use Clipboard in vbscript? - Stack Overflow
' https://stackoverflow.com/questions/19696308/how-can-i-use-clipboard-in-vbscript
' ----


