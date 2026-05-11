#InstallKeybdHook

#Persistent

TClick = 0

SetTimer Stroke,1



Stroke:

;   If (A_TimeIdlePhysical > 30 or A_TickCount-TClick < 125)
   If (A_TickCount-TClick < 1)

		SoundBeep , 2500 , 400
   
      Return
	  
	  
Return