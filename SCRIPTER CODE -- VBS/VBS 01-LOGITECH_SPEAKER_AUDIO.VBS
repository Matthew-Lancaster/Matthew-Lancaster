DIM X_TIMER

Set Sound = CreateObject("WMPlayer.OCX.7")
Sound.URL = "C:\SCRIPTER\SCRIPTER -- VBS\VBS 01-LOGITECH_SPEAKER_AUDIO 10Hz.wav"
Sound.Controls.play
DO WHILE Sound.currentmedia.duration = 0

'	wscript.sleep 100

	X_TIMER = NOW + TIMESERIAL(0,0,50)
	DO
	'SLEEP 10
	LOOP UNTIL X_TIMER < NOW

	' BEEP
LOOP

'wscript.sleep (int(Sound.currentmedia.duration)+1)*1000
'wscript.sleep 



' 2017_09_09 _ EDDIE BIRTHDAY DAY AFTER _ EVIE BIRTHDAY
' ----
' HOW TO STOP THE LOGITECH SPEAKER GOING INTO SLEEP MODE - Google Search
' https://www.google.co.uk/search?num=50&rlz=1C1CHBF_en-GBGB759GB759&q=HOW+TO+STOP+THE+LOGITECH+SPEAKER+GOING+INTO+SLEEP+MODE&spell=1&sa=X&ved=0ahUKEwje9rrbvZbWAhWDY1AKHVhZCdEQvwUIJSgA&biw=1536&bih=646
' --------
' Prevent Speaker's Standby Mode [Solved] - Peripherals - Audio
' http://www.tomsguide.com/answers/id-2417519/prevent-speaker-standby-mode.html
' ----