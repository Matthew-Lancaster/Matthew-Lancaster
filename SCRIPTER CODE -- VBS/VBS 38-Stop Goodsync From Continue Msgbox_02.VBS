'====================================================================
'# __ C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 38-Stop Goodsync From Continue Msgbox.VBS
'# __ 
'# __ VBS 38-Stop Goodsync From Continue Msgbox.VBS
'# __ 
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __ 
'# START      TIME [ Thu 18-Oct-2018 13:39:00 ] 
'# END LAST EDITOR [ Thu 18-Oct-2018 13:39:00 ] 
'# __ 
'====================================================================

'# ------------------------------------------------------------------
' DESCRIPTION 
'# ------------------------------------------------------------------
' THIS IS A SIMPLE ONE 
' BUT TOOK A WHILE TO THINKER
' IT PUT A STOP IN GOODSYNC
' WHEN RUN A JOB AND ADD THIS AS A SCRIPT
' AND START GOOD SYNC
' YOU ARE ABLE TO HAVE SOME JOB AT THE FRONT
' WHEN I WORK ON A DRIVE CRASH RECOVERY
' I WANTED AUTO OFF A LOT OF TIME
' BUT THERE ARE LIMIT OF WHAT ABLE TO DO WITH AUTO OFF
' AND TIME CONSUMING WORK
' AND MADE ALL THE BETTER WITH THIS ONE GOING
' AS YOU ARE ABLE TO LEAVE AUTO ON
' AND SCRIPT RUN THIS CODE
' MAKE HALF
' AS LONG AS SCRIPT LOADER IN GOODSYNC DOESN'T USE NOWAIT OPTION
' PLUS BONUS ALL THE TIMERS STILL WORK
' IDEA IS YOU DON'T PRESS IT ON THE MSGBOX SO THAT HALT GOODSYNC GOING FURTHER
' WITH THE TIMERS WORKING YOU ABEL TO HAVE LONG JOB RUNNING
' AND MAINTAIN SMALLER JOB AS IT INDICATE
' WHEN THEY ARE READY 
' ONLY THAT YOU HAVE TO DO THEM MANUALLY
' SO GET ALL YOUR SYNC IN ORDER 
' AND THEN MOVE ONTO NEXT LEVEL
' RECOVERY DONE
' MORE JOB GOING TO BACKUP A BIT MORE

' ONE LINER MSGBOX THAT IS ALL
' NEW TICK _ FOR TICK OR TEAT HALLOWEEN NIGHT

' RELEASE YOUR TIME BEING TAKEN UP
' SET JOB TO RUN ANALYZE AND THEN SYNC-ER
' AND RERUN GOODSYNC SO RUN ONE AT TOP

' ALSO IF YOU PRESS OK ON MESSAGE BOX YOUR HAVE TO TURN AUTO OFF
' BUT YOU ABEL TO RUN THE SCRIPT AGAIN AND PUT AUTO BACK ON 

' -------------------------------------------------------------------
' -------------------------------------------------------------------

'# ------------------------------------------------------------------
' LOCATTION ON-LINE 
'# ------------------------------------------------------------------
' -------------------------------------------------------------------
' GITHUB
' -------------------------------------------------------------------

'#-------------------------------------------------------------------
' 001 __ 
'#-------------------------------------------------------------------
' ADDER
' ADD CODE
' -------------------------------------------------------------------
' -------------------------------------------------------------------
' Thu 18-Oct-2018 13:39:00
' Thu 18-Oct-2018 13:39:00 __ 
'
' Thu 18-Oct-2018 18:20:00
' Thu 18-Oct-2018 18:22:00 __ WRITE THE HEADER
'#-------------------------------------------------------------------


' MSGBOX "GoodSync Script Command to Stop in Wait using a Messenger Box",

Dim lRet
' lRet = MsgBox("GoodSync Script Command to Stop in Wait using a Messenger Box"+VBCRLF+VBCRLF+"Goodsync Will Reset as is at The End Of a Cycle If CountDown Left", vbYesNo + vbMsgBoxSetForeground)

lRet = MsgBox("GoodSync xtreme Lock For a While Messenger Box"+VBCRLF+VBCRLF+"", vbOK + vbMsgBoxSetForeground)

Wscript.QUIT
