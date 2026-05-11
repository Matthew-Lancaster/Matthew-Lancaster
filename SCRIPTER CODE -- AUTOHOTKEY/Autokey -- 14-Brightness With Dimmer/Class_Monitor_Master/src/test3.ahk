#Persistent ; may not actually be necessary...
Loop, 255
Hotkey, % "~" Format("vk{:x}",A_Index), pressed, On
return

pressed:
Tooltip % "You pressed "  GetKeyName(SubStr(A_ThisHotkey,2))
return