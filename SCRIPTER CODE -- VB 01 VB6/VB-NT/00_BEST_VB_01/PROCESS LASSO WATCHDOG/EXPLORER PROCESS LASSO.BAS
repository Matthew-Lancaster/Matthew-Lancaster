Attribute VB_Name = "Module1"
Sub Main()

If App.PrevInstance = True Then End

I = ""
I = I + App.Path + "\"
I = I + vbCrLf
I = I + App.EXEName + ".EXE"
I = I + vbCrLf + vbCrLf
I = I + "PROCESS LASSO WATCH DOG HAS GOT EXPLOER OVER A CPU %"
MsgBox I, vbMsgBoxSetForeground, App.Path

End

End Sub
