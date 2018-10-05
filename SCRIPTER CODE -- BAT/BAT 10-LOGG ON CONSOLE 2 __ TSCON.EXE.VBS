REM DUPLICATE _ WORKING ON __ BAT 10-LOGG ON CONSOLE 2 __ TSCON.EXE.VBS

Set objShell = WScript.CreateObject("WScript.Shell")
REM objShell.Run """C:\Program Files\Sandboxie\Start.exe"" /box:NetBeans /wait ""C:\Program Files\NetBeans 7.3\bin\netbeans64.exe""", 1, True

REM strCommand = Quotes("C:\WINDOWS\System32\tscon.exe") & _
         REM " 1 /DEST:Console /PASSWORD:"" "" /V "

strCommand = Quotes("tscon.exe") & _
         " 1 /DEST:Console /PASSWORD:"" "" /V "

		 objShell.Run strCommand, 1, TRUE

Set objShell = Nothing



REM strCommand = Quotes("C:\WINDOWS\System32\tscon.exe") & _
         REM " /box:NetBeans /wait " &                            _
         REM Quotes("C:\Program Files\NetBeans 7.3\bin\netbeans64.exe")


Function Quotes(ByVal strValue)
    Quotes = Chr(34) & strValue & Chr(34)
End Function

REM Set objShell = CreateObject("WScript.Shell")
REM comspec = objShell.ExpandEnvironmentStrings("%comspec%")
REM Set objExec = objShell.Exec("tscon.exe" & "1 /DEST:Console /PASSWORD:"" "" /V")
REM Set objExec = objShell.Exec(comspec & " /c ipconfig")

REM scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

REM EXECUTE_PARAM = "/SELECT, """ + scriptdir + """"
REM EXECUTE_PARAM = """" + scriptdir + """"
REM EXECUTE_PARAM= scriptdir
REM REM MSGBOX EXECUTE_PARAM

REM Set objShell = CreateObject("Wscript.Shell")

REM EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE " + EXECUTE_PARAM 
REM objShell.Run EXECUTE_FILE_NAME, 0, True

REM MSGBOX EXECUTE_FILE_NAME

REM UNABLE TO DO PARAM YET

REM objShell.Run "C:\WINDOWS\EXPLORER.EXE", 0, True

REM objShell.Run "C:\WINDOWS\System32\tscon.exe 1 /DEST:Console /PASSWORD:"" "" /V"
REM objShell.Run ("C:\WINDOWS\System32\tscon.exe" & "1 /DEST:Console /PASSWORD:"" "" /V")
REM objShell.Run ("tscon.exe" & "1 /DEST:Console /PASSWORD:"" "" /V")

REM CreateObject("Wscript.Shell").Run """" & "tscon.exe" & "1 /DEST:Console /PASSWORD:"" "" /V" & """", 0, False


REM '"tscon.exe"

REM ---------------------------------------------------------
REM EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE /SELECT, C:\"
REM VBS CODE WILL NOT RUN FROM NOTEPAD++ -- NOT TRUE
REM ---------------------------------------------------------
