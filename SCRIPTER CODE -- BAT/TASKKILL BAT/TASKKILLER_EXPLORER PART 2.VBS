
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

REM EXECUTE_PARAM = "/SELECT, """ + scriptdir + """"
REM EXECUTE_PARAM = """" + scriptdir + """"
REM EXECUTE_PARAM= scriptdir
REM REM MSGBOX EXECUTE_PARAM

Set objShell = CreateObject("Wscript.Shell")

REM EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE " + EXECUTE_PARAM 
REM objShell.Run EXECUTE_FILE_NAME, 0, True

REM MSGBOX EXECUTE_FILE_NAME

REM UNABLE TO DO PARAM YET

objShell.Run "C:WINDOWS\EXPLORER.EXE", 0, True


REM ---------------------------------------------------------
REM EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE /SELECT, C:\"
REM VBS CODE WILL NOT RUN FROM NOTEPAD++ -- NOT TRUE
REM ---------------------------------------------------------
