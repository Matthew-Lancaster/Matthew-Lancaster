

Reg.exe add "HKCU\SOFTWARE\Microsoft\Visual Basic\6.0\Addins\VB6IDEMouseWheelAddin.Connect" /v "FriendlyName" /t REG_SZ /d "MouseWheel Fix" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Visual Basic\6.0\Addins\VB6IDEMouseWheelAddin.Connect" /v "LoadBehavior" /t REG_DWORD /d "3" /f
Reg.exe add "HKCU\SOFTWARE\Microsoft\Visual Basic\6.0\Addins\VB6IDEMouseWheelAddin.Connect" /v "CommandLineSafe" /t REG_DWORD /d "0" /f


Exit

