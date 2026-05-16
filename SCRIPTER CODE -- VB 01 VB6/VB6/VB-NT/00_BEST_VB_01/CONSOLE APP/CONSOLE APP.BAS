Attribute VB_Name = "Module1"


Sub Main_4()
Dim obj As Object
Dim TMP As String

obj = CreateObject("Console.Console")
obj.StartUp "My Console"
obj.ConsoleWriteLine "Hello World"
obj.ConsoleWriteLine "Press Enter To Exit"
TMP = obj.ConsoleReadLine
obj.Shutdown
End Sub




Sub Main()
    Dim FSO As New Scripting.FileSystemObject
    FSO.GetStandardStream(StdOut).WriteLine "Your argument was " & Command$()
End Sub


'Requires a reference to Microsoft Scripting Runtime.
Sub Main2()
    Dim FSO As New Scripting.FileSystemObject
    Dim sin As Scripting.TextStream
    Dim sout As Scripting.TextStream
    Dim strWord As String
     
    Set sin = FSO.GetStandardStream(StdIn)
    Set sout = FSO.GetStandardStream(StdOut)
    sout.WriteLine "Hello!"
    sout.WriteLine "What's the word?"
    strWord = sin.ReadLine()
    sout.WriteLine "So, the word is " & strWord
    Set sout = Nothing
    Set sin = Nothing
End Sub
