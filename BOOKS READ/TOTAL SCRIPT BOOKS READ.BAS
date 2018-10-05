Attribute VB_Name = "Module1"
Sub MAIN()


FR1 = FreeFile
Open "C:\SCRIPTER\BOOKS READ\TOTAL SCRIPT BOOKS READ.txt" For Input As FR1
Do
    Line Input #FR1, LINE_1
    If Mid(LINE_1, 1, 7) = "Count =" Then
        Line Input #FR1, LINE_1
        Line Input #FR1, LINE_1
        Line Input #FR1, LINE_1
        Counter = Counter + 1
        Debug.Print Format(Counter, "000") + " -- " + LINE_1
    End If
Loop Until EOF(FR1)
Close FR1


End Sub

