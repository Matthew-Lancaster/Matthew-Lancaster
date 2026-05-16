Attribute VB_Name = "modMain"
Option Explicit

Global StrConnect As String

Public Sub Main()

    Load MDIProcServ
    MDIProcServ.Show

End Sub

Public Sub Finish()

    Unload MDIProcServ
    End
    
End Sub
