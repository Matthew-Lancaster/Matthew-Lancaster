Attribute VB_Name = "RestarterMod"
Public Sub ReRunner()

If ReRun = True Then
    Activated = False

    Reset
    Load Ci
    'Ci.Show
    ReRun = False
End If

End Sub
