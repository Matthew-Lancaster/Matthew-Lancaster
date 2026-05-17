Attribute VB_Name = "modMain"
Option Explicit
Public appWorld As Excel.Application
Public wbWorld As Excel.Workbook
'
' Public shtWorld As Excel.Worksheet

Sub Setup()
    ' IMPORTANT: If your machine does not have Excel 97 installed,
    ' you must change the reference to the Excel 95 Object Library.
    ' Then, in the Declarations section above, change the variable
    ' declaration "wbWorld as Workbook" to "shtWorld As Worksheet."
    ' Then change all references to "wbWorld" to "shtWorld."

    On Error Resume Next 'ignore errors
    Set appWorld = GetObject(, "Excel.Application") 'look for a running copy of Excel
    If Err.Number <> 0 Then 'If Excel is not running then
        Set appWorld = CreateObject("Excel.Application") 'run it
    End If
    Err.Clear   ' Clear Err object in case error occurred.
    
    On Error GoTo 0 'Resume normal error processing

    Set wbWorld = appWorld.Workbooks.Open(App.Path & "\world.xls")
End Sub

' Set the objects to Nothing.
Sub CleanUp()
    ' This should force an unload of Microsoft Excel,
    ' providing no other applications or users have it loaded.
    wbWorld.Close
    Set appWorld = Nothing
    Set wbWorld = Nothing
End Sub

' Fill the Continents combo box with the names
' of the sheets in the workbook.
Sub FillContinentsList()
    Dim shtContinent As Excel.Worksheet
    
    ' Iterate through the collection of sheets and add
    ' the name of each sheet to the combo box.
    For Each shtContinent In wbWorld.Sheets
        frmMain.cmbContinents.AddItem shtContinent.Name
    Next
    ' Select the first item and display it in the combo box.
    frmMain.cmbContinents.Text = frmMain.cmbContinents.List(0)

    Set shtContinent = Nothing
End Sub

' Fill the Continents combo box with the names
' of the features corresponding to a given continent.
Sub FillFeaturesList()
    Dim shtContinent As Excel.Worksheet
    Dim rngFeatureList As Excel.Range
    Dim intFirstBlankCell As Integer
    Dim loop1 As Integer

    ' Hide the old ranking list.
    frmMain.lstTopRanking.Visible = False
    
    ' Get the sheet with the name of the continent selected in the Continents combo box.
    Set shtContinent = wbWorld.Sheets(frmMain.cmbContinents.Text)
    ' Assign the first row of this sheet to an object.
    Set rngFeatureList = shtContinent.Rows(1)
    
    ' See if it's an empty list.
    If (rngFeatureList.Cells(1, 1) = "") Then
        intFirstBlankCell = 0
    Else
        ' Search the row for the first blank cell.
        intFirstBlankCell = rngFeatureList.Find("").Column
    End If
    
    ' Empty the previous contents of the features combo box.
    frmMain.cmbFeatures.Clear
            
    ' Add the items to the features combo box.
    For loop1 = 1 To intFirstBlankCell
            frmMain.cmbFeatures.AddItem rngFeatureList.Cells(1, loop1)
    Next
    
    ' Select the first item and display it in the combo box.
    frmMain.cmbFeatures.Text = frmMain.cmbFeatures.List(0)

    ' Clean up.
    Set shtContinent = Nothing
    Set rngFeatureList = Nothing
End Sub

' Fill the list of ranking items.
Sub FillTopRankingList()
    Dim shtContinent As Excel.Worksheet
    Dim intColumOfFeature As Integer
    Dim rngRankedList As Excel.Range
    Dim intFirstBlankCell As Integer
    Dim loop1 As Integer
    
    ' Get the sheet with the name of the continent selected in the Continents combo box.
    Set shtContinent = wbWorld.Sheets(frmMain.cmbContinents.Text)
    
    ' Empty the previous contents of the ranking list box.
    frmMain.lstTopRanking.Clear
    
    ' If the feature selection is blank, do nothing.
    If (frmMain.cmbFeatures <> "") Then
        
        ' Look up the column of the selected feature in the first row of the spreadsheet.
        intColumOfFeature = shtContinent.Rows(1).Find(frmMain.cmbFeatures.Text).Column
        
        ' Assign the column to an object.
         Set rngRankedList = shtContinent.Columns(intColumOfFeature)
        
        ' See if it's a blank list.
        If (rngRankedList.Cells(1, 1) = "") Then
            intFirstBlankCell = 0
        Else
            ' Search the row for the first blank cell.
            intFirstBlankCell = rngRankedList.Find("").Row
        End If
                
        ' Add the items to the TopRanking ListBox.
        For loop1 = 2 To intFirstBlankCell
            frmMain.lstTopRanking.AddItem rngRankedList.Cells(loop1, 1)
        Next
    
        ' Show the new ranking list.
        frmMain.lstTopRanking.Visible = True
    
    End If
    
    ' Clean up.
    Set shtContinent = Nothing
    Set rngRankedList = Nothing
End Sub
