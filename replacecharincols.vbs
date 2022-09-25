Sub GoToLastColumnofSpreadsheet()

    Dim i As Long
    Dim finalColumn As Long
    
    finalColumn = ActiveSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    For i = 1 To finalColumn
    
       Sheet1.Cells(1, i).Value = Replace(Sheet1.Cells(1, i).Value, "@", "_")
    
    Next i

End Sub
