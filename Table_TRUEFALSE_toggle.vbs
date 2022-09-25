Private Sub worksheet_selectionchange(ByVal target As Range)
    
    Dim tb As ListObject
    Dim Currcell, i As Integer
    
    Set tb = Worksheets("Sheet1").ListObjects("Table1")
    
    ' Refer to columns of the table like so:
    If Not Intersect(tb.DataBodyRange.Columns("B:L"), Selection) Is Nothing Then
        For Each cell In Intersect(tb.DataBodyRange.Columns("B:L"), Selection).Cells
            
              ' Flip cell value & color on click
              If ActiveCell.Value = True Then
              ActiveCell.Value = False
              ActiveCell.Interior.Color = RGB(206, 32, 41)
              ElseIf ActiveCell.Value = False Then
                    ActiveCell.Value = True
                    ActiveCell.Interior.Color = RGB(0, 255, 64)
              End If
              
        Next cell
    End If
    
    'To refer to other columns of the table, add another code block and change the columns reference, like so :
    If Not Intersect(tb.DataBodyRange.Columns("S:U"), Selection) Is Nothing Then
        For Each cell In Intersect(tb.DataBodyRange.Columns("S:U"), Selection).Cells
            
              ' Flip cell value & color on click
              If ActiveCell.Value = True Then
              ActiveCell.Value = False
              ActiveCell.Interior.Color = RGB(206, 32, 41)
              ElseIf ActiveCell.Value = False Then
                    ActiveCell.Value = True
                    ActiveCell.Interior.Color = RGB(0, 255, 64)
              End If
              
        Next cell
    End If


End Sub
