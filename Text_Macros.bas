Attribute VB_Name = "Text_Macros"
Sub Uppercase()

    Dim cel As Range
    Dim selectedRange As Range

    Set selectedRange = Application.Selection

   ' Loop to cycle through each cell in the specified range.
   For Each cel In selectedRange.Cells
      ' Change the text in the range to uppercase letters.
      cel.Value = UCase(cel.Value)
   Next
End Sub
Sub Lowercase()

    Dim cel As Range
    Dim selectedRange As Range

    Set selectedRange = Application.Selection

   ' Loop to cycle through each cell in the specified range.
   For Each cel In selectedRange.Cells
      ' Change the text in the range to uppercase letters.
      cel.Value = LCase(cel.Value)
   Next
End Sub
