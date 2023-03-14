Sub clearTitle()
Dim i As Integer
Dim num As Integer
num = ThisWorkbook.Sheets.Count
'delete title from sheet3 to sheet_end
For i = 3 To num
Sheets(i).Rows("1:2").Delete Shift:=xlUp
Next
End Sub