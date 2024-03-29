Sub FindBySubj()

'目的：在sheet1（filterTab）的ABC列筛选任意n个受试者，contents之后所有的sheets自动显示这n个受试者的筛选结果

'Step1:在sheet1的C列的筛选结果放进array，需要先定义arr的元素个数，否则会溢出

Dim arr(2000)
Dim a As Integer
Dim b As Integer


'a为数组arr中元素的个数
'b为筛选的C列中最后一个筛选结果所在的行数
a = 0
b = Cells(Rows.Count, "C").End(xlUp).Row

'如果b=2那么说明仅筛选了C2单元格的受试者，所以a=1
If b = 2 Then
    a = 1
    arr(1) = Range("C2")
    
End If

'如果b <> 2那么把筛选的单元格放进array里。包含两种情况1.仅筛选了一个subj，且不是C2单元格的受试者；2.筛选了不止一个受试者
If b <> 2 Then

    For Each Rng In Sheet1.Range("C2:C" & b).Columns("C:C").SpecialCells(xlCellTypeVisible)
        a = a + 1
        C = Rng.Row
        arr(a) = Range("C" & C)
    Next

    
End If


'Step2:AE之后的受试者号均在C列(field=3)。
'注：Output listing时没有观测的domain的A2单元格内容为No Record Found，需要if判断

Dim num As Integer
Dim i As Integer

'取当前工作簿中工作表的个数
num = ThisWorkbook.Sheets.Count

For i = 3 To num

    If Sheets(i).Range("A2") <> "No Record Found" Then
        Sheets(i).Range("A1").AutoFilter field:=3, Criteria1:=arr, Operator:=xlFilterValues
    End If
    
Next



End Sub

Sub clear()

'目的：清除筛选结果

Dim i As Integer
Dim num As Integer

num = ThisWorkbook.Sheets.Count

For i = 3 To num

    If Sheets(i).Range("A2") <> "No Record Found" Then
        Sheets(i).Range("A1").AutoFilter
    End If
Next

End Sub

