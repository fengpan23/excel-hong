Sub stats() '多维 多条件 汇总统计'
Dim i, j, k, l, trow, statRows, statLines As Integer
Dim temp As Double

Dim sa() As String

statRows = Sheets(Sheets.Count).Range("a1").CurrentRegion.Rows.Count  '统计表 总行'
statLines = Sheets(Sheets.Count).UsedRange.Columns.Count '统计表 总列'
Sheets(Sheets.Count).Range("B3:IV65536").Clear  '清空表格数据'

For i = 1 To Sheets.Count - 1   '遍历需要统计的表'
    trow = Sheets(i).Range("a1").CurrentRegion.Rows.Count   '获取当前需要计算的表 行'
    For j = 2 To trow           '遍历行'
        For k = 3 To statRows       '遍历统计表 匹配 计数'
            If (Sheets(i).Cells(j, 1) = Sheets(Sheets.Count).Cells(k, 1)) Then
                For l = 1 To statLines
                    If (Sheets(i).Cells(j, 2) = Sheets(Sheets.Count).Cells(1, l) And Sheets(i).Cells(j, 3) = Sheets(Sheets.Count).Cells(2, l)) Then
                        temp = Sheets(Sheets.Count).Cells(k, l)
                        Sheets(Sheets.Count).Cells(k, l) = Sheets(i).Cells(j, 4) + temp
                    End If
                Next
            End If
        Next
    Next
Next
End Sub