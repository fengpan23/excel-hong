Sub stats() '多维 多条件 汇总统计'
Dim i, j, k, l, trow, statRows, statLines As Integer
Dim temp As Double

statRows = Sheets("每日合计").Range("a1").CurrentRegion.Rows.Count  '统计表 总行'
statLines = Sheets("每日合计").UsedRange.Columns.Count '统计表 总列'
Sheets("每日合计").Range(Sheets("每日合计").Cells(3, 2), Sheets("每日合计").Cells(statRows, statLines)).Clear  '清空统计表 数据'

For i = 1 To Sheets("每日合计").Index - 1   '遍历需要统计的表'
    trow = Sheets(i).Range("a1").CurrentRegion.Rows.Count   '获取当前需要计算的表 行'
    For j = 2 To trow           '遍历行'
        For k = 3 To statRows       '遍历统计表 匹配 计数'
            If (Sheets(i).Cells(j, 1) = Sheets("每日合计").Cells(k, 1)) Then
                For l = 2 To statLines
                    If (Sheets(i).Cells(j, 2) = Sheets("每日合计").Cells(1, l) And Sheets(i).Cells(j, 3) = Sheets("每日合计").Cells(2, l)) Then
                        temp = Sheets("每日合计").Cells(k, l)
                        Sheets("每日合计").Cells(k, l) = Sheets(i).Cells(j, 4) + temp
                    End If
                Next
            End If
        Next
    Next
Next
End Sub