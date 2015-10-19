Sub checkName() '数据字段匹配'
    Dim i, j, k, l, trow, statRows As Integer
    Dim flag As Boolean
    
    statRows = Sheets("原材料").Range("A65536").End(xlUp).Row  '第一列行数 用以遍历匹配'

    For i = 1 To Sheets("每日合计").Index - 1   '遍历个人详细表'
        trow = Sheets(i).Range("a1").CurrentRegion.Rows.Count   '获取当前需要计算的表 行'
        For j = 2 To trow           '遍历行'
            flag = True
            Sheets(i).Cells(j, 2).Font.Color = 0
            Sheets(i).Cells(j, 3).Font.Color = 0
            For k = 2 To statRows       '遍历统计表 匹配 计数'
                If (Sheets(i).Cells(j, 2) = Sheets("原材料").Cells(k, 1) And Sheets(i).Cells(j, 3) = Sheets("原材料").Cells(k, 2)) Then  '匹配产品名和规格'
                    flag = False
                End If
            Next
            If flag Then '没有找到匹配的产品规格或者尺寸'
                Sheets(i).Cells(j, 2).Font.Color = RGB(255, 0, 0)
                Sheets(i).Cells(j, 3).Font.Color = RGB(255, 0, 0)
            End If
        Next
    Next
End Sub

