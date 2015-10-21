Private Sub checkName() '数据字段匹配'
    Dim i, j, k, l, trow, statRows, indexA As Integer
    Dim flag As Boolean
    
    indexA = Sheets("标准表").index
    
    statRows = Sheets(indexA).Range("A65536").End(xlUp).Row  '第一列行数 用以遍历匹配'

    For i = 1 To Sheets("每日合计").index - 1   '遍历个人详细表'
        trow = Sheets(i).Range("a1").CurrentRegion.Rows.Count   '获取当前需要计算的表 行'
        For j = 2 To trow           '遍历行'
            flag = True
            Sheets(i).Cells(j, 2).Font.Color = 0
            Sheets(i).Cells(j, 3).Font.Color = 0
            
            For k = 2 To statRows       '遍历统计表 匹配 计数'
                If (Sheets(i).Cells(j, 2) = Sheets(indexA).Cells(k, 1) And Sheets(i).Cells(j, 3) = Sheets(indexA).Cells(k, 2)) Then  '匹配产品名和规格'
                    flag = False
                End If
            Next

            If flag Then '没有找到匹配的产品规格或者尺寸'
                Sheets(i).Cells(j, 2).Font.Color = RGB(255, 0, 0)
                Sheets(i).Cells(j, 3).Font.Color = RGB(255, 0, 0)
            End If
        Next
    Next
    
    trow = Sheets("成品出库").Range("a1").CurrentRegion.Rows.Count   '检测成品出库表  获取表 行'
    For j = 2 To trow           '遍历行'
        flag = True
        Sheets("成品出库").Cells(j, 2).Font.Color = 0
        Sheets("成品出库").Cells(j, 3).Font.Color = 0
        
        For k = 2 To statRows       '遍历统计表 匹配 计数'
            If (Sheets("成品出库").Cells(j, 2) = Sheets(indexA).Cells(k, 1) And Sheets("成品出库").Cells(j, 3) = Sheets(indexA).Cells(k, 2)) Then  '匹配产品名和规格'
                flag = False
            End If
        Next

        If flag Then '没有找到匹配的产品规格或者尺寸'
            Sheets("成品出库").Cells(j, 2).Font.Color = RGB(255, 0, 0)
            Sheets("成品出库").Cells(j, 3).Font.Color = RGB(255, 0, 0)
        End If
    Next
    
    trow = Sheets("库存量").Range("a1").CurrentRegion.Rows.Count   '检测库存量表  获取表 行'
    For j = 2 To trow           '遍历行'
        flag = True
        Sheets("库存量").Cells(j, 1).Font.Color = 0
        Sheets("库存量").Cells(j, 2).Font.Color = 0
        
        For k = 2 To statRows       '遍历统计表 匹配 计数'
            If (Sheets("库存量").Cells(j, 1) = Sheets(indexA).Cells(k, 1) And Sheets("库存量").Cells(j, 2) = Sheets(indexA).Cells(k, 2)) Then  '匹配产品名和规格'
                flag = False
            End If
        Next

        If flag Then '没有找到匹配的产品规格或者尺寸'
            Sheets("库存量").Cells(j, 1).Font.Color = RGB(255, 0, 0)
            Sheets("库存量").Cells(j, 2).Font.Color = RGB(255, 0, 0)
        End If
    Next
End Sub

Private Sub ruku() '入库'
    Dim i, j, trow, statLines As Integer

    trow = Sheets("库存量").Range("a1").CurrentRegion.Rows.Count   '检测库存量表  获取表 行'
    statLines = Sheets("每日合计").UsedRange.Columns.Count '统计表 总列'

    For i = 2 To trow
        For j = 2 To statLines
            if(Sheets("每日合计").Cells(1, j) = Sheets("库存量").Cells(i, 1) And Sheets("每日合计").Cells(2, j) = Sheets("库存量").Cells(i, 2)) Then
                Sheets("库存量").Cells(i, 3) = Sheets("每日合计").Cells(3, j)
            End if
        Next
    Next
End Sub