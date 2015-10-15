'自动生成 日期 产品规格'
Sub stats() '多维 多条件 汇总统计'
    Dim i, j, k, l, trow, statRows, statLines As Integer
    Dim dateFlag, pudFlag As Boolean,
    Dim temp As Double

    Sheets("每日合计").Range("A65536").Clear  '清空统计表 数据'

    For i = 1 To Sheets("每日合计").Index - 1   '遍历需要统计的表'
        trow = Sheets(i).Range("a1").CurrentRegion.Rows.Count   '获取当前需要计算的表 行'
        For j = 2 To trow           '遍历行'
            statRows = Sheets("每日合计").Range("A65536").End(xlUp).Row  '统计表 第一列行数 用以遍历以及清楚表格数据'
            dateFlag = True;
            For k = 3 To statRows       '遍历统计表 匹配 计数'
                If (Sheets(i).Cells(j, 1) = Sheets("每日合计").Cells(k, 1)) Then  '匹配日期'
                    dateFlag ＝ False
                    pudFlag ＝ True
                    For l = 2 To statLines  '遍历统计表列  匹配 产品名和规格'
                        If (Sheets(i).Cells(j, 2) = Sheets("每日合计").Cells(1, l) And Sheets(i).Cells(j, 3) = Sheets("每日合计").Cells(2, l)) Then
                            pudFlag ＝ False
                            temp = Sheets("每日合计").Cells(k, l)
                            Sheets("每日合计").Cells(k, l) = Sheets(i).Cells(j, 4) + temp
                        End If
                    Next

                    If pudFlag Then
                        Sheets("每日合计").Cells(1, l) ＝ Sheets(i).Cells(j, 2)  '赋值产品名'
                        Sheets("每日合计").Cells(2, l) ＝ Sheets(i).Cells(j, 3)  '赋值产品规格'
                        temp = Sheets("每日合计").Cells(k, l)
                        Sheets("每日合计").Cells(k, l) = Sheets(i).Cells(j, 4) + temp
                    End If
                End If
            Next

            If dateFlag Then '统计表中没找到对应日期 则 添加该日期'
                Sheets("每日合计").Cells(statRows, 1) = Sheets(i).Cells(j, 1) '赋值时间'
                pudFlag ＝ True
                For l = 2 To statLines  '遍历统计表列  匹配 产品名和规格'
                    If (Sheets(i).Cells(j, 2) = Sheets("每日合计").Cells(1, l) And Sheets(i).Cells(j, 3) = Sheets("每日合计").Cells(2, l)) Then
                        pudFlag ＝ False
                        temp = Sheets("每日合计").Cells(k, l)
                        Sheets("每日合计").Cells(k, l) = Sheets(i).Cells(j, 4) + temp
                    End If
                Next
                If pudFlag Then
                    Sheets("每日合计").Cells(1, l) ＝ Sheets(i).Cells(j, 2)  '赋值产品名'
                    Sheets("每日合计").Cells(2, l) ＝ Sheets(i).Cells(j, 3)  '赋值产品规格'
                    temp = Sheets("每日合计").Cells(k, l)
                    Sheets("每日合计").Cells(k, l) = Sheets(i).Cells(j, 4) + temp
                End If
            End If
        Next
    Next

    Sheets("每日合计").Range("A3").Sort Key1:=Sheets("每日合计").Range("A1")
End Sub