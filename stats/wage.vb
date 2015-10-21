Private Sub wage()
    Dim i, j, trow As Integer
    Dim mStr() As String
    Dim tsize, money, con, temp, temp1 As Double

    For i = 1 To Sheets("每日合计").index - 1   '遍历需要计算表'
        money = 0
        trow = Sheets(i).Range("a1").CurrentRegion.Rows.Count   '获取当前需要计算的表 行'

        Sheets(i).Cells(1, 10) = "计件工资结算"
        For j = 2 To trow
            mStr = Split(Sheets(i).Cells(j, 3), "*")
            If UBound(mStr) - LBound(mStr) + 1 = 2 Then
                tsize = Val(mStr(0)) * Val(mStr(1))
                temp = 0
                temp1 = money
    
                con = Sheets(i).Cells(j, 4) * 10000
                If tsize > (33 * 40) Then
                    If con > 35000 Then
                        temp = (con - 35000) * 0.003 + 28.5 * 3.5
                    Else
                        temp = con * 0.00285
                    End If
                Else
                    If con > 42000 Then
                        temp = (con - 42000) * 0.0026 + 24 * 4.2
                    Else
                        temp = con * 0.0024
                    End If
                End If
                Sheets(i).Cells(j, 10) = temp
                money = temp1 + temp
            End If
        Next
        'Sheets(i).Cells(trow + 1, 10) = money'
    Next
End Sub