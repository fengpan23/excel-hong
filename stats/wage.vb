Private Sub wage()
	Dim i, j, trow As Integer
	Dim mStr() As String
	Dim tsize, money, con, temp As Double

    For i = 1 To Sheets("每日合计").Index - 1   '遍历需要计算表'
        money = 0
    	trow = Sheets(i).Range("a1").CurrentRegion.Rows.Count   '获取当前需要计算的表 行'
    	For j = 2 To trow
    		mStr = Split(Sheets(i).cells(j, 3), "*")
    		tsize = Val(sa(0)) * Val(sa(1))
            temp ＝ money
            con = Sheets(i).cells(j, 4) * 10000
    		If tsize > (33 * 40) Then
    			If con > 35000 Then
    				money = temp + (con - 35000) * 0.003 + 28.5 * 3.5
                Else
                    money = temp + con * 0.00285
                End If
    		Else
                If con > 42000 Then
                    money = temp + (con - 42000) * 0.0026 + 24 * 4.2
                Else
                    money = temp + con * 0.0024
                End If
    		End If 
    	Next
    Next
End Sub

'33*40以下的规格 产量≤4.2万  24元/万  超出部分 26元/万
33*40以上的规格 产量≤3.5万  28.5元/万  超出部分 30元/万'