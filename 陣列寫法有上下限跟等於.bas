Attribute VB_Name = "陣列寫法有上下限跟等於"
Sub 陣列寫法()
Dim arr() '宣告陣列
Dim arrIdx As Integer '宣告陣列索引
Dim tagetValueUB, tagetValueLB, tagetValueB As Integer

tagetValueUB = CInt(InputBox("請輸入標記上限值(0-1000)")) '因為InputBox輸入為文字，故要轉型為整數
tagetValueLB = CInt(InputBox("請輸入標記下限值(0-1000)")) '因為InputBox輸入為文字，故要轉型為整數
tagetValueB = CInt(InputBox("請輸入標記等於值(0-1000)")) '因為InputBox輸入為文字，故要轉型為整數

Columns("b:b").Interior.Color = xlNone 'b欄位全部還原成無顏色
'B3欄位最後一列 = Cells(3,"B").End(xlDown)

arr = Range(Cells(3, "B"), Cells(3, "B").End(xlDown)) '陣列範圍從B3到B欄最後
For arrIdx = 1 To UBound(arr, 1) 'UBound = 陣列上限
    If arr(arrIdx, 1) > tagetValueUB Then '陣列第i個元素值 > 目標值
        Cells(arrIdx + 2, "B").Interior.Color = vbCyan '註記為天藍色
        
    ElseIf arr(arrIdx, 1) < tagetValueLB Then '陣列第i個元素值 < 目標值
        Cells(arrIdx + 2, "B").Interior.Color = vbBlue '註記為藍色
        
    ElseIf arr(arrIdx, 1) = tagetValueLB Then '陣列第i個元素值 = 目標值
        Cells(arrIdx + 2, "B").Interior.Color = vbRed '註記為紅色
    End If
Next

Range("a1").CurrentRegion.Borders.LineStyle = xlContinuous
End Sub
