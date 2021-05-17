Attribute VB_Name = "ArrDemo3"
Sub ArrDemo3()
Dim MyArray3(4) As Variant '宣告一個內含5個元素的陣列(索引值0-4)為萬用
MyArray3(0) = "疫苗" '第一個元素(索引值是0)
MyArray3(1) = 200 '第二個元素(索引值是1)
MyArray3(2) = 3.14159 '第三個元素(索引值是2)
MyArray3(3) = CBool(1) '第四個元素(索引值是3)
MyArray3(4) = 500 '第五個元素(索引值是4)


MsgBox (UBound(MyArray3)) '彈出視窗標示陣列上限
MsgBox (LBound(MyArray3)) '彈出視窗標示陣列下限
End Sub

