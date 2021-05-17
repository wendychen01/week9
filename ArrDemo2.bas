Attribute VB_Name = "ArrDemo2"
Sub ArrDemo2()
Dim MyArray2(4) As Variant '宣告一個內含5個元素的陣列(索引值0-4)為萬用
MyArray2(0) = "疫苗" '第一個元素(索引值是0)
MyArray2(1) = 200 '第二個元素(索引值是1)
MyArray2(2) = 3.14159 '第三個元素(索引值是2)
MyArray2(3) = CBool(1) '第四個元素(索引值是3)
MyArray2(4) = 500 '第五個元素(索引值是4)

Dim varStr2 As String '宣告陣列元素型態名稱變數為文字
varStr2 = TypeName(MyArray2(3)) '第四個陣列元素(索引值=3)值的資料型別
MsgBox (varStr2) '彈出視窗標示第四個陣列元素型態名稱
End Sub

