Attribute VB_Name = "ArrDemo4"
Sub ArrDemo4()
Dim MyArray4(4) As Variant '宣告一個內含5個元素的陣列(索引值0-4)為萬用
MyArray4(0) = "疫苗" '第一個元素(索引值是0)
MyArray4(1) = 200 '第二個元素(索引值是1)
MyArray4(2) = 3.14159 '第三個元素(索引值是2)
MyArray4(3) = CBool(1) '第四個元素(索引值是3)
MyArray4(4) = 500 '第五個元素(索引值是4)

Dim arrIdx As Integer
For arrIdx = LBound(MyArray4) To UBound(MyArray4)
    Dim varStr4 As String '宣告陣列元素型態名稱變數為文字
    varStr4 = TypeName(MyArray4(arrIdx)) '第四個陣列元素(索引值=3)值的資料型別
    MsgBox ("第" & arrIdx & "元素資料型態 = " & varStr4) '彈出視窗標示
Next
End Sub

