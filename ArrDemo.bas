Attribute VB_Name = "ArrDemo"
Sub ArrDemo()
Dim MyArray(4) As Integer '宣告一個內含5個元素的陣列(索引值0-4)
MyArray(0) = 100 '第一個元素(索引值是0)
MyArray(1) = 200 '第二個元素(索引值是1)
MyArray(2) = 300 '第三個元素(索引值是2)
MyArray(3) = 400 '第四個元素(索引值是3)
MyArray(4) = 500 '第五個元素(索引值是4)

Dim varStr As String '宣告陣列元素型態名稱變數為文字
varStr = TypeName(MyArray(0)) '第一個陣列元素(索引值=0)值的資料型別
MsgBox (varStr) '彈出視窗標示第一個陣列元素型態名稱
End Sub
