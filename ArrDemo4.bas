Attribute VB_Name = "ArrDemo4"
Sub ArrDemo4()
Dim MyArray4(4) As Variant '�ŧi�@�Ӥ��t5�Ӥ������}�C(���ޭ�0-4)���U��
MyArray4(0) = "�̭]" '�Ĥ@�Ӥ���(���ޭȬO0)
MyArray4(1) = 200 '�ĤG�Ӥ���(���ޭȬO1)
MyArray4(2) = 3.14159 '�ĤT�Ӥ���(���ޭȬO2)
MyArray4(3) = CBool(1) '�ĥ|�Ӥ���(���ޭȬO3)
MyArray4(4) = 500 '�Ĥ��Ӥ���(���ޭȬO4)

Dim arrIdx As Integer
For arrIdx = LBound(MyArray4) To UBound(MyArray4)
    Dim varStr4 As String '�ŧi�}�C�������A�W���ܼƬ���r
    varStr4 = TypeName(MyArray4(arrIdx)) '�ĥ|�Ӱ}�C����(���ޭ�=3)�Ȫ���ƫ��O
    MsgBox ("��" & arrIdx & "������ƫ��A = " & varStr4) '�u�X�����Х�
Next
End Sub

