Attribute VB_Name = "ArrDemo2"
Sub ArrDemo2()
Dim MyArray2(4) As Variant '�ŧi�@�Ӥ��t5�Ӥ������}�C(���ޭ�0-4)���U��
MyArray2(0) = "�̭]" '�Ĥ@�Ӥ���(���ޭȬO0)
MyArray2(1) = 200 '�ĤG�Ӥ���(���ޭȬO1)
MyArray2(2) = 3.14159 '�ĤT�Ӥ���(���ޭȬO2)
MyArray2(3) = CBool(1) '�ĥ|�Ӥ���(���ޭȬO3)
MyArray2(4) = 500 '�Ĥ��Ӥ���(���ޭȬO4)

Dim varStr2 As String '�ŧi�}�C�������A�W���ܼƬ���r
varStr2 = TypeName(MyArray2(3)) '�ĥ|�Ӱ}�C����(���ޭ�=3)�Ȫ���ƫ��O
MsgBox (varStr2) '�u�X�����Хܲĥ|�Ӱ}�C�������A�W��
End Sub

