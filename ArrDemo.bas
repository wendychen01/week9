Attribute VB_Name = "ArrDemo"
Sub ArrDemo()
Dim MyArray(4) As Integer '�ŧi�@�Ӥ��t5�Ӥ������}�C(���ޭ�0-4)
MyArray(0) = 100 '�Ĥ@�Ӥ���(���ޭȬO0)
MyArray(1) = 200 '�ĤG�Ӥ���(���ޭȬO1)
MyArray(2) = 300 '�ĤT�Ӥ���(���ޭȬO2)
MyArray(3) = 400 '�ĥ|�Ӥ���(���ޭȬO3)
MyArray(4) = 500 '�Ĥ��Ӥ���(���ޭȬO4)

Dim varStr As String '�ŧi�}�C�������A�W���ܼƬ���r
varStr = TypeName(MyArray(0)) '�Ĥ@�Ӱ}�C����(���ޭ�=0)�Ȫ���ƫ��O
MsgBox (varStr) '�u�X�����ХܲĤ@�Ӱ}�C�������A�W��
End Sub
