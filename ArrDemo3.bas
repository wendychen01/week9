Attribute VB_Name = "ArrDemo3"
Sub ArrDemo3()
Dim MyArray3(4) As Variant '�ŧi�@�Ӥ��t5�Ӥ������}�C(���ޭ�0-4)���U��
MyArray3(0) = "�̭]" '�Ĥ@�Ӥ���(���ޭȬO0)
MyArray3(1) = 200 '�ĤG�Ӥ���(���ޭȬO1)
MyArray3(2) = 3.14159 '�ĤT�Ӥ���(���ޭȬO2)
MyArray3(3) = CBool(1) '�ĥ|�Ӥ���(���ޭȬO3)
MyArray3(4) = 500 '�Ĥ��Ӥ���(���ޭȬO4)


MsgBox (UBound(MyArray3)) '�u�X�����Хܰ}�C�W��
MsgBox (LBound(MyArray3)) '�u�X�����Хܰ}�C�U��
End Sub

