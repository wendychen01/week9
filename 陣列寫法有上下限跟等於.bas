Attribute VB_Name = "�}�C�g�k���W�U���򵥩�"
Sub �}�C�g�k()
Dim arr() '�ŧi�}�C
Dim arrIdx As Integer '�ŧi�}�C����
Dim tagetValueUB, tagetValueLB, tagetValueB As Integer

tagetValueUB = CInt(InputBox("�п�J�аO�W����(0-1000)")) '�]��InputBox��J����r�A�G�n�૬�����
tagetValueLB = CInt(InputBox("�п�J�аO�U����(0-1000)")) '�]��InputBox��J����r�A�G�n�૬�����
tagetValueB = CInt(InputBox("�п�J�аO�����(0-1000)")) '�]��InputBox��J����r�A�G�n�૬�����

Columns("b:b").Interior.Color = xlNone 'b�������٭즨�L�C��
'B3���̫�@�C = Cells(3,"B").End(xlDown)

arr = Range(Cells(3, "B"), Cells(3, "B").End(xlDown)) '�}�C�d��qB3��B��̫�
For arrIdx = 1 To UBound(arr, 1) 'UBound = �}�C�W��
    If arr(arrIdx, 1) > tagetValueUB Then '�}�C��i�Ӥ����� > �ؼЭ�
        Cells(arrIdx + 2, "B").Interior.Color = vbCyan '���O�����Ŧ�
        
    ElseIf arr(arrIdx, 1) < tagetValueLB Then '�}�C��i�Ӥ����� < �ؼЭ�
        Cells(arrIdx + 2, "B").Interior.Color = vbBlue '���O���Ŧ�
        
    ElseIf arr(arrIdx, 1) = tagetValueLB Then '�}�C��i�Ӥ����� = �ؼЭ�
        Cells(arrIdx + 2, "B").Interior.Color = vbRed '���O������
    End If
Next

Range("a1").CurrentRegion.Borders.LineStyle = xlContinuous
End Sub
