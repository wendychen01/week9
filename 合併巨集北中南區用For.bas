Attribute VB_Name = "�X�֥����_���n�ϥ�For"
Sub �X�֥���2()
Dim shtIdx As Integer
For shtIdx = 1 To Sheets.Count - 1 'Sheeys Count-1 �O�]���`��ܾ�X���u�@��
    Sheets(shtIdx).Select
    Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("�`��").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
Next
End Sub

