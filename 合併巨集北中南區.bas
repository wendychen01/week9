Attribute VB_Name = "�X�֥����_���n��"
Sub �X�֥���()
Attribute �X�֥���.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �X�֥��� ����
    Sheets("�_��").Select
    Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("�`��").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    
    Sheets("����").Select
    Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("�`��").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    
    Sheets("�n��").Select
    Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("�`��").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    
End Sub
