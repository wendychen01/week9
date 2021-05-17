Attribute VB_Name = "合併巨集北中南區"
Sub 合併巨集()
Attribute 合併巨集.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 合併巨集 巨集
    Sheets("北區").Select
    Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("總店").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    
    Sheets("中區").Select
    Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("總店").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    
    Sheets("南區").Select
    Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("總店").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    
End Sub
