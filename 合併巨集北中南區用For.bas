Attribute VB_Name = "合併巨集北中南區用For"
Sub 合併巨集2()
Dim shtIdx As Integer
For shtIdx = 1 To Sheets.Count - 1 'Sheeys Count-1 是因為總表示整合的工作表
    Sheets(shtIdx).Select
    Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Sheets("總店").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
Next
End Sub

