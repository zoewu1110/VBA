Attribute VB_Name = "Module1"
Sub 由小到大()
Attribute 由小到大.VB_Description = "由小到大"
Attribute 由小到大.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' 由小到大 巨集
' 由小到大
'
' 快速鍵: Ctrl+a
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B50"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B50")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[49]C[-3])"
    Range("E2").Select
    ActiveWindow.SmallScroll Down:=-33
    Range("G1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=MAX(R[1]C[-5]:R[49]C[-5])"
    Range("G2").Select
    ActiveWindow.SmallScroll Down:=-36
    Range("I1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=MIN(R[1]C[-7]:R[49]C[-7])"
    Range("I2").Select
    ActiveWindow.SmallScroll Down:=-18
End Sub
Sub 由大到小()
Attribute 由大到小.VB_Description = "由大到小"
Attribute 由大到小.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' 由大到小 巨集
' 由大到小
'
' 快速鍵: Ctrl+w
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B50"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B50")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[49]C[-3])"
    Range("E2").Select
    ActiveWindow.SmallScroll Down:=-18
    Range("G1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=MAX(R[1]C[-5]:R[49]C[-5])"
    Range("G2").Select
    ActiveWindow.SmallScroll Down:=-21
    Range("I1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=MIN(R[1]C[-7]:R[49]C[-7])"
    Range("I2").Select
End Sub
