Attribute VB_Name = "Module1"
Sub �Ѥp��j()
Attribute �Ѥp��j.VB_Description = "�Ѥp��j"
Attribute �Ѥp��j.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' �Ѥp��j ����
' �Ѥp��j
'
' �ֳt��: Ctrl+a
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add2 Key:=Range("B2:B50"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
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
Sub �Ѥj��p()
Attribute �Ѥj��p.VB_Description = "�Ѥj��p"
Attribute �Ѥj��p.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' �Ѥj��p ����
' �Ѥj��p
'
' �ֳt��: Ctrl+w
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add2 Key:=Range("B2:B50"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
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
