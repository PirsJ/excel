Sub Tri_Poldat()
Attribute Tri_Poldat.VB_ProcData.VB_Invoke_Func = " \n14"

    ' Corrigé : la borne "205471" était figée (héritée d'un enregistrement de
    ' macro passé). "Poldat" ne fait plus que ~126 000 lignes aujourd'hui, mais
    ' le jour où elle dépassera 205 471 lignes, les lignes en trop auraient été
    ' silencieusement exclues du tri. Bornage recalculé dynamiquement.
    Dim wsPoldat As Worksheet
    Dim lastRow As Long
    Set wsPoldat = ActiveWorkbook.Worksheets("Poldat")
    wsPoldat.Select

    lastRow = wsPoldat.Cells(wsPoldat.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    Application.CutCopyMode = False
    Columns("A:G").Select
    wsPoldat.Sort.SortFields.Clear
    wsPoldat.Sort.SortFields.Add2 Key:=wsPoldat.Range( _
        "A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    wsPoldat.Sort.SortFields.Add2 Key:=wsPoldat.Range( _
        "B2:B" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With wsPoldat.Sort
        .SetRange wsPoldat.Range("A1:G" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
