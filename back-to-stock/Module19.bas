Sub Zone()
Attribute Zone.VB_ProcData.VB_Invoke_Func = " \n14"
    ' Corrigé : la boucle "Do While ActiveCell...Select" avançait cellule par
    ' cellule depuis D2 -> de plus en plus lente à mesure que "histo" grossit
    ' (23 000+ lignes mesurées). Remplacée par un accès direct à la 1ère ligne vide.
    Dim wsHisto As Worksheet
    Dim r As Long
    Set wsHisto = Sheets("histo")
    wsHisto.Select
    r = wsHisto.Cells(wsHisto.Rows.Count, "D").End(xlUp).Row + 1
    If r < 2 Then r = 2
    wsHisto.Cells(r, "D").Select
    ActiveCell.FormulaR1C1 = "=if(LEFT(RC[-1],1)=""A"",left(RC[-1],4),left(RC[-1],2))"
End Sub
Sub ItemType()
Attribute ItemType.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim wsHisto As Worksheet
    Dim r As Long
    Set wsHisto = Sheets("histo")
    wsHisto.Select
    r = wsHisto.Cells(wsHisto.Rows.Count, "E").End(xlUp).Row + 1
    If r < 2 Then r = 2
    wsHisto.Cells(r, "E").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(RC[-3],1)=""R"",""Label"",IF(LEFT(RC[-3],1)=""S"",""Solution"",IF(LEFT(RC[-3],1)=""T"",""Trial"",""Boxes"")))"
End Sub
Sub From()
Attribute From.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim wsHisto As Worksheet
    Dim r As Long
  Range("I15").Select
    Selection.Copy
    Set wsHisto = Sheets("histo")
    wsHisto.Select
    ActiveSheet.Paste
    r = wsHisto.Cells(wsHisto.Rows.Count, "G").End(xlUp).Row + 1
    If r < 2 Then r = 2
    wsHisto.Cells(r, "G").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub
