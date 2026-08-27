Sub manual()
Attribute manual.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim wsManual As Worksheet
    Dim destRow As Long

    Range("C24").Select
    Selection.Copy

    ' Corrigé : la boucle "Do While ActiveCell...Select" avançait cellule par
    ' cellule depuis B5 -> lente et de plus en plus lente à mesure que "Manual"
    ' grossit. Remplacée par un accès direct à la première ligne vide.
    Set wsManual = Sheets("Manual")
    destRow = wsManual.Cells(wsManual.Rows.Count, "B").End(xlUp).Row + 1
    If destRow < 5 Then destRow = 5

    wsManual.Cells(destRow, "B").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Scan Items").Select
    Application.CutCopyMode = False
    Selection.ClearContents
End Sub

' "enregihistodenis" a été retiré : il référençait la fenêtre "Back to stock 3.xlsm",
' un nom de fichier qui ne correspond plus au classeur actuel ("Back to stock 2.0 Test 4.xlsm").
' Cette macro plantait donc si elle était lancée (fenêtre introuvable). Elle faisait
' de toute façon exactement ce que fait "EndOfShift" (Module9), qui est la version
' à jour et fonctionnelle de cette sauvegarde vers l'historique réseau.
