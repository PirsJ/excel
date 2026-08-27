Sub EndOfShift()

    Dim wbSource As Workbook
    Dim wbTarget As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastRowSrc As Long
    Dim mdp As String

    mdp = "123"

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set wbTarget = Workbooks.Open(Filename:="P:\Rework\Operateurs\Historiques\Back to stock\Histo back to stock.xlsx")
    Set wbSource = Workbooks("Back to stock 2.0 Test 4.xlsm")

    For Each ws In wbSource.Worksheets
        ws.Unprotect mdp
    Next ws

    ' --- Historique des scans ---
    ' Corrigé : la borne "4995" était figée, alors que "histo" dépasse déjà largement
    ' cette taille (23 000+ lignes mesurées) -> tout ce qui dépassait 4995 n'était
    ' jamais archivé ni vidé. On calcule maintenant la dernière ligne réelle.
    Set wsSource = wbSource.Sheets("histo")
    Set wsTarget = wbTarget.Sheets("histo scan")

    lastRowSrc = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    If lastRowSrc >= 2 Then
        wsSource.Range("A2:H" & lastRowSrc).Copy
        lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1
        wsTarget.Cells(lastRow, 1).PasteSpecial Paste:=xlPasteValues
        wsTarget.Cells(lastRow, 1).PasteSpecial Paste:=xlPasteFormats
        wsSource.Range("A2:H" & lastRowSrc).ClearContents
    End If
    wsSource.Application.GoTo wsSource.Range("A1"), True

    ' --- Scan location ---
    ' Corrigé : la copie couvrait B:E mais l'effacement ne couvrait que B:C
    ' -> les colonnes D et E gardaient d'anciennes valeurs après la sauvegarde.
    Set wsSource = wbSource.Sheets("Scan location")
    Set wsTarget = wbTarget.Sheets("histo location")

    lastRowSrc = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    If lastRowSrc >= 5 Then
        wsSource.Range("B5:E" & lastRowSrc).Copy
        lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1
        wsTarget.Cells(lastRow, 1).PasteSpecial Paste:=xlPasteValues
        wsTarget.Cells(lastRow, 1).PasteSpecial Paste:=xlPasteFormats
        wsSource.Range("B5:E" & lastRowSrc).ClearContents
    End If
    wsSource.Application.GoTo wsSource.Range("A1"), True

    ' --- Manual ---
    ' Même correction : effacement aligné sur la plage réellement copiée (B:E).
    Set wsSource = wbSource.Sheets("Manual")
    Set wsTarget = wbTarget.Sheets("histo Manual")

    lastRowSrc = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    If lastRowSrc >= 5 Then
        wsSource.Range("B5:E" & lastRowSrc).Copy
        lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1
        wsTarget.Cells(lastRow, 1).PasteSpecial Paste:=xlPasteValues
        wsTarget.Cells(lastRow, 1).PasteSpecial Paste:=xlPasteFormats
        wsSource.Range("B5:E" & lastRowSrc).ClearContents
    End If
    wsSource.Application.GoTo wsSource.Range("A1"), True

    wbTarget.Save
    wbTarget.Close False

    With wbSource.Sheets("Scan Items")
        .Range("C8,D8,E8,F8,C16,D16").ClearContents

        .Range("E8:F8").Interior.Color = xlNone
    End With

    For Each ws In wbSource.Worksheets
        Select Case ws.Name
            Case "histo", "Poldat", "Traduction", "upc-item", "UPC - MDM", "Manual", "Repair"
            Case Else
                ws.Protect mdp
        End Select
    Next ws

    wbSource.Sheets("Scan Items").Activate

    wbSource.Save

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Sauvegarde effectuée"

End Sub
