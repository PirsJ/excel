Option Explicit
  
Sub GenererPlanningPacking()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsImport As Worksheet, wsCreation As Worksheet
    Dim wsSO As Worksheet, wsLO As Worksheet, wsHistory As Worksheet
    Dim wsTech As Worksheet, wsFormAM As Worksheet, wsFormPM As Worksheet
      
    ' =========================================================================
    ' 1. CONFIGURATION DES FEUILLES ET COLONNES
    ' =========================================================================
    Set wsImport = wb.Sheets("Import")
    Set wsCreation = wb.Sheets("Création")
    Set wsSO = wb.Sheets("Planning SO")
    Set wsLO = wb.Sheets("Planning LO")
      
    On Error Resume Next
    Set wsTech = wb.Sheets("Technical")
    Set wsFormAM = wb.Sheets("Formule AM + Day")
    Set wsFormPM = wb.Sheets("Formule PM")
    On Error GoTo 0
      
    Const COL_IMPORT_NOM As Integer = 1     ' Colonne A (Nom)
    Const COL_IMPORT_DATE As Integer = 2    ' Colonne B (Date par défaut)
    Const COL_IMPORT_SME As Integer = 18    ' Colonne R (Statut SME / Exclusion)
    Const COL_IMPORT_SHIFT As Integer = 34  ' Colonne AH (Shift)
    Const COL_IMPORT_ZONE As Integer = 44   ' Colonne AR (Zone)
      
    If wsImport Is Nothing Or wsCreation Is Nothing Or wsSO Is Nothing Or wsLO Is Nothing Then
        MsgBox "Une ou plusieurs feuilles de base sont introuvables.", vbCritical, "Erreur"
        Exit Sub
    End If
     
    ' Vérification / Création automatique de la feuille Historique
    On Error Resume Next
    Set wsHistory = wb.Sheets("Historique")
    If wsHistory Is Nothing Then
        Set wsHistory = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsHistory.Name = "Historique"
        wsHistory.Range("A1:D1").Value = Array("Date", "Nom", "Shift", "Role")
    End If
    On Error GoTo 0
      
    ' =========================================================================
    ' 2. UNIQUE ZONE DE DÉCLARATION DES VARIABLES (ZÉRO DOUBLON)
    ' =========================================================================
    Dim loSO As ListObject, loLO As ListObject
    Dim loBKSO As ListObject, loBKLO As ListObject
    Dim loExclude As ListObject, loForce As ListObject
    Dim row As ListRow
    Dim chkValue As Variant
    Dim dDate As Date
    Dim defaultDate As String
    Dim chosenDate As String
    Dim dateExiste As Boolean
    Dim reponse As VbMsgBoxResult
      
    Dim dictExclude As Object, dictForce As Object, dictHistory As Object, dictDayStations As Object, dictOpStation As Object
    Dim nmKey As String, opNom As String, opShift As String, opZone As String, opSME As String, keyData As String
    Dim forcedRoleString As String, bkText As String, finalRole As String
    Dim isForcedLO As Boolean, mixZones As Boolean, showNames As Boolean
    Dim lastRowHist As Long, nextRowHist As Long
    Dim histKey As String, histRows As Variant, item As Variant
      
    ' Quotas Généraux
    Dim QTA_SO_AM_MTO As Integer, QTA_SO_AM_SYNO As Integer
    Dim QTA_SO_PM_MTO As Integer, QTA_SO_PM_SYNO As Integer
    Dim QTA_LO_AM_HUGE As Integer, QTA_LO_AM_MERGING As Integer, QTA_LO_AM_REWORK As Integer
    Dim QTA_LO_PM_HUGE As Integer, QTA_LO_PM_MERGING As Integer, QTA_LO_PM_REWORK As Integer
      
    ' Quotas Back-up
    Dim BK_SO_AM_MTO As Integer, BK_SO_AM_SYNO As Integer
    Dim BK_SO_PM_MTO As Integer, BK_SO_PM_SYNO As Integer
    Dim BK_LO_AM_REWORK As Integer, BK_LO_AM_MERGING As Integer, BK_LO_AM_HUGE As Integer
    Dim BK_LO_PM_REWORK As Integer, BK_LO_PM_MERGING As Integer, BK_LO_PM_HUGE As Integer
      
    ' Index de boucles, tailles et calculs de quotas
    Dim lastRowImport As Long, i As Long
    Dim idxStrict As Long, idxRem As Long, idxLOStrict As Long, idxLORem As Long
    Dim idxPM As Long, idxPM_LO As Long
    Dim reqMTO As Integer, reqSYNO As Integer
      
    ' Tailles de stockage pour le mix des équipes
    Dim origSO_Strict_Count As Long, origLO_Strict_Count As Long
    Dim origSO_Rest_Count As Long, origLO_Rest_Count As Long
    Dim origSO_PM_Count As Long, origLO_PM_Count As Long
      
    ' Indicateurs de remplissage terrain
    Dim filledMTO As Integer, filledSYNO As Integer
    Dim filledPM_MTO As Integer, filledPM_SYNO As Integer, filledPM_SO As Integer
    Dim filledRework As Integer, filledMerging As Integer, filledHuge As Integer, filledMulti As Integer
    Dim filledPM_Rework As Integer, filledPM_Merging As Integer, filledPM_Huge As Integer, filledPM_Multi As Integer
    Dim allocatedBK_SO_MTO As Integer, allocatedBK_SO_SYNO As Integer
    Dim allocatedBK_LO_Rework As Integer, allocatedBK_LO_Merging As Integer, allocatedBK_LO_Huge As Integer
      
    Dim splitS() As String
      
    ' Instanciation des collections
    Dim listSO_AM_Forced As New Collection, listSO_PM_Forced As New Collection
    Dim listLO_AM_Forced As New Collection, listLO_PM_Forced As New Collection
    Dim listSO_AM_Pack_Strict As New Collection, listSO_AM_Pack_Rest As New Collection
    Dim listSO_AM_Ship As New Collection, listSO_PM_Pack As New Collection, listSO_PM_Ship As New Collection
    Dim listLO_AM_Strict As New Collection, listLO_AM_Rest As New Collection, listLO_PM As New Collection
    Dim listMix_AM_Strict As New Collection, listMix_AM_Rest As New Collection, listMix_PM As New Collection
    Dim listSO_AM_Pack_Remaining As New Collection, listLO_AM_Remaining As New Collection
    Dim listSO_AM_Final As New Collection, listSO_PM_Final As New Collection
    Dim listLO_AM_Final As New Collection, listLO_PM_Final As New Collection
    Dim listSO_AM_Temp As New Collection, listLO_AM_Temp As New Collection
    Dim listSO_PM_Temp As New Collection, listLO_PM_Temp As New Collection, listLO_PM_Temp2 As New Collection

    ' Un seul amorçage du générateur aléatoire pour toute l'exécution
    ' (évite des mélanges corrélés si ShuffleCollection est appelée plusieurs fois de suite)
    Randomize
      
    ' =========================================================================
    ' 3. DEMANDE DE LA DATE À L'UTILISATEUR (AVANT TOUT NETTOYAGE)
    ' =========================================================================
    If Trim(wsImport.Cells(2, COL_IMPORT_DATE).Text) <> "" Then
        defaultDate = Trim(wsImport.Cells(2, COL_IMPORT_DATE).Text)
    Else
        defaultDate = Format(Date, "dd/mm/yyyy")
    End If
      
    chosenDate = InputBox("Générer le planning pour la date :", "Date du planning", defaultDate)
      
    If Trim(chosenDate) = "" Then
        MsgBox "Génération annulée. Les plannings actuels n'ont pas été modifiés.", vbInformation, "Annulation"
        Exit Sub
    End If

    If IsDate(chosenDate) Then
        dDate = CDate(chosenDate)
    Else
        dDate = Date
    End If
    
    ' =========================================================================
    ' 4. PROTECTION CONTRE LES DOUBLONS DE DATE DANS L'HISTORIQUE
    ' =========================================================================
    dateExiste = False
    lastRowHist = wsHistory.Cells(wsHistory.Rows.Count, 1).End(xlUp).row
    If lastRowHist >= 2 Then
        For i = 2 To lastRowHist
            If wsHistory.Cells(i, 1).Value = dDate Then
                dateExiste = True
                Exit For
            End If
        Next i
    End If
     
    If dateExiste Then
        reponse = MsgBox("Un planning a déjà été généré dans l'historique pour la date du " & Format(dDate, "dd/mm/yyyy") & "." & vbCrLf & _
                         "Voulez-vous l'écraser et régénérer un nouveau planning pour cette date ?", vbYesNo + vbExclamation, "Planning déjà existant")
        If reponse = vbNo Then
            MsgBox "Génération annulée. Les plannings actuels n'ont pas été modifiés.", vbInformation, "Annulation"
            Exit Sub
        Else
            ' Nettoyage chirurgical de l'ancien historique pour cette date (boucle inversée indispensable)
            For i = lastRowHist To 2 Step -1
                If wsHistory.Cells(i, 1).Value = dDate Then
                    wsHistory.Rows(i).Delete
                End If
            Next i
        End If
    End If
      
    ' =========================================================================
    ' 5. LECTURE DES PARAMÈTRES, HISTORIQUE ET CONFIGURATIONS
    ' =========================================================================
    On Error Resume Next
    Set loSO = wsCreation.ListObjects("tblCapSO")
    Set loLO = wsCreation.ListObjects("tblCapLO")
    Set loBKSO = wsCreation.ListObjects("tblBackupSO")
    Set loBKLO = wsCreation.ListObjects("tblBackupLO")
    Set loExclude = wsCreation.ListObjects("tblExclude")
    Set loForce = wsCreation.ListObjects("tblForce")
    On Error GoTo 0
      
    If loSO Is Nothing Or loLO Is Nothing Then
        MsgBox "Impossible de trouver les tableaux de capacités globaux 'tblCapSO' ou 'tblCapLO'.", vbCritical, "Erreur"
        Exit Sub
    End If
      
    ' --- LECTURE DE LA CASE À COCHER MIX ALL ---
    mixZones = False
    On Error Resume Next
    chkValue = wsCreation.Range("chkMixZones").Value
    If Err.Number = 0 Then
        If chkValue = True Or UCase(Trim(CStr(chkValue))) = "TRUE" Then
            mixZones = True
        End If
    End If
    On Error GoTo 0
     
    ' --- LECTURE DE LA CASE À COCHER AFFICHAGE DES NOMS ---
    showNames = False
    On Error Resume Next
    chkValue = wsCreation.Range("chkShowNames").Value
    If Err.Number = 0 Then
        If chkValue = True Or UCase(Trim(CStr(chkValue))) = "TRUE" Then
            showNames = True
        End If
    End If
    On Error GoTo 0
      
    ' Chargement des Dictionnaires d'Exclusion et de Forçage
    Set dictExclude = CreateObject("Scripting.Dictionary")
    Set dictForce = CreateObject("Scripting.Dictionary")
      
    If Not loExclude Is Nothing Then
        For Each row In loExclude.ListRows
            nmKey = UCase(Trim(row.Range(1, 1).Value))
            If nmKey <> "" Then dictExclude(nmKey) = True
        Next row
    End If
    If Not loForce Is Nothing Then
        For Each row In loForce.ListRows
            nmKey = UCase(Trim(row.Range(1, 1).Value))
            If nmKey <> "" Then dictForce(nmKey) = Trim(row.Range(1, 2).Value)
        Next row
    End If
     
    ' --- CHARGEMENT DE L'HISTORIQUE SANS LA DATE ÉCRASÉE ---
    Set dictHistory = CreateObject("Scripting.Dictionary")
    lastRowHist = wsHistory.Cells(wsHistory.Rows.Count, 1).End(xlUp).row
    If lastRowHist >= 2 Then
        histRows = wsHistory.Range("A2:D" & lastRowHist).Value
        For i = 1 To UBound(histRows, 1)
            histKey = UCase(Trim(histRows(i, 2))) & "|" & UCase(Trim(histRows(i, 4)))
            If dictHistory.Exists(histKey) Then
                dictHistory(histKey) = dictHistory(histKey) + 1
            Else
                dictHistory(histKey) = 1
            End If
        Next i
    End If
      
    ' Chargement des volumes de capacités globales et backups
    For Each row In loLO.ListRows
        Select Case UCase(Trim(row.Range(1, 1).Value))
            Case "HUGE": QTA_LO_AM_HUGE = Val(row.Range(1, 2).Value): QTA_LO_PM_HUGE = Val(row.Range(1, 3).Value)
            Case "MERGING": QTA_LO_AM_MERGING = Val(row.Range(1, 2).Value): QTA_LO_PM_MERGING = Val(row.Range(1, 3).Value)
            Case "REWORK LO": QTA_LO_AM_REWORK = Val(row.Range(1, 2).Value): QTA_LO_PM_REWORK = Val(row.Range(1, 3).Value)
        End Select
    Next row
    For Each row In loSO.ListRows
        Select Case UCase(Trim(row.Range(1, 1).Value))
            Case "MTO": QTA_SO_AM_MTO = Val(row.Range(1, 2).Value): QTA_SO_PM_MTO = Val(row.Range(1, 3).Value)
            Case "SYNO": QTA_SO_AM_SYNO = Val(row.Range(1, 2).Value): QTA_SO_PM_SYNO = Val(row.Range(1, 3).Value)
        End Select
    Next row
    If Not loBKSO Is Nothing Then
        For Each row In loBKSO.ListRows
            Select Case UCase(Trim(row.Range(1, 1).Value))
                Case "MTO": BK_SO_AM_MTO = Val(row.Range(1, 2).Value): BK_SO_PM_MTO = Val(row.Range(1, 3).Value)
                Case "SYNO": BK_SO_AM_SYNO = Val(row.Range(1, 2).Value): BK_SO_PM_SYNO = Val(row.Range(1, 3).Value)
            End Select
        Next row
    End If
    If Not loBKLO Is Nothing Then
        For Each row In loBKLO.ListRows
            Select Case UCase(Trim(row.Range(1, 1).Value))
                Case "REWORK LO": BK_LO_AM_REWORK = Val(row.Range(1, 2).Value): BK_LO_PM_REWORK = Val(row.Range(1, 3).Value)
                Case "MERGING": BK_LO_AM_MERGING = Val(row.Range(1, 2).Value): BK_LO_PM_MERGING = Val(row.Range(1, 3).Value)
                Case "HUGE": BK_LO_AM_HUGE = Val(row.Range(1, 2).Value): BK_LO_PM_HUGE = Val(row.Range(1, 3).Value)
            End Select
        Next row
    End If
  
    ' =========================================================================
    ' 6. EXTRACTION, SÉCURITÉS FILTRES ET SEGMENTATION
    ' =========================================================================
    lastRowImport = wsImport.Cells(wsImport.Rows.Count, COL_IMPORT_NOM).End(xlUp).row
      
    For i = 2 To lastRowImport
        opNom = Trim(wsImport.Cells(i, COL_IMPORT_NOM).Value)
        opShift = UCase(Trim(wsImport.Cells(i, COL_IMPORT_SHIFT).Value))
        opZone = UCase(Trim(wsImport.Cells(i, COL_IMPORT_ZONE).Value))
        opSME = UCase(Trim(wsImport.Cells(i, COL_IMPORT_SME).Value))
        nmKey = UCase(opNom)
          
        If opNom <> "" And Not (opSME = "SME" Or opSME = "SME*") And Not dictExclude.Exists(nmKey) Then
            If dictForce.Exists(nmKey) Then
                forcedRoleString = UCase(Trim(dictForce(nmKey)))
                keyData = opNom & "|" & opShift & "|" & dictForce(nmKey)
                  
                If forcedRoleString Like "*HUGE*" Or forcedRoleString Like "*MERG*" Or forcedRoleString Like "*REWORK*" Or forcedRoleString Like "*TOTES*" Then
                    isForcedLO = True
                Else
                    isForcedLO = False
                End If
                  
                If isForcedLO Then
                    If opShift Like "*PM*" Then listLO_PM_Forced.Add keyData Else listLO_AM_Forced.Add keyData
                Else
                    If opShift Like "*PM*" Then listSO_PM_Forced.Add keyData Else listSO_AM_Forced.Add keyData
                End If
            Else
                keyData = opNom & "|" & opShift
                If opZone Like "*SMALL*" Then
                    If opShift Like "*PM*" Then
                        listSO_PM_Pack.Add keyData
                    ElseIf opShift = "AM" Then
                        listSO_AM_Pack_Strict.Add keyData
                    Else
                        listSO_AM_Pack_Rest.Add keyData
                    End If
                ElseIf opZone Like "*SHIP*" Then
                    If opShift Like "*PM*" Then listSO_PM_Ship.Add keyData Else listSO_AM_Ship.Add keyData
                ElseIf opZone Like "*LARGE*" Then
                    If opShift Like "*PM*" Then
                        listLO_PM.Add keyData
                    ElseIf opShift = "AM" Then
                        listLO_AM_Strict.Add keyData
                    Else
                        listLO_AM_Rest.Add keyData
                    End If
                End If
            End If
        End If
    Next i
      
    ' Logique de brassage et fusion globale des zones si l'option est cochée
    If mixZones = True Then
        origSO_Strict_Count = listSO_AM_Pack_Strict.Count
        origLO_Strict_Count = listLO_AM_Strict.Count
        origSO_Rest_Count = listSO_AM_Pack_Rest.Count
        origLO_Rest_Count = listLO_AM_Rest.Count
        origSO_PM_Count = listSO_PM_Pack.Count
        origLO_PM_Count = listLO_PM.Count
          
        For i = 1 To listSO_AM_Pack_Strict.Count: listMix_AM_Strict.Add listSO_AM_Pack_Strict(i): Next i
        For i = 1 To listLO_AM_Strict.Count: listMix_AM_Strict.Add listLO_AM_Strict(i): Next i
        For i = 1 To listSO_AM_Pack_Rest.Count: listMix_AM_Rest.Add listSO_AM_Pack_Rest(i): Next i
        For i = 1 To listLO_AM_Rest.Count: listMix_AM_Rest.Add listLO_AM_Rest(i): Next i
        For i = 1 To listSO_PM_Pack.Count: listMix_PM.Add listSO_PM_Pack(i): Next i
        For i = 1 To listLO_PM.Count: listMix_PM.Add listLO_PM(i): Next i
          
        Set listSO_AM_Pack_Strict = New Collection: Set listLO_AM_Strict = New Collection
        Set listSO_AM_Pack_Rest = New Collection: Set listLO_AM_Rest = New Collection
        Set listSO_PM_Pack = New Collection: Set listLO_PM = New Collection
          
        Set listMix_AM_Strict = ShuffleCollection(listMix_AM_Strict)
        Set listMix_AM_Rest = ShuffleCollection(listMix_AM_Rest)
        Set listMix_PM = ShuffleCollection(listMix_PM)
  
        For i = 1 To listMix_AM_Strict.Count
            If i <= origSO_Strict_Count Then
                listSO_AM_Pack_Strict.Add listMix_AM_Strict(i)
            Else
                listLO_AM_Strict.Add listMix_AM_Strict(i)
            End If
        Next i
          
        For i = 1 To listMix_AM_Rest.Count
            If i <= origSO_Rest_Count Then
                listSO_AM_Pack_Rest.Add listMix_AM_Rest(i)
            Else
                listLO_AM_Rest.Add listMix_AM_Rest(i)
            End If
        Next i
          
        For i = 1 To listMix_PM.Count
            If i <= origSO_PM_Count Then
                listSO_PM_Pack.Add listMix_PM(i)
            Else
                listLO_PM.Add listMix_PM(i)
            End If
        Next i
    End If
      
    ' Mélanges internes initiales
    Set listSO_AM_Pack_Strict = ShuffleCollection(listSO_AM_Pack_Strict)
    Set listSO_AM_Pack_Rest = ShuffleCollection(listSO_AM_Pack_Rest)
    Set listSO_AM_Ship = ShuffleCollection(listSO_AM_Ship)
    Set listSO_PM_Pack = ShuffleCollection(listSO_PM_Pack)
    Set listSO_PM_Ship = ShuffleCollection(listSO_PM_Ship)
    Set listLO_AM_Strict = ShuffleCollection(listLO_AM_Strict)
    Set listLO_AM_Rest = ShuffleCollection(listLO_AM_Rest)
    Set listLO_PM = ShuffleCollection(listLO_PM)
      
    ' =========================================================================
    ' 7. LOGIQUES ET ATTRIBUTIONS DES POSTES AVEC TOURNANTE EQUITABLE
    ' =========================================================================
    ' --- PLANNING SO - GAUCHE (AM / DAY / MI-TEMPS) ---
    For i = 1 To listSO_AM_Forced.Count
        splitS = Split(listSO_AM_Forced(i), "|"): finalRole = NormaliserRole(splitS(2), False)
        listSO_AM_Final.Add finalRole & "|" & splitS(0) & "|" & splitS(1) & "|" & ""
        If finalRole = "MTO" Then filledMTO = filledMTO + 1
        If finalRole = "SYNO" Then filledSYNO = filledSYNO + 1
    Next i
      
    reqMTO = Application.WorksheetFunction.RoundUp(QTA_SO_AM_MTO * 0.75, 0)
    reqSYNO = Application.WorksheetFunction.RoundUp(QTA_SO_AM_SYNO * 0.75, 0)
     
    ' Remplissage MTO Strict (priorité historique faible)
    Set listSO_AM_Pack_Strict = TrierParHistorique(listSO_AM_Pack_Strict, "MTO", dictHistory)
    idxStrict = 1
    Do While idxStrict <= listSO_AM_Pack_Strict.Count And filledMTO < reqMTO And filledMTO < QTA_SO_AM_MTO
        splitS = Split(listSO_AM_Pack_Strict(idxStrict), "|"): listSO_AM_Final.Add "MTO|" & splitS(0) & "|" & splitS(1) & "|" & ""
        filledMTO = filledMTO + 1: idxStrict = idxStrict + 1
    Loop
     
    ' Remplissage SYNO Strict sur le reste de la liste
    Set listSO_AM_Temp = New Collection
    Do While idxStrict <= listSO_AM_Pack_Strict.Count
        listSO_AM_Temp.Add listSO_AM_Pack_Strict(idxStrict)
        idxStrict = idxStrict + 1
    Loop
     
    Set listSO_AM_Temp = TrierParHistorique(listSO_AM_Temp, "SYNO", dictHistory)
    idxStrict = 1
    Do While idxStrict <= listSO_AM_Temp.Count And filledSYNO < reqSYNO And filledSYNO < QTA_SO_AM_SYNO
        splitS = Split(listSO_AM_Temp(idxStrict), "|"): listSO_AM_Final.Add "SYNO|" & splitS(0) & "|" & splitS(1) & "|" & ""
        filledSYNO = filledSYNO + 1: idxStrict = idxStrict + 1
    Loop
     
    ' On bascule les restants stricts + les "Rest" dans le pool Remaining
    Do While idxStrict <= listSO_AM_Temp.Count
        listSO_AM_Pack_Remaining.Add listSO_AM_Temp(idxStrict): idxStrict = idxStrict + 1
    Loop
    For i = 1 To listSO_AM_Pack_Rest.Count: listSO_AM_Pack_Remaining.Add listSO_AM_Pack_Rest(i): Next i
     
    ' Remplissage des quotas MTO Restants
    Set listSO_AM_Pack_Remaining = TrierParHistorique(listSO_AM_Pack_Remaining, "MTO", dictHistory)
    idxRem = 1
    Do While idxRem <= listSO_AM_Pack_Remaining.Count And filledMTO < QTA_SO_AM_MTO
        splitS = Split(listSO_AM_Pack_Remaining(idxRem), "|"): listSO_AM_Final.Add "MTO|" & splitS(0) & "|" & splitS(1) & "|" & ""
        filledMTO = filledMTO + 1: idxRem = idxRem + 1
    Loop
     
    ' Remplissage des quotas SYNO Restants
    Set listSO_AM_Temp = New Collection
    Do While idxRem <= listSO_AM_Pack_Remaining.Count
        listSO_AM_Temp.Add listSO_AM_Pack_Remaining(idxRem)
        idxRem = idxRem + 1
    Loop
    Set listSO_AM_Temp = TrierParHistorique(listSO_AM_Temp, "SYNO", dictHistory)
     
    idxRem = 1
    Do While idxRem <= listSO_AM_Temp.Count And filledSYNO < QTA_SO_AM_SYNO
        splitS = Split(listSO_AM_Temp(idxRem), "|"): listSO_AM_Final.Add "SYNO|" & splitS(0) & "|" & splitS(1) & "|" & ""
        filledSYNO = filledSYNO + 1: idxRem = idxRem + 1
    Loop
     
    ' Reste en SO / Backup
    Do While idxRem <= listSO_AM_Temp.Count
        splitS = Split(listSO_AM_Temp(idxRem), "|"): bkText = ""
        If splitS(1) = "AM" Then
            If allocatedBK_SO_MTO < BK_SO_AM_MTO Then
                bkText = "Back up MTO": allocatedBK_SO_MTO = allocatedBK_SO_MTO + 1
            ElseIf allocatedBK_SO_SYNO < BK_SO_AM_SYNO Then
                bkText = "Back up SYNO": allocatedBK_SO_SYNO = allocatedBK_SO_SYNO + 1
            End If
        End If
        listSO_AM_Final.Add "SO|" & splitS(0) & "|" & splitS(1) & "|" & bkText: idxRem = idxRem + 1
    Loop
    For i = 1 To listSO_AM_Ship.Count
        splitS = Split(listSO_AM_Ship(i), "|"): listSO_AM_Final.Add "Shipping|" & splitS(0) & "|" & splitS(1) & "|" & ""
    Next i
  
    ' --- PLANNING SO - DROITE (PM) ---
    For i = 1 To listSO_PM_Forced.Count
        splitS = Split(listSO_PM_Forced(i), "|"): finalRole = NormaliserRole(splitS(2), False)
        listSO_PM_Final.Add finalRole & "|" & splitS(0) & "|" & splitS(1) & "|" & ""
        If finalRole = "MTO" Then filledPM_MTO = filledPM_MTO + 1
        If finalRole = "SYNO" Then filledPM_SYNO = filledPM_SYNO + 1
    Next i
     
    ' PM - Tri MTO et remplissage
    Set listSO_PM_Pack = TrierParHistorique(listSO_PM_Pack, "MTO", dictHistory)
    idxPM = 1
    Do While idxPM <= listSO_PM_Pack.Count And filledPM_MTO < QTA_SO_PM_MTO
        splitS = Split(listSO_PM_Pack(idxPM), "|")
        listSO_PM_Final.Add "MTO|" & splitS(0) & "|" & splitS(1) & "|" & ""
        filledPM_MTO = filledPM_MTO + 1
        idxPM = idxPM + 1
    Loop
     
    ' PM - Tri SYNO sur le reste
    Set listSO_PM_Temp = New Collection
    Do While idxPM <= listSO_PM_Pack.Count
        listSO_PM_Temp.Add listSO_PM_Pack(idxPM)
        idxPM = idxPM + 1
    Loop
    Set listSO_PM_Temp = TrierParHistorique(listSO_PM_Temp, "SYNO", dictHistory)
     
    idxPM = 1
    Do While idxPM <= listSO_PM_Temp.Count And filledPM_SYNO < QTA_SO_PM_SYNO
        splitS = Split(listSO_PM_Temp(idxPM), "|")
        listSO_PM_Final.Add "SYNO|" & splitS(0) & "|" & splitS(1) & "|" & ""
        filledPM_SYNO = filledPM_SYNO + 1
        idxPM = idxPM + 1
    Loop
     
    ' PM - Le reste en SO classique
    Do While idxPM <= listSO_PM_Temp.Count
        splitS = Split(listSO_PM_Temp(idxPM), "|")
        filledPM_SO = filledPM_SO + 1: bkText = ""
        If filledPM_SO <= BK_SO_PM_MTO Then
            bkText = "Back up MTO"
        ElseIf filledPM_SO <= BK_SO_PM_MTO + BK_SO_PM_SYNO Then
            bkText = "Back up SYNO"
        End If
        listSO_PM_Final.Add "SO|" & splitS(0) & "|" & splitS(1) & "|" & bkText
        idxPM = idxPM + 1
    Loop
    For i = 1 To listSO_PM_Ship.Count
        splitS = Split(listSO_PM_Ship(i), "|"): listSO_PM_Final.Add "Shipping|" & splitS(0) & "|" & splitS(1) & "|" & ""
    Next i
  
    ' --- PLANNING LO - GAUCHE (AM / DAY / MI-TEMPS) ---
    For i = 1 To listLO_AM_Forced.Count
        splitS = Split(listLO_AM_Forced(i), "|"): finalRole = NormaliserRole(splitS(2), True)
        listLO_AM_Final.Add finalRole & "|" & splitS(0) & "|" & splitS(1) & "|" & ""
        If finalRole = "Rework LO" Then filledRework = filledRework + 1
        If finalRole = "Merging" Then filledMerging = filledMerging + 1
        If finalRole = "Huge Order" Then filledHuge = filledHuge + 1
    Next i
      
    ' Tri Rework LO
    Set listLO_AM_Strict = TrierParHistorique(listLO_AM_Strict, "Rework LO", dictHistory)
    idxLOStrict = 1
    Do While idxLOStrict <= listLO_AM_Strict.Count And filledRework < QTA_LO_AM_REWORK
        splitS = Split(listLO_AM_Strict(idxLOStrict), "|"): listLO_AM_Final.Add "Rework LO|" & splitS(0) & "|" & splitS(1) & "|" & ""
        filledRework = filledRework + 1: idxLOStrict = idxLOStrict + 1
    Loop
     
    ' Tri Merging
    Set listLO_AM_Temp = New Collection
    Do While idxLOStrict <= listLO_AM_Strict.Count
        listLO_AM_Temp.Add listLO_AM_Strict(idxLOStrict)
        idxLOStrict = idxLOStrict + 1
    Loop
    Set listLO_AM_Temp = TrierParHistorique(listLO_AM_Temp, "Merging", dictHistory)
     
    idxLOStrict = 1
    Do While idxLOStrict <= listLO_AM_Temp.Count And filledMerging < QTA_LO_AM_MERGING
        splitS = Split(listLO_AM_Temp(idxLOStrict), "|")
        listLO_AM_Final.Add "Merging|" & splitS(0) & "|" & splitS(1) & "|" & ""
        filledMerging = filledMerging + 1
        idxLOStrict = idxLOStrict + 1
    Loop
     
    ' On bascule le reste LO Strict + LO Rest dans Remaining
    Do While idxLOStrict <= listLO_AM_Temp.Count
        listLO_AM_Remaining.Add listLO_AM_Temp(idxLOStrict): idxLOStrict = idxLOStrict + 1
    Loop
    For i = 1 To listLO_AM_Rest.Count: listLO_AM_Remaining.Add listLO_AM_Rest(i): Next i
     
    ' Tri pour Huge Order
    Set listLO_AM_Remaining = TrierParHistorique(listLO_AM_Remaining, "Huge Order", dictHistory)
      
    idxLORem = 1
    Do While idxLORem <= listLO_AM_Remaining.Count
        splitS = Split(listLO_AM_Remaining(idxLORem), "|"): bkText = ""
        If filledHuge < QTA_LO_AM_HUGE Then
            filledHuge = filledHuge + 1
            If splitS(1) = "AM" Then
                If allocatedBK_LO_Rework < BK_LO_AM_REWORK Then
                    bkText = "Back up Rework": allocatedBK_LO_Rework = allocatedBK_LO_Rework + 1
                ElseIf allocatedBK_LO_Merging < BK_LO_AM_MERGING Then
                    bkText = "Back up Merging": allocatedBK_LO_Merging = allocatedBK_LO_Merging + 1
                End If
            End If
            listLO_AM_Final.Add "Huge Order|" & splitS(0) & "|" & splitS(1) & "|" & bkText
        Else
            filledMulti = filledMulti + 1
            If splitS(1) = "AM" Then
                If allocatedBK_LO_Huge < BK_LO_AM_HUGE Then
                    bkText = "Back up Huge order": allocatedBK_LO_Huge = allocatedBK_LO_Huge + 1
                End If
            End If
            listLO_AM_Final.Add "Multi Totes Orders|" & splitS(0) & "|" & splitS(1) & "|" & bkText
        End If
        idxLORem = idxLORem + 1
    Loop
  
    ' --- PLANNING LO - DROITE (PM) ---
    For i = 1 To listLO_PM_Forced.Count
        splitS = Split(listLO_PM_Forced(i), "|"): finalRole = NormaliserRole(splitS(2), True)
        listLO_PM_Final.Add finalRole & "|" & splitS(0) & "|" & splitS(1) & "|" & ""
        If finalRole = "Rework LO" Then filledPM_Rework = filledPM_Rework + 1
        If finalRole = "Merging" Then filledPM_Merging = filledPM_Merging + 1
        If finalRole = "Huge Order" Then filledPM_Huge = filledPM_Huge + 1
    Next i
     
    ' PM - Tri Rework LO
    Set listLO_PM = TrierParHistorique(listLO_PM, "Rework LO", dictHistory)
    idxPM_LO = 1
    Do While idxPM_LO <= listLO_PM.Count And filledPM_Rework < QTA_LO_PM_REWORK
        splitS = Split(listLO_PM(idxPM_LO), "|")
        listLO_PM_Final.Add "Rework LO|" & splitS(0) & "|" & splitS(1) & "|" & ""
        filledPM_Rework = filledPM_Rework + 1
        idxPM_LO = idxPM_LO + 1
    Loop
     
    ' PM - Tri Merging
    Set listLO_PM_Temp = New Collection
    Do While idxPM_LO <= listLO_PM.Count
        listLO_PM_Temp.Add listLO_PM(idxPM_LO)
        idxPM_LO = idxPM_LO + 1
    Loop
    Set listLO_PM_Temp = TrierParHistorique(listLO_PM_Temp, "Merging", dictHistory)
     
    idxPM_LO = 1
    Do While idxPM_LO <= listLO_PM_Temp.Count And filledPM_Merging < QTA_LO_PM_MERGING
        splitS = Split(listLO_PM_Temp(idxPM_LO), "|")
        listLO_PM_Final.Add "Merging|" & splitS(0) & "|" & splitS(1) & "|" & ""
        filledPM_Merging = filledPM_Merging + 1
        idxPM_LO = idxPM_LO + 1
    Loop
     
    ' PM - Tri Huge Order pour le reste
    Set listLO_PM_Temp2 = New Collection
    Do While idxPM_LO <= listLO_PM_Temp.Count
        listLO_PM_Temp2.Add listLO_PM_Temp(idxPM_LO)
        idxPM_LO = idxPM_LO + 1
    Loop
    Set listLO_PM_Temp2 = TrierParHistorique(listLO_PM_Temp2, "Huge Order", dictHistory)
     
    idxPM_LO = 1
    Do While idxPM_LO <= listLO_PM_Temp2.Count
        splitS = Split(listLO_PM_Temp2(idxPM_LO), "|"): bkText = ""
        If filledPM_Huge < QTA_LO_PM_HUGE Then
            filledPM_Huge = filledPM_Huge + 1
            If filledPM_Huge <= BK_LO_PM_REWORK Then
                bkText = "Back up Rework"
            ElseIf filledPM_Huge <= BK_LO_PM_REWORK + BK_LO_PM_MERGING Then
                bkText = "Back up Merging"
            End If
            listLO_PM_Final.Add "Huge Order|" & splitS(0) & "|" & splitS(1) & "|" & bkText
        Else
            filledPM_Multi = filledPM_Multi + 1
            If filledPM_Multi <= BK_LO_PM_HUGE Then bkText = "Back up Huge order"
            listLO_PM_Final.Add "Multi Totes Orders|" & splitS(0) & "|" & splitS(1) & "|" & bkText
        End If
        idxPM_LO = idxPM_LO + 1
    Loop
  
    ' =========================================================================
    ' 8. TRI ALPHABÉTIQUE DES ROLES
    ' =========================================================================
    Set listSO_AM_Final = TrierCollection(listSO_AM_Final, False)
    Set listSO_PM_Final = TrierCollection(listSO_PM_Final, False)
    Set listLO_AM_Final = TrierCollection(listLO_AM_Final, True)
    Set listLO_PM_Final = TrierCollection(listLO_PM_Final, True)
      
    ' =========================================================================
    ' 9. REMPLISSAGE FIXE DES PLANS DE BANCS DE BACKEND (FORMULES MAPS)
    ' =========================================================================
    Set dictDayStations = CreateObject("Scripting.Dictionary")
    Set dictOpStation = CreateObject("Scripting.Dictionary")
     
    Call RemplirPlanBanc(wsFormAM, listSO_AM_Final, wsTech, showNames, True, dictDayStations, dictOpStation)
    Call RemplirPlanBanc(wsFormPM, listSO_PM_Final, wsTech, showNames, False, dictDayStations, dictOpStation)
      
    ' Blocage des rafraîchissements pour l'écriture finale
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
     
    ' =========================================================================
    ' 9.5 RECONSTRUCTION DES EN-TÊTES ET ÉCRITURE DES DEUX FEUILLES FINALES
    ' =========================================================================
    ' --- TRAITEMENT DE LA FEUILLE DUAL SO ---
    Call PreparerFeuilleSO_Dual(wsSO, chosenDate, showNames)
     
    Call EcrireDonnees(wsSO, listSO_AM_Final, 1, 7, dictOpStation, showNames, False, False)
    Call EcrireDonnees(wsSO, listSO_PM_Final, 7, 7, dictOpStation, showNames, False, False)
     
    ' Recalcul et compression des hauteurs et polices sur le SO
    Dim maxSOCount As Long
    maxSOCount = listSO_AM_Final.Count
    If listSO_PM_Final.Count > maxSOCount Then maxSOCount = listSO_PM_Final.Count
     
    If maxSOCount > 0 Then
        Dim fSize As Integer, rHeight As Double
        Select Case maxSOCount
            Case Is <= 30:   fSize = 12: rHeight = 22
            Case 31 To 47:  fSize = 11: rHeight = 18
            Case 48 To 54:  fSize = 10: rHeight = 14
            Case 55 To 65:  fSize = 9:  rHeight = 11.5
            Case Else:       fSize = 8:  rHeight = 10
        End Select
         
        With wsSO.Range("A8:K" & (7 + maxSOCount)).Font
            .Size = fSize
        End With
        wsSO.Rows("8:" & (7 + maxSOCount)).RowHeight = rHeight
    End If
     
    ' --- TRAITEMENT DE LA FEUILLE LO ---
    Call PreparerFeuilleLO_Dual(wsLO, chosenDate, showNames)
     
    Call EcrireDonnees(wsLO, listLO_AM_Final, 1, 11, dictOpStation, showNames, True, True)
    Call EcrireDonnees(wsLO, listLO_PM_Final, 6, 16, dictOpStation, showNames, True, True)
     
    ' =========================================================================
    ' 10. ENREGISTREMENT AUTOMATIQUE DANS L'HISTORIQUE
    ' =========================================================================
    nextRowHist = wsHistory.Cells(wsHistory.Rows.Count, 1).End(xlUp).row + 1
     
    For Each item In listSO_AM_Final
        splitS = Split(item, "|")
        wsHistory.Cells(nextRowHist, 1).Value = dDate
        wsHistory.Cells(nextRowHist, 2).Value = splitS(1)
        wsHistory.Cells(nextRowHist, 3).Value = splitS(2)
        wsHistory.Cells(nextRowHist, 4).Value = splitS(0)
        nextRowHist = nextRowHist + 1
    Next item
    For Each item In listSO_PM_Final
        splitS = Split(item, "|")
        wsHistory.Cells(nextRowHist, 1).Value = dDate
        wsHistory.Cells(nextRowHist, 2).Value = splitS(1)
        wsHistory.Cells(nextRowHist, 3).Value = splitS(2)
        wsHistory.Cells(nextRowHist, 4).Value = splitS(0)
        nextRowHist = nextRowHist + 1
    Next item
    For Each item In listLO_AM_Final
        splitS = Split(item, "|")
        wsHistory.Cells(nextRowHist, 1).Value = dDate
        wsHistory.Cells(nextRowHist, 2).Value = splitS(1)
        wsHistory.Cells(nextRowHist, 3).Value = splitS(2)
        wsHistory.Cells(nextRowHist, 4).Value = splitS(0)
        nextRowHist = nextRowHist + 1
    Next item
    For Each item In listLO_PM_Final
        splitS = Split(item, "|")
        wsHistory.Cells(nextRowHist, 1).Value = dDate
        wsHistory.Cells(nextRowHist, 2).Value = splitS(1)
        wsHistory.Cells(nextRowHist, 3).Value = splitS(2)
        wsHistory.Cells(nextRowHist, 4).Value = splitS(0)
        nextRowHist = nextRowHist + 1
    Next item
      
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
      
    MsgBox "Plannings générés et enregistrés dans l'historique.", vbInformation, "Succès"
End Sub
  
' =========================================================================
' ROUTINES LOGICIELLES SECONDAIRES CENTRALISÉES
' =========================================================================
  
Sub PreparerFeuilleSO_Dual(ws As Worksheet, chosenDate As String, showNames As Boolean)
    On Error Resume Next
    ws.VPageBreaks.Clear
    On Error GoTo 0
     
    ws.Range("A7:Z120").ClearContents
    ws.Range("A8:Z120").Interior.ColorIndex = xlNone
    
    With ws.Range("A8:Z120").Font
        .Size = 11
        .Name = "Arial"
    End With
    ws.Rows("8:120").RowHeight = 18
     
    With ws.PageSetup
        .PaperSize = xlPaperA3
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
     
    ws.Cells(5, 1).Value = chosenDate
    ws.Cells(5, 7).Value = chosenDate
     
    Call InjecterHeaders(ws, 1, showNames, False)
    Call InjecterHeaders(ws, 7, showNames, False)
End Sub
  
Sub PreparerFeuilleLO_Dual(ws As Worksheet, chosenDate As String, showNames As Boolean)
    On Error Resume Next
    ws.VPageBreaks.Clear
    On Error GoTo 0
     
    ws.Range("A7:Z120").ClearContents
    ws.Range("A8:Z120").Interior.ColorIndex = xlNone
     
    With ws.PageSetup
        .PaperSize = xlPaperA3
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
     
    ws.Cells(5, 1).Value = chosenDate
    ws.Cells(5, 6).Value = chosenDate
     
    Call InjecterHeaders(ws, 1, showNames, True)
    Call InjecterHeaders(ws, 6, showNames, True)
End Sub
  
Sub InjecterHeaders(ws As Worksheet, c As Integer, showNames As Boolean, isLO As Boolean)
    ws.Cells(7, c).Value = "Taches"
    ws.Cells(7, c + 1).Value = "Operator"
    ws.Cells(7, c + 2).Value = "Back UP"
    ws.Cells(7, c + 3).Value = "Shift"
    
    If Not isLO And showNames Then
        ws.Cells(7, c + 4).Value = "Station"
    End If
End Sub
  
Sub EcrireDonnees(ws As Worksheet, col As Collection, block1Col As Integer, block2Col As Integer, dictOpStation As Object, showNames As Boolean, isLO As Boolean, allowSpill As Boolean)
    If col.Count = 0 Then Exit Sub
    Dim i As Long, r As Long, c As Integer
    Dim splitVal() As String, opName As String, stationName As String
    Dim maxRowsPerBlock As Integer, lastColOffset As Integer
     
    maxRowsPerBlock = 47
    lastColOffset = IIf(isLO, 3, 4)
      
    If allowSpill And col.Count > maxRowsPerBlock Then
        Call InjecterHeaders(ws, block2Col, showNames, isLO)
    End If
      
    For i = 1 To col.Count
        splitVal = Split(col(i), "|")
        opName = splitVal(1)
         
        If allowSpill Then
            If i <= maxRowsPerBlock Then
                r = 7 + i
                c = block1Col
            Else
                r = 7 + (i - maxRowsPerBlock)
                c = block2Col
            End If
        Else
            r = 7 + i
            c = block1Col
        End If
          
        ws.Cells(r, c).Value = splitVal(0)
        ws.Cells(r, c + 1).Value = opName
        ws.Cells(r, c + 2).Value = splitVal(3)
        ws.Cells(r, c + 3).Value = splitVal(2)
         
        If Not isLO And showNames Then
            stationName = ""
            If Not dictOpStation Is Nothing Then
                If dictOpStation.Exists(UCase(Trim(opName))) Then
                    stationName = dictOpStation(UCase(Trim(opName)))
                End If
            End If
            ws.Cells(r, c + 4).Value = stationName
        End If
         
        If r Mod 2 <> 0 Then
            ws.Range(ws.Cells(r, c), ws.Cells(r, c + lastColOffset)).Interior.Color = RGB(242, 242, 242)
        Else
            ws.Range(ws.Cells(r, c), ws.Cells(r, c + lastColOffset)).Interior.Color = RGB(255, 255, 255)
        End If
    Next i
End Sub
  
Sub RemplirPlanBanc(wsForm As Worksheet, colFinal As Collection, wsTech As Worksheet, showNames As Boolean, isAM As Boolean, dictDayStations As Object, dictOpStation As Object)
    If wsForm Is Nothing Or wsTech Is Nothing Then Exit Sub
    Dim r As Long
    Dim splitS() As String
    Dim colMTO As New Collection, colSYNO As New Collection, colSO As New Collection
    Dim item As Variant
    Dim idxMTO As Long: idxMTO = 1
    Dim idxSYNO As Long: idxSYNO = 1
    Dim idxSO As Long: idxSO = 1
    Dim isTech As Boolean, stationKey As String, assignedText As String, currentStation As String
      
    For Each item In colFinal
        splitS = Split(item, "|")
        Select Case UCase(splitS(0))
            Case "MTO": colMTO.Add item
            Case "SYNO": colSYNO.Add item
            Case "SO": colSO.Add item
        End Select
    Next item
      
    ' 1. BLOC MTO
    r = 3
    Do While Trim(wsForm.Cells(r, 1).Value) <> ""
        currentStation = Trim(wsForm.Cells(r, 1).Value)
        stationKey = UCase(currentStation)
        isTech = (UCase(Trim(wsTech.Cells(r, 2).Value)) = "TECHNICAL")
         
        If isTech Then
            wsForm.Cells(r, 2).Value = "Technical"
            wsForm.Cells(r, 3).Value = "Technical"
        Else
            wsForm.Cells(r, 2).Value = 0
            assignedText = "VIDE"
             
            If isAM Then
                If idxMTO <= colMTO.Count Then
                    splitS = Split(colMTO(idxMTO), "|")
                    assignedText = IIf(showNames, splitS(1), "OCCUPE")
                    If UCase(Trim(splitS(2))) Like "*DAY*" Then dictDayStations(stationKey) = splitS(1)
                    dictOpStation(UCase(Trim(splitS(1)))) = currentStation
                    idxMTO = idxMTO + 1
                End If
            Else
                If dictDayStations.Exists(stationKey) Then
                    assignedText = IIf(showNames, dictDayStations(stationKey), "OCCUPE")
                    dictOpStation(UCase(Trim(dictDayStations(stationKey)))) = currentStation
                Else
                    If idxMTO <= colMTO.Count Then
                        splitS = Split(colMTO(idxMTO), "|")
                        assignedText = IIf(showNames, splitS(1), "OCCUPE")
                        dictOpStation(UCase(Trim(splitS(1)))) = currentStation
                        idxMTO = idxMTO + 1
                    End If
                End If
            End If
            wsForm.Cells(r, 3).Value = assignedText
        End If
        r = r + 1
    Loop
      
    ' 2. BLOC SO
    r = 3
    Do While Trim(wsForm.Cells(r, 6).Value) <> ""
        currentStation = Trim(wsForm.Cells(r, 6).Value)
        stationKey = UCase(currentStation)
        isTech = (UCase(Trim(wsTech.Cells(r, 5).Value)) = "TECHNICAL")
         
        If isTech Then
            wsForm.Cells(r, 7).Value = "Technical"
            wsForm.Cells(r, 8).Value = "Technical"
        Else
            wsForm.Cells(r, 7).Value = 0
            assignedText = "VIDE"
             
            If isAM Then
                If idxSO <= colSO.Count Then
                    splitS = Split(colSO(idxSO), "|")
                    assignedText = IIf(showNames, splitS(1), "OCCUPE")
                    If UCase(Trim(splitS(2))) Like "*DAY*" Then dictDayStations(stationKey) = splitS(1)
                    dictOpStation(UCase(Trim(splitS(1)))) = currentStation
                    idxSO = idxSO + 1
                End If
            Else
                If dictDayStations.Exists(stationKey) Then
                    assignedText = IIf(showNames, dictDayStations(stationKey), "OCCUPE")
                    dictOpStation(UCase(Trim(dictDayStations(stationKey)))) = currentStation
                Else
                    If idxSO <= colSO.Count Then
                        splitS = Split(colSO(idxSO), "|")
                        assignedText = IIf(showNames, splitS(1), "OCCUPE")
                        dictOpStation(UCase(Trim(splitS(1)))) = currentStation
                        idxSO = idxSO + 1
                    End If
                End If
            End If
            wsForm.Cells(r, 8).Value = assignedText
        End If
        r = r + 1
    Loop
      
    ' 3. BLOC SYNO
    r = 3
    Do While Trim(wsForm.Cells(r, 11).Value) <> ""
        currentStation = Trim(wsForm.Cells(r, 11).Value)
        stationKey = UCase(currentStation)
        isTech = (UCase(Trim(wsTech.Cells(r, 8).Value)) = "TECHNICAL")
         
        If isTech Then
            wsForm.Cells(r, 12).Value = "Technical"
            wsForm.Cells(r, 13).Value = "Technical"
        Else
            wsForm.Cells(r, 12).Value = 0
            assignedText = "VIDE"
             
            If isAM Then
                If idxSYNO <= colSYNO.Count Then
                    splitS = Split(colSYNO(idxSYNO), "|")
                    assignedText = IIf(showNames, splitS(1), "OCCUPE")
                    If UCase(Trim(splitS(2))) Like "*DAY*" Then dictDayStations(stationKey) = splitS(1)
                    dictOpStation(UCase(Trim(splitS(1)))) = currentStation
                    idxSYNO = idxSYNO + 1
                End If
            Else
                If dictDayStations.Exists(stationKey) Then
                    assignedText = IIf(showNames, dictDayStations(stationKey), "OCCUPE")
                    dictOpStation(UCase(Trim(dictDayStations(stationKey)))) = currentStation
                Else
                    If idxSYNO <= colSYNO.Count Then
                        splitS = Split(colSYNO(idxSYNO), "|")
                        assignedText = IIf(showNames, splitS(1), "OCCUPE")
                        dictOpStation(UCase(Trim(splitS(1)))) = currentStation
                        idxSYNO = idxSYNO + 1
                    End If
                End If
            End If
            wsForm.Cells(r, 13).Value = assignedText
        End If
        r = r + 1
    Loop
End Sub
 
Function ShuffleCollection(col As Collection) As Collection
    Set ShuffleCollection = New Collection
    If col.Count = 0 Then Exit Function
    Dim arr() As String: ReDim arr(1 To col.Count)
    Dim i As Long, j As Long, temp As String
    For i = 1 To col.Count: arr(i) = col(i): Next i
    For i = col.Count To 2 Step -1
        j = Int((i * Rnd) + 1): temp = arr(i): arr(i) = arr(j): arr(j) = temp
    Next i
    For i = 1 To col.Count: ShuffleCollection.Add arr(i): Next i
End Function
 
Function TrierCollection(col As Collection, isLO As Boolean) As Collection
    Set TrierCollection = New Collection
    If col.Count = 0 Then Exit Function
    Dim arr() As String: ReDim arr(1 To col.Count)
    Dim i As Long, j As Long, temp As String, pI As Integer, pJ As Integer
    Dim splitI() As String, splitJ() As String
    For i = 1 To col.Count: arr(i) = col(i): Next i
      
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            splitI = Split(arr(i), "|"): splitJ = Split(arr(j), "|")
            pI = GetPriorityRank(splitI(0), isLO): pJ = GetPriorityRank(splitJ(0), isLO)
            If pI > pJ Then
                temp = arr(i): arr(i) = arr(j): arr(j) = temp
            ElseIf pI = pJ Then
                If UCase(Trim(splitI(1))) > UCase(Trim(splitJ(1))) Then
                    temp = arr(i): arr(i) = arr(j): arr(j) = temp
                End If
            End If
        Next j
    Next i
    For i = 1 To UBound(arr): TrierCollection.Add arr(i): Next i
End Function
  
Function GetPriorityRank(tache As String, isLO As Boolean) As Integer
    If isLO Then
        Select Case tache
            Case "Rework LO": GetPriorityRank = 1
            Case "Merging": GetPriorityRank = 2
            Case "Huge Order": GetPriorityRank = 3
            Case "Multi Totes Orders": GetPriorityRank = 4
            Case Else: GetPriorityRank = 99
        End Select
    Else
        Select Case tache
            Case "MTO": GetPriorityRank = 1
            Case "SYNO": GetPriorityRank = 2
            Case "SO": GetPriorityRank = 3
            Case "Shipping": GetPriorityRank = 4
            Case Else: GetPriorityRank = 99
        End Select
    End If
End Function
  
Function NormaliserRole(roleInput As String, isLO As Boolean) As String
    Dim r As String: r = UCase(Trim(roleInput))
    If isLO Then
        Select Case r
            Case "REWORK", "REWORK LO": NormaliserRole = "Rework LO"
            Case "MERGING": NormaliserRole = "Merging"
            Case "HUGE", "HUGE ORDER": NormaliserRole = "Huge Order"
            Case Else: NormaliserRole = "Multi Totes Orders"
        End Select
    Else
        Select Case r
            Case "MTO": NormaliserRole = "MTO"
            Case "SYNO": NormaliserRole = "SYNO"
            Case "SHIPPING": NormaliserRole = "Shipping"
            Case Else: NormaliserRole = "SO"
        End Select
    End If
End Function
 
Function TrierParHistorique(col As Collection, targetRole As String, dictHist As Object) As Collection
    Set TrierParHistorique = New Collection
    If col.Count = 0 Then Exit Function
      
    Dim tempCol As Collection
    Set tempCol = ShuffleCollection(col)
      
    Dim arr() As String: ReDim arr(1 To tempCol.Count)
    Dim i As Long, j As Long
    For i = 1 To tempCol.Count: arr(i) = tempCol(i): Next i
      
    Dim temp As String, rTarget As String
    Dim splitI() As String, splitJ() As String
    Dim nomI As String, nomJ As String
    Dim countI As Long, countJ As Long
    Dim keyI As String, keyJ As String
      
    rTarget = UCase(Trim(targetRole))
      
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            splitI = Split(arr(i), "|"): nomI = UCase(Trim(splitI(0)))
            splitJ = Split(arr(j), "|"): nomJ = UCase(Trim(splitJ(0)))
              
            keyI = nomI & "|" & rTarget
            keyJ = nomJ & "|" & rTarget
              
            countI = 0: If dictHist.Exists(keyI) Then countI = dictHist(keyI)
            countJ = 0: If dictHist.Exists(keyJ) Then countJ = dictHist(keyJ)
              
            If countI > countJ Then
                temp = arr(i): arr(i) = arr(j): arr(j) = temp
            End If
        Next j
    Next i
    For i = 1 To UBound(arr): TrierParHistorique.Add arr(i): Next i
End Function
