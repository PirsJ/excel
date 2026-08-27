Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cellScan As Range
    Set cellScan = Me.Range("C8")


    If Intersect(Target, cellScan) Is Nothing Then Exit Sub


    On Error GoTo ErrorHandler

    Application.EnableEvents = False


    Me.Unprotect Password:="123"


    If Not IsError(cellScan.Value) Then


        If Trim(CStr(cellScan.Value)) = "" Then

            cellScan.Offset(0, 1).Value = ""
            cellScan.Offset(0, 2).Value = ""
            cellScan.Offset(0, 3).Value = ""
            cellScan.Offset(0, 2).Interior.Color = xlNone
            cellScan.Offset(0, 3).Interior.Color = xlNone
            GoTo SafeExit
        End If


        TraiterScan cellScan.Value, cellScan

    End If

SafeExit:

    cellScan.NumberFormat = "@"
    cellScan.Select


    Me.Protect Password:="123"

    Application.EnableEvents = True
    Exit Sub

ErrorHandler:

    MsgBox "Erreur critique dans la feuille de scan : " & Err.description, vbCritical, "Alerte Bug"


    Me.Protect Password:="123"

    Application.EnableEvents = True
End Sub

