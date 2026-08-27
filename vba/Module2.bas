Option Explicit

Sub Clean()
    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("Effacer toutes les données (lignes 2 à la fin) de la feuille """ & ActiveSheet.Name & """ ?" & vbCrLf & _
                      "Cette action est irréversible.", vbYesNo + vbExclamation, "Confirmation")
    If reponse = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    ActiveSheet.Rows("2:" & ActiveSheet.Rows.Count).ClearContents

    Application.ScreenUpdating = True
End Sub
