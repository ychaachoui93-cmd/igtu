Attribute VB_Name = "modParametresNv"
Option Explicit

' Nouveau parametrage : Parametres!A2:A3 = produits (AUTO, EQ).
' tblProduits contient les colonnes Produit et Activité.
' Les colonnes Colonne2/Colonne1 sont vides dans le fichier fourni et ne sont pas des champs de saisie.
Public Function ChargerListeProduits(ByVal combo As Object) As Boolean
    Dim ws As Worksheet, lo As ListObject, c As Range
    On Error GoTo Echec
    Set ws = ThisWorkbook.Worksheets("Parametres")
    Set lo = ws.ListObjects("tblActivites")
    combo.Clear
    For Each c In lo.ListColumns("Produit").DataBodyRange.Cells
        If Len(Trim$(CStr(c.Value))) > 0 Then combo.AddItem Trim$(CStr(c.Value))
    Next c
    ChargerListeProduits = (combo.ListCount > 0)
    Exit Function
Echec:
    MsgBox "Impossible de lire la liste des produits dans tblActivites : " & Err.Description, vbCritical, "Parametres"
End Function

Public Function ChargerListeActivites(ByVal combo As Object, ByVal produit As String) As Boolean
    Dim lo As ListObject, i As Long, p As String, a As String
    On Error GoTo Echec
    Set lo = ThisWorkbook.Worksheets("Parametres").ListObjects("tblProduits")
    combo.Clear
    For i = 1 To lo.ListRows.Count
        p = Trim$(CStr(lo.ListColumns("Produit").DataBodyRange.Cells(i, 1).Value))
        a = Trim$(CStr(lo.ListColumns("Activité").DataBodyRange.Cells(i, 1).Value))
        If StrComp(p, Trim$(produit), vbTextCompare) = 0 And Len(a) > 0 Then combo.AddItem a
    Next i
    ChargerListeActivites = (combo.ListCount > 0)
    Exit Function
Echec:
    MsgBox "Impossible de lire les activités du produit " & produit & " : " & Err.Description, vbCritical, "Parametres"
End Function

Public Function ProduitActiviteValides(ByVal produit As String, ByVal activite As String) As Boolean
    Dim lo As ListObject, i As Long
    On Error GoTo Fin
    Set lo = ThisWorkbook.Worksheets("Parametres").ListObjects("tblProduits")
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.ListColumns("Produit").DataBodyRange.Cells(i, 1).Value)), Trim$(produit), vbTextCompare) = 0 _
        And StrComp(Trim$(CStr(lo.ListColumns("Activité").DataBodyRange.Cells(i, 1).Value)), Trim$(activite), vbTextCompare) = 0 Then
            ProduitActiviteValides = True: Exit Function
        End If
    Next i
Fin:
End Function
