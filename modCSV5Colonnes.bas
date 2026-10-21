Attribute VB_Name = "modCSV6Colonnes"
Option Explicit

Public Const CSV_HEADER_6 As String = "NumeroDossier;DateHeureSaisie;Agent;Produit;Activite;Commentaire"

' Chaque appel transmet les 6 colonnes dans l'ordre demandé.
Public Function LigneCSV5(ByVal numero As String, ByVal dateHeure As Date, ByVal agent As String, _
                          ByVal produit As String, ByVal activite As String, ByVal commentaire As String) As String
    LigneCSV5 = ChampCSV(numero) & ";" & ChampCSV(Format$(dateHeure, "yyyy-mm-dd hh:nn:ss")) & ";" & _
                ChampCSV(agent) & ";" & ChampCSV(produit) & ";" & ChampCSV(activite) & ";" & ChampCSV(commentaire)
End Function

Private Function ChampCSV(ByVal valeur As String) As String
    valeur = Replace$(valeur, """", """"")
    valeur = Replace$(valeur, vbCrLf, " ")
    valeur = Replace$(valeur, vbCr, " ")
    valeur = Replace$(valeur, vbLf, " ")
    ChampCSV = """" & valeur & """"
End Function

Public Function CSV6Conforme(ByVal ligne As String) As Boolean
    Dim i As Long, guillemets As Boolean, n As Long
    For i = 1 To Len(ligne)
        If Mid$(ligne, i, 1) = """" Then
            If guillemets And i < Len(ligne) And Mid$(ligne, i + 1, 1) = """" Then
                i = i + 1
            Else
                guillemets = Not guillemets
            End If
        ElseIf Mid$(ligne, i, 1) = ";" And Not guillemets Then
            n = n + 1
        End If
    Next i
    CSV6Conforme = (n = 5 And Not guillemets)
End Function
