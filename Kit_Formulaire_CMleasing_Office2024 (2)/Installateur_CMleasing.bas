Attribute VB_Name = "modInstallCMleasing"
Option Explicit

Private Const CT_MSFORM As Long = 3
Private Const CT_STANDARD_MODULE As Long = 1

Public Sub InstallerPremierFormulaireCMleasing()
    Dim wb As Workbook
    Dim projet As Object, composants As Object, nouveauForm As Object
    Dim designer As Object
    Dim cheminCode As String, codeForm As String, etape As String
    Dim choix As VbMsgBoxResult
    Dim erreurNum As Long, erreurDesc As String
    Dim fichier As Variant, ouvert As Variant, diagnostic As String

    On Error GoTo Echec
    Set wb = ThisWorkbook
    If wb.FileFormat <> xlOpenXMLWorkbookMacroEnabled Then
        MsgBox "Enregistrez une copie au format .xlsm, puis relancez l'installation.", vbExclamation, "CMleasing - Installation"
        Exit Sub
    End If
    If Len(wb.Path) = 0 Then
        MsgBox "Enregistrez le classeur .xlsm dans le dossier du kit avant l'installation.", vbExclamation, "CMleasing - Installation"
        Exit Sub
    End If

    etape = "verification des feuilles"
    If Not FeuilleExiste(wb, "Dashboard") Or Not FeuilleExiste(wb, "DATA_Temp") Or _
       Not FeuilleExiste(wb, "Historique") Or Not FeuilleExiste(wb, "Configuration") Then
        Err.Raise vbObjectError + 2103, , "Feuille requise absente : Dashboard, DATA_Temp, Historique ou Configuration."
    End If

    etape = "recherche du code de la UserForm"
    cheminCode = wb.Path & Application.PathSeparator & "Code_frmCMleasing_Statique.txt"
    If Len(Dir$(cheminCode)) = 0 Then
        fichier = Application.GetOpenFilename("Code texte (*.txt),*.txt", , "Selectionner Code_frmCMleasing_Statique.txt")
        If VarType(fichier) = vbBoolean Then Exit Sub
        cheminCode = CStr(fichier)
    End If
    If Len(Dir$(cheminCode)) = 0 Then Err.Raise vbObjectError + 2104, , "Code_frmCMleasing_Statique.txt introuvable. Extrayez le ZIP complet."
    codeForm = LireFichierTexte(cheminCode)
    If InStr(1, codeForm, "Private Sub UserForm_Initialize()", vbTextCompare) = 0 _
       Or InStr(1, codeForm, "Public Property Get DiagnosticInitialisation()", vbTextCompare) = 0 _
       Or InStr(1, codeForm, "VERSION 5.00", vbTextCompare) > 0 _
       Or InStr(1, codeForm, "Begin MSForms.", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 2101, , "Le fichier TXT choisi n'est pas le code source CMleasing du kit."
    End If

    etape = "verification de l'acces au projet VBA"
    Set projet = wb.VBProject
    Set composants = projet.VBComponents
    If composants.Count = 0 Then Err.Raise vbObjectError + 2102, , "Projet VBA inaccessible."

    etape = "confirmation du remplacement"
    If ComposantExiste(composants, "frmCMleasing") Then
        choix = MsgBox("Une UserForm frmCMleasing existe deja. La remplacer par la version du kit ?", vbYesNo + vbQuestion, "CMleasing - Remplacement")
        If choix <> vbYes Then Exit Sub
    End If

    etape = "nettoyage de l'ancienne UserForm"
    DechargerFormulaireCharge "frmCMleasing"
    SupprimerComposantSiPresent composants, "frmCMleasing_INSTALL"
    SupprimerComposantSiPresent composants, "frmCMleasing"
    SupprimerComposantSiPresent composants, "clsCMControls"

    etape = "creation de frmCMleasing"
    Set nouveauForm = composants.Add(CT_MSFORM)
    nouveauForm.Name = "frmCMleasing"

    etape = "creation des 49 controles"
    Set designer = nouveauForm.Designer
    CreerTousLesControles designer, etape
    ValiderControles designer.Controls
    If designer.Controls.Count <> 49 Then Err.Raise vbObjectError + 2110, , "Le formulaire ne contient pas 49 controles. Nombre cree : " & designer.Controls.Count

    etape = "liaison du code a la UserForm"
    nouveauForm.CodeModule.AddFromString codeForm

    etape = "configuration de l'ouverture par instance explicite"
    AssurerModuleOuverture composants
    wb.Save

    Set designer = Nothing
    Set nouveauForm = Nothing
    DoEvents

    etape = "premier affichage du formulaire"
    ouvert = Application.Run("'" & Replace(wb.Name, "'", "''") & "'!OuvrirFormulaireCMleasing")
    If CBool(ouvert) Then
        MsgBox "Installation terminee. Le formulaire et ses 49 controles sont crees; le classeur a ete enregistre.", vbInformation, "CMleasing"
    Else
        diagnostic = CStr(Application.Run("'" & Replace(wb.Name, "'", "''") & "'!LireDiagnosticCMleasing"))
        MsgBox "Le formulaire est cree, mais son initialisation a echoue :" & vbCrLf & vbCrLf & diagnostic, vbCritical, "CMleasing - Diagnostic"
    End If
    Exit Sub

Echec:
    erreurNum = Err.Number
    erreurDesc = Err.Description
    On Error Resume Next
    If Not nouveauForm Is Nothing Then composants.Remove nouveauForm
    On Error GoTo 0
    If etape = "verification de l'acces au projet VBA" Then
        MsgBox "Installation arretee : Excel refuse l'acces au projet VBA. Activez l'option d'acces au modele d'objet VBA dans le Centre de gestion de la confidentialite. Si elle est grisee, contactez le support informatique." & vbCrLf & vbCrLf & "Detail : " & erreurDesc, vbCritical, "CMleasing - Acces VBA"
    Else
        MsgBox "Installation arretee a l'etape : " & etape & vbCrLf & "Erreur " & erreurNum & " : " & erreurDesc, vbCritical, "CMleasing - Installation"
    End If
End Sub

Private Sub DechargerFormulaireCharge(ByVal nom As String)
    Dim f As Object
    On Error Resume Next
    For Each f In UserForms
        If StrComp(CStr(f.Name), nom, vbTextCompare) = 0 Then
            CallByName f, "AutoriserFermetureInstallation", VbMethod
            Unload f
            Exit For
        End If
    Next f
    On Error GoTo 0
End Sub

Private Sub ValiderControles(ByVal controles As Object)
    Dim noms As Variant, nom As Variant, ctl As Object
    noms = Array("fraSaisie", "fraRecap", "lblTitre", "lblSousTitre", "lblPersoID", "lblDateHeure", "lblDossier", "lblTache", "lblTypologie", "lblCommentaire", "lblResume", "lblAide", "lblH_Dossier", "lblH_Date", "lblH_Tache", "lblH_Typo", "lblH_Etat", "lblH_Comment", "txtPersoID", "txtDateHeure", "txtDossier", "txtCommentaire", "cboTache", "cboTypologie", "cmdAjouter", "cmdModifier", "cmdSupprimer", "cmdEnregistrer", "cmdDashboard", "cmdFermer", "lblCarteAjouteesTitre", "lblCarteAjouteesValeur", "lblCarteEnregistreesTitre", "lblCarteEnregistreesValeur", "lblCarteSupprimeesTitre", "lblCarteSupprimeesValeur", "lblCarteAttenteTitre", "lblCarteAttenteValeur", "lblGraphTitre", "lblGraphTask1", "lblGraphTask2", "lblGraphTrack1", "lblGraphTrack2", "lblGraphBar1", "lblGraphBar2", "lblGraphValue1", "lblGraphValue2", "lstRecap", "imgLogo")
    If controles.Count <> 49 Then Err.Raise vbObjectError + 2111, , "Le modele ne contient pas 49 controles."
    For Each nom In noms
        Set ctl = Nothing
        On Error Resume Next
        Set ctl = controles.Item(CStr(nom))
        On Error GoTo 0
        If ctl Is Nothing Then Err.Raise vbObjectError + 2112, , "Controle manquant : " & CStr(nom)
    Next nom
End Sub

Private Sub CreerTousLesControles(ByVal designer As Object, ByRef etape As String)
    AjouterControleDesigner designer, "Forms.Frame.1", "fraSaisie", 360, 3000, 9000, 9100, "", etape
    AjouterControleDesigner designer, "Forms.Frame.1", "fraRecap", 9680, 1680, 11320, 10480, "", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblTitre", 360, 200, 17800, 640, "CREDIT DU MAROC | LEASING & FACTORING", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblSousTitre", 360, 940, 17800, 400, "Saisie des realisations quotidiennes", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblPersoID", 360, 1560, 1700, 440, "PersoID", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblDateHeure", 6700, 1560, 1800, 440, "Date / heure", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblDossier", 360, 2240, 1700, 440, "N de dossier", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblTache", 6700, 2240, 1800, 440, "Tache", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblTypologie", 360, 2920, 1700, 440, "Typologie", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblCommentaire", 360, 3600, 1700, 440, "Commentaire", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblResume", 360, 5420, 17300, 400, "REALISATIONS DU JOUR", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblAide", 360, 14160, 17300, 560, "Selectionnez une ligne Ajoutee pour la modifier ou la supprimer. Le bouton Dashboard ouvre la feuille sans fermer le formulaire.", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblH_Dossier", 400, 8600, 2000, 400, "N de dossier", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblH_Date", 2400, 8600, 2100, 400, "Date / heure", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblH_Tache", 4500, 8600, 1400, 400, "Tache", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblH_Typo", 5900, 8600, 1800, 400, "Typologie", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblH_Etat", 7700, 8600, 1700, 400, "Etat", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblH_Comment", 9400, 8600, 7500, 400, "Commentaire", etape
    AjouterControleDesigner designer, "Forms.TextBox.1", "txtPersoID", 2160, 1500, 3900, 480, "", etape
    AjouterControleDesigner designer, "Forms.TextBox.1", "txtDateHeure", 8560, 1500, 4400, 480, "", etape
    AjouterControleDesigner designer, "Forms.TextBox.1", "txtDossier", 2160, 2180, 3900, 480, "", etape
    AjouterControleDesigner designer, "Forms.TextBox.1", "txtCommentaire", 2160, 3540, 10800, 840, "", etape
    AjouterControleDesigner designer, "Forms.ComboBox.1", "cboTache", 8560, 2180, 4400, 480, "", etape
    AjouterControleDesigner designer, "Forms.ComboBox.1", "cboTypologie", 2160, 2860, 10800, 480, "", etape
    AjouterControleDesigner designer, "Forms.CommandButton.1", "cmdAjouter", 360, 4600, 2760, 640, "Ajouter", etape
    AjouterControleDesigner designer, "Forms.CommandButton.1", "cmdModifier", 3280, 4600, 2760, 640, "Modifier", etape
    AjouterControleDesigner designer, "Forms.CommandButton.1", "cmdSupprimer", 6200, 4600, 2760, 640, "Supprimer", etape
    AjouterControleDesigner designer, "Forms.CommandButton.1", "cmdEnregistrer", 9120, 4600, 2760, 640, "Enregistrer la ligne", etape
    AjouterControleDesigner designer, "Forms.CommandButton.1", "cmdDashboard", 12040, 4600, 2760, 640, "Dashboard", etape
    AjouterControleDesigner designer, "Forms.CommandButton.1", "cmdFermer", 14960, 4600, 2760, 640, "Fermer", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblCarteAjouteesTitre", 360, 5920, 4100, 360, "AJOUTEES AUJOURD'HUI", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblCarteAjouteesValeur", 360, 6280, 4100, 500, "0", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblCarteEnregistreesTitre", 4740, 5920, 4100, 360, "ENREGISTREES AUJOURD'HUI", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblCarteEnregistreesValeur", 4740, 6280, 4100, 500, "0", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblCarteSupprimeesTitre", 9120, 5920, 4100, 360, "SUPPRIMEES AUJOURD'HUI", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblCarteSupprimeesValeur", 9120, 6280, 4100, 500, "0", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblCarteAttenteTitre", 13500, 5920, 4100, 360, "EN ATTENTE CSV", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblCarteAttenteValeur", 13500, 6280, 4100, 500, "0", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblGraphTitre", 360, 6880, 17300, 400, "DOSSIERS DU JOUR PAR ACTIVITE", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblGraphTask1", 360, 7380, 3200, 400, "Activite 1", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblGraphTask2", 360, 7980, 3200, 400, "Activite 2", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblGraphTrack1", 3700, 7440, 12400, 280, "", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblGraphTrack2", 3700, 8040, 12400, 280, "", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblGraphBar1", 3700, 7440, 20, 280, "", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblGraphBar2", 3700, 8040, 20, 280, "", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblGraphValue1", 16300, 7380, 1300, 400, "0", etape
    AjouterControleDesigner designer, "Forms.Label.1", "lblGraphValue2", 16300, 7980, 1300, 400, "0", etape
    AjouterControleDesigner designer, "Forms.ListBox.1", "lstRecap", 360, 9040, 17300, 4880, "", etape
    AjouterControleDesigner designer, "Forms.Image.1", "imgLogo", 360, 200, 1640, 1440, "", etape
End Sub

Private Sub AjouterControleDesigner(ByVal designer As Object, ByVal progID As String, ByVal nom As String, _
                                    ByVal leftTwips As Long, ByVal topTwips As Long, _
                                    ByVal widthTwips As Long, ByVal heightTwips As Long, _
                                    ByVal titre As String, ByRef etape As String)
    Dim ctl As Object
    etape = "creation du controle " & nom
    Set ctl = designer.Controls.Add(progID, nom, True)
    ctl.Left = leftTwips / 20: ctl.Top = topTwips / 20
    ctl.Width = widthTwips / 20: ctl.Height = heightTwips / 20
    If Len(titre) > 0 Then
        On Error Resume Next
        ctl.Caption = titre
        On Error GoTo 0
    End If
    StyliserControle ctl, nom, progID
End Sub

Private Sub StyliserControle(ByVal ctl As Object, ByVal nom As String, ByVal progID As String)
    On Error Resume Next
    ctl.Font.Name = "Arial": ctl.Font.Size = 9
    If InStr(1, progID, "Label", vbTextCompare) > 0 Then
        ctl.BackStyle = 0: ctl.ForeColor = RGB(35, 49, 66)
    ElseIf InStr(1, progID, "TextBox", vbTextCompare) > 0 Then
        ctl.BackColor = RGB(255, 255, 255): ctl.ForeColor = RGB(35, 49, 66)
        If nom = "txtPersoID" Or nom = "txtDateHeure" Then ctl.Locked = True
        If nom = "txtCommentaire" Then ctl.MultiLine = True: ctl.EnterKeyBehavior = True: ctl.ScrollBars = 2
    ElseIf InStr(1, progID, "ComboBox", vbTextCompare) > 0 Then
        ctl.Style = 2
    ElseIf InStr(1, progID, "CommandButton", vbTextCompare) > 0 Then
        ctl.Font.Bold = True: ctl.ForeColor = RGB(255, 255, 255): ctl.BackColor = RGB(0, 78, 135)
        If nom = "cmdSupprimer" Or nom = "cmdEnregistrer" Then ctl.BackColor = RGB(190, 35, 55)
        If nom = "cmdModifier" Or nom = "cmdFermer" Then ctl.BackColor = RGB(83, 101, 121)
    ElseIf InStr(1, progID, "Frame", vbTextCompare) > 0 Then
        ctl.Caption = "": ctl.BackColor = RGB(255, 255, 255): ctl.SpecialEffect = 0
    ElseIf InStr(1, progID, "ListBox", vbTextCompare) > 0 Then
        ctl.ColumnCount = 8: ctl.BoundColumn = 1: ctl.IntegralHeight = False: ctl.MultiSelect = 0
        ctl.ColumnWidths = "0 pt;66 pt;78 pt;58 pt;75 pt;60 pt;200 pt;0 pt"
        ctl.BackColor = RGB(255, 255, 255): ctl.ForeColor = RGB(38, 52, 69)
    End If
    On Error GoTo 0
End Sub

Private Function ControleExiste(ByVal controles As Object, ByVal nom As String) As Boolean
    Dim ctl As Object
    On Error Resume Next
    Set ctl = controles.Item(nom)
    ControleExiste = Not ctl Is Nothing
    On Error GoTo 0
End Function

Private Function FeuilleExiste(ByVal wb As Workbook, ByVal nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nom)
    FeuilleExiste = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function LireFichierTexte(ByVal chemin As String) As String
    Dim n As Integer
    n = FreeFile
    Open chemin For Input As #n
    LireFichierTexte = Input$(LOF(n), #n)
    Close #n
End Function

Private Function ComposantExiste(ByVal composants As Object, ByVal nom As String) As Boolean
    Dim i As Long, c As Object
    For i = 1 To composants.Count
        Set c = composants.Item(i)
        If StrComp(CStr(c.Name), nom, vbTextCompare) = 0 Then ComposantExiste = True: Exit Function
    Next i
End Function

Private Sub SupprimerComposantSiPresent(ByVal composants As Object, ByVal nom As String)
    Dim i As Long, c As Object
    For i = composants.Count To 1 Step -1
        Set c = composants.Item(i)
        If StrComp(CStr(c.Name), nom, vbTextCompare) = 0 Then composants.Remove c
    Next i
End Sub

Private Sub AssurerModuleOuverture(ByVal composants As Object)
    Dim c As Object, cm As Object, i As Long
    Dim moduleSource As String
    For i = 1 To composants.Count
        Set c = composants.Item(i)
        If c.Type = CT_STANDARD_MODULE And StrComp(CStr(c.Name), "modCMleasing", vbTextCompare) = 0 Then Set cm = c.CodeModule: Exit For
    Next i
    If cm Is Nothing Then
        Set c = composants.Add(CT_STANDARD_MODULE): c.Name = "modCMleasing": Set cm = c.CodeModule
    ElseIf cm.CountOfLines > 0 Then
        cm.DeleteLines 1, cm.CountOfLines
    End If
    moduleSource = "Option Explicit" & vbCrLf
    moduleSource = moduleSource & "Public CMInitInfo As String" & vbCrLf
    moduleSource = moduleSource & "Public CMForm As Object" & vbCrLf
    moduleSource = moduleSource & "Public Function OuvrirFormulaireCMleasing() As Boolean" & vbCrLf
    moduleSource = moduleSource & "    On Error Resume Next" & vbCrLf
    moduleSource = moduleSource & "    If Not CMForm Is Nothing Then" & vbCrLf
    moduleSource = moduleSource & "        CMForm.Show vbModeless" & vbCrLf
    moduleSource = moduleSource & "        If Err.Number = 0 Then" & vbCrLf
    moduleSource = moduleSource & "            OuvrirFormulaireCMleasing = CMForm.InitialisationOK" & vbCrLf
    moduleSource = moduleSource & "            CMInitInfo = CMForm.DiagnosticInitialisation" & vbCrLf
    moduleSource = moduleSource & "            Exit Function" & vbCrLf
    moduleSource = moduleSource & "        End If" & vbCrLf
    moduleSource = moduleSource & "        Set CMForm = Nothing: Err.Clear" & vbCrLf
    moduleSource = moduleSource & "    End If" & vbCrLf
    moduleSource = moduleSource & "    On Error GoTo EchecOuverture" & vbCrLf
    moduleSource = moduleSource & "    CMInitInfo = """"" & vbCrLf
    moduleSource = moduleSource & "    Set CMForm = UserForms.Add(""frmCMleasing"")" & vbCrLf
    moduleSource = moduleSource & "    CMForm.Show vbModeless" & vbCrLf
    moduleSource = moduleSource & "    OuvrirFormulaireCMleasing = CMForm.InitialisationOK" & vbCrLf
    moduleSource = moduleSource & "    CMInitInfo = CMForm.DiagnosticInitialisation" & vbCrLf
    moduleSource = moduleSource & "    If Not OuvrirFormulaireCMleasing Then CMForm.Hide" & vbCrLf
    moduleSource = moduleSource & "    Exit Function" & vbCrLf
    moduleSource = moduleSource & "EchecOuverture:" & vbCrLf
    moduleSource = moduleSource & "    CMInitInfo = ""Etape : creation/affichage de l'instance UserForm"" & vbCrLf & ""Erreur "" & Err.Number & "" : "" & Err.Description & "" / "" & Err.Source" & vbCrLf
    moduleSource = moduleSource & "    OuvrirFormulaireCMleasing = False" & vbCrLf
    moduleSource = moduleSource & "End Function" & vbCrLf
    moduleSource = moduleSource & "Public Function LireDiagnosticCMleasing() As String" & vbCrLf
    moduleSource = moduleSource & "    LireDiagnosticCMleasing = CMInitInfo" & vbCrLf
    moduleSource = moduleSource & "End Function" & vbCrLf
    moduleSource = moduleSource & "Public Sub Auto_Open()" & vbCrLf
    moduleSource = moduleSource & "    If Not OuvrirFormulaireCMleasing() And Len(CMInitInfo) > 0 Then MsgBox CMInitInfo, vbCritical, ""CMleasing - Diagnostic""" & vbCrLf
    moduleSource = moduleSource & "End Sub"
    cm.InsertLines 1, moduleSource
End Sub
