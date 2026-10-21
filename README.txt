ADAPTATION DU FORMULAIRE AU NOUVEAU PARAMETRAGE

Constat vérifié dans les deux fichiers fournis
- Produits : AUTO, EQ (Parametres!A2:A3).
- Associations : AUTO -> P1 à P19 ; EQ -> E1 à E4 (tblProduits, colonnes Produit / Activité).
- Les colonnes Colonne2 et Colonne1 sont vides. Elles ne déclenchent aucune saisie obligatoire.
- Les champs Type dossier, Décision et Dispo sont retirés. Le commentaire reste obligatoire à votre demande et est inclus dans le CSV.

Conditions obligatoires de saisie
1. Numéro de dossier non vide.
2. Produit sélectionné dans la liste de paramètres.
3. Activité sélectionnée et liée au produit sélectionné.
4. Commentaire obligatoire dans le formulaire et conservé dans la file temporaire.
5. Agent et horodatage sont remplis automatiquement.

CSV cible (6 colonnes, dans cet ordre)
NumeroDossier;DateHeureSaisie;Agent;Produit;Activite;Commentaire

IMPORTANT
Ce kit contient les routines adaptées et le code formulaire de référence. Le classeur joint contient un projet VBA compilé qui ne peut pas être remplacé proprement par édition de cellules : il faut importer/remplacer les modules dans Excel, puis reconstruire le UserForm avec les contrôles listés au début de UserForm1_Code_NV.txt. Le code fourni pour l'ajout temporaire attend les colonnes DATA_Temp indiquées dans ce fichier. Avant d'activer l'ouverture automatique, aligner également modCSV et modDashboard du classeur sur ce schéma.

Cette livraison n'inclut pas un classeur .xlsm final prêt à l'emploi : l'environnement actuel peut inspecter le VBA et les feuilles, mais ne peut pas compiler ni valider le projet VBA dans Excel.
