# Kamerix – Phase 5, lot 5.4a : OpenScreen, remplissage des listes, référentiel (31/10/2026)

## Principe des listes
Les listes s'affichent dans les userforms et sont chargées **directement depuis la base** :
`lstItems.Column = AdoRs.GetRows(...)` (la propriété Column attend l'orientation colonne × ligne de GetRows).
Pas de feuille intermédiaire. Les en-têtes de colonnes sont des libellés posés au-dessus de la liste.

## Contenu
| Fichier | Rôle |
|---|---|
| modKamerixListFill.bas | **nouveau** : FillListControl (ListBox ou ComboBox), RecordsetToColumns (GetRows, Null → "", formats) |
| modKamerixScreens.bas | **nouveau** : procédure unique **OpenScreen(nom)** : connexion, droits, ouverture du userform |
| ZmodFormBuilder.bas | + **frmReference** (marques, modèles, villes, quartiers, commissariats) ; 7 formulaires |
| modTestListFill.bas | **nouveau** : TestListFill (8 contrôles) |

## Écrans disponibles par OpenScreen
`OpenScreen "REFERENTIEL"`, `"COMPTES"`, `"MOT_DE_PASSE"`, `"CONNEXION"`, `"DECONNEXION"`
(fenêtre Exécution). Sans session, OpenScreen propose d'abord la connexion.

## Installation
1. `ReplaceKamerixModules` : les 4 fichiers.
2. Si tu avais déjà importé la version précédente : supprimer les modules **modKamerixListBuffer** et
   **modTestListBuffer**, et les feuilles masquées **kxList_…** s'il y en a.
3. `BuildKamerixForms` (7 formulaires, dont frmReference).
4. **Débogage > Compiler VBAProject**.
5. `TestListFill` (8/8).
6. Essais manuels : `OpenScreen "REFERENTIEL"` (connexion administrateur) :
   - Marques : ajouter « Toyota », « Nissan » ; doublon « toyota » → message rouge ;
   - Modèles : marque Toyota, ajouter « Corolla », « Yaris » ; renommer ; supprimer ;
   - Villes : « Douala » ; Quartiers : ville Douala, « Akwa » ; Commissariats : ville Douala, « Commissariat du 1er » ;
   - supprimer la ville Douala → refusé (utilisée) ;
   - en-têtes de colonnes alignés sur la liste ;
   - connecté en Vendeur : Ajouter, Renommer, Supprimer grisés.
