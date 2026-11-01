# Kamerix – Lot 5.6 : paramétrage de l'entreprise par écran (01/11/2026)

## Ce que fait le lot
- **Assistant de première configuration** : après la création du premier administrateur (et l'affichage du code de secours),
  `frmSettings` s'ouvre sur l'onglet **Entreprise** seul. Obligatoires : nom, adresse, téléphone ; facultatifs : e-mail,
  NIU, RCCM, pied de page, logo. « Plus tard » : l'assistant se rouvre à chaque connexion d'un administrateur.
- **Édition des documents bloquée** tant que l'entreprise n'est pas configurée et validée (message : « Les documents ne
  peuvent pas être édités tant que l'entreprise n'est pas configurée »). Après un encaissement, pas de proposition d'impression.
  Les autres rôles sont prévenus à la connexion.
- **Écran `OpenScreen "PARAMETRES"`** (administrateur seulement), 3 onglets ; **Enregistrer** valide les trois en une fois :
  | Onglet | Champs |
  |---|---|
  | Entreprise | nom *, adresse *, téléphone *, e-mail, NIU, RCCM, pied de page, logo (Parcourir…, Retirer, aperçu) |
  | Vente | taux de TVA (19,25 ou 19.25), modes de paiement (un par ligne), devise, décimales |
  | Sécurité | échecs avant blocage (1 à 20), durée du blocage (1 à 1 440 minutes) |
- **Logo** : copié dans `Ressources\logo_aaaammjj_hhnnss.ext` à côté du classeur ; le paramètre garde un **chemin relatif**
  (Kamerix peut être déplacé). PNG, JPG, BMP, GIF ; 5 Mo au plus. Aperçu par **votre `LoadPictureGDIPlus`**, appelée
  directement, transparence affichée sur fond blanc.
- **Logo de Kamerix** en haut à droite de frmSettings : `Images\AppLogo\Logo.png` à côté du classeur (rien ne s'affiche s'il est absent).
- **Garde-fous** : devise et décimales grisées dès qu'une vente existe ; un nouveau taux de TVA ne touche que les futures factures.
- Les valeurs d'exemple ne sont jamais proposées dans l'assistant (champs vides).

## Option Private Module
Tous les modules `.bas` l'ont désormais (17 mis en conformité). Conséquence : les macros n'apparaissent plus dans Alt+F8 ;
on les lance depuis la fenêtre Exécution (Ctrl+G : `StartKamerix`, `TestSettings`, `BuildKamerixForms`…) ou par F5 dans l'éditeur.

## Fichiers à importer avec `ReplaceKamerixModules` (23)
| Fichier | Changement |
|---|---|
| CompanyProfile.cls | **nouveau** : identité de l'entreprise, validation |
| GeneralSettings.cls | **nouveau** : réglages vente et sécurité, validation par onglet |
| modKamerixCompany.bas | **nouveau** : IsCompanyConfigured, Load/SaveCompanyProfile, Load/SaveAllSettings, ImportLogo, ResolveLogoPath, CurrencyLocked, TryParseRate, TryParseWholeNumber |
| modTestSettings.bas | **nouveau** : TestSettings (41 contrôles) |
| modKamerixDocuments.bas | blocage tant que non configurée ; logo en chemin relatif |
| modKamerixScreens.bas | écran PARAMETRES, ShowSettingsForm, CompanySetupIfNeeded, OfferDocument bloqué avant configuration |
| modKamerixLogin.bas | StartKamerix lance l'assistant si besoin |
| ZmodFormBuilder.bas | + frmSettings (onglets) ; ApplyFonts tolère le contrôle Image (**19 formulaires**) |
| modTestDocuments.bas | configuration exigée avant l'édition (**49 contrôles**) |
| ModulPictureGDIPlus.bas | **votre module, revu** : MsgBox d'erreur corrigé (le titre était passé à la place des boutons → « Type incompatible ») ; GDI+ et bitmap toujours libérés ; paramètre facultatif de couleur de fond (`GDIP_WHITE_BACKGROUND`) pour les PNG transparents ; Option Private Module ; appels existants inchangés |
| ZmodExportTool.bas, modKamerixSecurity.bas, modTestAccounts.bas, modTestCar.bas, modTestCatalog.bas, modTestListFill.bas, modTestLogin.bas, modTestPartyCarData.bas, modTestPurchase.bas, modTestReference.bas, modTestSales.bas, modTestSalesScreens.bas, modTestSecurity.bas | seulement `Option Private Module` ajouté |

**ZmodImportTool ne se remplace jamais lui-même** : ajoutez à la main la ligne `Option Private Module` sous `Option Explicit`
dans ce module (éditeur VBA), puis enregistrez.

Aucun changement de schéma (base en version 5).

## Installation
1. `ReplaceKamerixModules` : les 23 fichiers (tout le dossier Lot56 sauf ce LISEZMOI et le parcours).
2. ZmodImportTool : ajouter `Option Private Module` à la main.
3. `BuildKamerixForms` (19 formulaires).
4. **Débogage > Compiler VBAProject**.
5. Base sans utilisateur : `TestSettings` (**41/41**), puis `TestDocuments` (**49/49**).
6. Essais manuels : PARCOURS_PARAMETRES.md (7 étapes).

Si ModulPictureGDIPlus n'est pas encore dans le classeur Kamerix, il est ajouté par l'import (il est nécessaire à la compilation).
