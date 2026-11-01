# Kamerix – Phase 5, lot 5.1 : référentiel (31/10/2026)

## Règles de vente décidées
- Vente comptant : **plusieurs paiements possibles** jusqu'au solde (mouvement EN_COURS → SOLDE).
- Prix convenu **TTC** ; facture : HT = arrondi(TTC / 1,1925), taxe = TTC − HT.
- TVA **19,25 %** (paramètre TAUX_TAXE) ; devise FCFA, 0 décimale ; arrondi « moitié vers le haut ».
- Modes de paiement : **Espèces ; Mobile Money ; Virement bancaire** (paramètre MODES_PAIEMENT).
- Numéros de pièces : **FAC-2026-00001**, REC- (reçu), RMB- (remboursement) ; compteur par type et par année.

## Fichiers (10)
| Fichier | Changement |
|---|---|
| modKamerixConst.bas | statuts et types : mouvement, paiement, facture, CV, pièces, lieux ; erreurs 1210, 1600-1650, 1700 |
| modKamerixSchema.bas | nouvelles bases : TAUX_TAXE, MODES_PAIEMENT, DEVISE, DECIMALES_DEVISE |
| modKamerixValidation.bas | **nouveau** : IsValidPhone (partagé), TextOfValue |
| Party.cls | utilise IsValidPhone partagé (comportement inchangé) |
| modKamerixSettings.bas | **nouveau** : TaxRate, PaymentModes, IsValidPaymentMode, CurrencyCode, CurrencyDecimals, RoundAmount, SplitTaxIncluded |
| modKamerixNumbering.bas | **nouveau** : NextDocumentNumber (dans une transaction), FormatDocumentNumber |
| modKamerixReference.bas | **nouveau** : marques, modèles, lieux (ajout, renommage, suppression contrôlée, listes), ModelLabel, IntermediaryList |
| Intermediary.cls | **nouveau** : intermédiaire (nom, téléphone, ville) avec validation |
| DataAccess.cls | LoadIntermediary, SaveIntermediary |
| modTestReference.bas | **nouveau** : TestReference (46 contrôles) |

## Installation
1. `ReplaceKamerixModules` : sélectionner les 10 fichiers.
2. **Débogage > Compiler VBAProject**.
3. Supprimer la base, `TestDatabaseCreation` (nouveaux paramètres).
4. Tests : `TestReference` (46/46), puis `TestParty` (11/11, contrôle du téléphone déplacé),
   `TestAccounts` (40/40), `TestSql` (25/25), `TestDataAccess` (21/21).
