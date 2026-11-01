# Kamerix – Phase 5, lot 5.5 : documents imprimables et PDF (31/10/2026)

## Ce que fait le lot
| Document | Format | Contenu |
|---|---|---|
| **Facture** FAC-… | A4 entier | en-tête de l'entreprise, client (nom, adresse, téléphone, pièce), désignation figée de la voiture, kilométrage, HT, TVA 19,25 %, TTC, **montant en lettres** (« Arrêtée la présente facture à la somme de … toutes taxes comprises »), règlements reçus (REC-…), signature |
| **Reçu** REC-… | A4, **2 exemplaires pliables** | moitié haute : exemplaire client ; moitié basse : souche de l'entreprise. Client, somme en chiffres et en lettres, mode, objet (voiture), prix / total encaissé / reste à payer, facture citée sur le reçu qui solde la vente, signatures client et caissier |
| **Reçu de remboursement** RMB-… | A4, 2 exemplaires | même structure : sommes versées / retenue (frais du CV…) / montant remboursé, motif de l'annulation, « Le client (bon pour réception) » |

- **Original puis duplicata** : la 1ʳᵉ édition est l'original ; chaque réédition porte en rouge « DUPLICATA n° N – date heure ».
- **Facture annulée** : imprimable, avec « ANNULÉE » en filigrane.
- **PDF** : `Documents\AAAA\Factures`, `\Recus` ou `\Remboursements`, à côté du classeur (dossiers créés automatiquement).
  Original : `FAC-2026-00001.pdf` ; duplicata : `FAC-2026-00001_DUP2.pdf`. Un fichier existant n'est **jamais écrasé** (« (2) » ajouté).
- **Journal** : chaque édition est inscrite dans `JournalDocument` (type, pièce, PDF ou IMPRESSION, duplicata oui/non, utilisateur, date et heure).
- **Imprimer** = PDF enregistré **et** envoyé à l'imprimante par défaut de Windows. **PDF** = PDF enregistré puis ouvert.

## Écrans
| Où | Nouveauté |
|---|---|
| frmSales | bouton **Documents…** (vente sélectionnée) |
| frmSaleDocuments (**nouveau**) | liste facture, reçus et remboursements de la vente (type, numéro, date, montant, nombre d'éditions, remarque) ; boutons **Imprimer** et **PDF** ; double-clic = PDF |
| frmSalePayment | après l'encaissement : « Imprimer le reçu maintenant ? », puis « Imprimer la facture maintenant ? » si la vente est soldée |
| frmSaleCancel | après l'annulation avec remboursement : « Imprimer le reçu de remboursement maintenant ? » |

Droits : facture → lecture sur **Factures** ; reçus → lecture sur **Paiements** (les 4 rôles les ont par défaut).

## En-tête des documents (table Parametre)
Créés **automatiquement avec des valeurs d'exemple** au lancement (`StartKamerix`) ou à la première édition ; un paramètre existant n'est jamais remplacé.

| Clé | Valeur d'exemple | À remplacer par |
|---|---|---|
| SOCIETE_NOM | MA SOCIETE AUTO SARL | la raison sociale |
| SOCIETE_ADRESSE | Rue de l'Exemple, Akwa - BP 0000 Douala | l'adresse |
| SOCIETE_TELEPHONE | +237 6 00 00 00 00 | le téléphone |
| SOCIETE_EMAIL | contact@exemple.cm | l'e-mail (ou vide) |
| SOCIETE_NIU | M000000000000X | le NIU |
| SOCIETE_RCCM | RC/DLA/2026/B/0000 | le RCCM |
| SOCIETE_LOGO | (vide) | chemin complet d'une image (facultatif), par exemple `D:\TS-AUTO-PRO\Kamerix Pro\logo.png` |
| PIED_PAGE | Merci de votre confiance. | mentions libres du pied de facture |

Pour l'instant, modification directe dans Access (table Parametre, colonne Valeur) ; un écran des paramètres viendra avec l'administration.

## Fichiers (6)
| Fichier | Changement |
|---|---|
| modKamerixAmountWords.bas | **nouveau** : AmountInWords, NumberInWords, CapitalizeFirst (orthographe traditionnelle, jusqu'à 999 999 999 999) |
| modKamerixDocuments.bas | **nouveau** : ProduceDocument, DocumentData, DocumentRows, EnsureDocumentSettings, CompanySetting, EditionCount, PaymentIdOfReceipt, CanReadDocument |
| modTestDocuments.bas | **nouveau** : TestDocuments (48 contrôles) |
| modKamerixScreens.bas | OfferDocument (proposition d'impression après une opération) |
| modKamerixLogin.bas | StartKamerix appelle EnsureDocumentSettings |
| ZmodFormBuilder.bas | + frmSaleDocuments, bouton Documents… de frmSales, questions d'impression dans frmSalePayment et frmSaleCancel (**18 formulaires**) |

Aucun changement de schéma : la base reste en version 5 (la table JournalDocument existait déjà).

## Installation
1. `ReplaceKamerixModules` : les 6 fichiers.
2. `BuildKamerixForms` (18 formulaires).
3. **Débogage > Compiler VBAProject**.
4. Base sans utilisateur : `TestDocuments` (**48/48**), puis `TestSales` (46/46) et `TestSalesScreens` (14/14).

Prérequis : une imprimante installée dans Windows, au besoin « Microsoft Print to PDF » (Excel en a besoin pour la mise en page et l'export PDF).
Le test crée ses PDF dans `%TEMP%\KamerixTestDocs` et supprime ce dossier à la fin.

## Essais manuels (après les tests)
1. Base neuve : supprimer `Database\Kamerixdb.accdb`, `TestDatabaseCreation`, `StartKamerix` (administrateur), puis refaire rapidement les étapes 1 à 3 de PARCOURS_VENTE.md (référentiel, tiers avec **téléphone et pièce d'identité** pour l'acheteur, achat).
2. Dans Access, table **Parametre** : vérifier que les 8 clés SOCIETE_* et PIED_PAGE existent ; changer SOCIETE_NOM pour voir l'effet.
3. `OpenScreen "VENTES"` → Nouvelle vente 5 000 000 → Encaisser **2 000 000** Espèces → répondre **Oui** à « Imprimer le reçu ? ».
   **Attendu** : impression ; PDF `Documents\2026\Recus\REC-2026-00001.pdf`. Ouvrir le PDF : deux exemplaires, pointillé au milieu,
   « Deux millions de francs CFA », reste à payer 3 000 000.
4. Encaisser le solde **3 000 000** → « Imprimer le reçu ? » **Non** → « Imprimer la facture ? » **Non**.
5. Sélectionner la vente (filtre Soldées) → **Documents…** : 3 lignes (facture, 2 reçus). Facture → **PDF**.
   **Attendu** : le PDF s'ouvre ; TVA 19,25 %, TTC 5 000 000, « … cinq millions de francs CFA toutes taxes comprises », 2 règlements.
6. Facture → **PDF** une 2ᵉ fois. **Attendu** : `FAC-2026-00001_DUP2.pdf` avec « DUPLICATA n° 2 » ; colonne Éditions = 2.
7. Parcours d'annulation (étape 10 de PARCOURS_VENTE.md) → « Imprimer le reçu de remboursement ? » **Oui**.
   **Attendu** : RMB-2026-00001 : versé 2 500 000, retenue 60 000, remboursé 2 440 000.
   Puis Documents… de cette vente → facture → PDF : mention **ANNULÉE** en filigrane.
8. Imprimer un reçu sur papier, le plier au pointillé : les deux moitiés doivent être égales (sinon, me donner l'écart en mm).
9. Connecté en Vendeur : Documents… et PDF possibles (lecture), Encaisser grisé.
