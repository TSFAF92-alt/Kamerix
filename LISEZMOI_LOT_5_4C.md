# Kamerix – Phase 5, lot 5.4c : écran des ventes (31/10/2026)

## Écran `OpenScreen "VENTES"` (frmSales)
Liste des ventes comptant, les plus récentes d'abord, filtrée par statut (En cours par défaut, Soldées,
Annulées, Toutes) : date, client, voiture, prix, encaissé (remboursements déduits), solde, statut de la vente,
statut du CV.

| Bouton | Formulaire | Actif si |
|---|---|---|
| Nouvelle vente… | frmSaleStart : voiture (en stock **et** achetée), acheteur, prix TTC, intermédiaire, commission, lieu, date ; **coût de revient et marge prévue** affichés pendant la saisie (rouge si négative, confirmation demandée) | écriture sur Ventes |
| Encaisser… | frmSalePayment : rappel prix / encaissé / solde, montant (solde proposé), mode, date ; message avec le reçu, et la facture si la vente est soldée | écriture sur Paiements, vente en cours |
| Suivi du CV… | frmSaleCv : dépôt (commissariat, date, frais) puis signature (numéro, date) ; la signature rend la voiture « Vendue » | vente non annulée (boutons selon le statut du CV et les droits) |
| Annuler la vente… | frmSaleCancel : motif, case « désistement de l'acheteur », montant remboursé recalculé, mode de remboursement | droit VENTE_ANNULER, CV non signé |

## Fichiers (4)
| Fichier | Changement |
|---|---|
| modKamerixSales.bas | SaleRows (liste prête pour ListBox.Column, 11 colonnes), SaleableCarList, SaleLabel, SaleStatusLabel, CvStatusLabel |
| modKamerixScreens.bas | écran VENTES |
| ZmodFormBuilder.bas | + frmSales, frmSaleStart, frmSalePayment, frmSaleCv, frmSaleCancel (17 formulaires) |
| modTestSalesScreens.bas | **nouveau** : TestSalesScreens (14 contrôles) |

## Installation
1. `ReplaceKamerixModules` : les 4 fichiers.
2. `BuildKamerixForms` (17 formulaires).
3. **Débogage > Compiler VBAProject**.
4. Base sans utilisateur : `TestSalesScreens` (14/14), puis `TestSales` (46/46).

## Essais manuels (parcours complet d'une vente)
1. Référentiel (Toyota Corolla, ville, **commissariat**) ; tiers : un vendeur, un acheteur ; voiture achetée.
2. `OpenScreen "VENTES"` → Nouvelle vente : choisir la voiture, l'acheteur ; taper un prix sous le coût →
   marge rouge et confirmation ; puis un prix normal → marge verte ; enregistrer.
3. Encaisser une partie (Espèces) → reçu REC-… ; la liste montre l'encaissé et le solde.
4. Suivi du CV → « Déposer » grisé (vente non soldée).
5. Encaisser le solde (Mobile Money) → message avec la facture FAC-… ; la vente passe dans « Soldées ».
6. Suivi du CV → Déposer (commissariat, frais) → puis Signé (numéro) → la voiture est « Vendue » (écran Voitures).
7. Autre vente + un paiement → Annuler, cocher/décocher « désistement » (montant recalculé) → pièce RMB-… ;
   la voiture revient en stock.
8. Connecté en Vendeur : Nouvelle vente possible ; Encaisser et Annuler grisés.
