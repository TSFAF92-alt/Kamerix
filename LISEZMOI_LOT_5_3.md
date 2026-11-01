# Kamerix – Phase 5, lot 5.3 : vente comptant (31/10/2026)

## Cycle d'une vente comptant
| Opération (modKamerixSales) | Droit | Effet |
|---|---|---|
| StartCashSale | Ventes (écriture) | voiture en stock et achetée, acheteur marqué « acheteur » → voiture VENTE_EN_COURS, mouvement VENTE comptant EN_COURS, dossier de CV A_DEPOSER |
| RecordSalePayment | Paiements (écriture) | reçu REC-AAAA-NNNNN ; montant > 0 et ≤ solde ; mode du paramètre MODES_PAIEMENT ; **au dernier paiement** : mouvement SOLDE, **facture FAC-AAAA-NNNNN** (HT et taxe déduits du TTC), paiements rattachés à la facture |
| SubmitCv | Ventes (écriture) | vente soldée exigée ; lieu de type COMMISSARIAT ; date et frais → CV EN_COURS |
| SignCv | Ventes (écriture) | numéro du CV → CV SIGNE, voiture VENDU (définitif) |
| CancelSale | VENTE_ANNULER (+ REMBOURSER si remboursement) | CV non signé ; remboursement RMB-… de l'encaissé, **moins les frais du CV si l'acheteur se désiste** ; facture annulée si émise ; vente ANNULEE (date, motif) ; voiture remise en stock |
| LoadSaleContract | Ventes (lecture) | contrat vu par l'interface **IContract** (prix, encaissé, solde, soldé, CanReceive) |
| EstimatedMargin / SaleMargin | lecture | prix − coût de revient − commission de vente (− frais du CV) ; négative autorisée (avertissement à l'écran) |

## Fichiers (9)
| Fichier | Changement |
|---|---|
| modKamerixConst.bas | erreurs ERR_SALE (1820), ERR_PAYMENT_INVALID (1830), ERR_INVOICE_INVALID (1840), ERR_CV (1850) |
| IContract.cls | **nouveau** : interface d'une vente (polymorphisme : comptant maintenant, échelonnée en phase 6) |
| CashSale.cls | **nouveau** : vente comptant, `Implements IContract` |
| Payment.cls | **nouveau** : encaissement ou remboursement |
| Invoice.cls | **nouveau** : facture (HT, taxe, TTC ; annulation) |
| CvTracking.cls | **nouveau** : suivi du CV (Submit, MarkSigned) |
| DataAccess.cls | SavePayment, PaymentsTotal, LinkPaymentsToInvoice, LoadInvoiceOfMovement, SaveInvoice, LoadCvTrackingOfMovement, SaveCvTracking |
| modKamerixSales.bas | **nouveau** : service de vente comptant |
| modTestSales.bas | **nouveau** : TestSales (46 contrôles) |

## Installation
1. `ReplaceKamerixModules` : sélectionner les 9 fichiers (l'ordre n'a pas d'importance : on compile après).
2. **Débogage > Compiler VBAProject**.
3. Base sans utilisateur, créée avec le schéma du lot 5.1 (pas de changement de schéma).
4. Tests : `TestSales` (46/46), puis `TestPurchase` (30/30), `TestDataAccess` (21/21).
