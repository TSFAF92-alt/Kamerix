# Kamerix – Phase 5, lot 5.2 : achat d'une voiture (31/10/2026)

## Règles appliquées
- Droit : **écriture sur le module Voitures** (Administrateur, Responsable). Le Vendeur ne peut pas enregistrer d'achat.
- Le vendeur doit exister et être marqué « vendeur » (Party.IsSeller).
- La voiture doit être **en stock** et ne pas avoir déjà un achat (un achat annulé ne compte pas).
- Prix > 0 ; commission ≥ 0, inférieure au prix, **seulement avec un intermédiaire** ; intermédiaire et lieu existants.
- Date de l'achat : aujourd'hui par défaut, jamais dans le futur.
- Un achat est **réglé comptant et soldé dès sa création** ; l'utilisateur connecté est inscrit sur le mouvement.
- Tout ou rien : voiture et mouvement enregistrés dans la même transaction ; toutes les données sont validées
  avant la première écriture.
- Coût de revient : **prix d'achat + commission d'achat + charges** rattachées à la voiture (CarCostPrice).

## Fichiers (5)
| Fichier | Changement |
|---|---|
| modKamerixConst.bas | erreurs ERR_MOVEMENT_INVALID (1800), ERR_MOVEMENT_STATUS (1801), ERR_PURCHASE (1810) |
| Movement.cls | **nouveau** : mouvement d'achat ou de vente (données, validation, MarkSettled, Cancel) |
| DataAccess.cls | LoadMovement, LoadPurchaseOfCar, SaveMovement |
| modKamerixPurchase.bas | **nouveau** : RecordPurchase, PurchaseOfCar, CarCostPrice |
| modTestPurchase.bas | **nouveau** : TestPurchase (30 contrôles) |

## Installation
1. `ReplaceKamerixModules` : sélectionner les 5 fichiers.
2. **Débogage > Compiler VBAProject**.
3. Base sans utilisateur (pas de changement de schéma : la base actuelle convient si elle n'a pas d'utilisateur).
4. Tests : `TestPurchase` (30/30), puis `TestDataAccess` (21/21) et `TestReference` (46/46).
