# Kamerix – Phase 5, lot 5.4b : tiers, intermédiaires, voitures et achat (31/10/2026)

## Écrans (par la procédure unique OpenScreen)
| Commande | Écran | Contenu |
|---|---|---|
| `OpenScreen "TIERS"` | frmParties | liste des acheteurs/vendeurs (recherche par nom, filtre Tous/Acheteurs/Vendeurs) ou des intermédiaires ; Nouveau, Modifier (double-clic) → frmPartyEdit / frmIntermediaryEdit |
| `OpenScreen "VOITURES"` | frmCars | liste filtrée par statut (modèle, année, châssis, immatriculation, statut, prix d'achat) ; Nouvel achat, Modifier (double-clic), Retirer du stock |
| (depuis frmCars) | frmCarEdit | **nouvel achat** : voiture (marque → modèle, année, couleur, châssis, immatriculation, kilométrage) + achat (vendeur, prix, intermédiaire, commission, lieu, date) ; **fiche** d'une voiture existante : données modifiables, achat en consultation ; voiture vendue ou retirée : consultation seule |

Droits : lecture pour ouvrir (OpenScreen) ; boutons d'écriture grisés sans droit d'écriture (Tiers ou Voitures).
Les montants se saisissent avec ou sans espaces (« 3 500 000 »).

## Fichiers (8)
| Fichier | Changement |
|---|---|
| modKamerixParties.bas | **nouveau** : PartyList (recherche, rôle), Load/SavePartyChecked, Load/SaveIntermediaryChecked |
| modKamerixCars.bas | **nouveau** : CarList (statut, libellé, prix d'achat), CarStatusLabel, LoadCarChecked, SaveCarDetails, WithdrawCar |
| modKamerixListFill.bas | SelectById, SelectedIdOf, AddNoneItem |
| modKamerixValidation.bas | TryParseAmount |
| modKamerixReference.bas | BrandOfModel ; IntermediaryList avec recherche facultative |
| modKamerixScreens.bas | écrans TIERS et VOITURES |
| ZmodFormBuilder.bas | + frmParties, frmPartyEdit, frmIntermediaryEdit, frmCars, frmCarEdit (12 formulaires) |
| modTestCatalog.bas | **nouveau** : TestCatalog (31 contrôles) |

## Installation
1. `ReplaceKamerixModules` : les 8 fichiers.
2. `BuildKamerixForms` (12 formulaires).
3. **Débogage > Compiler VBAProject**.
4. Base sans utilisateur : `TestCatalog` (31/31), puis `TestReference`, `TestPurchase`, `TestListFill`.

## Essais manuels (après les tests)
1. `OpenScreen "REFERENTIEL"` : marque Toyota, modèle Corolla, ville Douala.
2. `OpenScreen "TIERS"` : créer un vendeur (case Vendeur), un acheteur ; rechercher ; filtrer ; double-clic → fiche.
   Liste Intermédiaires : créer un démarcheur (ville Douala).
3. `OpenScreen "VOITURES"` → Nouvel achat : Toyota Corolla 2015, châssis, vendeur, prix « 3 500 000 »,
   intermédiaire + commission « 150 000 » → la voiture apparaît « En stock » avec son prix d'achat.
   Erreurs à essayer : sans modèle, commission sans intermédiaire, châssis en double, date future.
4. Double-clic sur la voiture → fiche : modifier la couleur ; l'achat est grisé.
5. Retirer du stock (motif) → filtre « Retirée » ; la fiche est en consultation seule.
6. Connecté en Vendeur : écran Voitures en consultation (boutons grisés), écran Tiers modifiable.
