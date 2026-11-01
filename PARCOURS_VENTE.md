# Kamerix – Parcours complet d'une vente (essai manuel du lot 5.4c)

Toutes les commandes se tapent dans la fenêtre Exécution (Ctrl+G) puis Entrée.

## 0. Préparer une base neuve
1. Lancer les tests automatiques d'abord (ils exigent une base sans utilisateur) : `TestSalesScreens`, `TestSales`.
2. Supprimer `Database\Kamerixdb.accdb`, lancer `TestDatabaseCreation`.
3. `StartKamerix` → créer l'administrateur → **noter le code de secours** affiché.

## 1. Référentiel — `OpenScreen "REFERENTIEL"`
| Liste | Ajouter |
|---|---|
| Marques | Toyota |
| Modèles (marque Toyota) | Corolla |
| Villes | Douala |
| Commissariats (ville Douala) | Commissariat du 1er |

## 2. Tiers — `OpenScreen "TIERS"`
| Fiche | Nom | Rôle |
|---|---|---|
| Nouveau | Mballa Jean | case **Vendeur** |
| Nouveau | Kouam Paul | case **Acheteur** |
| Liste « Intermédiaires » → Nouveau | Ndongo Eric | ville Douala |

## 3. Achat — `OpenScreen "VOITURES"` → Nouvel achat…
Voiture : Toyota / Corolla / 2015 / Grise / châssis **JTDBR32E520012345** / immatriculation LT 123 AB / 85 000 km.
Achat : vendeur Mballa Jean / prix **3 500 000** / intermédiaire Ndongo Eric / commission **150 000** / lieu Douala.
**Attendu** : la voiture apparaît « En stock », prix d'achat 3 500 000 (coût de revient 3 650 000).

## 4. Nouvelle vente — `OpenScreen "VENTES"` → Nouvelle vente…
1. Voiture : Toyota Corolla - JTDBR32E520012345 ; acheteur : Kouam Paul. « Coût de revient : 3 650 000 » s'affiche.
2. Prix **3 000 000** → « Marge prévue : -650 000 » **en rouge**.
3. Prix **5 000 000**, sans intermédiaire → « Marge prévue : 1 350 000 » **en vert**. Enregistrer.
**Attendu** : la vente apparaît dans « En cours » : prix 5 000 000, encaissé 0, solde 5 000 000, CV « À déposer ».

## 5. Premier paiement — sélectionner la vente → Encaisser…
Montant **2 000 000** (remplacer le solde proposé), mode Espèces.
**Attendu** : message « reçu REC-2026-00001 » ; encaissé 2 000 000, solde 3 000 000.

## 6. CV avant le solde — Suivi du CV…
**Attendu** : bouton « Déposer » grisé, message « quand la vente sera soldée ». Fermer.

## 7. Paiement du solde — Encaisser…
Montant proposé 3 000 000, mode Mobile Money.
**Attendu** : « reçu REC-2026-00002 » **et** « facture FAC-2026-00001 émise (5 000 000 TTC) » ;
la vente quitte « En cours » → filtre « Soldées » : solde 0.

## 8. Dépôt du CV — (filtre Soldées) Suivi du CV…
Commissariat du 1er / date du jour / frais **75 000** → Déposer.
**Attendu** : CV « Déposé ».

## 9. Signature du CV — Suivi du CV…
Numéro **CV-2026-001** / date du jour → Signé → confirmer.
**Attendu** : CV « Signé » ; « Annuler la vente » grisé ; `OpenScreen "VOITURES"`, filtre « Vendue » : la Corolla y est.

## 10. Annulation avec désistement
1. Acheter une 2ᵉ voiture (étape 3, autre châssis, sans intermédiaire, prix 2 000 000).
2. Nouvelle vente de cette voiture à Kouam Paul, prix 2 500 000 ; Encaisser **2 500 000** (vente soldée, facture FAC-2026-00002).
3. Suivi du CV → Déposer avec frais **60 000**.
4. Annuler la vente… : motif « Désistement du client » ; case décochée → « Montant remboursé : 2 500 000 » ;
   case **cochée** → **2 440 000** ; mode Espèces → confirmer.
**Attendu** : « pièce RMB-2026-00001 » ; filtre « Annulées » : encaissé 60 000, solde vide ;
la voiture revient « En stock » (écran Voitures) et peut être revendue.

## 11. Droits d'un Vendeur
1. `OpenScreen "COMPTES"` → Nouveau compte : vendeur1, rôle Vendeur → noter le mot de passe provisoire.
2. `OpenScreen "DECONNEXION"`, puis `OpenScreen "VENTES"` → se connecter avec vendeur1 (changer le mot de passe).
**Attendu** : « Nouvelle vente » actif ; « Encaisser » et « Annuler la vente » grisés.
