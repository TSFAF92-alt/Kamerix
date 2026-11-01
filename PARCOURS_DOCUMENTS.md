# Kamerix – Parcours complet des documents (essais manuels du lot 5.5)

Toutes les commandes se tapent dans la fenêtre Exécution (Ctrl+G), puis Entrée.
Dossier des PDF : `<dossier du classeur>\Documents\2026\` (sous-dossiers Factures, Recus, Remboursements).

## 0. Préparer une base neuve
1. Les tests automatiques d'abord (ils exigent une base sans utilisateur) : `TestDocuments` → 48/48.
2. Fermer Access s'il est ouvert. Supprimer `Database\Kamerixdb.accdb`, puis lancer `TestDatabaseCreation`.
3. `StartKamerix` → créer l'administrateur → **noter le code de secours** affiché.
4. Supprimer le dossier `Documents` à côté du classeur s'il existe déjà, pour repartir de numéros propres.

## 1. Référentiel – `OpenScreen "REFERENTIEL"`
| Liste | Ajouter |
|---|---|
| Marques | Toyota |
| Modèles (marque Toyota) | Corolla |
| Villes | Douala |
| Commissariats (ville Douala) | Commissariat du 1er |

## 2. Tiers – `OpenScreen "TIERS"`
| Fiche | Nom | Compléments |
|---|---|---|
| Nouveau | Mballa Jean | case **Vendeur** |
| Nouveau | Kouam Paul | case **Acheteur** ; téléphone **699 00 00 00** ; adresse **Bonamoussadi, Douala** ; pièce **CNI**, numéro **123456789** |

## 3. Achats – `OpenScreen "VOITURES"` → Nouvel achat…
| Voiture | Achat |
|---|---|
| Toyota / Corolla / 2015 / Grise / châssis **JTDBR32E520012345** / LT 123 AB / 85 000 km | vendeur Mballa Jean / prix 3 500 000 / lieu Douala |
| Toyota / Corolla / 2012 / Noire / châssis **JTDBR32E520099999** / LT 456 CD / 120 000 km | vendeur Mballa Jean / prix 2 000 000 / lieu Douala |

**Attendu** : les deux voitures « En stock ».

## 4. En-tête des documents (dans Access)
1. Excel peut rester ouvert ; **Access devra être fermé à la fin de cette étape**.
2. Ouvrir `Kamerixdb.accdb` dans Access (mot de passe de développement) → table **Parametre**.
3. **Attendu** : 8 clés présentes : SOCIETE_NOM, SOCIETE_ADRESSE, SOCIETE_TELEPHONE, SOCIETE_EMAIL, SOCIETE_NIU,
   SOCIETE_RCCM, SOCIETE_LOGO (vide), PIED_PAGE.
4. Modifier **SOCIETE_NOM** : `TS AUTO PRO` (ou le nom de votre choix). Facultatif : **SOCIETE_LOGO** = chemin complet d'une
   image (ex. `D:\TS-AUTO-PRO\Kamerix Pro\logo.png`).
5. Fermer la table puis Access.

## 5. Vente et premier reçu (impression proposée) – `OpenScreen "VENTES"`
1. Nouvelle vente… : voiture Corolla **JTDBR32E520012345**, acheteur Kouam Paul, prix **5 000 000** → Enregistrer.
2. Sélectionner la vente → Encaisser… : montant **2 000 000**, mode Espèces → Encaisser.
3. Message « Paiement enregistré : reçu REC-2026-00001 » + « Imprimer le reçu maintenant ? » → **Oui**.
**Attendu** : impression sur l'imprimante par défaut ; fichier `Documents\2026\Recus\REC-2026-00001.pdf`.
Ouvrir ce PDF et vérifier :
- en-tête avec **TS AUTO PRO** (et le logo si renseigné) ; « REÇU DE PAIEMENT », N° REC-2026-00001, date du jour ;
- **deux exemplaires** : « Exemplaire client » en haut, « Souche – exemplaire de l'entreprise » en bas, pointillé au milieu ;
- Reçu de : Kouam Paul ; téléphone ; La somme de : **2 000 000 FCFA** ; En lettres : **Deux millions de francs CFA** ;
- Mode : Espèces ; Objet : Paiement sur la vente : Toyota Corolla 2015 - châssis … - LT 123 AB ;
- cases : Prix de vente **5 000 000** / Total encaissé **2 000 000** / Reste à payer **3 000 000** ;
- pas de mention de facture ; zones de signature « Le client » et « Le caissier » ; « Émis par : » votre nom.

## 6. Solde et facture (impression refusée)
1. Encaisser… : montant proposé **3 000 000**, mode Mobile Money → Encaisser.
2. « Imprimer le reçu maintenant ? » → **Non**. Puis « Imprimer la facture maintenant ? » → **Non**.
**Attendu** : aucun PDF créé pour REC-2026-00002 ni FAC-2026-00001 ; la vente passe dans « Soldées ».

## 7. Écran des documents – filtre « Soldées », sélectionner la vente → Documents…
**Attendu** : 3 lignes, dans cet ordre :
| Type | Numéro | Montant | Éditions |
|---|---|---|---|
| Facture | FAC-2026-00001 | 5 000 000 | 0 |
| Reçu | REC-2026-00001 | 2 000 000 | 1 |
| Reçu | REC-2026-00002 | 3 000 000 | 0 |

## 8. Facture en PDF – sélectionner la facture → **PDF**
**Attendu** : le PDF s'ouvre (`Documents\2026\Factures\FAC-2026-00001.pdf`) ; message vert « PDF enregistré : … » ;
colonne Éditions de la facture = 1. Vérifier dans le PDF :
- « FACTURE », N° FAC-2026-00001, date ; bloc « FACTURÉ À » : Kouam Paul, adresse, Tél., **CNI n° 123456789** ;
- désignation : Véhicule Toyota Corolla 2015 - châssis JTDBR32E520012345 - immatriculation LT 123 AB - Grise ; Kilométrage 85 000 km ;
- Total HT **4 192 872** ; TVA (19,25 %) **807 128** ; Total TTC **5 000 000** ;
- « Arrêtée la présente facture à la somme de **cinq millions de francs CFA** toutes taxes comprises. » ;
- règlements reçus : REC-2026-00001 (Espèces, 2 000 000), REC-2026-00002 (Mobile Money, 3 000 000), Total réglé 5 000 000 ;
- « Pour TS AUTO PRO », « Le vendeur : … » ; pied de page : « Merci de votre confiance. » + nom, NIU, RCCM.

## 9. Duplicata – facture toujours sélectionnée → **PDF** une 2ᵉ fois
**Attendu** : nouveau fichier `FAC-2026-00001_DUP2.pdf` (l'original reste intact) ; en rouge sous la date :
« DUPLICATA n° 2 – date heure » ; Éditions = 2.
Puis double-clic sur REC-2026-00002 : PDF ouvert, cases 5 000 000 / 5 000 000 / **0**, et la ligne
« Vente soldée : facture FAC-2026-00001 émise. »

## 10. Bouton Imprimer – sélectionner REC-2026-00002 → **Imprimer**
**Attendu** : impression ; message « Document envoyé à l'imprimante. PDF : … » ; comme c'est la 2ᵉ édition,
fichier `REC-2026-00002_DUP2.pdf` et mention DUPLICATA n° 2 sur les deux exemplaires.

## 11. Annulation avec remboursement
1. Nouvelle vente : Corolla **JTDBR32E520099999**, Kouam Paul, prix **2 500 000** ; Encaisser **2 500 000** Espèces
   (répondre Non aux questions d'impression).
2. Filtre Soldées → Suivi du CV… → Déposer : Commissariat du 1er, date du jour, frais **60 000**.
3. Annuler la vente… : motif « Désistement du client », case désistement **cochée** (2 440 000), mode Espèces → confirmer.
4. « Imprimer le reçu de remboursement maintenant ? » → **Oui**.
**Attendu** : `Documents\2026\Remboursements\RMB-2026-00001.pdf` : « REÇU DE REMBOURSEMENT », Remboursé à : Kouam Paul,
**2 440 000 FCFA**, « Deux millions quatre cent quarante mille francs CFA », Motif : Désistement du client,
cases Sommes versées 2 500 000 / Retenue 60 000 / Montant remboursé 2 440 000, « Le client (bon pour réception) ».
5. Filtre Annulées → Documents… : facture FAC-2026-00002 avec remarque « Annulée » → **PDF**.
**Attendu** : grande mention **ANNULÉE** rouge pâle, inclinée, sur la facture.

## 12. Pliage du reçu (papier)
Prendre le reçu imprimé à l'étape 5, plier au pointillé.
**Attendu** : les deux moitiés sont de même hauteur et le pointillé tombe sur le pli. Sinon, mesurer l'écart (en mm)
entre le pointillé et le milieu de la feuille et me le donner.

## 13. Droits d'un Vendeur
1. `OpenScreen "COMPTES"` → Nouveau compte : vendeur1, rôle Vendeur → noter le mot de passe provisoire.
2. `OpenScreen "DECONNEXION"`, puis `OpenScreen "VENTES"` → se connecter avec vendeur1 (changer le mot de passe).
3. Filtre Toutes → sélectionner la 1ʳᵉ vente → Documents… → facture → **PDF**.
**Attendu** : Documents… actif, PDF produit (DUPLICATA n° 3) ; dans l'écran des ventes, Encaisser et Annuler grisés.

## Résultat à me renvoyer
Le numéro de chaque étape réussie, et pour une étape en écart : ce qui est attendu, ce qui s'affiche
(message exact ou capture du PDF).
