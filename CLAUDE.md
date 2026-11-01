# Kamerix – Consignes pour Claude Code

Add-in Excel (VBA, POO) de gestion de **vente de véhicules d'occasion** (Cameroun), base **Access .accdb**
en arrière-plan (ADO/ACE). Interface 100 % Excel : userforms, puis ruban CustomUI. Destiné à être
converti en .xlam et commercialisé. Le développeur s'initie à la POO : être **précis**, expliquer les
choix de conception, livrer du code solide et testé.

## Fichiers de référence (à lire avant tout travail)
- `KAMERIX_JOURNAL.md` : point de reprise, décisions, conventions, anomalies, journal des échanges.
  **Le tenir à jour à chaque lot** (point de reprise + une ligne datée dans le journal des échanges).
- `Projet_vente_location_vehicules_resume_V11.docx` : résumé de cadrage (schéma, règles métier, planning).
  Produire une nouvelle version (V12, V13…) aux étapes importantes.
- `CORRESPONDANCE_NOMS.md` : anciens noms français → noms anglais.
- `LISEZMOI_*.md` et `PARCOURS_*.md` : notes de livraison et essais manuels par lot (même format pour la suite).

## Environnement : important
- Claude Code **ne peut pas exécuter le VBA** (pas d'Excel ici). Le développeur importe les modules,
  compile et lance les tests sous Windows, puis renvoie les résultats. Ne jamais prétendre qu'un test passe
  sans son retour.
- Import par l'outil `ZmodImportTool` (`ReplaceKamerixModules`) : il ne se remplace pas lui-même
  (modification à la main). Lister à chaque livraison les fichiers à importer.
- Toute modification ou import réinitialise le projet VBA : la session (variables de module) est perdue.

## Encodage des fichiers de code (impératif)
- `.bas` et `.cls` en **Windows-1252 (ANSI) avec fins de ligne CRLF**. Jamais d'UTF-8 : l'import VBA
  casse les accents. Vérifier avec `file` après chaque écriture.
- Préférer `ChrW(233)` etc. pour les accents dans les chaînes affichées (pratique déjà utilisée).
- Les fichiers `.cls` gardent leur en-tête (`VERSION 1.0 CLASS`, `Attribute VB_Name`…).
- Les userforms ne sont pas des fichiers : ils sont **générés par `ZmodFormBuilder`** (`BuildKamerixForms`).
  Pour changer un formulaire, modifier le générateur.

## Conventions de code
- **Tous les modules .bas : `Option Explicit` puis `Option Private Module`** (exigence ferme).
- Code en **anglais** (procédures, classes, variables, constantes, enums).
- Restent en **français** : schéma Access et SQL, codes de statut stockés (EN_STOCK…), messages, commentaires.
- Constantes de module en tête, avant toute procédure.
- SQL uniquement via `modKamerixSql` (requêtes paramétrées `?`, transactions) et les couches d'accès.
  Éviter les sous-requêtes EXISTS paramétrées : préférer une jointure.
- Erreurs métier : `vbObjectError + n`, une plage par domaine, constantes dans `modKamerixConst`
  (1100 Car, 1200 Party, 1300 données, 1400 session, 1500 base/migration, 1600 référentiel/paramètres,
  1700 numérotation, 1800 mouvements/vente/paiement/facture/CV, 1900 documents). Choisir une plage libre
  pour la phase 6 et la noter dans le journal.
- Droits : vérifier `RequireRead/RequireWrite/RequireRight` (modKamerixSession) dans les services.
- Chaque lot livre son module de test `modTestXxx` (`CheckTrue`, rapport en MsgBox et dans la fenêtre
  Exécution) et, si des écrans changent, un `PARCOURS_XXX.md`.
- Changement de schéma : nouvelle version + procédure `MigrateToVersionN` dans `modKamerixMigration`.

## État (01/11/2026)
- Phases 1 à 5 terminées : base v5 (27 tables), sécurité complète, vente comptant, documents PDF,
  paramétrage de l'entreprise. 555 contrôles automatiques réussis.
- **Phase 6 à démarrer** : vente échelonnée (`InstallmentSale` implémentant `IContract`), réservation avec
  avance, prorogations, charges. Règles déjà décidées : voir le journal et le résumé (prorogations : 2 si marge
  ≥ 500 000, sinon 1, 30 jours max, marge recalculée à chaque demande ; 75 % payés pour y avoir droit ;
  CV signé après le dernier paiement ; une facture au total, un reçu par mensualité ; charge générale).
- **À trancher avec le développeur avant de coder** : acompte minimum, nombre maximal de mensualités,
  pénalités de retard.
- Avant toute livraison commerciale : MigrateToVersion5, mot de passe définitif de la base, références VBA,
  retrait des modules Zmod* et modTest*, ruban vérifiant `IsLoggedIn()` à chaque clic.

## Façon de travailler
- Répondre en **français**.
- Découper en lots courts, valider chaque lot (tests + parcours manuel) avant le suivant.
- Poser les questions de règles métier avant d'écrire le code ; ne pas inventer de règle.
