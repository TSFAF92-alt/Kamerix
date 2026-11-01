# Kamerix – Installation du code renommé (31/10/2026)

## Remplacer des modules ensuite (outil ZmodImportTool)
1. Importer **une seule fois, à la main**, `ZmodImportTool.bas` (Fichier > Importer un fichier…).
2. À chaque livraison : lancer `ReplaceKamerixModules` (F5), sélectionner les fichiers reçus
   (autant que nécessaire), valider.
3. Lire le rapport (aussi dans la fenêtre Exécution, Ctrl+G), puis **Débogage > Compiler VBAProject**
   et relancer les tests concernés.
Chaque module remplacé est d'abord exporté dans `Backup_VBA_aaaammjj_hhnnss` à côté du classeur.
L'outil ne remplace ni lui-même, ni ThisWorkbook, ni les feuilles.

---

Le dossier `code` contient 22 modules : 19 modules standard (.bas) et 3 classes (.cls).
ThisWorkbook (ZyKmxPro) et la feuille ShMain ne changent pas : ne pas les réimporter.

## 1. Sauvegarde
Faire une copie du classeur .xlsm avant de commencer.

## 2. Supprimer les anciens modules
Dans l'éditeur VBA, clic droit > Supprimer (répondre « Non » à la proposition d'export) sur :

modKemerixConst, modKemerixDb, modKemerixSql, modKemerixSchema, modKemerixMigration,
modKemerixCrypto, modKemerixSecurite, modKemerixSession, modKemerixTests,
ModKmerixProjRefs, ZmodOutilExport, modTestConnexion, modTestSql, modTestVoiture,
modTestTiersVoitureDonnees, modTestDonnees, modTestSecuriteSeed, modTestSecurite,
modTestSession, et les classes Voiture, Tiers, DonneesAccess.

## 3. Importer les nouveaux
Fichier > Importer un fichier… : sélectionner les 22 fichiers du dossier `code`.
Puis **Débogage > Compiler VBAProject** : aucune erreur ne doit apparaître.

## 4. Base de données
Le fichier s'appelle désormais `Kamerixdb.accdb`. Le schéma ne change pas :
renommer `Database\Kemerixdb.accdb` en `Database\Kamerixdb.accdb`, ou le supprimer
et lancer `TestDatabaseCreation` pour le recréer.
TestSecurity et TestSession exigent une base sans utilisateur.

## 5. Relancer les tests (F5 dans chaque procédure)

| Ancien nom | Nouveau nom | Module |
|---|---|---|
| TestCreationBase | TestDatabaseCreation | modKamerixDb |
| TesterConnexion | TestConnection | modTestConnection |
| TestOuvertureBase | TestDatabaseOpening | modTestDatabase |
| TestSeedSecurite | TestSecuritySeed | modTestSecuritySeed |
| TestCrypto | TestCrypto | modKamerixCrypto |
| TesterSql | TestSql | modTestSql |
| TesterVoiture | TestCar | modTestCar |
| TesterTiers | TestParty | modTestPartyCarData |
| TesterVoitureDonnees | TestCarData | modTestPartyCarData |
| TesterDonnees | TestDataAccess | modTestDataAccess |
| TestSecurite | TestSecurity | modTestSecurity |
| TestSession | TestSession | modTestSession |

Le détail des nouveaux noms est dans CORRESPONDANCE_NOMS.md.
