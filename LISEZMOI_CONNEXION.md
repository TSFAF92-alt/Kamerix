# Kamerix – Écran de connexion (lot 2, 31/10/2026)

## Fichiers
| Fichier | Rôle |
|---|---|
| modKamerixSecurity.bas | **modifié** : remise à niveau du hash après connexion (le hash suit le réglage d'itérations en vigueur) ; `PasswordRulesText` |
| modKamerixLogin.bas | **nouveau** : `StartKamerix`, `ChangeMyPasswordInteractive`, `Logout`, `LoginResultMessage` |
| ZmodFormBuilder.bas | **nouveau** (outil) : `BuildKamerixForms` crée frmLogin, frmFirstAdmin, frmChangePassword |
| modTestLogin.bas | **nouveau** : `TestLogin` (18 contrôles) |

## Installation
1. `ReplaceKamerixModules` (ZmodImportTool) : sélectionner les 4 fichiers.
2. Lancer `BuildKamerixForms` (F5, module ZmodFormBuilder) : les 3 userforms sont créés.
3. **Débogage > Compiler VBAProject**.
4. Base sans utilisateur, puis `TestLogin` (18/18 attendu), puis `TestSecurity` et `TestSession`
   (toujours 59/59 : le comportement de la connexion ne change pas à réglage égal).

## Essais manuels (base de développement recréée, sans utilisateur)
Lancer `StartKamerix` (module modKamerixLogin) et vérifier :
1. **Premier lancement** : l'écran « Création de l'administrateur » s'ouvre.
   - Confirmation différente → message rouge, rien n'est créé.
   - Mot de passe de 7 caractères → message rouge (règle affichée sous les champs).
   - Annuler (ou la croix, ou Échap) → l'écran se ferme ; `?IsLoggedIn` dans la fenêtre Exécution → Faux.
   - Saisie correcte → l'écran se ferme ; `?IsLoggedIn` → Vrai, `?CurrentRoleName` → Administrateur.
2. `Logout`, puis `StartKamerix` : l'écran « Connexion » s'ouvre.
   - Mauvais mot de passe → « Identifiant ou mot de passe incorrect. », champ mot de passe vidé.
   - 5 échecs → message de blocage (15 minutes).
   - Bon mot de passe (avant blocage) → connecté.
   - Entrée valide, Échap annule.
3. Compte créé par l'administrateur (dans la fenêtre Exécution, connecté en admin) :
   `?CreateAccount("vendeur1", "Vendeur Test", GetRoleIdByName("Vendeur"), "Temporaire1")`
   puis `Logout`, `StartKamerix`, connexion avec vendeur1 / Temporaire1 :
   - l'écran « Changement du mot de passe » s'ouvre avec le bouton « Quitter » ;
   - Quitter → message « Le mot de passe doit être changé… », `?IsLoggedIn` → Faux ;
   - recommencer et changer le mot de passe → connecté, `?CanWrite("Ventes")` → Vrai.
4. Connecté : `ChangeMyPasswordInteractive` → écran avec « Annuler » ; changement → message de confirmation.

## Lancement automatique (facultatif, à décider)
Pour afficher la connexion à l'ouverture du classeur, ajouter dans ThisWorkbook (ZyKmxPro),
après `Call InitializeEnvironment` :

    StartKamerix

Le ruban (phase 8) proposera aussi « Se connecter », « Changer mon mot de passe » et « Se déconnecter ».
