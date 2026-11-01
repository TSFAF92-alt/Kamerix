# Kamerix – Gestion des comptes et code de secours (lot 3, 31/10/2026)

## Fichiers (7)
| Fichier | Changement |
|---|---|
| modKamerixSchema.bas | nouvelles bases : paramètres MAX_ECHECS_CONNEXION = 5 et DUREE_BLOCAGE_MINUTES = 15 |
| modKamerixSecurity.bas | UserList, RoleList, SetUserActive, ChangeUserRole (garde-fou « dernier administrateur »), GenerateTemporaryPassword, code de secours (CreateRecoveryCode, RecoveryCodeExists, ResetAdminWithRecoveryCode), règle du mot de passe avec accents ; erreurs ERR_SEC_LAST_ADMIN (1406), ERR_SEC_RECOVERY_INVALID (1407) |
| modKamerixSession.bas | ListAccounts, CreateAccountWithTemporaryPassword, ResetAccountWithTemporaryPassword, SetAccountActive, ChangeAccountRole (refusés sur son propre compte), RenewRecoveryCode ; RefreshSession relit Actif et DoitChangerMdp et ferme la session d'un compte désactivé |
| modKamerixLogin.bas | code de secours montré au premier lancement ; ManageAccounts ; ShowRecoveryForm ; ShowTemporaryPassword ; ShowRecoveryCode |
| ZmodFormBuilder.bas | 6 formulaires : frmLogin (+ lien « Mot de passe oublié ? »), frmFirstAdmin, frmChangePassword, frmRecovery, frmUsers, frmUserEdit |
| modTestSecurity.bas | test du paramètre adapté (la clé existe désormais dans les nouvelles bases) |
| modTestAccounts.bas | **nouveau** : TestAccounts (40 contrôles) |

## Installation
1. `ReplaceKamerixModules` : sélectionner les 7 fichiers.
2. `BuildKamerixForms` : les 6 formulaires sont (re)créés.
3. **Débogage > Compiler VBAProject**.
4. Supprimer `Database\Kamerixdb.accdb`, lancer `TestDatabaseCreation` (le schéma inscrit les 2 paramètres).
5. Tests : `TestAccounts` (40/40), puis `TestLogin` (18/18), `TestSecurity` (59/59), `TestSession` (59/59),
   `TestSecuritySeed` (9/9), `TestSql` (25/25).

## Essais manuels (après les tests, base sans utilisateur)
1. `StartKamerix` : créer l'administrateur → **un code de secours s'affiche** : le noter (Ctrl+C copie le message).
2. `ManageAccounts` :
   - « Nouveau compte… » : vendeur1, rôle Vendeur → le mot de passe provisoire s'affiche une fois ;
   - la liste montre vendeur1 « Mdp provisoire » ;
   - « Changer le rôle… » → Comptable ; « Désactiver » (confirmation) → « Inactif », bouton devient « Activer » ;
   - sélectionner son propre compte « admin » → « Désactiver » → message rouge (refusé) ;
   - « Nouveau mot de passe » sur vendeur1 → nouveau mot de passe provisoire affiché ;
   - « Nouveau code de secours » → nouveau code affiché (l'ancien ne marche plus).
3. `Logout`, `StartKamerix` → lien **« Mot de passe oublié ? »** → code de secours + identifiant admin +
   nouveau mot de passe → nouveau code de secours affiché → connexion avec le nouveau mot de passe.
4. Avec vendeur1 (réactivé) connecté : `ManageAccounts` → « Votre rôle ne donne pas accès… ».
