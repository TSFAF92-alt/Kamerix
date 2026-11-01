# Kamerix – Journal de reprise

Source : résumé de cadrage V6 (30/10/2026, schéma v5), puis code complet reçu le 31/10/2026.
Résumé Word à jour : **Projet_vente_location_vehicules_resume_V11.docx** (01/11/2026, phase 5 terminée).

## POINT DE REPRISE (01/11/2026, matin)
- **Phase 5 terminée** : lot 5.6 validé (TestSettings 41/41, TestDocuments 49/49, PARCOURS_PARAMETRES) ; 555 contrôles.
- Message de bienvenue ajouté à StartKamerix (modKamerixLogin).
- **La phase 6 se poursuit dans Claude Code** (voir CLAUDE.md) : d'abord trancher acompte minimum, nombre de
  mensualités et pénalités de retard ; puis vente échelonnée, réservation, prorogations, charges.
- Pour reprendre : l'archive Kamerix_EN_20261101_PHASE5.zip (code, ce journal, résumé V11, CLAUDE.md).

## Point de reprise précédent (01/11/2026, nuit)
- **Lot 5.6 (paramétrage de l'entreprise par écran) livré** : l'utilisateur importe les 23 fichiers, ajoute à la main
  Option Private Module dans ZmodImportTool, BuildKamerixForms (19 formulaires), compile, TestSettings (41) et
  TestDocuments (49), puis PARCOURS_PARAMETRES.md (7 étapes), et revient avec les résultats.
- Ensuite : phase 6 (vente échelonnée, réservation, prorogations, charges).
- Pour reprendre : résumé V10, ce journal et l'archive Kamerix_EN_20261101_LOT56.zip.

## Point de reprise précédent (01/11/2026)
- **Lot 5.5 validé** : TestDocuments 48/48 (513 contrôles au total) et parcours PARCOURS_DOCUMENTS.md (13 étapes) réussi,
  rendu des PDF jugé parfait. **Phase 5 terminée : première version utilisable.**
- Suite : renseigner les vrais paramètres SOCIETE_* ; puis phase 6 (vente échelonnée, réservation, prorogations, charges).

## Ancien point de reprise (31/10/2026, nuit)
- Parcours manuel PARCOURS_VENTE.md (11 étapes) : **validé** par l'utilisateur ; lot 5.4c terminé.
- **Lot 5.5 (documents) livré** : l'utilisateur installe les 6 fichiers, lance BuildKamerixForms (18 formulaires),
  compile, lance TestDocuments (48 contrôles) puis les essais manuels de LISEZMOI_LOT_5_5.md, et revient avec les résultats.
- Ensuite : renseigner les vrais paramètres SOCIETE_* ; puis phase 6 (vente échelonnée, réservation, prorogations, charges).
- Pour reprendre : fournir le résumé V9, ce journal et l'archive Kamerix_EN_20261031_LOT55.zip.

## Conventions (décidées le 31/10/2026, complétées le 01/11/2026)
- **Tous les modules .bas ont Option Private Module** (exigence de l'utilisateur, 01/11/2026) : macros lancées depuis
  la fenêtre Exécution ou par F5, plus depuis Alt+F8. ZmodImportTool ne se remplace pas lui-même (modification à la main).
- Nom unique du projet : **Kamerix** (modules modKamerix*, base Kamerixdb.accdb).
- **Code en anglais** : procédures, propriétés, classes, modules, constantes, enums, variables.
- **Restent en français** : schéma Access et SQL, codes de statut stockés (EN_STOCK…),
  messages affichés, commentaires.
- Classes : Car (ex-Voiture), Party (ex-Tiers), DataAccess (ex-DonneesAccess).
- Correspondance complète : CORRESPONDANCE_NOMS.md.
- Outils de développement : ZmodExportTool (ExportKamerixModules) et ZmodImportTool
  (ReplaceKamerixModules : remplace ou ajoute N modules choisis, sauvegarde préalable).

## État au 31/10/2026
- Phase 3 (fondations) : faite.
- Phase 4 (sécurité) : rôles, modules, permissions, droits spéciaux (schéma v5), modKamerixCrypto
  (PBKDF2 CNG, **DEFAULT_ITERATIONS = 310 000**, décision prise), modKamerixSecurity,
  modKamerixSession et leurs tests : faits, tous les tests réussis avant le renommage.
- TestSession : 59 contrôles réussis (capture du 31/10).
- **Renommage anglais et Kamerix + lot 1 : importés avec ZmodImportTool et validés.**
  Tous les tests réussis : base 8, connexion 5, SQL 25, Car 9, CarData 13, Party 11,
  DataAccess 21, sécurité/seed 9, crypto 13, sécurité 59, Session 59.
- Rappel : la fenêtre Exécution ne garde que les 200 dernières lignes ; lancer les tests un par un
  pour lire un rapport complet.

## Écran de connexion (lot 2, décisions du 31/10/2026)
- Rehash automatique après une connexion réussie : le hash suit le réglage d'itérations en vigueur,
  à la hausse comme à la baisse (comparaison dans modKamerixSecurity ; NeedsRehash de
  modKamerixCrypto, qui ne gère que la hausse, n'est pas utilisé).
- Annuler ou la croix : l'écran se ferme sans session. Mot de passe provisoire : renoncer au
  changement ferme la session.
- Userforms construits par ZmodFormBuilder (BuildKamerixForms) ; appelés par leur nom
  (VBA.UserForms.Add) depuis modKamerixLogin, qui se compile sans eux.
- Premier administrateur : connecté directement après sa création.

## Gestion des comptes (lot 3, décisions du 31/10/2026)
- Mots de passe provisoires **générés** (format Kx7m-Pq4z-R2wt, sans caractères ambigus) et affichés une
  seule fois à l'administrateur ; l'utilisateur doit les changer à la connexion.
- **Code de secours** : créé et affiché une fois à la création du premier administrateur ; stocké haché dans
  Parametre (CODE_SECOURS_HASH, CODE_SECOURS_SEL) ; utilisable depuis frmLogin (« Mot de passe oublié ? »)
  pour un compte Administrateur ; consommé à l'usage (nouveau code affiché) ; renouvelable dans frmUsers.
- Garde-fous : dernier administrateur actif non désactivable et non rétrogradable ; on ne peut ni désactiver
  son propre compte ni changer son propre rôle.
- RefreshSession relit Actif et DoitChangerMdp (anomalie 5 corrigée).
- Règle du mot de passe affichée avec accents. Lancement automatique de StartKamerix : plus tard, avec le ruban.

## Phase 5 – règles de vente (décidées le 31/10/2026)
- Vente comptant : plusieurs paiements possibles jusqu'au solde ; un reçu par paiement.
- Prix convenu TTC ; facture : HT = arrondi(TTC / (1 + taux)), taxe = TTC − HT.
- TVA 19,25 % (paramètre TAUX_TAXE) ; devise FCFA (DEVISE), 0 décimale (DECIMALES_DEVISE) ;
  arrondi moitié vers le haut (RoundAmount), jamais l'arrondi bancaire de Round.
- Modes de paiement : Espèces, Mobile Money, Virement bancaire (MODES_PAIEMENT, séparés par ;).
- Pièces : FAC-AAAA-NNNNN, REC-…, RMB-… ; compteur par type et par année, attribué dans la
  transaction qui enregistre la pièce.
- Lieux : VILLE (sans parent), QUARTIER (parent ville obligatoire), COMMISSARIAT (parent ville facultatif).

## Achat (lot 5.2, 31/10/2026)
- Droit : écriture sur Voitures. Vendeur existant et marqué vendeur ; voiture en stock sans achat préalable ;
  prix > 0 ; commission seulement avec un intermédiaire et inférieure au prix ; date jamais future.
- Achat réglé comptant et soldé à sa création ; voiture et mouvement en une transaction ; validation complète
  avant la première écriture (un service ne laisse rien dans la transaction d'un appelant en cas d'erreur).
- Coût de revient = prix d'achat + commission d'achat + charges (décidé le 31/10/2026).

## Vente comptant (lot 5.3, décisions du 31/10/2026)
- Facture émise **au solde** (dernier paiement), une seule par vente, pour le prix total TTC.
- Annulation (avant signature du CV) : remboursement de la totalité des sommes versées ; si l'acheteur se
  **désiste**, les frais du CV déjà payés restent à sa charge (déduits du remboursement).
- Vente sous le coût de revient acceptée, avec avertissement (marge négative affichée avant validation).

## Écrans (lot 5.4, décisions du 31/10/2026)
- Listes affichées dans des **userforms**, chargées **directement depuis la base** :
  ListBox.Column = AdoRs.GetRows(...) (orientation colonne × ligne), Null remplacés par "", formats appliqués
  (modKamerixListFill). Pas de feuille tampon (idée abandonnée le 31/10 : les données sont déjà dans la base).
  En-têtes de colonnes : libellés au-dessus de la liste.
- Ouverture des écrans : **une procédure unique OpenScreen(nom)** qui contrôle la connexion et les droits
  puis ouvre le userform ; le ruban (phase 8) l'appellera.
- Écrans de la première version : référentiel ; tiers et intermédiaires ; voitures et achat ; vente comptant.
- Requêtes SQL : seulement des fonctions du moteur ACE via ADO (IIf, IsNull, Trim, Switch) ; **Nz n'existe
  que dans l'application Access**.
- Constantes de module (Private Const) toujours en tête de module, avant toute procédure.
- Éviter les sous-requêtes EXISTS paramétrées (SaleableCarList renvoyait une liste vide) : préférer une jointure.

## Documents (lot 5.5, décisions du 31/10/2026)
- Construits **par code** dans un classeur temporaire (A4 portrait, une page), exportés en PDF (ExportAsFixedFormat),
  imprimés si demandé, classeur fermé sans enregistrer. Pas de feuille modèle dans l'add-in.
- Facture : A4 entier. Reçu et reçu de remboursement : **A4 avec deux exemplaires pliables** (client en haut, souche en bas ;
  moitié haute complétée à 392,25 pt pour que le pli tombe au milieu de la page).
- Montant en lettres : **orthographe traditionnelle** (vingt et un, quatre-vingts, deux cent mille, deux cents millions,
  « de francs CFA » après un nombre rond de millions).
- En-tête : paramètres SOCIETE_NOM, _ADRESSE, _TELEPHONE, _EMAIL, _NIU, _RCCM, _LOGO et PIED_PAGE ; **valeurs d'exemple**
  créées par EnsureDocumentSettings (StartKamerix et avant chaque édition), à remplacer par l'utilisateur.
- Original puis DUPLICATA n° N (2ᵉ édition et suivantes) ; facture annulée imprimable avec ANNULÉE en filigrane.
- PDF : Documents\AAAA\Factures|Recus|Remboursements ; FAC-…pdf, duplicata FAC-…_DUP2.pdf ; jamais d'écrasement.
- Chaque édition → JournalDocument (TypeDocument = FAC/REC/RMB, DocumentID = FactureID ou PaiementID, Action PDF/IMPRESSION).
- Écrans : bouton Documents… (frmSaleDocuments : Imprimer, PDF) ; impression proposée après encaissement, solde, remboursement.
- Données d'un document : Scripting.Dictionary (DocumentData), séparées de la mise en page (testables sans Excel).
  L'interface IDocument prévue au cadrage est reportée (phase 6 : contrat, échéancier, relance).
- Erreurs : ERR_DOCUMENT = vbObjectError + 1900 ; ERR_AMOUNT_WORDS = vbObjectError + 1910 (plage 1900 à 1999 : documents).
- Paramètres lus une seule fois par édition (cache), pour éviter une connexion par champ.

## Paramétrage (lot 5.6, décisions du 01/11/2026)
- Paramètres saisis dans des **userforms**, jamais dans Access. frmSettings sert à la fois d'**assistant de première
  configuration** (onglet Entreprise seul, après la création du premier administrateur ; « Plus tard » : rouvert à chaque
  connexion d'un administrateur) et d'**écran PARAMETRES** (3 onglets : Entreprise, Vente, Sécurité ; écriture Administration).
- **Édition des documents bloquée** tant que CONFIG_ENTREPRISE_FAITE <> 1 (décidé avec l'utilisateur) ; OfferDocument
  ne propose pas l'impression ; les autres rôles sont prévenus à la connexion.
- Classes **CompanyProfile** (identité, validation ; NIU et RCCM en majuscules) et **GeneralSettings** (vente et sécurité,
  SaleErrors et SecurityErrors par onglet). Service modKamerixCompany.
- Logo de l'entreprise copié dans **Ressources** (nom horodaté), chemin **relatif** stocké ; aperçu par
  **LoadPictureGDIPlus** (module ModulPictureGDIPlus de l'utilisateur, appel direct, fond blanc).
- `Images\AppLogo\Logo.png` (à côté du classeur) = **logo de Kamerix** (application), distinct du logo de l'entreprise ;
  affiché en haut de frmSettings.
- ModulPictureGDIPlus revu : MsgBox d'erreur (titre passé à la place des boutons), libération GDI+ et bitmap sur toutes
  les sorties, paramètre facultatif de couleur de fond (GDIP_WHITE_BACKGROUND).
- Devise et décimales **verrouillées** dès qu'une vente existe (même annulée) ; TVA modifiable (factures émises inchangées).
- TVA stockée avec Str$ (point décimal). Erreurs : ERR_COMPANY_INVALID 1660, ERR_SETTINGS_INVALID 1661, ERR_LOGO 1662.

## API Session (nouveaux noms)
OpenSession / CloseSession / RefreshSession, IsLoggedIn, CurrentUserId / CurrentUserLogin /
CurrentUserName, CurrentRoleId / CurrentRoleName, MustChangePassword, CanRead / CanWrite /
CanDelete, HasRight, RequireRead / RequireWrite / RequireDelete / RequireRight, ChangeMyPassword,
CreateAccount, ResetAccount, UnlockAccount. Erreurs : ERR_SESSION_NOT_LOGGED_IN (1410),
ERR_SESSION_DENIED (1411). Résultat de connexion : KxLoginResult (kxLoginSuccess,
kxLoginDenied, kxLoginLocked).

## Anomalies relevées à la relecture
1. CORRIGÉ (lot 1) : MigrateToVersion4 teste désormais la version 3 (et sort si la base est déjà en v4 ou plus).
2. CORRIGÉ (lot 1) : nouvelle constante ERR_MIGRATION = ERR_BASE + 500 (plage 1500 à 1599).
3. CORRIGÉ en partie (lot 1) : modKamerixReferences vise ThisWorkbook.VBProject et sort sans erreur
   si l'accès au modèle objet VBA n'est pas approuvé. À décider au packaging : garder ou non
   l'ajout automatique des références, et alléger la liste (Common Controls 6.0, Shell, RegExp,
   ADOX, DAO ne servent pas aujourd'hui).
4. CORRIGÉ (lot 1) : TestSession, TestSecurity, TestSecuritySeed et TestCrypto écrivent aussi le
   rapport complet dans la fenêtre Exécution (Debug.Print, Ctrl+G).
5. Session en variables de module (perdue après une réinitialisation VBA) : connu, l'interface teste IsLoggedIn.
   RefreshSession : CORRIGÉ (lot 3).

## Prochaines étapes
(Écran de connexion : FAIT et validé le 31/10/2026.)
(Gestion des comptes, code de secours, paramètres de blocage : FAIT et validé le 31/10/2026,
 tests et essais manuels. **Phase 4 (sécurité) terminée.** 290 contrôles automatiques.)
1. Avant toute livraison : MigrateToVersion5 ; mot de passe définitif de la base ; références (anomalie 3).
2. Phase 5 : vente comptant, en 5 lots :
   5.1 référentiel (**fait et validé** : TestReference 46/46 ; 336 contrôles au total) ;
   5.2 achat (**fait et validé** : TestPurchase 30/30 ; 366 contrôles au total) ; 5.3 vente comptant (**fait et validé** : TestSales 46/46 ; 412 contrôles au total) ;
   5.4 écrans, en 3 sous-lots : 5.4a OpenScreen + remplissage des listes + référentiel (**fait et validé** : TestListFill 8/8, frmReference essayé) ;
       5.4b tiers, intermédiaires, voitures et achat (**fait et validé** : TestCatalog 31/31, essais manuels 1 à 5 réussis) ; 5.4c vente comptant (TestSalesScreens 14/14 et parcours manuel PARCOURS_VENTE.md **validés**) ;
   5.5 documents (facture, reçu, PDF) : **fait et validé** (TestDocuments 48/48, PARCOURS_DOCUMENTS 13 étapes ; 513 contrôles).
   5.6 paramétrage de l'entreprise : **fait et validé** (TestSettings 41/41, TestDocuments 49/49, PARCOURS_PARAMETRES ; 555 contrôles).
   **Phase 5 terminée.**
3. Configuration réelle de l'entreprise (OpenScreen "PARAMETRES") sur la base de travail.
4. Phase 6 : vente échelonnée, réservation, prorogations, charges (décider d'abord : acompte minimum, nombre de mensualités, pénalités de retard).

## Journal des échanges
- 31/10/2026 : reprise à partir du résumé V6 ; Session relue ; code complet reçu (24 fichiers) ;
  renommage Kamerix et noms anglais livré ; anomalies notées.
- 31/10/2026 : correctif v2 : chaînes "Tiers"/"Voiture" (tables et module de droits) rétablies dans modKamerixSchema, modKamerixSession et modKamerixMigration ; TestSession 55/59 avant correctif.
- 31/10/2026 : TestSession 59/59 après correctif v2 ; lot 1 livré (anomalies 1 à 4).
- 31/10/2026 : outil ZmodImportTool livré (remplacement dynamique de modules).
- 31/10/2026 : lot 1 et outil d'import validés ; tous les tests réussis.
- 31/10/2026 : choix écran de connexion (rehash oui, Annuler sans session, userforms par VBA) ; lot 2 livré.
- 31/10/2026 : lot 2 importé ; frmLogin, frmFirstAdmin, frmChangePassword créés sans erreur.
- 31/10/2026 : lot 2 compilé sans erreur.
- 31/10/2026 : TestLogin 18/18, TestSecurity 59/59, TestSession 59/59 (lot 2).
- 31/10/2026 : rendu des 3 formulaires validé ; police irrégulière sur certains libellés → ZmodFormBuilder corrigé (police du formulaire + ApplyFonts), à relancer.
- 31/10/2026 : StartKamerix était une Function (invisible dans Alt+F8, non lançable par F5) → transformée en Sub.
- 31/10/2026 : écran de connexion validé (tests et essais manuels) ; décisions lot 3 ; lot 3 livré.
- 31/10/2026 : lot 3 : tous les tests automatiques réussis.
- 31/10/2026 : essais manuels du lot 3 réussis (code de secours, comptes, mot de passe oublié). Phase 4 terminée.
- 31/10/2026 : règles de vente décidées ; phase 5 découpée en 5 lots ; lot 5.1 (référentiel) livré.
- 31/10/2026 : résumé Word V7 produit (Projet_vente_location_vehicules_resume_V7.docx) ; total des tests validés corrigé : 290 (et non 300).
- 31/10/2026 : lot 5.1 validé (TestReference 46/46, TestParty, TestAccounts, TestSql, TestDataAccess) ; 336 contrôles.
- 31/10/2026 : lot 5.2 (achat) livré ; point ouvert : commission d'achat dans le coût de revient ?
- 31/10/2026 : décision : la commission d'achat fait partie du coût de revient (CarCostPrice et TestPurchase mis à jour).
- 31/10/2026 : lot 5.2 validé (TestPurchase 30/30, TestDataAccess 21/21, TestReference 46/46) ; 366 contrôles.
- 31/10/2026 : décisions vente comptant (facture au solde, remboursement moins frais du CV si désistement, vente à perte avec avertissement) ; lot 5.3 livré.
- 31/10/2026 : lot 5.3 validé (compilation sans erreur, TestSales 46/46, TestPurchase 30/30, TestDataAccess 21/21) ; 412 contrôles.
- 31/10/2026 : décisions écrans (userforms + feuille tampon xlVeryHidden, procédure unique OpenScreen, 3 sous-lots).
- 31/10/2026 : lot 5.4a livré (OpenScreen, tampon des listes, frmReference).
- 31/10/2026 : feuille tampon abandonnée (remarque de l'utilisateur) : ListBox.Column = GetRows ; lot 5.4a relivré.
- 31/10/2026 : TestListFill 8/8 (420 contrôles au total).
- 31/10/2026 : frmReference essayé avec succès (ajouts, doublons, suppressions refusées, mise en page) ; lot 5.4a terminé.
- 31/10/2026 : lot 5.4b livré (écrans TIERS et VOITURES, fiche voiture et achat).
- 31/10/2026 : lot 5.4b : TestCatalog 31/31, TestReference 46/46, TestPurchase 30/30, TestListFill 8/8 ; 451 contrôles.
- 31/10/2026 : essais manuels du lot 5.4b réussis (tiers, intermédiaires, achat, fiche, retrait, droits) ; lot 5.4b terminé.
- 31/10/2026 : lot 5.4c livré (écran VENTES : nouvelle vente avec marge, encaissement, CV, annulation).
- 31/10/2026 : OpenScreen : un nom d'écran inconnu affiche la liste des écrans disponibles (ScreenNames).
- 31/10/2026 : TestSalesScreens 11/14 (SaleableCarList vide : sous-requête EXISTS paramétrée) → remplacée par une jointure ; TestSales en échec car la base n'existait pas au moment du test.
- 31/10/2026 : après correctif, TestSalesScreens 14/14 et TestSales 46/46 ; 465 contrôles.
- 31/10/2026 : résumé V8 produit ; archive de fin de journée Kamerix_EN_20261031_FINAL.zip.
- 31/10/2026 (soir) : reprise sur l'archive FINAL ; parcours manuel des 11 étapes confirmé réussi par l'utilisateur ; lot 5.4c terminé.
- 31/10/2026 : décisions lot 5.5 (documents par code, en-tête en paramètres, original/duplicata, PDF classés, bouton Documents…) ;
  reçu en A4 à deux exemplaires pliables ; lettres en orthographe traditionnelle ; valeurs d'exemple pour l'en-tête.
- 31/10/2026 : lot 5.5 livré (6 fichiers, 18 formulaires, TestDocuments 48 contrôles) ; résumé V9 ; archive Kamerix_EN_20261031_LOT55.zip.
- 31/10/2026 : TestDocuments 47/48 (test : libellé voiture sans l'année) → test corrigé → 48/48.
- 01/11/2026 : parcours manuel des documents (13 étapes) réussi ; phase 5 terminée.
- 01/11/2026 : l'utilisateur demande que les paramètres de l'entreprise se saisissent par userforms après la première
  ouverture ; lot 5.6 décidé (assistant + écran PARAMETRES à 3 onglets, documents bloqués avant configuration, logo dans
  Ressources, verrou de la devise). Il dispose déjà de LoadPictureGDIPlus pour l'aperçu des PNG.
- 01/11/2026 : convention Option Private Module dans tous les .bas ; 17 modules mis en conformité.
- 01/11/2026 : lot 5.6 livré (22 fichiers, 19 formulaires, TestSettings 41 contrôles, TestDocuments 49).
- 01/11/2026 : module ModulPictureGDIPlus fourni par l'utilisateur, revu et intégré (23 fichiers) ; logo de Kamerix = Images\AppLogo\Logo.png.
- 01/11/2026 : parcours PARAMETRES : étape 1 faite (« Enregistrer » cliqué, l'étape 2 « Plus tard » n'est plus testable sur cette base) ; session perdue après une réinitialisation VBA (anomalie 5, pas un défaut) ; StartKamerix affiche désormais un message de bienvenue (nom, identifiant, rôle) après la connexion (modKamerixLogin).
- 01/11/2026 : PARCOURS_PARAMETRES réussi (étapes 3 à 7 ; étape 2 non rejouée) ; message de bienvenue validé ; lot 5.6 et phase 5 terminés ; résumé V11.
