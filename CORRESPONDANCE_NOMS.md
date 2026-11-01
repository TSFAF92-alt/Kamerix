# Kamerix – Correspondance des noms (français → anglais)

Renommage du 31/10/2026. Le schéma Access, le SQL, les codes de statut stockés (EN_STOCK…), les messages affichés et les commentaires restent en français. Les variables locales et paramètres ont aussi été renommés ; ils ne sont pas listés ici.

## Modules et classes

| Ancien | Nouveau |
|---|---|
| modKemerixConst | modKamerixConst |
| modKemerixDb | modKamerixDb |
| modKemerixSql | modKamerixSql |
| modKemerixSchema | modKamerixSchema |
| modKemerixMigration | modKamerixMigration |
| modKemerixCrypto | modKamerixCrypto |
| modKemerixSecurite | modKamerixSecurity |
| modKemerixSession | modKamerixSession |
| modKemerixTests | modTestDatabase |
| modTestConnexion | modTestConnection |
| modTestVoiture | modTestCar |
| modTestTiersVoitureDonnees | modTestPartyCarData |
| modTestDonnees | modTestDataAccess |
| modTestSecuriteSeed | modTestSecuritySeed |
| modTestSecurite | modTestSecurity |
| ModKmerixProjRefs | modKamerixReferences |
| ZmodOutilExport | ZmodExportTool |
| DonneesAccess | DataAccess |
| Voiture | Car |
| Tiers | Party |
| Fichier Kemerixdb.accdb | Kamerixdb.accdb |

## DataAccess (ancien DonneesAccess)

| Ancien | Nouveau |
|---|---|
| ChargerVoiture | LoadCar |
| EnregistrerVoiture | SaveCar |
| ChargerTiers | LoadParty |
| EnregistrerTiers | SaveParty |
| TexteDe | TextOf |
| ExigerChassisUnique | RequireUniqueChassis |

## Party (ancien Tiers)

| Ancien | Nouveau |
|---|---|
| LONG_NOM | MAX_NAME_LENGTH |
| LONG_TELEPHONE | MAX_PHONE_LENGTH |
| LONG_TYPE_PIECE | MAX_ID_TYPE_LENGTH |
| LONG_NUMERO_PIECE | MAX_ID_NUMBER_LENGTH |
| LONG_ADRESSE | MAX_ADDRESS_LENGTH |
| Nom | Name |
| Telephone | Phone |
| TypePiece | IdType |
| NumeroPiece | IdNumber |
| Adresse | Address |
| EstAcheteur | IsBuyer |
| EstVendeur | IsSeller |
| Hydrater | Hydrate |
| DefinirId | SetId |
| Erreurs | ValidationErrors |
| Verifier | Validate |
| TelephoneValide | IsValidPhone |
| TexteDe | TextOf |
| BoolDe | BoolOf |

## Car (ancien Voiture)

| Ancien | Nouveau |
|---|---|
| LONG_COULEUR | MAX_COLOR_LENGTH |
| LONG_CHASSIS | MAX_CHASSIS_LENGTH |
| LONG_IMMATRICULATION | MAX_PLATE_LENGTH |
| ANNEE_MINIMALE | MIN_MODEL_YEAR |
| LONG_MOTIF | MAX_REASON_LENGTH |
| Statut | Status |
| MotifRetrait | WithdrawalReason |
| ModeleId | ModelId |
| Annee | ModelYear |
| Couleur | Color |
| Immatriculation | PlateNumber |
| Kilometrage | Mileage |
| Hydrater | Hydrate |
| HydraterDonnees | HydrateData |
| DefinirId | SetId |
| Erreurs | ValidationErrors |
| Verifier | Validate |
| PeutPasserA | CanTransitionTo |
| Reserver | Reserve |
| AnnulerReservation | CancelReservation |
| ConfirmerVente | ConfirmSale |
| DemarrerVenteDirecte | StartDirectSale |
| MarquerVendu | MarkAsSold |
| AnnulerVente | CancelSale |
| Retirer | Withdraw |
| Transiter | TransitionTo |
| ExigerStatut | RequireStatus |
| ExigerMotif | RequireReason |
| ExigerModifiable | RequireEditable |
| StatutValide | IsValidStatus |
| ChassisValide | IsValidChassis |
| NombreOuNull | NumberOrNull |
| TexteDe | TextOf |

## ZmodExportTool (ancien ZmodOutilExport)

| Ancien | Nouveau |
|---|---|
| ExporterModulesKemerix | ExportKamerixModules |

## modKamerixConst (ancien modKemerixConst)

| Ancien | Nouveau |
|---|---|
| STATUT_VOITURE_EN_STOCK | CAR_STATUS_IN_STOCK |
| STATUT_VOITURE_RESERVE | CAR_STATUS_RESERVED |
| STATUT_VOITURE_VENTE_EN_COURS | CAR_STATUS_SALE_IN_PROGRESS |
| STATUT_VOITURE_VENDU | CAR_STATUS_SOLD |
| STATUT_VOITURE_RETIRE | CAR_STATUS_WITHDRAWN |
| ERR_VOITURE_TRANSITION | ERR_CAR_TRANSITION |
| ERR_VOITURE_MOTIF | ERR_CAR_REASON |
| ERR_VOITURE_STATUT | ERR_CAR_STATUS |
| ERR_VOITURE_DONNEES | ERR_CAR_DATA |
| ERR_VOITURE_VERROUILLEE | ERR_CAR_LOCKED |
| ERR_TIERS_INVALIDE | ERR_PARTY_INVALID |
| ERR_DONNEES_DOUBLON | ERR_DATA_DUPLICATE |
| ERR_DONNEES_INTROUVABLE | ERR_DATA_NOT_FOUND |

## modKamerixDb (ancien modKemerixDb)

| Ancien | Nouveau |
|---|---|
| TestCreationBase | TestDatabaseCreation |

## modKamerixMigration (ancien modKemerixMigration)

| Ancien | Nouveau |
|---|---|
| MigrerVersion4 | MigrateToVersion4 |
| ColonneExiste | ColumnExists |

## modKamerixSchema (ancien modKemerixSchema)

| Ancien | Nouveau |
|---|---|
| DefineTablesReferentiel | DefineReferenceTables |
| DefineTablesPersonnes | DefinePeopleTables |
| DefineTablesMetier | DefineBusinessTables |
| DefineTablesFacturation | DefineBillingTables |
| DefineTablesSecurite | DefineSecurityTables |

## modKamerixSecurity (ancien modKemerixSecurite)

| Ancien | Nouveau |
|---|---|
| ERR_SECU_DONNEES | ERR_SEC_INVALID_DATA |
| ERR_SECU_LOGIN_EXISTE | ERR_SEC_LOGIN_EXISTS |
| ERR_SECU_INTROUVABLE | ERR_SEC_NOT_FOUND |
| ERR_SECU_MDP_INCORRECT | ERR_SEC_WRONG_PASSWORD |
| ERR_SECU_ADMIN_EXISTE | ERR_SEC_ADMIN_EXISTS |
| ERR_SECU_ROLE_INCONNU | ERR_SEC_UNKNOWN_ROLE |
| ROLE_ADMINISTRATEUR | ROLE_ADMINISTRATOR |
| PARAM_MAX_ECHECS | PARAM_MAX_FAILURES |
| PARAM_DUREE_BLOCAGE | PARAM_LOCKOUT_MINUTES |
| DEFAUT_MAX_ECHECS | DEFAULT_MAX_FAILURES |
| DEFAUT_DUREE_BLOCAGE | DEFAULT_LOCKOUT_MINUTES |
| LONG_MIN_LOGIN | MIN_LOGIN_LENGTH |
| LONG_MAX_LOGIN | MAX_LOGIN_LENGTH |
| LONG_MAX_NOM | MAX_NAME_LENGTH |
| LONG_MIN_MDP | MIN_PASSWORD_LENGTH |
| LONG_MAX_MDP | MAX_PASSWORD_LENGTH |
| RES_SUCCES | RESULT_SUCCESS |
| RES_ECHEC | RESULT_FAILURE |
| RES_BLOQUE | RESULT_LOCKED |
| KxConnexion | KxLoginResult |
| AdminExiste | AdminExists |
| CreerPremierAdmin | CreateFirstAdmin |
| CreerUtilisateur | CreateUser |
| RoleIdParNom | GetRoleIdByName |
| Connecter | AuthenticateUser |
| NbEchecsMax | MaxFailedAttempts |
| DureeBlocageMinutes | LockoutMinutes |
| ValiderMotDePasse | ValidatePasswordRules |
| ChangerMotDePasse | ChangePassword |
| ReinitialiserMotDePasse | ResetPassword |
| DebloquerUtilisateur | UnlockUser |
| DefinirIterationsPourTests | SetIterationsForTests |
| IterationsCourantes | CurrentIterations |
| InsererUtilisateur | InsertUser |
| DefinirMotDePasse | SetPassword |
| LoginDe | LoginOf |
| LoginExiste | LoginExists |
| RoleExiste | RoleExists |
| ParametreEntier | GetIntegerParameter |
| JournaliserConnexion | LogLoginAttempt |
| NomPoste | WorkstationName |
| DepenserTempsCalcul | SpendHashTime |
| kxConnSucces | kxLoginSuccess |
| kxConnRefusee | kxLoginDenied |
| kxConnBloquee | kxLoginLocked |

## modKamerixSession (ancien modKemerixSession)

| Ancien | Nouveau |
|---|---|
| ERR_SESSION_NON_CONNECTE | ERR_SESSION_NOT_LOGGED_IN |
| ERR_SESSION_REFUSE | ERR_SESSION_DENIED |
| MOD_VOITURES | MOD_CARS |
| MOD_TIERS | MOD_PARTIES |
| MOD_VENTES | MOD_SALES |
| MOD_PAIEMENTS | MOD_PAYMENTS |
| MOD_FACTURES | MOD_INVOICES |
| MOD_CHARGES | MOD_EXPENSES |
| MOD_RAPPORTS | MOD_REPORTS |
| DROIT_VENTE_ANNULER | RIGHT_CANCEL_SALE |
| DROIT_RESERVATION_ANNULER | RIGHT_CANCEL_RESERVATION |
| DROIT_FACTURE_ANNULER | RIGHT_CANCEL_INVOICE |
| DROIT_PAIEMENT_ANNULER | RIGHT_CANCEL_PAYMENT |
| DROIT_REMBOURSER | RIGHT_REFUND |
| DROIT_REMISE_STOCK_VALIDER | RIGHT_APPROVE_RESTOCK |
| MASQUE_LECTURE | MASK_READ |
| MASQUE_ECRITURE | MASK_WRITE |
| MASQUE_SUPPRESSION | MASK_DELETE |
| OuvrirSession | OpenSession |
| FermerSession | CloseSession |
| ActualiserSession | RefreshSession |
| EstConnecte | IsLoggedIn |
| UtilisateurCourantId | CurrentUserId |
| UtilisateurCourantLogin | CurrentUserLogin |
| UtilisateurCourantNom | CurrentUserName |
| RoleCourantId | CurrentRoleId |
| RoleCourantNom | CurrentRoleName |
| DoitChangerMotDePasse | MustChangePassword |
| PeutLire | CanRead |
| PeutEcrire | CanWrite |
| PeutSupprimer | CanDelete |
| ADroit | HasRight |
| ExigerLecture | RequireRead |
| ExigerEcriture | RequireWrite |
| ExigerSuppression | RequireDelete |
| ExigerDroit | RequireRight |
| ChangerMonMotDePasse | ChangeMyPassword |
| CreerCompte | CreateAccount |
| ReinitialiserCompte | ResetAccount |
| DebloquerCompte | UnlockAccount |
| ChargerSession | LoadSession |
| ChargerUtilisateur | LoadUser |
| ChargerDroits | LoadRights |
| AAcces | HasAccess |
| MasqueModule | ModuleMask |
| ExigerAcces | RequireAccess |
| ExigerConnexionActive | RequireActiveSession |
| Cle | NormalizeKey |

## modTestDatabase (ancien modKemerixTests)

| Ancien | Nouveau |
|---|---|
| TestOuvertureBase | TestDatabaseOpening |
| ExempleLectureSpecialites | ExampleReadSpecialties |
| Test_FichierExiste | Test_FileExists |
| Test_OuvertureAvecMotDePasse | Test_OpenWithPassword |
| Test_NombreDeTables | Test_TableCount |
| Test_DonneesReference | Test_ReferenceData |
| Test_LectureSeule | Test_ReadOnly |
| Test_EcritureAvecRollback | Test_WriteWithRollback |
| Test_RefusSansMotDePasse | Test_RejectWithoutPassword |
| Test_RefusMauvaisMotDePasse | Test_RejectWrongPassword |

## modTestConnection (ancien modTestConnexion)

| Ancien | Nouveau |
|---|---|
| Verifier (procédure de test) | Check |
| TesterConnexion | TestConnection |

## modTestDataAccess (ancien modTestDonnees)

| Ancien | Nouveau |
|---|---|
| Verifier (procédure de test) | Check |
| EgalA | IsEqualTo |
| NouvelleVoiture | NewCar |
| TesterDonnees | TestDataAccess |

## modTestSecurity (ancien modTestSecurite)

| Ancien | Nouveau |
|---|---|
| MDP_ADMIN | PWD_ADMIN |
| MDP_TEMP | PWD_TEMP |
| MDP_NOUVEAU | PWD_NEW |
| MDP_PROVISOIRE | PWD_PROVISIONAL |
| TestSecurite | TestSecurity |
| Tente | TryAuthenticate |

## modTestSecuritySeed (ancien modTestSecuriteSeed)

| Ancien | Nouveau |
|---|---|
| TestSeedSecurite | TestSecuritySeed |

## modTestSession (ancien modTestSession)

| Ancien | Nouveau |
|---|---|
| MDP_ADMIN | PWD_ADMIN |
| MDP_TEMP | PWD_TEMP |
| MDP_NOUVEAU | PWD_NEW |
| MDP_PROVISOIRE | PWD_PROVISIONAL |
| VerifierMatrice | CheckMatrix |
| VerifierDroits | CheckRights |

## modTestSql (ancien modTestSql)

| Ancien | Nouveau |
|---|---|
| Verifier (procédure de test) | Check |
| TexteDe | TextOf |
| EstVrai | IsTrueValue |
| EstFaux | IsFalseValue |
| TesterSql | TestSql |

## modTestPartyCarData (ancien modTestTiersVoitureDonnees)

| Ancien | Nouveau |
|---|---|
| Verifier (procédure de test) | Check |
| TiersValide | NewValidParty |
| VoitureValide | NewValidCar |
| TesterTiers | TestParty |
| TesterVoitureDonnees | TestCarData |

## modTestCar (ancien modTestVoiture)

| Ancien | Nouveau |
|---|---|
| Verifier (procédure de test) | Check |
| NouvelleVoiture | NewCar |
| TesterVoiture | TestCar |
