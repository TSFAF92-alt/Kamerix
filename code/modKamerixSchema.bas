Attribute VB_Name = "modKamerixSchema"
Option Explicit
Option Private Module
'==============================================================================
' KAMERIX - Schema de la base : tables, index, relations, donnees de reference
'------------------------------------------------------------------------------
' Regles :
'  - Tout le DDL est ici. modKamerixDb ne contient aucune instruction de structure.
'  - Les relations (cles etrangeres) sont ajoutees APRES la creation de toutes
'    les tables : l'ordre de creation n'a pas d'importance, les auto-references
'    (Lieu, Facture) fonctionnent.
'  - CHECK sur les valeurs arretees : statuts de Voiture, Mouvement, Facture,
'    SuiviCV, Reservation, Prorogation ; TypeMouvement, ModeReglement,
'    TypeRetenue (Reservation), Sens (Paiement) ; montants des colonnes
'    obligatoires. Les CHECK sont ecrits avec des OR (pas de IN).
'  - Les transitions de statut ne sont PAS dans la base : elles sont appliquees
'    par le code (classes). La base garantit seulement les valeurs possibles.
'  - Regles appliquees par le code (la base ne peut pas les exprimer) :
'      * un Paiement est lie a un Mouvement OU a une Reservation (au moins un) ;
'      * MontantAvance <= PrixConvenu ; une seule Reservation ACTIVE par Voiture ;
'      * retenue (Reservation, et Mouvement echelonne) non modifiable apres
'        la signature ; TypeRetenue et ValeurRetenue dans {FIXE, POURCENTAGE} ;
'      * Mouvement ECHELONNE : TypeRetenue, ValeurRetenue, DateLimiteSolde requis ;
'      * un SuiviCV ne passe a EN_COURS que si le Mouvement est SOLDE ;
'      * Prorogation acceptee seulement si TauxPaye >= SEUIL_PROROGATION_POURCENT
'        (table Parametre) ; refus du client = resiliation du Mouvement ;
'      * Mouvement ANNULE : DateAnnulation et MotifAnnulation renseignes ;
'        annulation impossible une fois le CV signe.
'  - Identifiants entre crochets dans le DDL ([Role], [Module], [Action],
'    [Description]) : penser aux crochets dans toute requete SQL sur ces noms.
'  - Evolution : incrementer SCHEMA_VERSION et ecrire la migration
'    correspondante. La version appliquee est lue dans la table VersionSchema.
'==============================================================================

Public Const SCHEMA_VERSION As Long = 5

'==============================================================================
' API PUBLIQUE (appelee par modKamerixDb)
'==============================================================================

'Liste des tables que doit contenir la base
Public Function ExpectedTables() As Variant
    ExpectedTables = Array( _
        "Parametre", "Compteur", "VersionSchema", "Lieu", "Marque", "Modele", _
        "Specialite", "CategorieCharge", "Tiers", "Intermediaire", "Prestataire", _
        "Voiture", "Reservation", "Mouvement", "Echeance", "Prorogation", "Paiement", "SuiviCV", "Charge", _
        "Facture", "JournalDocument", "Role", "Module", "Permission", "DroitSpecial", _
        "Utilisateur", "JournalConnexion")
End Function

' Cree tables, index et relations, puis insere les donnees de reference
' et la version du schema. Adocn = connexion ADO ouverte en ecriture.
Public Sub ApplySchema(ByVal Adocn As Object)
    Dim colSql As Collection
    Dim v As Variant

    Set colSql = New Collection

    DefineReferenceTables colSql
    DefinePeopleTables colSql
    DefineBusinessTables colSql
    DefineBillingTables colSql
    DefineSecurityTables colSql
    DefineIndexes colSql
    DefineForeignKeys colSql

    For Each v In colSql
        ExecSql Adocn, CStr(v)
    Next v

    SeedReferenceData Adocn
    RecordSchemaVersion Adocn
End Sub

'==============================================================================
' DEFINITION DES TABLES
'==============================================================================

Private Sub DefineReferenceTables(ByVal c As Collection)
    c.Add "CREATE TABLE [Parametre] (" & _
        "ParametreID COUNTER, " & _
        "Cle TEXT(60) NOT NULL, " & _
        "Valeur MEMO, " & _
        "CONSTRAINT pk_Parametre PRIMARY KEY (ParametreID))"

    c.Add "CREATE TABLE [Compteur] (" & _
        "CompteurID COUNTER, " & _
        "TypeDocument TEXT(20) NOT NULL, " & _
        "Annee LONG NOT NULL, " & _
        "DerniereValeur LONG NOT NULL, " & _
        "CONSTRAINT pk_Compteur PRIMARY KEY (CompteurID))"

    c.Add "CREATE TABLE [VersionSchema] (" & _
        "VersionID COUNTER, " & _
        "Numero LONG NOT NULL, " & _
        "DateApplication DATETIME NOT NULL, " & _
        "[Description] TEXT(255), " & _
        "CONSTRAINT pk_VersionSchema PRIMARY KEY (VersionID))"

    c.Add "CREATE TABLE [Lieu] (" & _
        "LieuID COUNTER, " & _
        "Nom TEXT(100) NOT NULL, " & _
        "TypeLieu TEXT(20) NOT NULL, " & _
        "LieuParentID LONG, " & _
        "CONSTRAINT pk_Lieu PRIMARY KEY (LieuID))"

    c.Add "CREATE TABLE [Marque] (" & _
        "MarqueID COUNTER, " & _
        "Nom TEXT(60) NOT NULL, " & _
        "CONSTRAINT pk_Marque PRIMARY KEY (MarqueID))"

    c.Add "CREATE TABLE [Modele] (" & _
        "ModeleID COUNTER, " & _
        "MarqueID LONG NOT NULL, " & _
        "Nom TEXT(60) NOT NULL, " & _
        "CONSTRAINT pk_Modele PRIMARY KEY (ModeleID))"

    c.Add "CREATE TABLE [Specialite] (" & _
        "SpecialiteID COUNTER, " & _
        "Libelle TEXT(60) NOT NULL, " & _
        "CONSTRAINT pk_Specialite PRIMARY KEY (SpecialiteID))"

    c.Add "CREATE TABLE [CategorieCharge] (" & _
        "CategorieID COUNTER, " & _
        "Libelle TEXT(60) NOT NULL, " & _
        "CONSTRAINT pk_CategorieCharge PRIMARY KEY (CategorieID))"
End Sub

Private Sub DefinePeopleTables(ByVal c As Collection)
    c.Add "CREATE TABLE [Tiers] (" & _
        "TiersID COUNTER, " & _
        "Nom TEXT(120) NOT NULL, " & _
        "Telephone TEXT(30), " & _
        "TypePiece TEXT(30), " & _
        "NumeroPiece TEXT(40), " & _
        "Adresse TEXT(200), " & _
        "EstAcheteur YESNO, " & _
        "EstVendeur YESNO, " & _
        "CONSTRAINT pk_Tiers PRIMARY KEY (TiersID))"

    c.Add "CREATE TABLE [Intermediaire] (" & _
        "IntermediaireID COUNTER, " & _
        "Nom TEXT(120) NOT NULL, " & _
        "Telephone TEXT(30), " & _
        "VilleID LONG NOT NULL, " & _
        "CONSTRAINT pk_Intermediaire PRIMARY KEY (IntermediaireID))"

    c.Add "CREATE TABLE [Prestataire] (" & _
        "PrestataireID COUNTER, " & _
        "Nom TEXT(120) NOT NULL, " & _
        "Telephone TEXT(30), " & _
        "SpecialiteID LONG, " & _
        "Actif YESNO, " & _
        "CONSTRAINT pk_Prestataire PRIMARY KEY (PrestataireID))"
End Sub

Private Sub DefineBusinessTables(ByVal c As Collection)
    c.Add "CREATE TABLE [Voiture] (" & _
        "VoitureID COUNTER, " & _
        "ModeleID LONG NOT NULL, " & _
        "Annee LONG, " & _
        "Couleur TEXT(30), " & _
        "Chassis TEXT(30) NOT NULL, " & _
        "Immatriculation TEXT(20), " & _
        "Kilometrage LONG, " & _
        "Statut TEXT(20) NOT NULL, " & _
        "MotifRetrait TEXT(100), " & _
        "CONSTRAINT pk_Voiture PRIMARY KEY (VoitureID), " & _
        "CONSTRAINT ck_Voiture_Statut CHECK (Statut = 'EN_STOCK' OR Statut = 'RESERVE' OR " & _
        "Statut = 'VENTE_EN_COURS' OR Statut = 'VENDU' OR Statut = 'RETIRE'))"

    c.Add "CREATE TABLE [Reservation] (" & _
        "ReservationID COUNTER, " & _
        "VoitureID LONG NOT NULL, " & _
        "TiersID LONG NOT NULL, " & _
        "DateReservation DATETIME NOT NULL, " & _
        "PrixConvenu CURRENCY NOT NULL, " & _
        "MontantAvance CURRENCY NOT NULL, " & _
        "DateLimiteSolde DATETIME NOT NULL, " & _
        "TypeRetenue TEXT(12) NOT NULL, " & _
        "ValeurRetenue DOUBLE NOT NULL, " & _
        "Statut TEXT(12) NOT NULL, " & _
        "MouvementID LONG, " & _
        "DateCloture DATETIME, " & _
        "UtilisateurID LONG NOT NULL, " & _
        "CONSTRAINT pk_Reservation PRIMARY KEY (ReservationID), " & _
        "CONSTRAINT ck_Reservation_Montants CHECK (PrixConvenu >= 0 AND MontantAvance > 0 AND ValeurRetenue >= 0), " & _
        "CONSTRAINT ck_Reservation_TypeRetenue CHECK (TypeRetenue = 'FIXE' OR TypeRetenue = 'POURCENTAGE'), " & _
        "CONSTRAINT ck_Reservation_Statut CHECK (Statut = 'ACTIVE' OR Statut = 'CONVERTIE' OR " & _
        "Statut = 'ANNULEE' OR Statut = 'EXPIREE'))"

    c.Add "CREATE TABLE [Mouvement] (" & _
        "MouvementID COUNTER, " & _
        "TypeMouvement TEXT(10) NOT NULL, " & _
        "VoitureID LONG NOT NULL, TiersID LONG NOT NULL, " & _
        "IntermediaireID LONG, LieuID LONG, " & _
        "DateMouvement DATETIME NOT NULL, " & _
        "Prix CURRENCY NOT NULL, Commission CURRENCY NOT NULL, " & _
        "ModeReglement TEXT(12) NOT NULL, " & _
        "Statut TEXT(12) NOT NULL, " & _
        "TypeRetenue TEXT(12), ValeurRetenue DOUBLE, " & _
        "DateLimiteSolde DATETIME, " & _
        "DateAnnulation DATETIME, MotifAnnulation TEXT(100), " & _
        "UtilisateurID LONG NOT NULL, " & _
        "CONSTRAINT pk_Mouvement PRIMARY KEY (MouvementID), " & _
        "CONSTRAINT ck_Mouvement_Montants CHECK (Prix >= 0 AND Commission >= 0), " & _
        "CONSTRAINT ck_Mouvement_Type CHECK (TypeMouvement = 'ACHAT' OR TypeMouvement = 'VENTE'), " & _
        "CONSTRAINT ck_Mouvement_Reglement CHECK (ModeReglement = 'COMPTANT' OR ModeReglement = 'ECHELONNE'), " & _
        "CONSTRAINT ck_Mouvement_Statut CHECK (Statut = 'EN_COURS' OR Statut = 'SOLDE' OR Statut = 'ANNULE'))"

    c.Add "CREATE TABLE [Echeance] (" & _
        "EcheanceID COUNTER, " & _
        "MouvementID LONG NOT NULL, " & _
        "Numero LONG NOT NULL, " & _
        "DatePrevue DATETIME NOT NULL, " & _
        "MontantDu CURRENCY NOT NULL, " & _
        "CONSTRAINT pk_Echeance PRIMARY KEY (EcheanceID), " & _
        "CONSTRAINT ck_Echeance_Montant CHECK (MontantDu >= 0))"

    c.Add "CREATE TABLE [Prorogation] (" & _
        "ProrogationID COUNTER, " & _
        "MouvementID LONG NOT NULL, " & _
        "DateDecision DATETIME NOT NULL, " & _
        "TauxPaye DOUBLE NOT NULL, " & _
        "AncienneDateLimite DATETIME NOT NULL, " & _
        "NouvelleDateLimite DATETIME NOT NULL, " & _
        "Statut TEXT(12) NOT NULL, " & _
        "UtilisateurID LONG NOT NULL, " & _
        "CONSTRAINT pk_Prorogation PRIMARY KEY (ProrogationID), " & _
        "CONSTRAINT ck_Prorogation_Taux CHECK (TauxPaye >= 0 AND TauxPaye <= 100), " & _
        "CONSTRAINT ck_Prorogation_Statut CHECK (Statut = 'ACCEPTEE' OR Statut = 'REFUSEE'))"

    c.Add "CREATE TABLE [Paiement] (" & _
        "PaiementID COUNTER, " & _
        "MouvementID LONG, " & _
        "ReservationID LONG, " & _
        "EcheanceID LONG, " & _
        "FactureID LONG, " & _
        "NumeroRecu TEXT(30) NOT NULL, " & _
        "Sens TEXT(15) NOT NULL, " & _
        "DatePaiement DATETIME NOT NULL, " & _
        "Montant CURRENCY NOT NULL, " & _
        "ModePaiement TEXT(20) NOT NULL, " & _
        "UtilisateurID LONG NOT NULL, " & _
        "CONSTRAINT pk_Paiement PRIMARY KEY (PaiementID), " & _
        "CONSTRAINT ck_Paiement_Montant CHECK (Montant > 0), " & _
        "CONSTRAINT ck_Paiement_Sens CHECK (Sens = 'ENCAISSEMENT' OR Sens = 'REMBOURSEMENT'))"

    c.Add "CREATE TABLE [SuiviCV] (" & _
        "SuiviCVID COUNTER, " & _
        "MouvementID LONG NOT NULL, " & _
        "NumeroCV TEXT(40), " & _
        "CommissariatID LONG, " & _
        "DateDepot DATETIME, " & _
        "DateSignature DATETIME, " & _
        "FraisPayes CURRENCY NOT NULL, " & _
        "Statut TEXT(20) NOT NULL, " & _
        "CONSTRAINT pk_SuiviCV PRIMARY KEY (SuiviCVID), " & _
        "CONSTRAINT ck_SuiviCV_Frais CHECK (FraisPayes >= 0), " & _
        "CONSTRAINT ck_SuiviCV_Statut CHECK (Statut = 'A_DEPOSER' OR Statut = 'EN_COURS' OR Statut = 'SIGNE'))"

    c.Add "CREATE TABLE [Charge] (" & _
        "ChargeID COUNTER, " & _
        "VoitureID LONG, " & _
        "CategorieID LONG NOT NULL, " & _
        "PrestataireID LONG, " & _
        "DateCharge DATETIME NOT NULL, " & _
        "[Description] TEXT(255), " & _
        "Montant CURRENCY NOT NULL, " & _
        "ModePaiement TEXT(20), " & _
        "UtilisateurID LONG NOT NULL, " & _
        "CONSTRAINT pk_Charge PRIMARY KEY (ChargeID), " & _
        "CONSTRAINT ck_Charge_Montant CHECK (Montant >= 0))"
End Sub

Private Sub DefineBillingTables(ByVal c As Collection)
    c.Add "CREATE TABLE [Facture] (" & _
        "FactureID COUNTER, " & _
        "NumeroFacture TEXT(30) NOT NULL, " & _
        "MouvementID LONG NOT NULL, " & _
        "TiersID LONG NOT NULL, " & _
        "DateFacture DATETIME NOT NULL, " & _
        "Designation TEXT(255) NOT NULL, " & _
        "MontantHT CURRENCY NOT NULL, " & _
        "TauxTaxe DOUBLE NOT NULL, " & _
        "MontantTaxe CURRENCY NOT NULL, " & _
        "MontantTTC CURRENCY NOT NULL, " & _
        "Statut TEXT(12) NOT NULL, " & _
        "FactureOrigineID LONG, " & _
        "UtilisateurID LONG NOT NULL, " & _
        "CONSTRAINT pk_Facture PRIMARY KEY (FactureID), " & _
        "CONSTRAINT ck_Facture_Montants CHECK (MontantHT >= 0 AND TauxTaxe >= 0 AND MontantTaxe >= 0 AND MontantTTC >= 0), " & _
        "CONSTRAINT ck_Facture_Statut CHECK (Statut = 'EMISE' OR Statut = 'ANNULEE'))"

    c.Add "CREATE TABLE [JournalDocument] (" & _
        "JournalDocID COUNTER, " & _
        "TypeDocument TEXT(20) NOT NULL, " & _
        "DocumentID LONG NOT NULL, " & _
        "[Action] TEXT(20) NOT NULL, " & _
        "Duplicata YESNO, " & _
        "UtilisateurID LONG NOT NULL, " & _
        "DateHeure DATETIME NOT NULL, " & _
        "CONSTRAINT pk_JournalDocument PRIMARY KEY (JournalDocID))"
End Sub

Private Sub DefineSecurityTables(ByVal c As Collection)
    c.Add "CREATE TABLE [Role] (" & _
        "RoleID COUNTER, " & _
        "Nom TEXT(60) NOT NULL, " & _
        "CONSTRAINT pk_Role PRIMARY KEY (RoleID))"

    c.Add "CREATE TABLE [Module] (" & _
        "ModuleID COUNTER, " & _
        "Nom TEXT(60) NOT NULL, " & _
        "CONSTRAINT pk_Module PRIMARY KEY (ModuleID))"

    c.Add "CREATE TABLE [Permission] (" & _
        "PermissionID COUNTER, " & _
        "RoleID LONG NOT NULL, " & _
        "ModuleID LONG NOT NULL, " & _
        "Lecture YESNO, " & _
        "Ecriture YESNO, " & _
        "Suppression YESNO, " & _
        "CONSTRAINT pk_Permission PRIMARY KEY (PermissionID))"

    c.Add "CREATE TABLE [Utilisateur] (" & _
        "UtilisateurID COUNTER, " & _
        "Login TEXT(50) NOT NULL, " & _
        "NomComplet TEXT(120) NOT NULL, " & _
        "MotDePasseHash TEXT(128) NOT NULL, " & _
        "Sel TEXT(64) NOT NULL, " & _
        "RoleID LONG NOT NULL, " & _
        "Actif YESNO, " & _
        "DoitChangerMdp YESNO, " & _
        "NbEchecs LONG NOT NULL, " & _
        "DateBlocage DATETIME, " & _
        "CONSTRAINT pk_Utilisateur PRIMARY KEY (UtilisateurID), " & _
        "CONSTRAINT ck_Utilisateur_Echecs CHECK (NbEchecs >= 0))"

    c.Add "CREATE TABLE [JournalConnexion] (" & _
        "JournalID COUNTER, " & _
        "UtilisateurID LONG, " & _
        "LoginSaisi TEXT(50), " & _
        "DateHeure DATETIME NOT NULL, " & _
        "Resultat TEXT(10) NOT NULL, " & _
        "Poste TEXT(100), " & _
        "CONSTRAINT pk_JournalConnexion PRIMARY KEY (JournalID))"
        
    c.Add "CREATE TABLE [DroitSpecial] (" & _
        "DroitID COUNTER, " & _
        "RoleID LONG NOT NULL, " & _
        "Code TEXT(40) NOT NULL, " & _
        "CONSTRAINT pk_DroitSpecial PRIMARY KEY (DroitID))"
End Sub

'==============================================================================
' INDEX
'==============================================================================

Private Sub DefineIndexes(ByVal c As Collection)
    ' Unicite
    AddIndex c, True, "ux_Parametre_Cle", "Parametre", "Cle"
    AddIndex c, True, "ux_Compteur_TypeAnnee", "Compteur", "TypeDocument, Annee"
    AddIndex c, True, "ux_VersionSchema_Numero", "VersionSchema", "Numero"
    AddIndex c, True, "ux_Marque_Nom", "Marque", "Nom"
    AddIndex c, True, "ux_Modele_MarqueNom", "Modele", "MarqueID, Nom"
    AddIndex c, True, "ux_Specialite_Libelle", "Specialite", "Libelle"
    AddIndex c, True, "ux_CategorieCharge_Libelle", "CategorieCharge", "Libelle"
    AddIndex c, True, "ux_Voiture_Chassis", "Voiture", "Chassis"
    AddIndex c, True, "ux_Echeance_MouvementNumero", "Echeance", "MouvementID, Numero"
    AddIndex c, True, "ux_Paiement_NumeroRecu", "Paiement", "NumeroRecu"
    AddIndex c, True, "ux_SuiviCV_Mouvement", "SuiviCV", "MouvementID"
    AddIndex c, True, "ux_Facture_NumeroFacture", "Facture", "NumeroFacture"
    AddIndex c, True, "ux_Role_Nom", "Role", "Nom"
    AddIndex c, True, "ux_Module_Nom", "Module", "Nom"
    AddIndex c, True, "ux_Permission_RoleModule", "Permission", "RoleID, ModuleID"
    AddIndex c, True, "ux_Utilisateur_Login", "Utilisateur", "Login"
    AddIndex c, True, "ux_DroitSpecial_RoleCode", "DroitSpecial", "RoleID, Code"
    
    ' Performance (recherches frequentes)
    AddIndex c, False, "ix_Mouvement_Date", "Mouvement", "DateMouvement"
    AddIndex c, False, "ix_Paiement_Date", "Paiement", "DatePaiement"
    AddIndex c, False, "ix_Reservation_DateLimite", "Reservation", "DateLimiteSolde"
    AddIndex c, False, "ix_Charge_Date", "Charge", "DateCharge"
    AddIndex c, False, "ix_JournalConnexion_Date", "JournalConnexion", "DateHeure"
    AddIndex c, False, "ix_JournalDocument_Date", "JournalDocument", "DateHeure"
End Sub

'==============================================================================
' RELATIONS (cles etrangeres, sans suppression ni mise a jour en cascade)
'==============================================================================

Private Sub DefineForeignKeys(ByVal c As Collection)
    AddForeignKey c, "Modele", "MarqueID", "Marque", "MarqueID"
    AddForeignKey c, "Voiture", "ModeleID", "Modele", "ModeleID"

    AddForeignKey c, "Lieu", "LieuParentID", "Lieu", "LieuID"
    AddForeignKey c, "Intermediaire", "VilleID", "Lieu", "LieuID"
    AddForeignKey c, "Prestataire", "SpecialiteID", "Specialite", "SpecialiteID"

    AddForeignKey c, "Mouvement", "VoitureID", "Voiture", "VoitureID"
    AddForeignKey c, "Mouvement", "TiersID", "Tiers", "TiersID"
    AddForeignKey c, "Mouvement", "IntermediaireID", "Intermediaire", "IntermediaireID"
    AddForeignKey c, "Mouvement", "LieuID", "Lieu", "LieuID"
    AddForeignKey c, "Mouvement", "UtilisateurID", "Utilisateur", "UtilisateurID"

    AddForeignKey c, "Echeance", "MouvementID", "Mouvement", "MouvementID"

    AddForeignKey c, "Prorogation", "MouvementID", "Mouvement", "MouvementID"
    AddForeignKey c, "Prorogation", "UtilisateurID", "Utilisateur", "UtilisateurID"

    AddForeignKey c, "Reservation", "VoitureID", "Voiture", "VoitureID"
    AddForeignKey c, "Reservation", "TiersID", "Tiers", "TiersID"
    AddForeignKey c, "Reservation", "MouvementID", "Mouvement", "MouvementID"
    AddForeignKey c, "Reservation", "UtilisateurID", "Utilisateur", "UtilisateurID"

    AddForeignKey c, "Paiement", "MouvementID", "Mouvement", "MouvementID"
    AddForeignKey c, "Paiement", "ReservationID", "Reservation", "ReservationID"
    AddForeignKey c, "Paiement", "EcheanceID", "Echeance", "EcheanceID"
    AddForeignKey c, "Paiement", "FactureID", "Facture", "FactureID"
    AddForeignKey c, "Paiement", "UtilisateurID", "Utilisateur", "UtilisateurID"

    AddForeignKey c, "SuiviCV", "MouvementID", "Mouvement", "MouvementID"
    AddForeignKey c, "SuiviCV", "CommissariatID", "Lieu", "LieuID"

    AddForeignKey c, "Charge", "VoitureID", "Voiture", "VoitureID"
    AddForeignKey c, "Charge", "CategorieID", "CategorieCharge", "CategorieID"
    AddForeignKey c, "Charge", "PrestataireID", "Prestataire", "PrestataireID"
    AddForeignKey c, "Charge", "UtilisateurID", "Utilisateur", "UtilisateurID"

    AddForeignKey c, "Facture", "MouvementID", "Mouvement", "MouvementID"
    AddForeignKey c, "Facture", "TiersID", "Tiers", "TiersID"
    AddForeignKey c, "Facture", "FactureOrigineID", "Facture", "FactureID"
    AddForeignKey c, "Facture", "UtilisateurID", "Utilisateur", "UtilisateurID"

    AddForeignKey c, "JournalDocument", "UtilisateurID", "Utilisateur", "UtilisateurID"

    AddForeignKey c, "Permission", "RoleID", "Role", "RoleID"
    AddForeignKey c, "Permission", "ModuleID", "Module", "ModuleID"
    AddForeignKey c, "Utilisateur", "RoleID", "Role", "RoleID"
    AddForeignKey c, "JournalConnexion", "UtilisateurID", "Utilisateur", "UtilisateurID"
    
     AddForeignKey c, "DroitSpecial", "RoleID", "Role", "RoleID"
End Sub

'==============================================================================
' DONNEES DE REFERENCE ET VERSION
'==============================================================================

Private Sub SeedReferenceData(ByVal Adocn As Object)
    ' Metiers de techniciens
    InsertLabel Adocn, "Specialite", "Libelle", "M" & ChrW(233) & "canicien"
    InsertLabel Adocn, "Specialite", "Libelle", "T" & ChrW(244) & "lier-peintre"
    InsertLabel Adocn, "Specialite", "Libelle", ChrW(201) & "lectricien"
    InsertLabel Adocn, "Specialite", "Libelle", "Frigoriste"
    InsertLabel Adocn, "Specialite", "Libelle", "Colleur"
    InsertLabel Adocn, "Specialite", "Libelle", "D" & ChrW(233) & "corateur"

    ' Categories de charges
    InsertLabel Adocn, "CategorieCharge", "Libelle", "Intervention technique"
    InsertLabel Adocn, "CategorieCharge", "Libelle", "Gardiennage"
    InsertLabel Adocn, "CategorieCharge", "Libelle", "Lavage"
    InsertLabel Adocn, "CategorieCharge", "Libelle", "Carburant"
    InsertLabel Adocn, "CategorieCharge", "Libelle", "Plateforme (boost des ventes)"
    InsertLabel Adocn, "CategorieCharge", "Libelle", "Autre"

    ' Parametres de gestion (modifiables sans toucher au code)
    ExecSql Adocn, "INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES ('SEUIL_PROROGATION_POURCENT', '75')"
    ExecSql Adocn, "INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES ('MAX_ECHECS_CONNEXION', '5')"
    ExecSql Adocn, "INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES ('DUREE_BLOCAGE_MINUTES', '15')"
    ' Vente : taxe sur le prix TTC, modes de paiement (separes par ;), devise des montants
    ExecSql Adocn, "INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES ('TAUX_TAXE', '19.25')"
    ExecSql Adocn, "INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES ('MODES_PAIEMENT', '" & _
        "Esp" & ChrW(232) & "ces;Mobile Money;Virement bancaire')"
    ExecSql Adocn, "INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES ('DEVISE', 'FCFA')"
    ExecSql Adocn, "INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES ('DECIMALES_DEVISE', '0')"
    SeedSecurityData Adocn
End Sub

Private Sub RecordSchemaVersion(ByVal Adocn As Object)
    ExecSql Adocn, "INSERT INTO [VersionSchema] ([Numero], [DateApplication], [Description]) VALUES (" & _
        SCHEMA_VERSION & ", #" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#, 'Creation initiale')"
End Sub

' Roles, modules, matrice de permissions (Lecture / Ecriture / Suppression)
' et droits speciaux. Le premier compte Administrateur n'est PAS cree ici :
' il est cree par le code de securite au premier lancement (mot de passe choisi
' a ce moment-la, jamais ecrit dans le code).
Private Sub SeedSecurityData(ByVal Adocn As Object)
    Dim vRoles As Variant
    Dim vModules As Variant
    Dim vMatrix As Variant
    Dim vCells As Variant
    Dim vAll As Variant
    Dim vAccounting As Variant
    Dim i As Long
    Dim j As Long

    vRoles = Array("Administrateur", "Responsable", "Vendeur", "Comptable")

    ' Ordre des modules = ordre des colonnes de la matrice ci-dessous
    vModules = Array("Voitures", "Tiers", "R" & ChrW(233) & "servations", "Ventes", _
                     "Paiements", "Factures", "Charges", "Rapports", "Administration")

    ' Une ligne par role, une cellule par module (meme ordre que vModules).
    ' Chaque cellule = 3 caracteres : L (lecture) E (ecriture) S (suppression),
    ' "-" si le droit n'est pas accorde.
    vMatrix = Array( _
        "LES|LES|LES|LES|LES|LES|LES|LES|LES", _
        "LE-|LE-|LE-|LE-|LE-|LE-|LE-|LE-|---", _
        "L--|LE-|LE-|LE-|L--|L--|---|---|---", _
        "L--|L--|L--|L--|LE-|LE-|LE-|L--|---")

    For i = LBound(vRoles) To UBound(vRoles)
        InsertLabel Adocn, "Role", "Nom", CStr(vRoles(i))
    Next i

    For j = LBound(vModules) To UBound(vModules)
        InsertLabel Adocn, "Module", "Nom", CStr(vModules(j))
    Next j

    For i = LBound(vRoles) To UBound(vRoles)
        vCells = Split(CStr(vMatrix(i)), "|")
        For j = LBound(vModules) To UBound(vModules)
            AddPermission Adocn, CStr(vRoles(i)), CStr(vModules(j)), CStr(vCells(j))
        Next j
    Next i

    ' Droits speciaux (verifies par leur code dans le code de l'application)
    vAll = Array("VENTE_ANNULER", "RESERVATION_ANNULER", "FACTURE_ANNULER", _
                 "PAIEMENT_ANNULER", "REMBOURSER", "REMISE_STOCK_VALIDER")
    vAccounting = Array("PAIEMENT_ANNULER", "FACTURE_ANNULER", "REMBOURSER")

    AddSpecialRights Adocn, "Administrateur", vAll
    AddSpecialRights Adocn, "Responsable", vAll
    AddSpecialRights Adocn, "Comptable", vAccounting
    ' Vendeur : aucun droit special
End Sub

' Ajoute une ligne de permission pour un role et un module (identifies par leur nom)
Private Sub AddPermission(ByVal Adocn As Object, ByVal strRole As String, _
                          ByVal strModule As String, ByVal strFlags As String)
    ExecSql Adocn, "INSERT INTO [Permission] ([RoleID], [ModuleID], [Lecture], [Ecriture], [Suppression]) " & _
        "SELECT r.[RoleID], m.[ModuleID], " & _
        SqlBool(Mid$(strFlags, 1, 1) = "L") & ", " & _
        SqlBool(Mid$(strFlags, 2, 1) = "E") & ", " & _
        SqlBool(Mid$(strFlags, 3, 1) = "S") & " " & _
        "FROM [Role] AS r, [Module] AS m " & _
        "WHERE r.[Nom] = '" & Replace(strRole, "'", "''") & "' " & _
        "AND m.[Nom] = '" & Replace(strModule, "'", "''") & "'"
End Sub

' Accorde une liste de droits speciaux a un role
Private Sub AddSpecialRights(ByVal Adocn As Object, ByVal strRole As String, ByVal vCodes As Variant)
    Dim i As Long

    For i = LBound(vCodes) To UBound(vCodes)
        ExecSql Adocn, "INSERT INTO [DroitSpecial] ([RoleID], [Code]) " & _
            "SELECT r.[RoleID], '" & Replace(CStr(vCodes(i)), "'", "''") & "' " & _
            "FROM [Role] AS r WHERE r.[Nom] = '" & Replace(strRole, "'", "''") & "'"
    Next i
End Sub

Private Function SqlBool(ByVal blnValue As Boolean) As String
    If blnValue Then SqlBool = "True" Else SqlBool = "False"
End Function

'==============================================================================
' OUTILS INTERNES
'==============================================================================

Private Sub AddIndex(ByVal c As Collection, ByVal blnUnique As Boolean, ByVal strName As String, _
                     ByVal strTable As String, ByVal strCols As String)
    Dim strSql As String

    strSql = "CREATE "
    If blnUnique Then strSql = strSql & "UNIQUE "
    strSql = strSql & "INDEX " & strName & " ON [" & strTable & "] (" & BracketList(strCols) & ")"
    c.Add strSql
End Sub

Private Sub AddForeignKey(ByVal c As Collection, ByVal strChild As String, ByVal strCol As String, _
                          ByVal strParent As String, ByVal strParentCol As String)
    c.Add "ALTER TABLE [" & strChild & "] ADD CONSTRAINT fk_" & strChild & "_" & strCol & _
          " FOREIGN KEY ([" & strCol & "]) REFERENCES [" & strParent & "] ([" & strParentCol & "])"
End Sub

' "A, B" -> "[A], [B]"
Private Function BracketList(ByVal strCols As String) As String
    Dim a() As String
    Dim i As Long

    a = Split(strCols, ",")
    For i = LBound(a) To UBound(a)
        a(i) = "[" & Trim$(a(i)) & "]"
    Next i
    BracketList = Join(a, ", ")
End Function

Private Sub InsertLabel(ByVal Adocn As Object, ByVal strTable As String, ByVal strColumn As String, _
                        ByVal strValue As String)
    ExecSql Adocn, "INSERT INTO [" & strTable & "] ([" & strColumn & "]) VALUES ('" & _
        Replace(strValue, "'", "''") & "')"
End Sub

' Execute une instruction SQL ; en cas d'echec, l'erreur indique l'instruction fautive
Private Sub ExecSql(ByVal Adocn As Object, ByVal strSql As String)
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo EH

    Adocn.Execute strSql, , 129             ' 129 = adCmdText + adExecuteNoRecords
    Exit Sub

EH:
    lngErr = Err.Number
    strErr = Err.Description
    Err.Raise ERR_SQL, "modKamerixSchema.ExecSql", _
        "Echec SQL (erreur d'origine " & lngErr & ") : " & strErr & vbCrLf & _
        "Instruction : " & Left$(strSql, 300)
End Sub
