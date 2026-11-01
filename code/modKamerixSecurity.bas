Attribute VB_Name = "modKamerixSecurity"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Securite : comptes utilisateurs, connexion, mots de passe
'------------------------------------------------------------------------------
' Role        : creer les comptes, controler une connexion (avec blocage
'               temporaire apres des echecs repetes et journal), changer ou
'               reinitialiser un mot de passe. Ne contient aucun etat de
'               session : la Session (utilisateur connecte, droits par module)
'               sera un module separe qui appellera AuthenticateUser.
' Mots de passe : jamais stockes en clair (modKamerixCrypto : PBKDF2-HMAC-SHA256
'               avec sel). Un mot de passe ne se relit pas, il se remplace.
' Connexion   : un seul message cote ecran pour un identifiant inconnu ou un
'               mauvais mot de passe (kxLoginDenied). Seul le blocage est
'               annonce (kxLoginLocked).
' Blocage     : apres MAX_ECHECS_CONNEXION echecs consecutifs (5 par defaut),
'               le compte est bloque DUREE_BLOCAGE_MINUTES minutes (15 par
'               defaut). Les deux valeurs se lisent dans la table Parametre ;
'               si la cle est absente, les valeurs par defaut s'appliquent.
' Droits      : qui a le droit d'appeler CreateUser, ResetPassword
'               ou UnlockUser se controle dans la couche Session,
'               pas ici.
' Erreurs     : Err.Raise avec les numeros ERR_SEC_* (plage 1400 a 1499).
' Dependances : modKamerixSql, modKamerixCrypto, modKamerixDb (ERR_BASE).
'==============================================================================

Public Const ERR_SEC_INVALID_DATA As Long = ERR_BASE + 400        ' donnees invalides
Public Const ERR_SEC_LOGIN_EXISTS As Long = ERR_BASE + 401   ' identifiant deja utilise
Public Const ERR_SEC_NOT_FOUND As Long = ERR_BASE + 402    ' utilisateur introuvable
Public Const ERR_SEC_WRONG_PASSWORD As Long = ERR_BASE + 403  ' ancien mot de passe incorrect
Public Const ERR_SEC_ADMIN_EXISTS As Long = ERR_BASE + 404   ' un administrateur existe deja
Public Const ERR_SEC_UNKNOWN_ROLE As Long = ERR_BASE + 405   ' role inexistant
Public Const ERR_SEC_LAST_ADMIN As Long = ERR_BASE + 406     ' dernier administrateur actif
Public Const ERR_SEC_RECOVERY_INVALID As Long = ERR_BASE + 407 ' code de secours refuse

Public Const ROLE_ADMINISTRATOR As String = "Administrateur"
Public Const PARAM_MAX_FAILURES As String = "MAX_ECHECS_CONNEXION"
Public Const PARAM_LOCKOUT_MINUTES As String = "DUREE_BLOCAGE_MINUTES"

Private Const DEFAULT_MAX_FAILURES As Long = 5
Private Const DEFAULT_LOCKOUT_MINUTES As Long = 15

' Code de secours (table Parametre) et caracteres des codes generes
Private Const PARAM_RECOVERY_HASH As String = "CODE_SECOURS_HASH"
Private Const PARAM_RECOVERY_SALT As String = "CODE_SECOURS_SEL"
Private Const CHARS_PASSWORD As String = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789"
Private Const CHARS_RECOVERY As String = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"

Private Const MIN_LOGIN_LENGTH As Long = 3
Private Const MAX_LOGIN_LENGTH As Long = 50
Private Const MAX_NAME_LENGTH As Long = 120
Private Const MIN_PASSWORD_LENGTH As Long = 8
Private Const MAX_PASSWORD_LENGTH As Long = 100

' Valeurs de JournalConnexion.Resultat (10 caracteres au plus)
Private Const RESULT_SUCCESS As String = "Succes"
Private Const RESULT_FAILURE As String = "Echec"
Private Const RESULT_LOCKED As String = "Bloque"
Private Const RESULT_RECOVERY As String = "Secours"

' Resultat d'une tentative de connexion
Public Enum KxLoginResult
    kxLoginSuccess = 1    ' identifiant et mot de passe corrects
    kxLoginDenied = 2   ' identifiant inconnu, mot de passe incorrect ou compte inactif
    kxLoginLocked = 3   ' compte temporairement bloque
End Enum

' Nombre d'iterations du hachage : 0 = valeur standard (DEFAULT_ITERATIONS).
' Reserve aux tests, pour qu'ils ne durent pas des minutes.
Private mlngIterations As Long

'==============================================================================
' COMPTES
'==============================================================================

' Vrai s'il existe au moins un compte Administrateur actif
Public Function AdminExists() As Boolean
    Dim lngCount As Long

    lngCount = CLng(QueryValue( _
        "SELECT COUNT(*) FROM [Utilisateur] AS u INNER JOIN [Role] AS r ON u.[RoleID] = r.[RoleID] " & _
        "WHERE r.[Nom] = ? AND u.[Actif] = ?", ROLE_ADMINISTRATOR, True))
    AdminExists = (lngCount > 0)
End Function

' Cree le premier compte Administrateur (premier lancement). Refuse s'il en existe deja un.
' L'administrateur choisit son propre mot de passe : pas de changement obligatoire.
' Retourne l'identifiant du compte.
Public Function CreateFirstAdmin(ByVal pLogin As String, ByVal pFullName As String, _
                                  ByVal pPassword As String) As Long
    If AdminExists() Then
        Err.Raise ERR_SEC_ADMIN_EXISTS, "modKamerixSecurity.CreateFirstAdmin", _
            "Un administrateur existe deja."
    End If

    CreateFirstAdmin = InsertUser(pLogin, pFullName, GetRoleIdByName(ROLE_ADMINISTRATOR), _
                                           pPassword, False)
End Function

' Cree un compte avec un mot de passe temporaire : l'utilisateur devra le changer
' a sa premiere connexion. Retourne l'identifiant du compte.
Public Function CreateUser(ByVal pLogin As String, ByVal pFullName As String, _
                                 ByVal pRoleId As Long, ByVal pTempPassword As String) As Long
    If Not RoleExists(pRoleId) Then
        Err.Raise ERR_SEC_UNKNOWN_ROLE, "modKamerixSecurity.CreateUser", _
            "Role inconnu (identifiant " & pRoleId & ")."
    End If

    CreateUser = InsertUser(pLogin, pFullName, pRoleId, pTempPassword, True)
End Function

' Identifiant d'un role d'apres son nom ; erreur si le role n'existe pas
Public Function GetRoleIdByName(ByVal pName As String) As Long
    Dim vId As Variant

    vId = QueryValue("SELECT [RoleID] FROM [Role] WHERE [Nom] = ?", pName)
    If IsNull(vId) Then
        Err.Raise ERR_SEC_UNKNOWN_ROLE, "modKamerixSecurity.GetRoleIdByName", "Role inconnu : " & pName
    End If
    GetRoleIdByName = CLng(vId)
End Function

'==============================================================================
' CONNEXION
'==============================================================================

' Controle une tentative de connexion et l'inscrit dans le journal.
'   pUserId  : en sortie, identifiant du compte (0 si la connexion echoue)
'   pMustChangePassword : en sortie, Vrai si l'utilisateur doit changer son mot de passe
Public Function AuthenticateUser(ByVal pLogin As String, ByVal pPassword As String, _
                          ByRef pUserId As Long, ByRef pMustChangePassword As Boolean) As KxLoginResult
    Dim AdoRs As ADODB.Recordset
    Dim strLogin As String
    Dim lngId As Long
    Dim strHash As String
    Dim strSalt As String
    Dim blnActive As Boolean
    Dim blnMustChange As Boolean
    Dim lngFailures As Long
    Dim vLockedUntil As Variant

    pUserId = 0
    pMustChangePassword = False
    strLogin = Trim$(pLogin)

    ' Saisie vide : echec, apres le meme temps de calcul qu'une vraie verification
    If Len(strLogin) = 0 Or Len(pPassword) = 0 Then
        SpendHashTime pPassword
        LogLoginAttempt Null, strLogin, RESULT_FAILURE
        AuthenticateUser = kxLoginDenied
        Exit Function
    End If

    Set AdoRs = QueryRecordset( _
        "SELECT [UtilisateurID], [MotDePasseHash], [Sel], [Actif], [DoitChangerMdp], [NbEchecs], [DateBlocage] " & _
        "FROM [Utilisateur] WHERE [Login] = ?", strLogin)

    ' Identifiant inconnu : meme duree et meme reponse qu'un mauvais mot de passe
    If AdoRs.EOF Then
        AdoRs.Close
        SpendHashTime pPassword
        LogLoginAttempt Null, strLogin, RESULT_FAILURE
        AuthenticateUser = kxLoginDenied
        Exit Function
    End If

    lngId = CLng(AdoRs.Fields("UtilisateurID").Value)
    strHash = CStr(AdoRs.Fields("MotDePasseHash").Value)
    strSalt = CStr(AdoRs.Fields("Sel").Value)
    blnActive = CBool(AdoRs.Fields("Actif").Value)
    blnMustChange = CBool(AdoRs.Fields("DoitChangerMdp").Value)
    lngFailures = CLng(AdoRs.Fields("NbEchecs").Value)
    vLockedUntil = AdoRs.Fields("DateBlocage").Value
    AdoRs.Close

    ' Compte bloque : on ne verifie meme pas le mot de passe
    If Not IsNull(vLockedUntil) Then
        If CDate(vLockedUntil) > Now Then
            LogLoginAttempt lngId, strLogin, RESULT_LOCKED
            AuthenticateUser = kxLoginLocked
            Exit Function
        End If
        lngFailures = 0       ' blocage expire : on repart de zero
    End If

    If VerifyPassword(pPassword, strSalt, strHash) And blnActive Then
        ExecuteNonQuery "UPDATE [Utilisateur] SET [NbEchecs] = 0, [DateBlocage] = NULL WHERE [UtilisateurID] = ?", lngId
        LogLoginAttempt lngId, strLogin, RESULT_SUCCESS
        ' Le hash suit le reglage en vigueur (nombre d'iterations augmente ou diminue)
        If StoredIterations(strHash) <> CurrentIterations() Then RehashPassword lngId, pPassword
        pUserId = lngId
        pMustChangePassword = blnMustChange
        AuthenticateUser = kxLoginSuccess
        Exit Function
    End If

    ' Echec : on compte, et on bloque si le maximum est atteint
    lngFailures = lngFailures + 1
    If lngFailures >= MaxFailedAttempts() Then
        ExecuteNonQuery "UPDATE [Utilisateur] SET [NbEchecs] = 0, [DateBlocage] = ? WHERE [UtilisateurID] = ?", _
            DateAdd("n", LockoutMinutes(), Now), lngId
        LogLoginAttempt lngId, strLogin, RESULT_LOCKED
        AuthenticateUser = kxLoginLocked
    Else
        ExecuteNonQuery "UPDATE [Utilisateur] SET [NbEchecs] = ?, [DateBlocage] = NULL WHERE [UtilisateurID] = ?", _
            lngFailures, lngId
        LogLoginAttempt lngId, strLogin, RESULT_FAILURE
        AuthenticateUser = kxLoginDenied
    End If
End Function

' Texte des regles du mot de passe, pour les ecrans de saisie
Public Function PasswordRulesText() As String
    PasswordRulesText = "Au moins " & MIN_PASSWORD_LENGTH & " caract" & ChrW(232) & "res (" & MAX_PASSWORD_LENGTH & _
                        " au plus), diff" & ChrW(233) & "rent de l'identifiant."
End Function

' Nombre d'echecs consecutifs avant blocage (table Parametre, 5 par defaut)
Public Function MaxFailedAttempts() As Long
    MaxFailedAttempts = GetIntegerParameter(PARAM_MAX_FAILURES, DEFAULT_MAX_FAILURES)
End Function

' Duree du blocage en minutes (table Parametre, 15 par defaut)
Public Function LockoutMinutes() As Long
    LockoutMinutes = GetIntegerParameter(PARAM_LOCKOUT_MINUTES, DEFAULT_LOCKOUT_MINUTES)
End Function

'==============================================================================
' MOTS DE PASSE
'==============================================================================

' Regles du mot de passe. Retourne un message d'erreur, ou une chaine vide si le
' mot de passe est accepte. A utiliser aussi dans les formulaires (controle a la saisie).
Public Function ValidatePasswordRules(ByVal pPassword As String, Optional ByVal pLogin As String = "") As String
    If Len(Trim$(pPassword)) = 0 Then
        ValidatePasswordRules = "Le mot de passe ne peut pas etre vide."
    ElseIf Len(pPassword) < MIN_PASSWORD_LENGTH Then
        ValidatePasswordRules = "Le mot de passe doit contenir au moins " & MIN_PASSWORD_LENGTH & " caracteres."
    ElseIf Len(pPassword) > MAX_PASSWORD_LENGTH Then
        ValidatePasswordRules = "Le mot de passe doit contenir au plus " & MAX_PASSWORD_LENGTH & " caracteres."
    ElseIf Len(Trim$(pLogin)) > 0 And StrComp(pPassword, Trim$(pLogin), vbTextCompare) = 0 Then
        ValidatePasswordRules = "Le mot de passe doit etre different de l'identifiant."
    End If
End Function

' Changement par l'utilisateur lui-meme : il doit fournir son ancien mot de passe.
' Remet a zero les echecs et supprime l'obligation de changer.
Public Sub ChangePassword(ByVal pUserId As Long, ByVal pOldPassword As String, _
                             ByVal pNewPassword As String)
    Dim AdoRs As ADODB.Recordset
    Dim strLogin As String
    Dim strHash As String
    Dim strSalt As String
    Dim strMsg As String

    Set AdoRs = QueryRecordset("SELECT [Login], [MotDePasseHash], [Sel] FROM [Utilisateur] WHERE [UtilisateurID] = ?", _
                            pUserId)
    If AdoRs.EOF Then
        AdoRs.Close
        Err.Raise ERR_SEC_NOT_FOUND, "modKamerixSecurity.ChangePassword", _
            "Utilisateur introuvable (identifiant " & pUserId & ")."
    End If
    strLogin = CStr(AdoRs.Fields("Login").Value)
    strHash = CStr(AdoRs.Fields("MotDePasseHash").Value)
    strSalt = CStr(AdoRs.Fields("Sel").Value)
    AdoRs.Close

    If Not VerifyPassword(pOldPassword, strSalt, strHash) Then
        Err.Raise ERR_SEC_WRONG_PASSWORD, "modKamerixSecurity.ChangePassword", _
            "Ancien mot de passe incorrect."
    End If

    strMsg = ValidatePasswordRules(pNewPassword, strLogin)
    If Len(strMsg) > 0 Then Err.Raise ERR_SEC_INVALID_DATA, "modKamerixSecurity.ChangePassword", strMsg

    If pNewPassword = pOldPassword Then
        Err.Raise ERR_SEC_INVALID_DATA, "modKamerixSecurity.ChangePassword", _
            "Le nouveau mot de passe doit etre different de l'ancien."
    End If

    SetPassword pUserId, pNewPassword, False
End Sub

' Reinitialisation par un administrateur : mot de passe temporaire, a changer a la
' prochaine connexion. Debloque aussi le compte. (Droit a controler dans la Session.)
Public Sub ResetPassword(ByVal pUserId As Long, ByVal pTempPassword As String)
    Dim strMsg As String

    strMsg = ValidatePasswordRules(pTempPassword, LoginOf(pUserId))
    If Len(strMsg) > 0 Then Err.Raise ERR_SEC_INVALID_DATA, "modKamerixSecurity.ResetPassword", strMsg

    SetPassword pUserId, pTempPassword, True
End Sub

' Leve le blocage d'un compte sans changer son mot de passe. (Droit a controler dans la Session.)
Public Sub UnlockUser(ByVal pUserId As Long)
    Dim lngRows As Long

    lngRows = ExecuteNonQuery( _
        "UPDATE [Utilisateur] SET [NbEchecs] = 0, [DateBlocage] = NULL WHERE [UtilisateurID] = ?", pUserId)
    If lngRows = 0 Then
        Err.Raise ERR_SEC_NOT_FOUND, "modKamerixSecurity.UnlockUser", _
            "Utilisateur introuvable (identifiant " & pUserId & ")."
    End If
End Sub

'==============================================================================
' GESTION DES COMPTES (droits a controler dans la Session)
'==============================================================================

' Liste des comptes pour l'ecran d'administration, triee par identifiant.
' Colonnes : UtilisateurID, Login, NomComplet, RoleID, NomRole, Actif, DoitChangerMdp, DateBlocage
Public Function UserList() As ADODB.Recordset
    Set UserList = QueryRecordset( _
        "SELECT u.[UtilisateurID], u.[Login], u.[NomComplet], u.[RoleID], r.[Nom] AS NomRole, " & _
        "u.[Actif], u.[DoitChangerMdp], u.[DateBlocage] " & _
        "FROM [Utilisateur] AS u INNER JOIN [Role] AS r ON u.[RoleID] = r.[RoleID] " & _
        "ORDER BY u.[Login]")
End Function

' Liste des roles (RoleID, Nom), dans l'ordre de creation
Public Function RoleList() As ADODB.Recordset
    Set RoleList = QueryRecordset("SELECT [RoleID], [Nom] FROM [Role] ORDER BY [RoleID]")
End Function

' Active ou desactive un compte. Un compte inactif ne peut plus se connecter.
' Le dernier administrateur actif ne peut pas etre desactive.
Public Sub SetUserActive(ByVal pUserId As Long, ByVal pActive As Boolean)
    LoginOf pUserId          ' erreur ERR_SEC_NOT_FOUND si le compte n'existe pas

    If Not pActive Then
        If IsActiveAdmin(pUserId) And ActiveAdminCount() <= 1 Then
            Err.Raise ERR_SEC_LAST_ADMIN, "modKamerixSecurity.SetUserActive", _
                "Impossible : c'est le dernier administrateur actif."
        End If
    End If

    ExecuteNonQuery "UPDATE [Utilisateur] SET [Actif] = ? WHERE [UtilisateurID] = ?", pActive, pUserId
End Sub

' Change le role d'un compte. Le dernier administrateur actif ne peut pas changer de role.
Public Sub ChangeUserRole(ByVal pUserId As Long, ByVal pRoleId As Long)
    LoginOf pUserId
    If Not RoleExists(pRoleId) Then
        Err.Raise ERR_SEC_UNKNOWN_ROLE, "modKamerixSecurity.ChangeUserRole", _
            "Role inconnu (identifiant " & pRoleId & ")."
    End If

    If pRoleId <> GetRoleIdByName(ROLE_ADMINISTRATOR) Then
        If IsActiveAdmin(pUserId) And ActiveAdminCount() <= 1 Then
            Err.Raise ERR_SEC_LAST_ADMIN, "modKamerixSecurity.ChangeUserRole", _
                "Impossible : c'est le dernier administrateur actif."
        End If
    End If

    ExecuteNonQuery "UPDATE [Utilisateur] SET [RoleID] = ? WHERE [UtilisateurID] = ?", pRoleId, pUserId
End Sub

' Mot de passe provisoire aleatoire, lisible (sans 0/O, 1/l/I) : "Kx7m-Pq4z-R2wt"
Public Function GenerateTemporaryPassword() As String
    GenerateTemporaryPassword = RandomCode(CHARS_PASSWORD, 3, 4)
End Function

'==============================================================================
' CODE DE SECOURS (redefinir le mot de passe d'un administrateur)
'------------------------------------------------------------------------------
' Un seul code par base, affiche une seule fois a sa creation, stocke hache
' (comme un mot de passe) dans la table Parametre. Utiliser le code le consomme :
' un nouveau code est cree et doit etre note a son tour.
'==============================================================================

' Vrai si un code de secours est enregistre
Public Function RecoveryCodeExists() As Boolean
    RecoveryCodeExists = Not IsNull(QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [Cle] = ?", PARAM_RECOVERY_HASH))
End Function

' Cree un nouveau code de secours (l'ancien ne marche plus) et le retourne EN CLAIR :
' c'est la seule fois ou il est visible. Format : "ABCD-EFGH-JKLM-NPQR-STUV".
Public Function CreateRecoveryCode() As String
    Dim strCode As String
    Dim strSalt As String

    strCode = RandomCode(CHARS_RECOVERY, 5, 4)
    strSalt = GenerateSalt()
    SetParameter PARAM_RECOVERY_SALT, strSalt
    SetParameter PARAM_RECOVERY_HASH, HashPassword(NormalizeRecoveryCode(strCode), strSalt, CurrentIterations())
    CreateRecoveryCode = strCode
End Function

' Redefinit le mot de passe d'un administrateur avec le code de secours.
' Le compte est aussi debloque et reactive. Le code est consomme : la fonction
' retourne le NOUVEAU code de secours, a afficher et a noter.
' Refus generique (ERR_SEC_RECOVERY_INVALID) si le code, l'identifiant ou le role ne conviennent pas.
Public Function ResetAdminWithRecoveryCode(ByVal pCode As String, ByVal pLogin As String, _
                                           ByVal pNewPassword As String) As String
    Dim AdoRs As ADODB.Recordset
    Dim vUserId As Variant
    Dim strLogin As String
    Dim strRole As String
    Dim blnCodeOk As Boolean
    Dim blnOwnTransaction As Boolean
    Dim strMsg As String
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    strLogin = Trim$(pLogin)
    vUserId = Null

    strMsg = ValidatePasswordRules(pNewPassword, strLogin)
    If Len(strMsg) > 0 Then Err.Raise ERR_SEC_INVALID_DATA, "modKamerixSecurity.ResetAdminWithRecoveryCode", strMsg

    If Not RecoveryCodeExists() Then
        Err.Raise ERR_SEC_RECOVERY_INVALID, "modKamerixSecurity.ResetAdminWithRecoveryCode", _
            "Aucun code de secours n'est enregistre dans cette base."
    End If

    ' Le code est toujours verifie (meme duree de reponse quel que soit l'identifiant)
    blnCodeOk = VerifyPassword(NormalizeRecoveryCode(pCode), _
                               CStr(QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [Cle] = ?", PARAM_RECOVERY_SALT)), _
                               CStr(QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [Cle] = ?", PARAM_RECOVERY_HASH)))

    Set AdoRs = QueryRecordset( _
        "SELECT u.[UtilisateurID], r.[Nom] AS NomRole FROM [Utilisateur] AS u " & _
        "INNER JOIN [Role] AS r ON u.[RoleID] = r.[RoleID] WHERE u.[Login] = ?", strLogin)
    If Not AdoRs.EOF Then
        vUserId = CLng(AdoRs.Fields("UtilisateurID").Value)
        strRole = CStr(AdoRs.Fields("NomRole").Value)
    End If
    AdoRs.Close

    If Not blnCodeOk Or IsNull(vUserId) Or strRole <> ROLE_ADMINISTRATOR Then
        LogLoginAttempt vUserId, strLogin, RESULT_FAILURE
        Err.Raise ERR_SEC_RECOVERY_INVALID, "modKamerixSecurity.ResetAdminWithRecoveryCode", _
            "Code de secours ou identifiant d'administrateur incorrect."
    End If

    ' Tout ou rien : mot de passe, reactivation et nouveau code
    blnOwnTransaction = Not TransactionActive()
    If blnOwnTransaction Then BeginTransaction
    On Error GoTo EH

    SetPassword CLng(vUserId), pNewPassword, False
    ExecuteNonQuery "UPDATE [Utilisateur] SET [Actif] = ? WHERE [UtilisateurID] = ?", True, CLng(vUserId)
    LogLoginAttempt vUserId, strLogin, RESULT_RECOVERY
    ResetAdminWithRecoveryCode = CreateRecoveryCode()

    If blnOwnTransaction Then CommitTransaction
    Exit Function

EH:
    lngErr = Err.Number
    strErr = Err.Description
    strSrc = Err.Source
    If blnOwnTransaction Then RollbackTransaction
    Err.Raise lngErr, strSrc, strErr
End Function

'==============================================================================
' TESTS
'==============================================================================

' Fixe le nombre d'iterations du hachage pour les tests (minimum 1000) ; 0 = valeur standard.
' A ne jamais appeler en production.
Public Sub SetIterationsForTests(ByVal pIterations As Long)
    If pIterations <> 0 And pIterations < 1000 Then
        Err.Raise ERR_SEC_INVALID_DATA, "modKamerixSecurity.SetIterationsForTests", _
            "Le nombre d'iterations doit etre au moins 1000 (ou 0 pour la valeur standard)."
    End If
    mlngIterations = pIterations
End Sub

'==============================================================================
' PROCEDURES INTERNES
'==============================================================================

' Vrai si le compte est un administrateur actif
Private Function IsActiveAdmin(ByVal pUserId As Long) As Boolean
    IsActiveAdmin = (CLng(QueryValue( _
        "SELECT COUNT(*) FROM [Utilisateur] AS u INNER JOIN [Role] AS r ON u.[RoleID] = r.[RoleID] " & _
        "WHERE u.[UtilisateurID] = ? AND r.[Nom] = ? AND u.[Actif] = ?", pUserId, ROLE_ADMINISTRATOR, True)) > 0)
End Function

' Nombre d'administrateurs actifs
Private Function ActiveAdminCount() As Long
    ActiveAdminCount = CLng(QueryValue( _
        "SELECT COUNT(*) FROM [Utilisateur] AS u INNER JOIN [Role] AS r ON u.[RoleID] = r.[RoleID] " & _
        "WHERE r.[Nom] = ? AND u.[Actif] = ?", ROLE_ADMINISTRATOR, True))
End Function

' Code aleatoire : pGroups groupes de pGroupLength caracteres de pChars, separes par "-".
' Tirage sans biais : les octets au-dela du dernier multiple de Len(pChars) sont rejetes.
Private Function RandomCode(ByVal pChars As String, ByVal pGroups As Long, ByVal pGroupLength As Long) As String
    Dim bytRandom() As Byte
    Dim lngLimit As Long
    Dim lngNeeded As Long
    Dim strOut As String
    Dim i As Long

    lngLimit = (256 \ Len(pChars)) * Len(pChars)
    lngNeeded = pGroups * pGroupLength

    Do While Len(Replace(strOut, "-", "")) < lngNeeded
        bytRandom = RandomBytes(lngNeeded * 2)
        For i = LBound(bytRandom) To UBound(bytRandom)
            If bytRandom(i) < lngLimit Then
                If Len(strOut) > 0 And (Len(Replace(strOut, "-", "")) Mod pGroupLength) = 0 Then strOut = strOut & "-"
                strOut = strOut & Mid$(pChars, (bytRandom(i) Mod Len(pChars)) + 1, 1)
                If Len(Replace(strOut, "-", "")) = lngNeeded Then Exit For
            End If
        Next i
    Loop

    RandomCode = strOut
End Function

' Code de secours saisi -> forme comparee : majuscules, sans tirets ni espaces
Private Function NormalizeRecoveryCode(ByVal pCode As String) As String
    NormalizeRecoveryCode = UCase$(Replace(Replace(Trim$(pCode), "-", ""), " ", ""))
End Function

' Ecrit une valeur dans la table Parametre (mise a jour, ou creation de la cle)
Private Sub SetParameter(ByVal pKey As String, ByVal pValue As String)
    If ExecuteNonQuery("UPDATE [Parametre] SET [Valeur] = ? WHERE [Cle] = ?", pValue, pKey) = 0 Then
        ExecuteNonQuery "INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES (?, ?)", pKey, pValue
    End If
End Sub

' Nombre d'iterations d'une valeur stockee "iterations$hash" (0 si mal formee)
Private Function StoredIterations(ByVal pStored As String) As Long
    Dim lngPos As Long

    lngPos = InStr(pStored, "$")
    If lngPos < 2 Then Exit Function
    If Not IsNumeric(Left$(pStored, lngPos - 1)) Then Exit Function
    StoredIterations = CLng(Left$(pStored, lngPos - 1))
End Function

' Recalcule le hash avec le reglage en vigueur apres une connexion reussie.
' Nouveau sel ; ne touche ni a DoitChangerMdp ni au blocage. Un echec ici ne
' doit pas empecher la connexion : l'ancien hash reste valable.
Private Sub RehashPassword(ByVal pUserId As Long, ByVal pPassword As String)
    Dim strSalt As String
    Dim strHash As String

    On Error GoTo EH
    strSalt = GenerateSalt()
    strHash = HashPassword(pPassword, strSalt, CurrentIterations())
    ExecuteNonQuery "UPDATE [Utilisateur] SET [MotDePasseHash] = ?, [Sel] = ? WHERE [UtilisateurID] = ?", _
        strHash, strSalt, pUserId
    Exit Sub

EH:
    Debug.Print "RehashPassword (utilisateur " & pUserId & ") : erreur " & Err.Number & " - " & Err.Description
End Sub

Private Function CurrentIterations() As Long
    If mlngIterations = 0 Then
        CurrentIterations = DEFAULT_ITERATIONS
    Else
        CurrentIterations = mlngIterations
    End If
End Function

' Valide les donnees puis insere le compte (Actif, sans echec, non bloque)
Private Function InsertUser(ByVal pLogin As String, ByVal pFullName As String, _
                                    ByVal pRoleId As Long, ByVal pPassword As String, _
                                    ByVal pMustChange As Boolean) As Long
    Dim strLogin As String
    Dim strName As String
    Dim strMsg As String
    Dim strSalt As String
    Dim strHash As String

    strLogin = Trim$(pLogin)
    strName = Trim$(pFullName)

    If Len(strLogin) < MIN_LOGIN_LENGTH Or Len(strLogin) > MAX_LOGIN_LENGTH Then
        Err.Raise ERR_SEC_INVALID_DATA, "modKamerixSecurity.InsertUser", _
            "L'identifiant doit contenir de " & MIN_LOGIN_LENGTH & " a " & MAX_LOGIN_LENGTH & " caracteres."
    End If
    If InStr(strLogin, " ") > 0 Then
        Err.Raise ERR_SEC_INVALID_DATA, "modKamerixSecurity.InsertUser", _
            "L'identifiant ne doit pas contenir d'espace."
    End If
    If Len(strName) = 0 Or Len(strName) > MAX_NAME_LENGTH Then
        Err.Raise ERR_SEC_INVALID_DATA, "modKamerixSecurity.InsertUser", _
            "Le nom complet est obligatoire (" & MAX_NAME_LENGTH & " caracteres au plus)."
    End If

    strMsg = ValidatePasswordRules(pPassword, strLogin)
    If Len(strMsg) > 0 Then Err.Raise ERR_SEC_INVALID_DATA, "modKamerixSecurity.InsertUser", strMsg

    If LoginExists(strLogin) Then
        Err.Raise ERR_SEC_LOGIN_EXISTS, "modKamerixSecurity.InsertUser", _
            "Cet identifiant est deja utilise."
    End If

    strSalt = GenerateSalt()
    strHash = HashPassword(pPassword, strSalt, CurrentIterations())

    InsertUser = ExecuteInsert( _
        "INSERT INTO [Utilisateur] ([Login], [NomComplet], [MotDePasseHash], [Sel], [RoleID], [Actif], " & _
        "[DoitChangerMdp], [NbEchecs], [DateBlocage]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", _
        strLogin, strName, strHash, strSalt, pRoleId, True, pMustChange, 0, Null)
End Function

' Nouveau sel et nouveau hash ; remet les echecs a zero et leve le blocage
Private Sub SetPassword(ByVal pUserId As Long, ByVal pPassword As String, _
                              ByVal pMustChange As Boolean)
    Dim strSalt As String
    Dim strHash As String
    Dim lngRows As Long

    strSalt = GenerateSalt()
    strHash = HashPassword(pPassword, strSalt, CurrentIterations())

    lngRows = ExecuteNonQuery( _
        "UPDATE [Utilisateur] SET [MotDePasseHash] = ?, [Sel] = ?, [DoitChangerMdp] = ?, " & _
        "[NbEchecs] = 0, [DateBlocage] = NULL WHERE [UtilisateurID] = ?", _
        strHash, strSalt, pMustChange, pUserId)
    If lngRows = 0 Then
        Err.Raise ERR_SEC_NOT_FOUND, "modKamerixSecurity.SetPassword", _
            "Utilisateur introuvable (identifiant " & pUserId & ")."
    End If
End Sub

' Identifiant (login) d'un utilisateur ; erreur s'il n'existe pas
Private Function LoginOf(ByVal pUserId As Long) As String
    Dim vLogin As Variant

    vLogin = QueryValue("SELECT [Login] FROM [Utilisateur] WHERE [UtilisateurID] = ?", pUserId)
    If IsNull(vLogin) Then
        Err.Raise ERR_SEC_NOT_FOUND, "modKamerixSecurity.LoginOf", _
            "Utilisateur introuvable (identifiant " & pUserId & ")."
    End If
    LoginOf = CStr(vLogin)
End Function

' Vrai si l'identifiant est deja pris (comparaison insensible a la casse, comme Access)
Private Function LoginExists(ByVal pLogin As String) As Boolean
    LoginExists = (CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur] WHERE [Login] = ?", pLogin)) > 0)
End Function

Private Function RoleExists(ByVal pRoleId As Long) As Boolean
    RoleExists = (CLng(QueryValue("SELECT COUNT(*) FROM [Role] WHERE [RoleID] = ?", pRoleId)) > 0)
End Function

' Valeur entiere positive d'un parametre ; valeur par defaut si absent ou invalide
Private Function GetIntegerParameter(ByVal pKey As String, ByVal pDefault As Long) As Long
    Dim vValue As Variant

    GetIntegerParameter = pDefault
    vValue = QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [Cle] = ?", pKey)
    If IsNull(vValue) Then Exit Function
    If Not IsNumeric(vValue) Then Exit Function
    If CLng(vValue) < 1 Then Exit Function
    GetIntegerParameter = CLng(vValue)
End Function

' Inscrit une tentative dans JournalConnexion (pUserId = Null si l'identifiant est inconnu)
Private Sub LogLoginAttempt(ByVal pUserId As Variant, ByVal pLogin As String, _
                                 ByVal pResult As String)
    ExecuteNonQuery _
        "INSERT INTO [JournalConnexion] ([UtilisateurID], [LoginSaisi], [DateHeure], [Resultat], [Poste]) " & _
        "VALUES (?, ?, ?, ?, ?)", _
        pUserId, NullIfEmpty(Left$(pLogin, MAX_LOGIN_LENGTH)), Now, pResult, NullIfEmpty(WorkstationName())
End Sub

Private Function WorkstationName() As String
    WorkstationName = Left$(Environ$("COMPUTERNAME"), 100)
End Function

' Calcule un hash sans rien en faire : un identifiant inconnu prend le meme temps
' de reponse qu'un identifiant connu, ce qui ne revele pas quels comptes existent.
Private Sub SpendHashTime(ByVal pPassword As String)
    Dim strUnused As String

    If Len(pPassword) = 0 Then pPassword = "x"
    strUnused = HashPassword(pPassword, String$(64, "0"), CurrentIterations())
End Sub
