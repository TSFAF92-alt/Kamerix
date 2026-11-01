Attribute VB_Name = "modTestSecurity"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Tests du module de securite (modKamerixSecurity)
'------------------------------------------------------------------------------
' A lancer (F5 sur TestSecurity) sur une base de developpement SANS utilisateur.
' Tout se passe dans une transaction annulee a la fin : la base reste intacte.
' Le hachage utilise 1000 iterations pendant le test (au lieu de 600 000) pour
' que le test dure quelques secondes ; la valeur standard est ensuite remise.
'==============================================================================

Private Const PWD_ADMIN As String = "Secret#2026"
Private Const PWD_TEMP As String = "Temporaire1"
Private Const PWD_NEW As String = "NouveauMdp#1"
Private Const PWD_PROVISIONAL As String = "Provisoire#9"

Public Sub TestSecurity()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim lngAdmin As Long
    Dim lngSeller As Long
    Dim lngId As Long
    Dim blnMustChange As Boolean
    Dim lngErr As Long
    Dim strErrText As String
    Dim strStoredHash As String
    Dim i As Long
    Dim eRes As KxLoginResult
    Dim blnTransaction As Boolean

    On Error GoTo EH

    If CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")) > 0 Then
        MsgBox "Ce test demande une base sans utilisateur." & vbCrLf & _
               "Supprimez la base de developpement et recreez-la, puis relancez le test.", _
               vbExclamation, "Kamerix - tests de securite"
        Exit Sub
    End If

    SetIterationsForTests 1000
    BeginTransaction
    blnTransaction = True

    ' ---------- Regles du mot de passe ----------
    CheckTrue "Premier lancement : aucun administrateur", Not AdminExists(), lngOk, lngTotal, strReport
    CheckTrue "Mot de passe trop court refuse", Len(ValidatePasswordRules("abc12", "bob")) > 0, lngOk, lngTotal, strReport
    CheckTrue "Mot de passe de 8 caracteres accepte", Len(ValidatePasswordRules("abcd1234", "bob")) = 0, lngOk, lngTotal, strReport
    CheckTrue "Mot de passe egal a l'identifiant refuse (casse ignoree)", _
        Len(ValidatePasswordRules("Administrateur", "administrateur")) > 0, lngOk, lngTotal, strReport
    CheckTrue "Mot de passe fait d'espaces refuse", Len(ValidatePasswordRules("        ", "bob")) > 0, lngOk, lngTotal, strReport

    ' ---------- Premier administrateur ----------
    lngAdmin = CreateFirstAdmin("admin", "Administrateur Test", PWD_ADMIN)
    CheckTrue "Premier administrateur cree", lngAdmin > 0, lngOk, lngTotal, strReport
    CheckTrue "AdminExists apres creation", AdminExists(), lngOk, lngTotal, strReport

    On Error Resume Next
    CreateFirstAdmin "admin2", "Autre Admin", PWD_ADMIN
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Un second premier administrateur est refuse", lngErr, ERR_SEC_ADMIN_EXISTS, lngOk, lngTotal, strReport

    strStoredHash = CStr(QueryValue("SELECT [MotDePasseHash] FROM [Utilisateur] WHERE [UtilisateurID] = ?", lngAdmin))
    CheckTrue "Hash stocke au format iterations$hash", (Left$(strStoredHash, 5) = "1000$"), lngOk, lngTotal, strReport
    CheckTrue "Le mot de passe n'apparait pas dans la base", (InStr(strStoredHash, PWD_ADMIN) = 0), lngOk, lngTotal, strReport
    CheckTrue "L'administrateur n'a pas a changer son mot de passe", _
        Not CBool(QueryValue("SELECT [DoitChangerMdp] FROM [Utilisateur] WHERE [UtilisateurID] = ?", lngAdmin)), _
        lngOk, lngTotal, strReport

    ' ---------- Connexions simples ----------
    eRes = TryAuthenticate("admin", PWD_ADMIN, lngId, blnMustChange)
    CheckTrue "Connexion correcte acceptee", eRes = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckEq "Identifiant retourne", lngId, lngAdmin, lngOk, lngTotal, strReport
    CheckTrue "Pas de changement de mot de passe exige", Not blnMustChange, lngOk, lngTotal, strReport

    eRes = TryAuthenticate("ADMIN", PWD_ADMIN)
    CheckTrue "Identifiant insensible a la casse", eRes = kxLoginSuccess, lngOk, lngTotal, strReport

    eRes = TryAuthenticate("admin", "mauvais-mot-de-passe", lngId)
    CheckTrue "Mauvais mot de passe refuse", eRes = kxLoginDenied, lngOk, lngTotal, strReport
    CheckEq "Aucun identifiant retourne apres un echec", lngId, 0, lngOk, lngTotal, strReport

    eRes = TryAuthenticate("inconnu", PWD_ADMIN)
    CheckTrue "Identifiant inconnu refuse", eRes = kxLoginDenied, lngOk, lngTotal, strReport

    eRes = TryAuthenticate("", "")
    CheckTrue "Saisie vide refusee", eRes = kxLoginDenied, lngOk, lngTotal, strReport

    CheckEq "Journal : 2 connexions reussies", _
        CLng(QueryValue("SELECT COUNT(*) FROM [JournalConnexion] WHERE [Resultat] = ?", "Succes")), 2, _
        lngOk, lngTotal, strReport
    CheckEq "Journal : 3 echecs (dont l'identifiant inconnu et la saisie vide)", _
        CLng(QueryValue("SELECT COUNT(*) FROM [JournalConnexion] WHERE [Resultat] = ?", "Echec")), 3, _
        lngOk, lngTotal, strReport
    CheckEq "Journal : l'identifiant inconnu est conserve", _
        CLng(QueryValue("SELECT COUNT(*) FROM [JournalConnexion] WHERE [LoginSaisi] = ? AND [UtilisateurID] IS NULL", "inconnu")), 1, _
        lngOk, lngTotal, strReport

    ' ---------- Blocage ----------
    TryAuthenticate "admin", PWD_ADMIN                      ' remet le compteur d'echecs a zero
    For i = 1 To 5
        eRes = TryAuthenticate("admin", "mauvais-" & i)
        If i < 5 Then
            CheckTrue "Echec " & i & " sur 5 : refuse sans blocage", eRes = kxLoginDenied, lngOk, lngTotal, strReport
        Else
            CheckTrue "5e echec : compte bloque", eRes = kxLoginLocked, lngOk, lngTotal, strReport
        End If
    Next i

    eRes = TryAuthenticate("admin", PWD_ADMIN)
    CheckTrue "Bon mot de passe refuse pendant le blocage", eRes = kxLoginLocked, lngOk, lngTotal, strReport
    CheckTrue "Date de blocage enregistree", _
        Not IsNull(QueryValue("SELECT [DateBlocage] FROM [Utilisateur] WHERE [UtilisateurID] = ?", lngAdmin)), _
        lngOk, lngTotal, strReport
    CheckEq "Journal : 2 lignes 'Bloque'", _
        CLng(QueryValue("SELECT COUNT(*) FROM [JournalConnexion] WHERE [Resultat] = ?", "Bloque")), 2, _
        lngOk, lngTotal, strReport

    ' Fin du blocage simulee en reculant la date d'une minute
    ExecuteNonQuery "UPDATE [Utilisateur] SET [DateBlocage] = ? WHERE [UtilisateurID] = ?", DateAdd("n", -1, Now), lngAdmin
    eRes = TryAuthenticate("admin", PWD_ADMIN)
    CheckTrue "Connexion de nouveau possible apres le blocage", eRes = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckTrue "Blocage et echecs remis a zero", _
        IsNull(QueryValue("SELECT [DateBlocage] FROM [Utilisateur] WHERE [UtilisateurID] = ?", lngAdmin)) And _
        CLng(QueryValue("SELECT [NbEchecs] FROM [Utilisateur] WHERE [UtilisateurID] = ?", lngAdmin)) = 0, _
        lngOk, lngTotal, strReport

    ' Deblocage par un administrateur
    For i = 1 To 5
        eRes = TryAuthenticate("admin", "mauvais-" & i)
    Next i
    CheckTrue "Compte de nouveau bloque", eRes = kxLoginLocked, lngOk, lngTotal, strReport
    UnlockUser lngAdmin
    eRes = TryAuthenticate("admin", PWD_ADMIN)
    CheckTrue "UnlockUser rend la connexion possible", eRes = kxLoginSuccess, lngOk, lngTotal, strReport

    ' ---------- Compte inactif ----------
    ExecuteNonQuery "UPDATE [Utilisateur] SET [Actif] = ? WHERE [UtilisateurID] = ?", False, lngAdmin
    eRes = TryAuthenticate("admin", PWD_ADMIN)
    CheckTrue "Compte inactif refuse meme avec le bon mot de passe", eRes = kxLoginDenied, lngOk, lngTotal, strReport
    ExecuteNonQuery "UPDATE [Utilisateur] SET [Actif] = ?, [NbEchecs] = 0 WHERE [UtilisateurID] = ?", True, lngAdmin

    ' ---------- Creation d'un utilisateur ----------
    lngSeller = CreateUser("vendeur1", "Vendeur Test", GetRoleIdByName("Vendeur"), PWD_TEMP)
    CheckTrue "Utilisateur Vendeur cree", lngSeller > 0, lngOk, lngTotal, strReport
    CheckTrue "Mot de passe temporaire : changement exige", _
        CBool(QueryValue("SELECT [DoitChangerMdp] FROM [Utilisateur] WHERE [UtilisateurID] = ?", lngSeller)), _
        lngOk, lngTotal, strReport

    eRes = TryAuthenticate("vendeur1", PWD_TEMP, lngId, blnMustChange)
    CheckTrue "Connexion avec le mot de passe temporaire", eRes = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckTrue "L'ecran sera prevenu : changement obligatoire", blnMustChange, lngOk, lngTotal, strReport

    On Error Resume Next
    CreateUser "VENDEUR1", "Doublon", GetRoleIdByName("Vendeur"), PWD_TEMP
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Identifiant en double refuse (casse ignoree)", lngErr, ERR_SEC_LOGIN_EXISTS, lngOk, lngTotal, strReport

    On Error Resume Next
    CreateUser "a b", "Espace", GetRoleIdByName("Vendeur"), PWD_TEMP
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Identifiant avec espace refuse", lngErr, ERR_SEC_INVALID_DATA, lngOk, lngTotal, strReport

    On Error Resume Next
    CreateUser "court1", "Mdp court", GetRoleIdByName("Vendeur"), "abc"
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Mot de passe trop court refuse a la creation", lngErr, ERR_SEC_INVALID_DATA, lngOk, lngTotal, strReport

    On Error Resume Next
    CreateUser "fantome", "Role inconnu", 99999, PWD_TEMP
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Role inexistant refuse", lngErr, ERR_SEC_UNKNOWN_ROLE, lngOk, lngTotal, strReport

    ' ---------- Changement de mot de passe ----------
    On Error Resume Next
    ChangePassword lngSeller, "pas-le-bon", PWD_NEW
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Ancien mot de passe incorrect refuse", lngErr, ERR_SEC_WRONG_PASSWORD, lngOk, lngTotal, strReport

    On Error Resume Next
    ChangePassword lngSeller, PWD_TEMP, "court1"
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Nouveau mot de passe trop court refuse", lngErr, ERR_SEC_INVALID_DATA, lngOk, lngTotal, strReport

    On Error Resume Next
    ChangePassword lngSeller, PWD_TEMP, PWD_TEMP
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Nouveau mot de passe identique a l'ancien refuse", lngErr, ERR_SEC_INVALID_DATA, lngOk, lngTotal, strReport

    ChangePassword lngSeller, PWD_TEMP, PWD_NEW
    eRes = TryAuthenticate("vendeur1", PWD_TEMP)
    CheckTrue "L'ancien mot de passe ne marche plus", eRes = kxLoginDenied, lngOk, lngTotal, strReport
    eRes = TryAuthenticate("vendeur1", PWD_NEW, lngId, blnMustChange)
    CheckTrue "Le nouveau mot de passe marche", eRes = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckTrue "Plus de changement exige apres le changement", Not blnMustChange, lngOk, lngTotal, strReport

    ' ---------- Reinitialisation ----------
    For i = 1 To 5
        eRes = TryAuthenticate("vendeur1", "mauvais-" & i)
    Next i
    CheckTrue "Vendeur bloque avant la reinitialisation", eRes = kxLoginLocked, lngOk, lngTotal, strReport

    ResetPassword lngSeller, PWD_PROVISIONAL
    eRes = TryAuthenticate("vendeur1", PWD_PROVISIONAL, lngId, blnMustChange)
    CheckTrue "La reinitialisation debloque et le mot de passe provisoire marche", eRes = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckTrue "Changement exige apres une reinitialisation", blnMustChange, lngOk, lngTotal, strReport
    eRes = TryAuthenticate("vendeur1", PWD_NEW)
    CheckTrue "L'ancien mot de passe ne marche plus apres la reinitialisation", eRes = kxLoginDenied, lngOk, lngTotal, strReport

    On Error Resume Next
    ResetPassword 99999, PWD_PROVISIONAL
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Reinitialisation d'un utilisateur inconnu refusee", lngErr, ERR_SEC_NOT_FOUND, lngOk, lngTotal, strReport

    ' ---------- Parametres ----------
    CheckEq "Maximum d'echecs : 5 par defaut", MaxFailedAttempts(), 5, lngOk, lngTotal, strReport
    CheckEq "Duree du blocage : 15 minutes par defaut", LockoutMinutes(), 15, lngOk, lngTotal, strReport

    ' La cle existe dans les bases creees depuis le lot 3 (5 par defaut) : on la modifie, sinon on la cree
    If IsNull(QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [Cle] = ?", PARAM_MAX_FAILURES)) Then
        ExecuteNonQuery "INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES (?, ?)", PARAM_MAX_FAILURES, "3"
    Else
        ExecuteNonQuery "UPDATE [Parametre] SET [Valeur] = ? WHERE [Cle] = ?", "3", PARAM_MAX_FAILURES
    End If
    CheckEq "Maximum d'echecs lu dans la table Parametre", MaxFailedAttempts(), 3, lngOk, lngTotal, strReport

    ' ---------- Fin : on annule tout ----------
    RollbackTransaction
    blnTransaction = False
    SetIterationsForTests 0

    CheckEq "La base est restee intacte (aucun utilisateur)", _
        CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")), 0, lngOk, lngTotal, strReport
    CheckEq "La base est restee intacte (journal vide)", _
        CLng(QueryValue("SELECT COUNT(*) FROM [JournalConnexion]")), 0, lngOk, lngTotal, strReport

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport, _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), "Kamerix - tests de securite"
    Exit Sub

EH:
    lngErr = Err.Number
    strErrText = Err.Description
    On Error Resume Next
    If blnTransaction Then RollbackTransaction
    SetIterationsForTests 0
    On Error GoTo 0
    MsgBox "Echec (erreur " & lngErr & ") :" & vbCrLf & vbCrLf & strErrText & vbCrLf & vbCrLf & _
           "Resultats deja obtenus :" & vbCrLf & strReport, vbCritical, "Kamerix - tests de securite"
End Sub

' Raccourci : tente une connexion en ignorant les sorties si elles ne sont pas demandees
Private Function TryAuthenticate(ByVal pLogin As String, ByVal pPassword As String, _
                       Optional ByRef pId As Long, Optional ByRef pMustChange As Boolean) As KxLoginResult
    TryAuthenticate = AuthenticateUser(pLogin, pPassword, pId, pMustChange)
End Function

' Comparaison de deux valeurs (converties en texte) ; affiche obtenu et attendu en cas d'echec
Private Sub CheckEq(ByVal strLabel As String, ByVal vActual As Variant, ByVal vExpected As Variant, _
                    ByRef lngOk As Long, ByRef lngTotal As Long, ByRef strReport As String)
    lngTotal = lngTotal + 1
    If CStr(vActual) = CStr(vExpected) Then
        lngOk = lngOk + 1
        strReport = strReport & "OK     " & strLabel & vbCrLf
    Else
        strReport = strReport & "ECHEC  " & strLabel & " (obtenu " & CStr(vActual) & ", attendu " & CStr(vExpected) & ")" & vbCrLf
    End If
End Sub

' Verification d'une condition vraie (independant de la langue de Windows/Office)
Private Sub CheckTrue(ByVal strLabel As String, ByVal blnCondition As Boolean, _
                      ByRef lngOk As Long, ByRef lngTotal As Long, ByRef strReport As String)
    lngTotal = lngTotal + 1
    If blnCondition Then
        lngOk = lngOk + 1
        strReport = strReport & "OK     " & strLabel & vbCrLf
    Else
        strReport = strReport & "ECHEC  " & strLabel & vbCrLf
    End If
End Sub
