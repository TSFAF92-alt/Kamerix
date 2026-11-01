Attribute VB_Name = "modTestAccounts"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Tests de la gestion des comptes et du code de secours
'------------------------------------------------------------------------------
' A lancer (F5 sur TestAccounts) sur une base de developpement SANS utilisateur,
' creee avec le schema du lot 3 (parametres de blocage inscrits).
' Tout se passe dans une transaction annulee a la fin : la base reste intacte.
'==============================================================================

Private Const PWD_ADMIN As String = "Secret#2026"
Private Const PWD_ADMIN2 As String = "Secret#2027"
Private Const PWD_NEW As String = "NouveauMdp#1"

Public Sub TestAccounts()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim lngAdmin As Long
    Dim lngAdmin2 As Long
    Dim lngSeller As Long
    Dim strTemp As String
    Dim strTemp2 As String
    Dim strCode As String
    Dim strCode2 As String
    Dim lngErr As Long
    Dim strErrText As String
    Dim blnTransaction As Boolean
    Dim i As Long

    On Error GoTo EH

    If CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")) > 0 Then
        MsgBox "Ce test demande une base sans utilisateur." & vbCrLf & _
               "Supprimez la base de developpement et recreez-la, puis relancez le test.", _
               vbExclamation, "Kamerix - tests des comptes"
        Exit Sub
    End If

    ' ---------- Parametres de blocage inscrits dans la base ----------
    CheckEq "Parametre MAX_ECHECS_CONNEXION = 5", _
        QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [Cle] = ?", PARAM_MAX_FAILURES), "5", lngOk, lngTotal, strReport
    CheckEq "Parametre DUREE_BLOCAGE_MINUTES = 15", _
        QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [Cle] = ?", PARAM_LOCKOUT_MINUTES), "15", lngOk, lngTotal, strReport

    ' ---------- Codes generes ----------
    strTemp = GenerateTemporaryPassword()
    strTemp2 = GenerateTemporaryPassword()
    CheckTrue "Mot de passe provisoire : format XXXX-XXXX-XXXX", strTemp Like "????-????-????", lngOk, lngTotal, strReport
    CheckEq "Mot de passe provisoire : conforme aux regles", ValidatePasswordRules(strTemp, "admin"), "", lngOk, lngTotal, strReport
    CheckTrue "Mot de passe provisoire : different a chaque tirage", strTemp <> strTemp2, lngOk, lngTotal, strReport
    CheckTrue "Mot de passe provisoire : sans caractere ambigu (0, O, 1, l, I)", _
        Not (strTemp Like "*[0O1lI]*"), lngOk, lngTotal, strReport

    CloseSession
    SetIterationsForTests 1000
    BeginTransaction
    blnTransaction = True

    ' ---------- Comptes crees par l'administrateur ----------
    lngAdmin = CreateFirstAdmin("admin", "Administrateur Test", PWD_ADMIN)
    CheckTrue "Pas de code de secours avant sa creation", Not RecoveryCodeExists(), lngOk, lngTotal, strReport
    OpenSession "admin", PWD_ADMIN

    lngSeller = CreateAccountWithTemporaryPassword("vendeur1", "Vendeur Test", GetRoleIdByName("Vendeur"), strTemp)
    CheckTrue "Compte cree avec un mot de passe provisoire genere", lngSeller > 0 And Len(strTemp) = 14, lngOk, lngTotal, strReport
    CheckTrue "ListAccounts : 2 comptes", RecordCount(ListAccounts()) = 2, lngOk, lngTotal, strReport

    CheckTrue "Le vendeur se connecte avec le mot de passe provisoire", _
        OpenSession("vendeur1", strTemp) = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckTrue "Le vendeur doit changer son mot de passe", MustChangePassword(), lngOk, lngTotal, strReport
    ChangeMyPassword strTemp, PWD_NEW

    CheckEq "Vendeur : liste des comptes refusee", ErrorOf("ListAccounts"), ERR_SESSION_DENIED, lngOk, lngTotal, strReport
    CheckEq "Vendeur : nouveau code de secours refuse", ErrorOf("RenewRecoveryCode"), ERR_SESSION_DENIED, lngOk, lngTotal, strReport

    ' ---------- Reinitialisation, desactivation, role ----------
    OpenSession "admin", PWD_ADMIN
    strTemp2 = ResetAccountWithTemporaryPassword(lngSeller)
    CheckTrue "Reinitialisation : l'ancien mot de passe ne marche plus", _
        OpenSession("vendeur1", PWD_NEW) = kxLoginDenied, lngOk, lngTotal, strReport
    CheckTrue "Reinitialisation : le mot de passe provisoire marche", _
        OpenSession("vendeur1", strTemp2) = kxLoginSuccess And MustChangePassword(), lngOk, lngTotal, strReport

    OpenSession "admin", PWD_ADMIN
    SetAccountActive lngSeller, False
    CheckTrue "Compte desactive : connexion refusee", OpenSession("vendeur1", strTemp2) = kxLoginDenied, lngOk, lngTotal, strReport
    OpenSession "admin", PWD_ADMIN
    SetAccountActive lngSeller, True
    CheckTrue "Compte reactive : connexion acceptee", OpenSession("vendeur1", strTemp2) = kxLoginSuccess, lngOk, lngTotal, strReport

    OpenSession "admin", PWD_ADMIN
    ChangeAccountRole lngSeller, GetRoleIdByName("Comptable")
    OpenSession "vendeur1", strTemp2
    CheckEq "Changement de role pris en compte a la connexion", CurrentRoleName(), "Comptable", lngOk, lngTotal, strReport

    ' ---------- Garde-fous ----------
    OpenSession "admin", PWD_ADMIN
    CheckEq "On ne peut pas desactiver son propre compte", ErrorOf("SelfDeactivate"), ERR_SESSION_DENIED, lngOk, lngTotal, strReport
    CheckEq "On ne peut pas changer son propre role", ErrorOf("SelfRole"), ERR_SESSION_DENIED, lngOk, lngTotal, strReport
    CheckEq "Dernier administrateur : desactivation refusee", ErrorOf("LastAdminDeactivate", lngAdmin), _
        ERR_SEC_LAST_ADMIN, lngOk, lngTotal, strReport
    CheckEq "Dernier administrateur : changement de role refuse", ErrorOf("LastAdminRole", lngAdmin), _
        ERR_SEC_LAST_ADMIN, lngOk, lngTotal, strReport

    lngAdmin2 = CreateUser("admin2", "Second Administrateur", GetRoleIdByName(ROLE_ADMINISTRATOR), PWD_ADMIN2)
    SetUserActive lngAdmin, False
    CheckTrue "Avec deux administrateurs, l'un peut etre desactive", _
        OpenSession("admin", PWD_ADMIN) = kxLoginDenied, lngOk, lngTotal, strReport
    SetUserActive lngAdmin, True

    ' ---------- RefreshSession detecte un compte desactive ----------
    OpenSession "vendeur1", strTemp2
    SetUserActive lngSeller, False
    RefreshSession
    CheckTrue "RefreshSession : compte desactive -> session fermee", Not IsLoggedIn(), lngOk, lngTotal, strReport
    SetUserActive lngSeller, True

    ' ---------- Code de secours ----------
    OpenSession "admin", PWD_ADMIN
    strCode = RenewRecoveryCode()
    CheckTrue "Code de secours : format XXXX-XXXX-XXXX-XXXX-XXXX", _
        strCode Like "????-????-????-????-????", lngOk, lngTotal, strReport
    CheckTrue "Code de secours enregistre", RecoveryCodeExists(), lngOk, lngTotal, strReport
    CheckTrue "Code de secours stocke hache (absent en clair)", _
        InStr(CStr(QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [Cle] = ?", "CODE_SECOURS_HASH")), _
              Replace(strCode, "-", "")) = 0, lngOk, lngTotal, strReport
    CloseSession

    CheckEq "Secours : mauvais code refuse", _
        RecoveryError("AAAA-BBBB-CCCC-DDDD-EEEE", "admin", PWD_NEW), ERR_SEC_RECOVERY_INVALID, lngOk, lngTotal, strReport
    CheckEq "Secours : compte non administrateur refuse", _
        RecoveryError(strCode, "vendeur1", PWD_NEW), ERR_SEC_RECOVERY_INVALID, lngOk, lngTotal, strReport
    CheckEq "Secours : identifiant inconnu refuse", _
        RecoveryError(strCode, "inconnu", PWD_NEW), ERR_SEC_RECOVERY_INVALID, lngOk, lngTotal, strReport
    CheckEq "Secours : mot de passe trop court refuse", _
        RecoveryError(strCode, "admin", "court"), ERR_SEC_INVALID_DATA, lngOk, lngTotal, strReport

    ' L'administrateur s'est bloque puis utilise le code (saisi en minuscules, avec espaces)
    For i = 1 To 5
        OpenSession "admin", "mauvais-" & i
    Next i
    CheckTrue "Administrateur bloque avant le secours", OpenSession("admin", PWD_ADMIN) = kxLoginLocked, lngOk, lngTotal, strReport
    strCode2 = ResetAdminWithRecoveryCode(" " & LCase$(Replace(strCode, "-", " ")) & " ", "admin", PWD_NEW)
    CheckTrue "Secours : nouveau code remis, different de l'ancien", _
        strCode2 Like "????-????-????-????-????" And strCode2 <> strCode, lngOk, lngTotal, strReport
    CheckTrue "Secours : connexion avec le nouveau mot de passe (compte debloque)", _
        OpenSession("admin", PWD_NEW) = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckTrue "Secours : pas de changement de mot de passe exige", Not MustChangePassword(), lngOk, lngTotal, strReport
    CloseSession
    CheckEq "Secours : l'ancien code ne marche plus", _
        RecoveryError(strCode, "admin", PWD_ADMIN), ERR_SEC_RECOVERY_INVALID, lngOk, lngTotal, strReport
    CheckTrue "Secours : journal (ligne Secours)", _
        CLng(QueryValue("SELECT COUNT(*) FROM [JournalConnexion] WHERE [Resultat] = ?", "Secours")) = 1, _
        lngOk, lngTotal, strReport

    ' ---------- Fin : on annule tout ----------
    CloseSession
    RollbackTransaction
    blnTransaction = False
    SetIterationsForTests 0

    CheckEq "La base est restee intacte (aucun utilisateur)", _
        CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")), 0, lngOk, lngTotal, strReport
    CheckEq "La base est restee intacte (journal vide)", _
        CLng(QueryValue("SELECT COUNT(*) FROM [JournalConnexion]")), 0, lngOk, lngTotal, strReport
    CheckTrue "La base est restee intacte (aucun code de secours)", Not RecoveryCodeExists(), lngOk, lngTotal, strReport

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport, _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), "Kamerix - tests des comptes"
    Exit Sub

EH:
    lngErr = Err.Number
    strErrText = Err.Description
    On Error Resume Next
    CloseSession
    If blnTransaction Then RollbackTransaction
    SetIterationsForTests 0
    On Error GoTo 0
    MsgBox "Echec (erreur " & lngErr & ") :" & vbCrLf & vbCrLf & strErrText & vbCrLf & vbCrLf & _
           "Resultats deja obtenus :" & vbCrLf & strReport, vbCritical, "Kamerix - tests des comptes"
End Sub

' Numero d'erreur levee par une operation (0 si aucune)
Private Function ErrorOf(ByVal pOperation As String, Optional ByVal pUserId As Long) As Long
    Dim AdoRs As ADODB.Recordset

    On Error Resume Next
    Select Case pOperation
        Case "ListAccounts":        Set AdoRs = ListAccounts()
        Case "RenewRecoveryCode":   RenewRecoveryCode
        Case "SelfDeactivate":      SetAccountActive CurrentUserId(), False
        Case "SelfRole":            ChangeAccountRole CurrentUserId(), GetRoleIdByName("Vendeur")
        Case "LastAdminDeactivate": SetUserActive pUserId, False
        Case "LastAdminRole":       ChangeUserRole pUserId, GetRoleIdByName("Vendeur")
    End Select
    ErrorOf = Err.Number
    Err.Clear
    If Not AdoRs Is Nothing Then AdoRs.Close
End Function

' Numero d'erreur d'une tentative de secours (0 si acceptee)
Private Function RecoveryError(ByVal pCode As String, ByVal pLogin As String, ByVal pNewPassword As String) As Long
    On Error Resume Next
    ResetAdminWithRecoveryCode pCode, pLogin, pNewPassword
    RecoveryError = Err.Number
    Err.Clear
End Function

' Nombre de lignes d'un recordset (qui est ensuite ferme)
Private Function RecordCount(ByVal pRs As ADODB.Recordset) As Long
    Do While Not pRs.EOF
        RecordCount = RecordCount + 1
        pRs.MoveNext
    Loop
    pRs.Close
End Function

' Comparaison de deux valeurs (converties en texte) ; affiche obtenu et attendu en cas d'echec
Private Sub CheckEq(ByVal strLabel As String, ByVal vActual As Variant, ByVal vExpected As Variant, _
                    ByRef lngOk As Long, ByRef lngTotal As Long, ByRef strReport As String)
    lngTotal = lngTotal + 1
    If IsNull(vActual) Then vActual = "(Null)"
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
