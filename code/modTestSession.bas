Attribute VB_Name = "modTestSession"
Option Explicit
Option Private Module
'==============================================================================
' KAMERIX - Tests de la Session (modKamerixSession)
'------------------------------------------------------------------------------
' A lancer (F5 sur TestSession) sur une base de developpement SANS utilisateur.
' Tout se passe dans une transaction annulee a la fin : la base reste intacte.
' Le hachage utilise 1000 iterations pendant le test ; la valeur standard est
' remise ensuite.
'==============================================================================

Private Const PWD_ADMIN As String = "Secret#2026"
Private Const PWD_TEMP As String = "Temporaire1"
Private Const PWD_NEW As String = "NouveauMdp#1"
Private Const PWD_PROVISIONAL As String = "Provisoire#9"

Public Sub TestSession()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim vModules As Variant
    Dim vRights As Variant
    Dim lngAdmin As Long
    Dim lngSeller As Long
    Dim lngAccountant As Long
    Dim lngManager As Long
    Dim lngCarsModuleId As Long
    Dim lngErr As Long
    Dim strErrText As String
    Dim i As Long
    Dim eRes As KxLoginResult
    Dim blnTransaction As Boolean

    On Error GoTo EH

    If CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")) > 0 Then
        MsgBox "Ce test demande une base sans utilisateur." & vbCrLf & _
               "Supprimez la base de developpement et recreez-la, puis relancez le test.", _
               vbExclamation, "Kamerix - tests de la Session"
        Exit Sub
    End If

    vModules = Array(MOD_CARS, MOD_PARTIES, MOD_RESERVATIONS, MOD_SALES, MOD_PAYMENTS, _
                     MOD_INVOICES, MOD_EXPENSES, MOD_REPORTS, MOD_ADMINISTRATION)
    vRights = Array(RIGHT_CANCEL_SALE, RIGHT_CANCEL_RESERVATION, RIGHT_CANCEL_INVOICE, _
                    RIGHT_CANCEL_PAYMENT, RIGHT_REFUND, RIGHT_APPROVE_RESTOCK)

    CloseSession
    SetIterationsForTests 1000
    BeginTransaction
    blnTransaction = True

    ' ---------- Sans connexion : tout est refuse ----------
    CheckTrue "Sans connexion : personne n'est connecte", Not IsLoggedIn(), lngOk, lngTotal, strReport
    CheckTrue "Sans connexion : lecture refusee", Not CanRead(MOD_CARS), lngOk, lngTotal, strReport
    CheckTrue "Sans connexion : ecriture refusee", Not CanWrite(MOD_CARS), lngOk, lngTotal, strReport
    CheckTrue "Sans connexion : suppression refusee", Not CanDelete(MOD_CARS), lngOk, lngTotal, strReport
    CheckTrue "Sans connexion : droit special refuse", Not HasRight(RIGHT_REFUND), lngOk, lngTotal, strReport

    On Error Resume Next
    RequireRead MOD_CARS
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Sans connexion : RequireRead leve ERR_SESSION_NOT_LOGGED_IN", lngErr, ERR_SESSION_NOT_LOGGED_IN, lngOk, lngTotal, strReport

    On Error Resume Next
    ChangeMyPassword PWD_ADMIN, PWD_NEW
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Sans connexion : ChangeMyPassword refuse", lngErr, ERR_SESSION_NOT_LOGGED_IN, lngOk, lngTotal, strReport

    ' ---------- Administrateur ----------
    lngAdmin = CreateFirstAdmin("admin", "Administrateur Test", PWD_ADMIN)
    eRes = OpenSession("admin", PWD_ADMIN)
    CheckTrue "Administrateur : session ouverte", eRes = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckTrue "Administrateur : IsLoggedIn", IsLoggedIn(), lngOk, lngTotal, strReport
    CheckEq "Administrateur : identifiant de la session", CurrentUserId(), lngAdmin, lngOk, lngTotal, strReport
    CheckEq "Administrateur : identifiant saisi", CurrentUserLogin(), "admin", lngOk, lngTotal, strReport
    CheckEq "Administrateur : nom du role", CurrentRoleName(), "Administrateur", lngOk, lngTotal, strReport
    CheckTrue "Administrateur : pas de changement de mot de passe exige", Not MustChangePassword(), lngOk, lngTotal, strReport
    CheckMatrix "Administrateur", vModules, _
        Array("LES", "LES", "LES", "LES", "LES", "LES", "LES", "LES", "LES"), lngOk, lngTotal, strReport
    CheckRights "Administrateur", vRights, vRights, lngOk, lngTotal, strReport
    CheckTrue "Le nom du module s'ecrit avec ou sans accent", _
        CanRead("R" & ChrW(233) & "servations") And CanRead("RESERVATIONS"), lngOk, lngTotal, strReport
    CheckTrue "Module inconnu : refuse", Not CanRead("Inexistant"), lngOk, lngTotal, strReport
    CheckTrue "Droit special inconnu : refuse", Not HasRight("DROIT_INCONNU"), lngOk, lngTotal, strReport

    ' ---------- Creation de comptes par l'administrateur ----------
    lngSeller = CreateAccount("vendeur1", "Vendeur Test", GetRoleIdByName("Vendeur"), PWD_TEMP)
    lngAccountant = CreateAccount("compta1", "Comptable Test", GetRoleIdByName("Comptable"), PWD_TEMP)
    lngManager = CreateAccount("resp1", "Responsable Test", GetRoleIdByName("Responsable"), PWD_TEMP)
    CheckTrue "L'administrateur cree trois comptes", lngSeller > 0 And lngAccountant > 0 And lngManager > 0, lngOk, lngTotal, strReport
    ChangePassword lngAccountant, PWD_TEMP, PWD_NEW        ' pour ouvrir ces sessions sans passer par l'ecran de changement
    ChangePassword lngManager, PWD_TEMP, PWD_NEW

    CloseSession
    CheckTrue "CloseSession : plus de connexion ni de droit", Not IsLoggedIn() And Not CanRead(MOD_CARS), lngOk, lngTotal, strReport
    CheckEq "CloseSession : role efface", CurrentRoleName(), "", lngOk, lngTotal, strReport

    ' ---------- Vendeur : mot de passe temporaire ----------
    eRes = OpenSession("vendeur1", PWD_TEMP)
    CheckTrue "Vendeur : session ouverte avec le mot de passe temporaire", eRes = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckTrue "Vendeur : changement de mot de passe exige", MustChangePassword(), lngOk, lngTotal, strReport
    CheckTrue "Vendeur : aucun droit tant que le mot de passe n'est pas change", _
        Not CanRead(MOD_CARS) And Not CanRead(MOD_PARTIES), lngOk, lngTotal, strReport

    On Error Resume Next
    RequireRead MOD_CARS
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Vendeur : RequireRead refuse tant que le mot de passe n'est pas change", lngErr, ERR_SESSION_DENIED, lngOk, lngTotal, strReport

    ChangeMyPassword PWD_TEMP, PWD_NEW
    CheckTrue "Vendeur : plus de changement exige apres le changement", Not MustChangePassword(), lngOk, lngTotal, strReport
    CheckMatrix "Vendeur", vModules, _
        Array("L--", "LE-", "LE-", "LE-", "L--", "L--", "---", "---", "---"), lngOk, lngTotal, strReport
    CheckRights "Vendeur", vRights, Array(), lngOk, lngTotal, strReport

    On Error Resume Next
    RequireWrite MOD_CARS
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Vendeur : RequireWrite refuse sur Voitures", lngErr, ERR_SESSION_DENIED, lngOk, lngTotal, strReport

    On Error Resume Next
    RequireRead MOD_CARS
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Vendeur : RequireRead accepte sur Voitures", lngErr, 0, lngOk, lngTotal, strReport

    On Error Resume Next
    RequireRight RIGHT_REFUND
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Vendeur : RequireRight refuse le remboursement", lngErr, ERR_SESSION_DENIED, lngOk, lngTotal, strReport

    On Error Resume Next
    CreateAccount "pirate1", "Pirate", GetRoleIdByName("Administrateur"), PWD_TEMP
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Vendeur : creation de compte refusee", lngErr, ERR_SESSION_DENIED, lngOk, lngTotal, strReport
    CheckEq "Vendeur : aucun compte pirate cree", _
        CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur] WHERE [Login] = ?", "pirate1")), 0, lngOk, lngTotal, strReport

    On Error Resume Next
    ResetAccount lngManager, PWD_PROVISIONAL
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Vendeur : reinitialisation de mot de passe refusee", lngErr, ERR_SESSION_DENIED, lngOk, lngTotal, strReport

    On Error Resume Next
    UnlockAccount lngManager
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Vendeur : deblocage de compte refuse", lngErr, ERR_SESSION_DENIED, lngOk, lngTotal, strReport

    ' Les droits sont lus a l'ouverture : une modification ne compte qu'apres RefreshSession
    lngCarsModuleId = CLng(QueryValue("SELECT [ModuleID] FROM [Module] WHERE [Nom] = ?", "Voitures"))
    ExecuteNonQuery "UPDATE [Permission] SET [Ecriture] = ? WHERE [RoleID] = ? AND [ModuleID] = ?", _
        True, CurrentRoleId(), lngCarsModuleId
    CheckTrue "Droits en cache : pas encore modifies", Not CanWrite(MOD_CARS), lngOk, lngTotal, strReport
    RefreshSession
    CheckTrue "RefreshSession : le nouveau droit est pris en compte", CanWrite(MOD_CARS), lngOk, lngTotal, strReport
    CheckTrue "RefreshSession : la session reste ouverte", IsLoggedIn() And CurrentRoleName() = "Vendeur", lngOk, lngTotal, strReport

    ' ---------- Comptable ----------
    eRes = OpenSession("compta1", PWD_NEW)
    CheckTrue "Comptable : session ouverte", eRes = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckEq "Comptable : nom du role", CurrentRoleName(), "Comptable", lngOk, lngTotal, strReport
    CheckMatrix "Comptable", vModules, _
        Array("L--", "L--", "L--", "L--", "LE-", "LE-", "LE-", "L--", "---"), lngOk, lngTotal, strReport
    CheckRights "Comptable", vRights, Array(RIGHT_CANCEL_PAYMENT, RIGHT_CANCEL_INVOICE, RIGHT_REFUND), _
        lngOk, lngTotal, strReport

    ' ---------- Responsable ----------
    eRes = OpenSession("resp1", PWD_NEW)
    CheckTrue "Responsable : session ouverte", eRes = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckEq "Responsable : nom du role", CurrentRoleName(), "Responsable", lngOk, lngTotal, strReport
    CheckMatrix "Responsable", vModules, _
        Array("LE-", "LE-", "LE-", "LE-", "LE-", "LE-", "LE-", "LE-", "---"), lngOk, lngTotal, strReport
    CheckRights "Responsable", vRights, vRights, lngOk, lngTotal, strReport

    On Error Resume Next
    CreateAccount "pirate2", "Pirate", GetRoleIdByName("Responsable"), PWD_TEMP
    lngErr = Err.Number
    Err.Clear
    On Error GoTo EH
    CheckEq "Responsable : creation de compte refusee (pas de droit sur Administration)", lngErr, ERR_SESSION_DENIED, lngOk, lngTotal, strReport

    ' ---------- Echecs de connexion et administration des comptes ----------
    eRes = OpenSession("resp1", "mauvais-mot-de-passe")
    CheckTrue "Mauvais mot de passe : connexion refusee", eRes = kxLoginDenied, lngOk, lngTotal, strReport
    CheckTrue "Apres un echec : plus aucune session ouverte", Not IsLoggedIn() And Not CanRead(MOD_CARS), lngOk, lngTotal, strReport

    For i = 1 To 5
        eRes = OpenSession("vendeur1", "mauvais-" & i)
    Next i
    CheckTrue "Cinq echecs : compte bloque", eRes = kxLoginLocked, lngOk, lngTotal, strReport
    CheckTrue "Compte bloque : aucune session", Not IsLoggedIn(), lngOk, lngTotal, strReport

    eRes = OpenSession("admin", PWD_ADMIN)
    CheckTrue "Administrateur : session rouverte", eRes = kxLoginSuccess, lngOk, lngTotal, strReport
    UnlockAccount lngSeller
    ResetAccount lngSeller, PWD_PROVISIONAL

    eRes = OpenSession("vendeur1", PWD_PROVISIONAL)
    CheckTrue "Vendeur : connexion avec le mot de passe provisoire apres deblocage", eRes = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckTrue "Vendeur : changement exige apres reinitialisation", MustChangePassword(), lngOk, lngTotal, strReport
    CheckEq "Vendeur : role relu depuis la base", CurrentRoleName(), "Vendeur", lngOk, lngTotal, strReport
    CheckEq "Vendeur : nom complet", CurrentUserName(), "Vendeur Test", lngOk, lngTotal, strReport

    ' ---------- Fin : on annule tout ----------
    CloseSession
    RollbackTransaction
    blnTransaction = False
    SetIterationsForTests 0

    CheckEq "La base est restee intacte (aucun utilisateur)", _
        CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")), 0, lngOk, lngTotal, strReport
    CheckEq "La base est restee intacte (journal vide)", _
        CLng(QueryValue("SELECT COUNT(*) FROM [JournalConnexion]")), 0, lngOk, lngTotal, strReport
    CheckTrue "La base est restee intacte (droits d'origine : le Vendeur ne peut pas ecrire dans Voitures)", _
        CLng(QueryValue("SELECT COUNT(*) FROM ([Permission] AS p INNER JOIN [Role] AS r ON p.[RoleID] = r.[RoleID]) " & _
        "INNER JOIN [Module] AS m ON p.[ModuleID] = m.[ModuleID] " & _
        "WHERE r.[Nom] = ? AND m.[Nom] = ? AND p.[Ecriture] = ?", "Vendeur", "Voitures", True)) = 0, _
        lngOk, lngTotal, strReport

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport, _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), "Kamerix - tests de la Session"
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
           "Resultats deja obtenus :" & vbCrLf & strReport, vbCritical, "Kamerix - tests de la Session"
End Sub

' Compare, pour chaque module, les droits de la session avec la matrice attendue
' (3 caracteres par module : L, E, S ou "-")
Private Sub CheckMatrix(ByVal pRole As String, ByVal vModules As Variant, ByVal vMatrix As Variant, _
                            ByRef lngOk As Long, ByRef lngTotal As Long, ByRef strReport As String)
    Dim j As Long
    Dim strFlags As String
    Dim strMismatches As String

    For j = LBound(vModules) To UBound(vModules)
        strFlags = CStr(vMatrix(j))
        If CanRead(CStr(vModules(j))) <> (Mid$(strFlags, 1, 1) = "L") Then strMismatches = strMismatches & vModules(j) & ":lecture "
        If CanWrite(CStr(vModules(j))) <> (Mid$(strFlags, 2, 1) = "E") Then strMismatches = strMismatches & vModules(j) & ":ecriture "
        If CanDelete(CStr(vModules(j))) <> (Mid$(strFlags, 3, 1) = "S") Then strMismatches = strMismatches & vModules(j) & ":suppression "
    Next j

    CheckEq "Matrice des droits conforme : " & pRole, strMismatches, "", lngOk, lngTotal, strReport
End Sub

' Compare les droits speciaux de la session avec la liste attendue
Private Sub CheckRights(ByVal pRole As String, ByVal vAllRights As Variant, ByVal vExpectedRights As Variant, _
                           ByRef lngOk As Long, ByRef lngTotal As Long, ByRef strReport As String)
    Dim j As Long
    Dim strExpectedList As String
    Dim strMismatches As String

    strExpectedList = "," & Join(vExpectedRights, ",") & ","
    For j = LBound(vAllRights) To UBound(vAllRights)
        If HasRight(CStr(vAllRights(j))) <> (InStr(strExpectedList, "," & CStr(vAllRights(j)) & ",") > 0) Then
            strMismatches = strMismatches & vAllRights(j) & " "
        End If
    Next j

    CheckEq "Droits speciaux conformes : " & pRole, strMismatches, "", lngOk, lngTotal, strReport
End Sub

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
