Attribute VB_Name = "modKamerixSession"
Option Explicit
Option Private Module
'==============================================================================
' KAMERIX - Session : utilisateur connecte et controle des droits
'------------------------------------------------------------------------------
' Role        : garder l'utilisateur connecte et son role, dire ce qu'il a le
'               droit de faire (CanRead, CanWrite, CanDelete par module,
'               HasRight pour un droit special) et controler l'acces aux
'               operations d'administration des comptes.
' Principe    : TOUT EST REFUSE par defaut. Sans utilisateur connecte, ou si le
'               module ou le droit est inconnu, ou tant que l'utilisateur doit
'               changer son mot de passe, les fonctions retournent Faux.
'               L'ecran de premiere utilisation (creation du premier
'               administrateur) n'a pas besoin de la Session : il appelle
'               AdminExists et CreateFirstAdmin (modKamerixSecurity).
' Cache       : les droits du role sont lus dans la base a l'ouverture de la
'               session. Apres une modification des droits ou du role de
'               l'utilisateur connecte, appeler RefreshSession.
' Usage       : If CanWrite(MOD_SALES) Then ...      (pour griser un bouton)
'               RequireWrite MOD_SALES                (dans la couche metier)
'               RequireRight RIGHT_REFUND
' Modules     : les noms sont compares sans tenir compte de la casse ni des
'               accents (le module "Reservations" peut s'ecrire avec ou sans accent).
' Erreurs     : ERR_SESSION_NOT_LOGGED_IN et ERR_SESSION_DENIED (1410, 1411).
' Dependances : modKamerixSecurity, modKamerixSql, modKamerixDb (ERR_BASE).
'==============================================================================

Public Const ERR_SESSION_NOT_LOGGED_IN As Long = ERR_BASE + 410
Public Const ERR_SESSION_DENIED As Long = ERR_BASE + 411

' Noms des modules (table Module)
Public Const MOD_CARS As String = "Voitures"
Public Const MOD_PARTIES As String = "Tiers"
Public Const MOD_RESERVATIONS As String = "Reservations"
Public Const MOD_SALES As String = "Ventes"
Public Const MOD_PAYMENTS As String = "Paiements"
Public Const MOD_INVOICES As String = "Factures"
Public Const MOD_EXPENSES As String = "Charges"
Public Const MOD_REPORTS As String = "Rapports"
Public Const MOD_ADMINISTRATION As String = "Administration"

' Codes des droits speciaux (table DroitSpecial)
Public Const RIGHT_CANCEL_SALE As String = "VENTE_ANNULER"
Public Const RIGHT_CANCEL_RESERVATION As String = "RESERVATION_ANNULER"
Public Const RIGHT_CANCEL_INVOICE As String = "FACTURE_ANNULER"
Public Const RIGHT_CANCEL_PAYMENT As String = "PAIEMENT_ANNULER"
Public Const RIGHT_REFUND As String = "REMBOURSER"
Public Const RIGHT_APPROVE_RESTOCK As String = "REMISE_STOCK_VALIDER"

Private Const MASK_READ As Long = 1
Private Const MASK_WRITE As Long = 2
Private Const MASK_DELETE As Long = 4

Private mlngUserId As Long
Private mstrLogin As String
Private mstrFullName As String
Private mlngRoleId As Long
Private mstrRoleName As String
Private mblnMustChange As Boolean
Private mcolModules As Collection     ' cle = nom du module, valeur = masque de droits
Private mcolRights As Collection      ' cle = code du droit special

'==============================================================================
' OUVERTURE ET FERMETURE DES SESSIONS UTILISATEUR
'==============================================================================
' Ouvre la session si l'identifiant et le mot de passe sont corrects.
' Ferme d'abord toute session precedente : apres un echec, personne n'est connecte.
' Retourne le resultat de AuthenticateUser (kxLoginSuccess, kxLoginDenied ou kxLoginLocked).
' Apres un succes, tester MustChangePassword : si Vrai, afficher l'ecran de
' changement (ChangeMyPassword) avant tout autre usage.
Public Function OpenSession(ByVal pLogin As String, ByVal pPassword As String) As KxLoginResult
    Dim lngId As Long
    Dim blnMustChange As Boolean
    Dim eRes As KxLoginResult

    CloseSession

    eRes = AuthenticateUser(pLogin, pPassword, lngId, blnMustChange)
    If eRes = kxLoginSuccess Then LoadSession lngId, blnMustChange

    OpenSession = eRes
End Function

' Ferme la session : plus aucun droit
Public Sub CloseSession()
    mlngUserId = 0
    mstrLogin = ""
    mstrFullName = ""
    mlngRoleId = 0
    mstrRoleName = ""
    mblnMustChange = False
    Set mcolModules = Nothing
    Set mcolRights = Nothing
End Sub

' Relit le compte de l'utilisateur connecte dans la base : role, droits et
' obligation de changer le mot de passe. Si le compte a ete desactive (ou
' supprime) entre-temps, la session est fermee.
Public Sub RefreshSession()
    Dim lngId As Long
    Dim vActive As Variant
    Dim vMustChange As Variant
    Dim AdoRs As ADODB.Recordset

    If Not IsLoggedIn() Then Exit Sub

    lngId = mlngUserId
    Set AdoRs = QueryRecordset("SELECT [Actif], [DoitChangerMdp] FROM [Utilisateur] WHERE [UtilisateurID] = ?", lngId)
    If Not AdoRs.EOF Then
        vActive = AdoRs.Fields("Actif").Value
        vMustChange = AdoRs.Fields("DoitChangerMdp").Value
    End If
    AdoRs.Close

    If IsEmpty(vActive) Or IsNull(vActive) Then
        CloseSession
    ElseIf Not CBool(vActive) Then
        CloseSession
    Else
        LoadSession lngId, CBool(Nz0(vMustChange))
    End If
End Sub

'==============================================================================
' ETAT DE LA SESSION
'==============================================================================

Public Function IsLoggedIn() As Boolean
    IsLoggedIn = (mlngUserId <> 0)
End Function

Public Function CurrentUserId() As Long
    CurrentUserId = mlngUserId
End Function

Public Function CurrentUserLogin() As String
    CurrentUserLogin = mstrLogin
End Function

Public Function CurrentUserName() As String
    CurrentUserName = mstrFullName
End Function

Public Function CurrentRoleId() As Long
    CurrentRoleId = mlngRoleId
End Function

Public Function CurrentRoleName() As String
    CurrentRoleName = mstrRoleName
End Function

' Vrai si l'utilisateur doit changer son mot de passe avant de pouvoir travailler
Public Function MustChangePassword() As Boolean
    MustChangePassword = IsLoggedIn() And mblnMustChange
End Function

'==============================================================================
' DROITS (retournent Faux au moindre doute)
'==============================================================================

Public Function CanRead(ByVal pModule As String) As Boolean
    CanRead = HasAccess(pModule, MASK_READ)
End Function

Public Function CanWrite(ByVal pModule As String) As Boolean
    CanWrite = HasAccess(pModule, MASK_WRITE)
End Function

Public Function CanDelete(ByVal pModule As String) As Boolean
    CanDelete = HasAccess(pModule, MASK_DELETE)
End Function

' Droit special (annuler une vente, rembourser, ...)
Public Function HasRight(ByVal pCode As String) As Boolean
    If Not IsLoggedIn() Then Exit Function
    If mblnMustChange Then Exit Function
    If mcolRights Is Nothing Then Exit Function

    On Error Resume Next
    HasRight = CBool(mcolRights.Item(NormalizeKey(pCode)))
    On Error GoTo 0
End Function

'==============================================================================
' CONTROLES QUI LEVENT UNE ERREUR (pour la couche metier)
'==============================================================================

Public Sub RequireRead(ByVal pModule As String)
    RequireAccess pModule, MASK_READ, "lire"
End Sub

Public Sub RequireWrite(ByVal pModule As String)
    RequireAccess pModule, MASK_WRITE, "modifier"
End Sub

Public Sub RequireDelete(ByVal pModule As String)
    RequireAccess pModule, MASK_DELETE, "supprimer"
End Sub

Public Sub RequireRight(ByVal pCode As String)
    RequireActiveSession "modKamerixSession.RequireRight"
    If Not HasRight(pCode) Then
        Err.Raise ERR_SESSION_DENIED, "modKamerixSession.RequireRight", _
            "Acces refuse : le role " & mstrRoleName & " n'a pas le droit " & pCode & "."
    End If
End Sub

'==============================================================================
' MOTS DE PASSE ET COMPTES (avec controle des droits)
'==============================================================================

' Changement de son propre mot de passe (ancien mot de passe exige).
' Leve l'obligation de changer si elle existait.
Public Sub ChangeMyPassword(ByVal pOldPassword As String, ByVal pNewPassword As String)
    If Not IsLoggedIn() Then
        Err.Raise ERR_SESSION_NOT_LOGGED_IN, "modKamerixSession.ChangeMyPassword", _
            "Aucun utilisateur connecte."
    End If

    ChangePassword mlngUserId, pOldPassword, pNewPassword
    mblnMustChange = False
End Sub

' Creation d'un compte avec mot de passe temporaire. Reserve a qui peut ecrire dans le module Administration.
Public Function CreateAccount(ByVal pLogin As String, ByVal pFullName As String, _
                            ByVal pRoleId As Long, ByVal pTempPassword As String) As Long
    RequireWrite MOD_ADMINISTRATION
    CreateAccount = CreateUser(pLogin, pFullName, pRoleId, pTempPassword)
End Function

' Reinitialisation du mot de passe d'un compte (mot de passe temporaire). Reserve a l'administration.
Public Sub ResetAccount(ByVal pUserId As Long, ByVal pTempPassword As String)
    RequireWrite MOD_ADMINISTRATION
    ResetPassword pUserId, pTempPassword
End Sub

' Deblocage d'un compte. Reserve a l'administration.
Public Sub UnlockAccount(ByVal pUserId As Long)
    RequireWrite MOD_ADMINISTRATION
    UnlockUser pUserId
End Sub

' Liste des comptes (voir UserList). Reserve a qui peut lire le module Administration.
Public Function ListAccounts() As ADODB.Recordset
    RequireRead MOD_ADMINISTRATION
    Set ListAccounts = UserList()
End Function

' Creation d'un compte avec un mot de passe provisoire genere.
' pTempPassword recoit ce mot de passe, a communiquer a l'utilisateur (il devra le changer).
Public Function CreateAccountWithTemporaryPassword(ByVal pLogin As String, ByVal pFullName As String, _
                                                   ByVal pRoleId As Long, ByRef pTempPassword As String) As Long
    Dim strTemp As String

    RequireWrite MOD_ADMINISTRATION
    strTemp = GenerateTemporaryPassword()
    CreateAccountWithTemporaryPassword = CreateUser(pLogin, pFullName, pRoleId, strTemp)
    pTempPassword = strTemp
End Function

' Reinitialisation avec un mot de passe provisoire genere, retourne pour etre communique.
' Le compte est aussi debloque.
Public Function ResetAccountWithTemporaryPassword(ByVal pUserId As Long) As String
    Dim strTemp As String

    RequireWrite MOD_ADMINISTRATION
    strTemp = GenerateTemporaryPassword()
    ResetPassword pUserId, strTemp
    ResetAccountWithTemporaryPassword = strTemp
End Function

' Activation ou desactivation d'un compte. On ne peut pas desactiver son propre compte.
Public Sub SetAccountActive(ByVal pUserId As Long, ByVal pActive As Boolean)
    RequireWrite MOD_ADMINISTRATION
    If Not pActive And pUserId = mlngUserId Then
        Err.Raise ERR_SESSION_DENIED, "modKamerixSession.SetAccountActive", _
            "Vous ne pouvez pas desactiver votre propre compte."
    End If
    SetUserActive pUserId, pActive
End Sub

' Changement de role d'un compte. On ne peut pas changer son propre role.
Public Sub ChangeAccountRole(ByVal pUserId As Long, ByVal pRoleId As Long)
    RequireWrite MOD_ADMINISTRATION
    If pUserId = mlngUserId Then
        Err.Raise ERR_SESSION_DENIED, "modKamerixSession.ChangeAccountRole", _
            "Vous ne pouvez pas changer votre propre role."
    End If
    ChangeUserRole pUserId, pRoleId
End Sub

' Nouveau code de secours (l'ancien ne marche plus), retourne en clair pour etre note.
Public Function RenewRecoveryCode() As String
    RequireWrite MOD_ADMINISTRATION
    RenewRecoveryCode = CreateRecoveryCode()
End Function

'==============================================================================
' PROCEDURES INTERNES
'==============================================================================

' Charge l'utilisateur, son role et les droits de ce role. En cas d'erreur, la session est fermee.
Private Sub LoadSession(ByVal pUserId As Long, ByVal pMustChange As Boolean)
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    On Error GoTo EH

    LoadUser pUserId
    mblnMustChange = pMustChange
    LoadRights
    Exit Sub

EH:
    lngErr = Err.Number
    strErr = Err.Description
    strSrc = Err.Source

    CloseSession
    Err.Raise lngErr, strSrc, strErr
End Sub

Private Sub LoadUser(ByVal pUserId As Long)
    Dim AdoRs As ADODB.Recordset

    Set AdoRs = QueryRecordset( _
        "SELECT u.[Login], u.[NomComplet], u.[RoleID], r.[Nom] AS NomRole " & _
        "FROM [Utilisateur] AS u INNER JOIN [Role] AS r ON u.[RoleID] = r.[RoleID] " & _
        "WHERE u.[UtilisateurID] = ?", pUserId)

    If AdoRs.EOF Then
        AdoRs.Close
        Err.Raise ERR_SEC_NOT_FOUND, "modKamerixSession.LoadUser", _
            "Utilisateur introuvable (identifiant " & pUserId & ")."
    End If

    mlngUserId = pUserId
    mstrLogin = CStr(AdoRs.Fields("Login").Value)
    mstrFullName = CStr(AdoRs.Fields("NomComplet").Value)
    mlngRoleId = CLng(AdoRs.Fields("RoleID").Value)
    mstrRoleName = CStr(AdoRs.Fields("NomRole").Value)
    AdoRs.Close
End Sub

' Lit les droits du role : un masque par module et la liste des droits speciaux
Private Sub LoadRights()
    Dim AdoRs As ADODB.Recordset
    Dim lngMask As Long

    Set mcolModules = New Collection
    Set mcolRights = New Collection

    Set AdoRs = QueryRecordset( _
        "SELECT m.[Nom], p.[Lecture], p.[Ecriture], p.[Suppression] " & _
        "FROM [Permission] AS p INNER JOIN [Module] AS m ON p.[ModuleID] = m.[ModuleID] " & _
        "WHERE p.[RoleID] = ?", mlngRoleId)
    Do While Not AdoRs.EOF
        lngMask = 0
        If CBool(AdoRs.Fields("Lecture").Value) Then lngMask = lngMask + MASK_READ
        If CBool(AdoRs.Fields("Ecriture").Value) Then lngMask = lngMask + MASK_WRITE
        If CBool(AdoRs.Fields("Suppression").Value) Then lngMask = lngMask + MASK_DELETE
        mcolModules.Add lngMask, NormalizeKey(CStr(AdoRs.Fields("Nom").Value))
        AdoRs.MoveNext
    Loop
    AdoRs.Close

    Set AdoRs = QueryRecordset("SELECT [Code] FROM [DroitSpecial] WHERE [RoleID] = ?", mlngRoleId)
    Do While Not AdoRs.EOF
        mcolRights.Add True, NormalizeKey(CStr(AdoRs.Fields("Code").Value))
        AdoRs.MoveNext
    Loop
    AdoRs.Close
End Sub

Private Function HasAccess(ByVal pModule As String, ByVal pMask As Long) As Boolean
    If Not IsLoggedIn() Then Exit Function
    If mblnMustChange Then Exit Function

    HasAccess = ((ModuleMask(pModule) And pMask) <> 0)
End Function

' Masque de droits du role sur le module (0 si le module est inconnu)
Private Function ModuleMask(ByVal pModule As String) As Long
    If mcolModules Is Nothing Then Exit Function

    On Error Resume Next
    ModuleMask = CLng(mcolModules.Item(NormalizeKey(pModule)))
    On Error GoTo 0
End Function

Private Sub RequireAccess(ByVal pModule As String, ByVal pMask As Long, ByVal pVerb As String)
    RequireActiveSession "modKamerixSession.RequireAccess"

    If (ModuleMask(pModule) And pMask) = 0 Then
        Err.Raise ERR_SESSION_DENIED, "modKamerixSession.RequireAccess", _
            "Accès refusé : le rôle " & mstrRoleName & " ne peut pas " & pVerb & " le module " & pModule & "."
    End If
End Sub

' Leve une erreur si personne n'est connecte ou si le changement de mot de passe est en attente
Private Sub RequireActiveSession(ByVal pSource As String)
    If Not IsLoggedIn() Then
        Err.Raise ERR_SESSION_NOT_LOGGED_IN, pSource, "Aucun utilisateur connecté."
    End If
    If mblnMustChange Then
        Err.Raise ERR_SESSION_DENIED, pSource, "Accès refusé : le mot de passe doit d'abord être changé."
    End If
End Sub

' Null ou vide -> Faux (colonne Oui/Non)
Private Function Nz0(ByVal pValue As Variant) As Variant
    If IsNull(pValue) Or IsEmpty(pValue) Then Nz0 = False Else Nz0 = pValue
End Function

' NormalizeKey : cle de recherche : minuscules, sans accents courants
Private Function NormalizeKey(ByVal pText As String) As String
    Dim strKey As String

    strKey = LCase$(Trim$(pText))
    strKey = Replace(strKey, ChrW(233), "e")
    strKey = Replace(strKey, ChrW(232), "e")
    strKey = Replace(strKey, ChrW(234), "e")
    NormalizeKey = strKey
End Function
