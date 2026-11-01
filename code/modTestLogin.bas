Attribute VB_Name = "modTestLogin"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Tests de la connexion : messages et remise a niveau du hash
'------------------------------------------------------------------------------
' A lancer (F5 sur TestLogin) sur une base de developpement SANS utilisateur.
' Tout se passe dans une transaction annulee a la fin : la base reste intacte.
' Les ecrans eux-memes (frmLogin, frmFirstAdmin, frmChangePassword) se testent
' a la main : voir la liste de controles dans LISEZMOI_CONNEXION.md.
'==============================================================================

Private Const PWD_ADMIN As String = "Secret#2026"

Public Sub TestLogin()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim lngAdmin As Long
    Dim strHashBefore As String
    Dim strSaltBefore As String
    Dim lngErr As Long
    Dim strErrText As String
    Dim blnTransaction As Boolean

    On Error GoTo EH

    If CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")) > 0 Then
        MsgBox "Ce test demande une base sans utilisateur." & vbCrLf & _
               "Supprimez la base de developpement et recreez-la, puis relancez le test.", _
               vbExclamation, "Kamerix - tests de la connexion"
        Exit Sub
    End If

    ' ---------- Messages de connexion ----------
    CheckEq "Message apres un succes : vide", LoginResultMessage(kxLoginSuccess), "", lngOk, lngTotal, strReport
    CheckEq "Message apres un refus : generique", LoginResultMessage(kxLoginDenied), _
        "Identifiant ou mot de passe incorrect.", lngOk, lngTotal, strReport
    CheckTrue "Message apres un blocage : indique la duree", _
        InStr(LoginResultMessage(kxLoginLocked), CStr(LockoutMinutes()) & " minutes") > 0, lngOk, lngTotal, strReport
    CheckTrue "Regles du mot de passe : texte disponible pour les ecrans", _
        InStr(PasswordRulesText(), "8") > 0, lngOk, lngTotal, strReport

    ' ---------- Remise a niveau du hash ----------
    CloseSession
    SetIterationsForTests 1000
    BeginTransaction
    blnTransaction = True

    lngAdmin = CreateFirstAdmin("admin", "Administrateur Test", PWD_ADMIN)
    strHashBefore = StoredHash(lngAdmin)
    CheckTrue "Compte cree avec 1000 iterations", Left$(strHashBefore, 5) = "1000$", lngOk, lngTotal, strReport

    ' Meme reglage : le hash ne change pas
    CheckTrue "Connexion avec le reglage inchange", OpenSession("admin", PWD_ADMIN) = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckEq "Reglage inchange : hash conserve", StoredHash(lngAdmin), strHashBefore, lngOk, lngTotal, strReport

    ' Reglage augmente : remise a niveau apres la connexion
    SetIterationsForTests 2000
    strSaltBefore = StoredSalt(lngAdmin)
    CheckTrue "Connexion apres augmentation du reglage", OpenSession("admin", PWD_ADMIN) = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckTrue "Hash remis a niveau : 2000 iterations", Left$(StoredHash(lngAdmin), 5) = "2000$", lngOk, lngTotal, strReport
    CheckTrue "Nouveau sel apres la remise a niveau", StoredSalt(lngAdmin) <> strSaltBefore, lngOk, lngTotal, strReport
    CheckTrue "Le mot de passe marche encore apres la remise a niveau", _
        OpenSession("admin", PWD_ADMIN) = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckTrue "Pas de changement de mot de passe exige apres la remise a niveau", _
        Not MustChangePassword(), lngOk, lngTotal, strReport

    ' Reglage diminue : le hash suit aussi
    SetIterationsForTests 1000
    CheckTrue "Connexion apres diminution du reglage", OpenSession("admin", PWD_ADMIN) = kxLoginSuccess, lngOk, lngTotal, strReport
    CheckTrue "Hash aligne sur le reglage : 1000 iterations", Left$(StoredHash(lngAdmin), 5) = "1000$", lngOk, lngTotal, strReport

    ' Echec de connexion : aucune remise a niveau
    SetIterationsForTests 3000
    strHashBefore = StoredHash(lngAdmin)
    CheckTrue "Mauvais mot de passe refuse", OpenSession("admin", "mauvais-mdp") = kxLoginDenied, lngOk, lngTotal, strReport
    CheckEq "Echec : hash inchange", StoredHash(lngAdmin), strHashBefore, lngOk, lngTotal, strReport

    ' ---------- Fin : on annule tout ----------
    CloseSession
    RollbackTransaction
    blnTransaction = False
    SetIterationsForTests 0

    CheckEq "La base est restee intacte (aucun utilisateur)", _
        CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")), 0, lngOk, lngTotal, strReport
    CheckEq "La base est restee intacte (journal vide)", _
        CLng(QueryValue("SELECT COUNT(*) FROM [JournalConnexion]")), 0, lngOk, lngTotal, strReport

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport, _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), "Kamerix - tests de la connexion"
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
           "Resultats deja obtenus :" & vbCrLf & strReport, vbCritical, "Kamerix - tests de la connexion"
End Sub

Private Function StoredHash(ByVal pUserId As Long) As String
    StoredHash = CStr(QueryValue("SELECT [MotDePasseHash] FROM [Utilisateur] WHERE [UtilisateurID] = ?", pUserId))
End Function

Private Function StoredSalt(ByVal pUserId As Long) As String
    StoredSalt = CStr(QueryValue("SELECT [Sel] FROM [Utilisateur] WHERE [UtilisateurID] = ?", pUserId))
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
