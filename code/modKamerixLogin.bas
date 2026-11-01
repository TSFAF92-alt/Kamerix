Attribute VB_Name = "modKamerixLogin"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Demarrage et ecrans de connexion
'------------------------------------------------------------------------------
' Role        : enchainer les ecrans au lancement :
'                 1. base absente : creee (EnsureDatabase) ;
'                 2. aucun administrateur : frmFirstAdmin (creation puis
'                    ouverture de la session) ;
'                 3. sinon : frmLogin ;
'                 4. mot de passe provisoire : frmChangePassword obligatoire,
'                    la session est fermee si l'utilisateur renonce ;
'                 5. entreprise pas encore configuree : assistant frmSettings pour
'                    un administrateur (CompanySetupIfNeeded, lot 5.6).
' Annuler     : un ecran ferme par Annuler ou par la croix laisse l'application
'               sans session : tout ce qui est protege reste refuse.
' Formulaires : appeles par leur nom (VBA.UserForms.Add) : ce module se compile
'               meme si les formulaires n'existent pas encore. Ils sont crees
'               par ZmodFormBuilder (BuildKamerixForms). Chacun expose une
'               propriete Succeeded, lue apres sa fermeture.
' Usage       : StartKamerix                 (Workbook_Open, bouton du ruban)
'               ChangeMyPasswordInteractive  (changement volontaire)
'               ManageAccounts               (ecran de gestion des comptes)
'               Logout                       (deconnexion)
' Secours     : a la creation du premier administrateur, un code de secours est
'               affiche une seule fois. Il sert depuis frmLogin ("Mot de passe
'               oublie ?") a redefinir le mot de passe d'un administrateur.
' Dependances : modKamerixDb, modKamerixSecurity, modKamerixSession.
'==============================================================================

Public Const FORM_LOGIN As String = "frmLogin"
Public Const FORM_FIRST_ADMIN As String = "frmFirstAdmin"
Public Const FORM_CHANGE_PASSWORD As String = "frmChangePassword"
Public Const FORM_RECOVERY As String = "frmRecovery"
Public Const FORM_USERS As String = "frmUsers"

Private Const APP_TITLE As String = "Kamerix"

'==============================================================================
' API PUBLIQUE
'==============================================================================

' Lance la sequence de connexion (F5, Alt+F8, Workbook_Open ou bouton du ruban).
' Si une session est deja ouverte, elle est conservee. Pour savoir ensuite si
' quelqu'un est connecte : IsLoggedIn().
Public Sub StartKamerix()
    On Error GoTo EH

    If IsLoggedIn() Then Exit Sub

    EnsureDatabase
    EnsureDocumentSettings          ' en-tete des documents : valeurs d'exemple si absentes (lot 5.5)

    ' Premier lancement : creation du premier administrateur (il est connecte ensuite)
    If Not AdminExists() Then
        ShowForm FORM_FIRST_ADMIN
        If Not AdminExists() Then Exit Sub
    End If

    ' Code de secours : cree et montre une seule fois a l'administrateur connecte
    If IsLoggedIn() And Not RecoveryCodeExists() Then
        ShowRecoveryCode RenewRecoveryCode(), _
            "Code de secours de Kamerix : il permet de red" & ChrW(233) & "finir le mot de passe d'un " & _
            "administrateur s'il est perdu (" & ChrW(171) & " Mot de passe oubli" & ChrW(233) & " ? " & ChrW(187) & _
            " sur l'" & ChrW(233) & "cran de connexion)."
    End If

    ' Connexion (sauf si l'ecran precedent a deja ouvert la session)
    If Not IsLoggedIn() Then
        If Not ShowForm(FORM_LOGIN) Then Exit Sub
    End If

    ' Mot de passe provisoire : changement obligatoire avant tout usage
    If MustChangePassword() Then
        If Not ShowChangePasswordForm(True) Then
            CloseSession
            MsgBox "Le mot de passe doit etre change pour utiliser Kamerix." & vbCrLf & _
                   "Vous n'etes pas connecte.", vbExclamation, APP_TITLE
            Exit Sub
        End If
    End If

    ' Confirmation de la connexion (sinon StartKamerix se termine sans aucun retour)
    MsgBox "Bienvenue, " & CurrentUserName() & "." & vbCrLf & _
           "Connect" & ChrW(233) & " en tant que " & CurrentUserLogin() & " (" & CurrentRoleName() & ").", _
           vbInformation, APP_TITLE

    ' Premiere configuration de l'entreprise (en-tete des documents)
    CompanySetupIfNeeded
    Exit Sub

EH:
    CloseSession
    MsgBox "Kamerix ne peut pas demarrer (erreur " & Err.Number & ") :" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, APP_TITLE
End Sub

' Changement volontaire du mot de passe par l'utilisateur connecte
Public Sub ChangeMyPasswordInteractive()
    If Not IsLoggedIn() Then
        MsgBox "Connectez-vous d'abord.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    If ShowChangePasswordForm(False) Then
        MsgBox "Votre mot de passe a ete change.", vbInformation, APP_TITLE
    End If
End Sub

' Ecran de gestion des comptes (consultation : lecture du module Administration ;
' actions : ecriture du module Administration)
Public Sub ManageAccounts()
    On Error GoTo EH

    If Not IsLoggedIn() Then
        MsgBox "Connectez-vous d'abord.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    If Not CanRead(MOD_ADMINISTRATION) Then
        MsgBox "Votre r" & ChrW(244) & "le ne donne pas acc" & ChrW(232) & "s " & ChrW(224) & " la gestion des comptes.", _
               vbExclamation, APP_TITLE
        Exit Sub
    End If

    ShowForm FORM_USERS
    Exit Sub

EH:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, APP_TITLE
End Sub

' Ecran "Mot de passe oublie ?" (appele par frmLogin)
Public Sub ShowRecoveryForm()
    On Error GoTo EH

    If Not RecoveryCodeExists() Then
        MsgBox "Aucun code de secours n'est enregistr" & ChrW(233) & "." & vbCrLf & _
               "Demandez " & ChrW(224) & " un administrateur de r" & ChrW(233) & "initialiser votre mot de passe.", _
               vbInformation, APP_TITLE
        Exit Sub
    End If

    ShowForm FORM_RECOVERY
    Exit Sub

EH:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, APP_TITLE
End Sub

' Affiche un mot de passe provisoire a transmettre (une seule fois)
Public Sub ShowTemporaryPassword(ByVal pLogin As String, ByVal pTempPassword As String)
    MsgBox "Mot de passe provisoire de " & pLogin & " :" & vbCrLf & vbCrLf & _
           "        " & pTempPassword & vbCrLf & vbCrLf & _
           "Transmettez-le " & ChrW(224) & " l'utilisateur : il devra le changer " & ChrW(224) & " sa prochaine connexion." & vbCrLf & _
           "Ce mot de passe ne sera plus affich" & ChrW(233) & ". (Ctrl+C copie ce message.)", _
           vbInformation, APP_TITLE
End Sub

' Affiche un code de secours a noter (une seule fois)
Public Sub ShowRecoveryCode(ByVal pCode As String, ByVal pIntro As String)
    MsgBox pIntro & vbCrLf & vbCrLf & _
           "        " & pCode & vbCrLf & vbCrLf & _
           "Notez-le et gardez-le en lieu s" & ChrW(251) & "r, hors de l'ordinateur : il ne sera plus affich" & ChrW(233) & "." & vbCrLf & _
           "(Ctrl+C copie ce message.)", vbExclamation, APP_TITLE
End Sub

' Deconnexion : plus aucun droit jusqu'a la prochaine connexion
Public Sub Logout()
    CloseSession
End Sub

' Message a afficher pour un resultat de connexion (chaine vide en cas de succes).
' Un identifiant inconnu et un mauvais mot de passe recoivent le meme message.
Public Function LoginResultMessage(ByVal eResult As KxLoginResult) As String
    Select Case eResult
        Case kxLoginSuccess
            LoginResultMessage = ""
        Case kxLoginLocked
            LoginResultMessage = "Compte temporairement bloque apres plusieurs echecs. " & _
                                 "Reessayez dans " & LockoutMinutes() & " minutes au plus."
        Case Else
            LoginResultMessage = "Identifiant ou mot de passe incorrect."
    End Select
End Function

'==============================================================================
' PROCEDURES INTERNES
'==============================================================================

' Affiche un formulaire modal et retourne sa propriete Succeeded
Private Function ShowForm(ByVal strFormName As String) As Boolean
    Dim frm As Object

    Set frm = LoadForm(strFormName)
    frm.Show
    ShowForm = frm.Succeeded
    Unload frm
End Function

' Ecran de changement de mot de passe : pForced = Vrai pour un mot de passe provisoire
Private Function ShowChangePasswordForm(ByVal pForced As Boolean) As Boolean
    Dim frm As Object

    Set frm = LoadForm(FORM_CHANGE_PASSWORD)
    frm.Forced = pForced
    frm.Show
    ShowChangePasswordForm = frm.Succeeded
    Unload frm
End Function

' Charge un formulaire par son nom ; message clair s'il n'a pas encore ete cree
Private Function LoadForm(ByVal strFormName As String) As Object
    On Error GoTo EH
    Set LoadForm = VBA.UserForms.Add(strFormName)
    Exit Function

EH:
    Err.Raise Err.Number, "modKamerixLogin.LoadForm", _
        "Le formulaire " & strFormName & " est introuvable. Lancez BuildKamerixForms " & _
        "(module ZmodFormBuilder) pour creer les ecrans de connexion." & vbCrLf & _
        "(" & Err.Description & ")"
End Function
