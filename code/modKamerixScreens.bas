Attribute VB_Name = "modKamerixScreens"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Ouverture des ecrans : procedure unique OpenScreen
'------------------------------------------------------------------------------
' Role        : point d'entree unique de tous les ecrans. OpenScreen verifie
'               qu'un utilisateur est connecte (sinon lance la connexion), que
'               son role peut lire le module concerne, puis ouvre le userform.
'               Le ruban (phase 8) appellera OpenScreen avec le nom de l'ecran.
' Usage       : OpenScreen SCREEN_REFERENCE      (fenetre Execution ou code)
'               OpenScreen "REFERENTIEL"
' Ecrans      : REFERENTIEL (lot 5.4a) ; TIERS, VOITURES (lot 5.4b) ; VENTES (lot 5.4c) ;
'               COMPTES, MOT_DE_PASSE, CONNEXION, DECONNEXION (lots 2 et 3) ;
'               PARAMETRES (lot 5.6 : entreprise, vente, securite ; ecriture Administration).
' Configuration : CompanySetupIfNeeded (appele par StartKamerix) ouvre l'assistant de
'               premiere configuration de l'entreprise pour un administrateur, tant
'               qu'elle n'est pas faite.
' Documents   : OfferDocument propose d'imprimer une piece qui vient d'etre creee
'               (recu apres un encaissement, facture au solde, recu de remboursement).
' Dependances : modKamerixSession, modKamerixLogin, modKamerixDocuments.
'==============================================================================

Public Const SCREEN_REFERENCE As String = "REFERENTIEL"
Public Const SCREEN_PARTIES As String = "TIERS"
Public Const SCREEN_CARS As String = "VOITURES"
Public Const SCREEN_SALES As String = "VENTES"
Public Const SCREEN_ACCOUNTS As String = "COMPTES"
Public Const SCREEN_PASSWORD As String = "MOT_DE_PASSE"
Public Const SCREEN_LOGIN As String = "CONNEXION"
Public Const SCREEN_LOGOUT As String = "DECONNEXION"
Public Const SCREEN_SETTINGS As String = "PARAMETRES"
Public Const FORM_SETTINGS As String = "frmSettings"

Private Const APP_TITLE As String = "Kamerix"

' Ouvre l'ecran demande (nom insensible a la casse)
Public Sub OpenScreen(ByVal pScreen As String)
    Dim strScreen As String

    On Error GoTo EH
    strScreen = UCase$(Trim$(pScreen))

    ' Ecrans accessibles sans session
    Select Case strScreen
        Case SCREEN_LOGIN
            StartKamerix
            Exit Sub
        Case SCREEN_LOGOUT
            Logout
            Exit Sub
    End Select

    ' Les autres exigent une session : on propose la connexion
    If Not IsLoggedIn() Then StartKamerix
    If Not IsLoggedIn() Then Exit Sub

    Select Case strScreen
        Case SCREEN_REFERENCE
            If HasReadAccess(MOD_CARS) Then ShowModal "frmReference"
        Case SCREEN_PARTIES
            If HasReadAccess(MOD_PARTIES) Then ShowModal "frmParties"
        Case SCREEN_CARS
            If HasReadAccess(MOD_CARS) Then ShowModal "frmCars"
        Case SCREEN_SALES
            If HasReadAccess(MOD_SALES) Then ShowModal "frmSales"
        Case SCREEN_ACCOUNTS
            ManageAccounts
        Case SCREEN_PASSWORD
            ChangeMyPasswordInteractive
        Case SCREEN_SETTINGS
            If CanWrite(MOD_ADMINISTRATION) Then
                ShowSettingsForm False
            Else
                MsgBox "Seul un administrateur peut modifier les param" & ChrW(232) & "tres de Kamerix.", vbExclamation, APP_TITLE
            End If
        Case Else
            MsgBox ChrW(201) & "cran inconnu : " & pScreen & vbCrLf & vbCrLf & _
                   ChrW(201) & "crans disponibles :" & vbCrLf & ScreenNames(), vbExclamation, APP_TITLE
    End Select
    Exit Sub

EH:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, APP_TITLE
End Sub

' Liste des noms d'ecrans acceptes par OpenScreen (un par ligne)
Public Function ScreenNames() As String
    ScreenNames = SCREEN_SALES & vbCrLf & SCREEN_CARS & vbCrLf & SCREEN_PARTIES & vbCrLf & SCREEN_REFERENCE & vbCrLf & _
                  SCREEN_ACCOUNTS & vbCrLf & SCREEN_SETTINGS & vbCrLf & SCREEN_PASSWORD & vbCrLf & SCREEN_LOGIN & vbCrLf & SCREEN_LOGOUT
End Function

' Vrai si le role connecte peut lire le module ; sinon message
Private Function HasReadAccess(ByVal pModule As String) As Boolean
    HasReadAccess = CanRead(pModule)
    If Not HasReadAccess Then
        MsgBox "Votre r" & ChrW(244) & "le ne donne pas acc" & ChrW(232) & "s " & ChrW(224) & " cet " & ChrW(233) & "cran (" & _
               pModule & ").", vbExclamation, APP_TITLE
    End If
End Function

' Affiche un formulaire modal par son nom, puis le decharge
Private Sub ShowModal(ByVal pFormName As String)
    Dim frm As Object

    Set frm = VBA.UserForms.Add(pFormName)
    frm.Show
    Unload frm
End Sub

' Affiche pMessage et propose d'imprimer le document (pQuestion, Oui/Non). Sans droit de lecture
' sur ce document, ou s'il est introuvable, seul le message s'affiche. Une erreur d'impression
' n'annule rien : le document reste editable par le bouton "Documents..." de l'ecran des ventes.
Public Sub OfferDocument(ByVal pDocType As String, ByVal pDocId As Long, ByVal pMessage As String, _
                         ByVal pQuestion As String)
    On Error GoTo EH
    If Not IsCompanyConfigured() Then
        MsgBox pMessage & vbCrLf & vbCrLf & "Impression impossible : l'entreprise n'est pas encore configur" & ChrW(233) & "e " & _
               "(OpenScreen ""PARAMETRES"", r" & ChrW(233) & "serv" & ChrW(233) & " " & ChrW(224) & " un administrateur).", _
               vbInformation, APP_TITLE
        Exit Sub
    End If
    If pDocId <= 0 Or Not CanReadDocument(pDocType) Then
        MsgBox pMessage, vbInformation, APP_TITLE
        Exit Sub
    End If
    If MsgBox(pMessage & vbCrLf & vbCrLf & pQuestion, vbQuestion + vbYesNo, APP_TITLE) = vbYes Then
        ProduceDocument pDocType, pDocId, kxDocPrint
    End If
    Exit Sub

EH:
    MsgBox "Le document n'a pas pu " & ChrW(234) & "tre imprim" & ChrW(233) & " (erreur " & Err.Number & ") :" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Vous pourrez le r" & ChrW(233) & ChrW(233) & "diter avec le bouton " & ChrW(171) & " Documents... " & ChrW(187) & _
           " de l'" & ChrW(233) & "cran des ventes.", vbExclamation, APP_TITLE
End Sub

' Ouvre frmSettings : pSetup = Vrai pour l'assistant de premiere configuration (onglet
' Entreprise seul), Faux pour l'ecran PARAMETRES complet. Retourne Vrai si enregistre.
Public Function ShowSettingsForm(ByVal pSetup As Boolean) As Boolean
    Dim frm As Object

    Set frm = VBA.UserForms.Add(FORM_SETTINGS)
    frm.Prepare pSetup
    frm.Show
    ShowSettingsForm = frm.Succeeded
    Unload frm
End Function

' Apres la connexion : tant que l'entreprise n'est pas configuree, l'assistant s'ouvre pour
' un administrateur ; les autres roles sont seulement prevenus (documents bloques).
Public Sub CompanySetupIfNeeded()
    On Error GoTo EH
    If Not IsLoggedIn() Then Exit Sub
    If IsCompanyConfigured() Then Exit Sub

    If Not CanWrite(MOD_ADMINISTRATION) Then
        MsgBox "L'entreprise n'est pas encore configur" & ChrW(233) & "e : les factures et les re" & ChrW(231) & "us ne pourront " & _
               "pas " & ChrW(234) & "tre " & ChrW(233) & "dit" & ChrW(233) & "s tant qu'un administrateur ne l'aura pas fait.", _
               vbInformation, APP_TITLE
        Exit Sub
    End If

    If Not ShowSettingsForm(True) Then
        MsgBox "Configuration report" & ChrW(233) & "e : elle vous sera propos" & ChrW(233) & "e " & ChrW(224) & " nouveau " & _
               ChrW(224) & " la prochaine connexion." & vbCrLf & _
               "D'ici l" & ChrW(224) & ", les factures et les re" & ChrW(231) & "us ne peuvent pas " & ChrW(234) & "tre " & _
               ChrW(233) & "dit" & ChrW(233) & "s (OpenScreen ""PARAMETRES"" pour la faire).", vbInformation, APP_TITLE
    End If
    Exit Sub

EH:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, APP_TITLE
End Sub
